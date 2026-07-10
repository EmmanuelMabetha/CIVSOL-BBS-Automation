"""
dxf_bbs_extractor.py
=====================

General-purpose tool that reads bar-mark callouts off a DXF drawing
(e.g. "Extractor.dxf") and turns them into a clean, review-ready Bar
Bending Schedule (BBS) starting point in Excel.

WHY THIS EXISTS
----------------
On a reinforcement drawing, every bar/link group is normally labelled
with a text callout like:

    6Y12-15 (3T,3B)        -> 6 no. Y12 bars, mark 15, 3 Top / 3 Bottom
    20Y8-01-300 LINKS       -> 20 no. Y8 links, mark 01, spacing 300 c/c
    3Y12-32                -> 3 no. Y12 bars, mark 32

Copying these labels into a BBS by hand (drawing -> spreadsheet) is
exactly where transcription errors creep in: a "6" read as "8", a
mark number skipped, a diameter mistyped. This script extracts every
such label straight from the DXF geometry (no re-typing), and:

  1. Produces a raw, traceable "extraction report" (one row per label
     found, with its source layer/handle/location) so every entry can
     be checked against the drawing.
  2. Aggregates repeated bar marks (the same mark can appear more than
     once on a drawing, e.g. mirrored footings) into totals.
  3. Flags anything suspicious automatically -- e.g. the same bar mark
     number used with two different diameters/types, which is almost
     always a drafting error worth catching before it reaches site.
  4. Writes a BBS workbook in the same layout/formula style as a
     standard SABS/SANS 10100 schedule (matching the uploaded
     BBS01 template), with the objective fields (mark, type, diameter,
     count) filled in automatically.

WHAT IT DELIBERATELY DOES **NOT** GUESS
-----------------------------------------
The individual bar LENGTH and the BS8666 SHAPE CODE (and its A/B/C/D/n
bend dimensions) are left blank for the engineer to complete. A text
label like "20Y8-01-300 LINKS" tells you the bar type, diameter, mark
and spacing -- it does NOT reliably tell you the leg lengths of the
link or the true individual bar length, and guessing those from a
regex would risk putting wrong numbers into a structural document.
Where the DXF *does* carry a usable geometric length (the dimension's
measured value), it is carried through as a separate reference column
so you can use it as a sanity check, not as gospel.

USAGE
-----
    python dxf_bbs_extractor.py Extractor.dxf \\
        --template BBS01_-_FOUNDATION_REINFORCEMENT.xlsx \\
        --outdir /mnt/user-data/outputs \\
        --member "FOOTING"

    # restrict to specific layers (default: every layer in the file --
    # the idea is you pre-filter by only copying the layers you want
    # into the extractor DXF in the first place)
    python dxf_bbs_extractor.py Extractor.dxf --layers DIMS,REINFORCEMENT

Outputs (written to --outdir):
    extraction_report.csv   raw, one row per label found on the drawing
    bar_mark_summary.csv    aggregated per bar mark, with conflict flags
    BBS_from_DXF.xlsx       BBS-formatted workbook, ready to check/complete
"""

from __future__ import annotations

import argparse
import glob
import math
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import ezdxf
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# sans282_shape_codes.py is kept as a SEPARATE file, imported here, rather
# than pasted into this script. Reasoning: it's reference DATA (formulas,
# bend allowances) with its own "confirmed vs unverified" provenance that
# needs to be checked against the actual standard independently of how
# this extractor works -- keeping it separate means updating/correcting
# shape-code data never risks touching extraction logic, and vice versa.
# Both files just need to sit in the same folder.
try:
    import sans282_shape_codes as shape_codes
except ImportError:
    shape_codes = None


# ---------------------------------------------------------------------------
# 0. FILE LOCATION -- the only thing you should need to edit day-to-day
# ---------------------------------------------------------------------------
#
# Set this ONCE to the folder where you drop your drawings. Then just run:
#     python dxf_bbs_extractor.py
# with no arguments, and it will automatically find the .dxf (and .xlsx
# template, if one is present) sitting in that folder -- no need to type
# a file path or filename every time.
#
# You can still override this from the command line if you want
# (see --dxf / --template / --outdir below), but for normal day-to-day
# use, just change FOLDER once and drop files into it.

FOLDER = r"C:\Users\T14s\OneDrive - University of Cape Town\Desktop\CIVSOL\Automation"


def find_file(folder: str, extension: str, label: str) -> Optional[str]:
    """Look for exactly one file of the given extension in `folder`.

    If there's more than one, asks you to pick. If there's none, returns
    None (the caller decides whether that's fatal).
    """
    matches = sorted(glob.glob(os.path.join(folder, f"*{extension}")))
    if not matches:
        return None
    if len(matches) == 1:
        print(f"Using {label}: {matches[0]}")
        return matches[0]
    print(f"Multiple {extension} files found in {folder}:")
    for i, f in enumerate(matches, 1):
        print(f"  {i}. {os.path.basename(f)}")
    choice = input(f"Enter the number of the {label} to use: ").strip()
    try:
        return matches[int(choice) - 1]
    except (ValueError, IndexError):
        print("Invalid choice, using the first match.")
        return matches[0]


# ---------------------------------------------------------------------------
# 1. CONFIGURATION -- tweak these for a different drawing/labelling style
# ---------------------------------------------------------------------------

# Matches:  <count><type letter(s)><diameter>-<mark>[-<spacing>] [suffix]
# e.g. "6Y12-15 (3T,3B)" / "20Y8-01-300 LINKS" / "3H16-07"
BAR_LABEL_PATTERN = re.compile(
    r"""^\s*
        (?P<count>\d+)\s*
        (?P<type>[A-Za-z]+)\s*
        (?P<diameter>\d+)\s*
        -\s*
        (?P<mark>[A-Za-z0-9]+)
        (?:\s*-\s*(?P<spacing>\d+))?
        \s*(?P<suffix>.*?)\s*$
    """,
    re.VERBOSE,
)

# Entity types to inspect for text content. DIMENSION text overrides are
# the primary source on most reinforcement drawings; TEXT/MTEXT are kept
# as a fallback for drawings that label bars with plain text instead.
TEXT_BEARING_TYPES = ("DIMENSION", "TEXT", "MTEXT")

# Words that indicate a link/stirrup rather than a straight bar. Used only
# to leave a helpful hint in the "suggested shape code" column -- never
# auto-committed as a final engineering decision.
LINK_KEYWORDS = ("LINK", "STIRRUP", "TIE")

# DXF dimensions with near-zero measured length usually mean the label was
# placed as a leader/callout with no real geometry behind it (defpoint
# coincident with the text), not an actual zero-length bar. Anything below
# this is treated as "no usable length" rather than a real measurement.
MIN_VALID_LENGTH_MM = 50

# ---------------------------------------------------------------------------
# 1b. BAR LENGTH FROM LEADER-DOT GEOMETRY
# ---------------------------------------------------------------------------
# A DIMENSION's own "measured value" (used above) turned out to often be
# unrelated to the actual bar -- it just measures whatever two points the
# original dimension happened to be drawn between, which is frequently a
# leftover distance from something else on the drawing entirely.
#
# What IS reliable on this drawing: every straight-bar callout has a solid
# filled circle ("dot") sitting almost exactly on top of the bar it refers
# to (a HATCH entity with solid_fill=1). So instead of trusting the
# DIMENSION's own measured value, we:
#   1. Find every dot (solid HATCH) on DOT_LAYER.
#   2. Find every candidate bar LINE on BAR_LAYER.
#   3. For each dot, find the closest bar LINE -- if it's within
#      BAR_MATCH_TOL_MM of the dot, that LINE's length is treated as the
#      real bar length (this distance is normally under ~2mm when it's a
#      genuine match, so this tolerance is deliberately tight).
#   4. Separately, work out which bar-mark label that dot belongs to.
#
# Step 4 is the part that needed care. The straightforward approach --
# "whichever label's defpoint is closest to the dot in a straight line" --
# gets fooled when two different bars' dots sit near each other (e.g. two
# parallel bars a short distance apart): the WRONG label can simply happen
# to be a few mm closer in raw distance than the CORRECT one.
#
# The fix: every leader on this drawing is drawn axis-aligned (either
# purely horizontal or purely vertical -- standard orthogonal drafting
# practice). That means a dot and its TRUE label share one coordinate
# almost exactly (to a fraction of a mm) -- the dot and the label's
# defpoint sit at the same X (for a vertical leader) or the same Y (for a
# horizontal one), with the *other* coordinate representing the leader's
# length. A wrong, merely-nearby label will NOT share either coordinate
# precisely, even if its raw straight-line distance happens to be smaller.
# So instead of minimising raw distance, we minimise
# min(|dot.x - label.x|, |dot.y - label.y|) -- the aligned-axis offset --
# and only accept a match where that's near zero.
#
# Links are excluded from this matching entirely (see is_link handling
# elsewhere) -- this is for straight bars only, for now.

DOT_LAYER = "DIMENSION"       # layer carrying the solid filled circle markers
BAR_LAYER = "REINFORCEMENT"   # layer carrying the actual bar LINE geometry
DOT_LABEL_TOL_MM = 5          # max axis-aligned offset between a dot and its label's defpoint
BAR_MATCH_TOL_MM = 10         # max distance from a dot to the bar LINE it sits on


# ---------------------------------------------------------------------------
# 2. DATA STRUCTURES
# ---------------------------------------------------------------------------

@dataclass
class BarLabel:
    layer: str
    entity_type: str
    handle: str
    raw_text: str
    x: float
    y: float
    measured_length: Optional[float]  # DXF dimension geometric value -- kept for reference only, unreliable
    defpoint: Optional[tuple] = None  # DIMENSION's own defpoint, used to match against leader-dot geometry
    count: Optional[int] = None
    bar_type: Optional[str] = None
    diameter: Optional[int] = None
    mark: Optional[str] = None
    spacing: Optional[int] = None
    suffix: str = ""
    parsed_ok: bool = False


# ---------------------------------------------------------------------------
# 3. TEXT CLEANUP + PARSING
# ---------------------------------------------------------------------------

def clean_dxf_text(raw: str) -> str:
    """Strip AutoCAD MTEXT/DIMENSION formatting control codes.

    Dimension text overrides and MTEXT can carry inline formatting such
    as \\P (paragraph break), \\W (width factor), \\H (height), font/colour
    codes, and {..} grouping braces. We only want the visible text.
    """
    if raw is None:
        return ""
    text = raw
    text = re.sub(r"\\P", " ", text)          # paragraph break -> space
    text = re.sub(r"\\[A-Za-z][^;]*;", "", text)  # \Xvalue; formatting codes
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_bar_label(raw_text: str) -> dict:
    """Parse a cleaned label string into its structured fields.

    Returns a dict with parsed_ok=False if the text doesn't match the
    expected bar-mark pattern (e.g. it's an unrelated dimension/note).
    """
    text = clean_dxf_text(raw_text)
    m = BAR_LABEL_PATTERN.match(text)
    if not m:
        return {"parsed_ok": False}

    gd = m.groupdict()
    return {
        "parsed_ok": True,
        "count": int(gd["count"]),
        "bar_type": gd["type"].upper(),
        "diameter": int(gd["diameter"]),
        "mark": gd["mark"].upper(),
        "spacing": int(gd["spacing"]) if gd["spacing"] else None,
        "suffix": gd["suffix"].strip(),
    }


def suggest_shape_code(suffix: str) -> Optional[str]:
    """Very light heuristic hint only -- an engineer must confirm this.

    Returns the literal string "LINK" when the label text says
    LINK/STIRRUP/TIE -- NOT a numeric shape code. Vector-traced geometry
    from this project's own drawing (301-00825-01-171 Rev 1) confirms 60
    genuinely IS the closed rectangular link/stirrup shape (legs A, B,
    rounded corners, short lapped tail) -- so the shape NUMBER isn't the
    problem. The problem is the label text alone ("6Y8-01-300 LINKS")
    never tells us the actual A/B leg lengths, and this tool doesn't
    currently trace closed/looped chains to measure them either -- so
    writing "60" into the schedule would still be asserting a length-
    bearing shape code with no dimensions behind it. "LINK" is a flag for
    the engineer to fill in A/B and confirm the code by hand, not a
    length calculation.
    """
    upper = suffix.upper()
    if any(k in upper for k in LINK_KEYWORDS):
        return "LINK"
    return None


# ---------------------------------------------------------------------------
# 4. DXF EXTRACTION
# ---------------------------------------------------------------------------

def get_entity_text(entity) -> Optional[str]:
    etype = entity.dxftype()
    if etype == "DIMENSION":
        # An empty/space override ("") means "use the measured value" --
        # not a bar-mark label, so skip it.
        override = entity.dxf.get("text", None)
        if override in (None, "", " "):
            return None
        return override
    if etype == "TEXT":
        return entity.dxf.text
    if etype == "MTEXT":
        # ezdxf exposes a plain-text helper that strips formatting codes
        try:
            return entity.plain_text()
        except Exception:
            return entity.text
    return None


def get_entity_position(entity) -> tuple[float, float]:
    etype = entity.dxftype()
    try:
        if etype == "DIMENSION":
            p = entity.dxf.text_midpoint if entity.dxf.hasattr("text_midpoint") else entity.dxf.defpoint
        elif etype == "TEXT":
            p = entity.dxf.insert
        elif etype == "MTEXT":
            p = entity.dxf.insert
        else:
            return (0.0, 0.0)
        return (float(p[0]), float(p[1]))
    except Exception:
        return (0.0, 0.0)


def get_measured_length(entity) -> Optional[float]:
    if entity.dxftype() != "DIMENSION":
        return None
    try:
        return float(entity.get_measurement())
    except Exception:
        return None


def extract_labels(doc, layers: Optional[list[str]] = None) -> list[BarLabel]:
    """Walk modelspace, pull every text-bearing entity, parse bar labels."""
    msp = doc.modelspace()

    layer_filter = set(l.upper() for l in layers) if layers else None

    labels: list[BarLabel] = []
    for entity in msp:
        if entity.dxftype() not in TEXT_BEARING_TYPES:
            continue
        layer = entity.dxf.layer
        if layer_filter is not None and layer.upper() not in layer_filter:
            continue

        raw = get_entity_text(entity)
        if not raw or not raw.strip():
            continue

        cleaned = clean_dxf_text(raw)
        parsed = parse_bar_label(cleaned)
        x, y = get_entity_position(entity)

        defpoint = None
        if entity.dxftype() == "DIMENSION":
            try:
                defpoint = (float(entity.dxf.defpoint[0]), float(entity.dxf.defpoint[1]))
            except Exception:
                defpoint = None

        label = BarLabel(
            layer=layer,
            entity_type=entity.dxftype(),
            handle=entity.dxf.handle,
            raw_text=cleaned,
            x=round(x, 1),
            y=round(y, 1),
            measured_length=get_measured_length(entity),
            defpoint=defpoint,
        )
        if parsed.get("parsed_ok"):
            label.parsed_ok = True
            label.count = parsed["count"]
            label.bar_type = parsed["bar_type"]
            label.diameter = parsed["diameter"]
            label.mark = parsed["mark"]
            label.spacing = parsed["spacing"]
            label.suffix = parsed["suffix"]

        labels.append(label)

    return labels


# ---------------------------------------------------------------------------
# 4b. LEADER-DOT -> BAR LENGTH MATCHING
# ---------------------------------------------------------------------------

def _dist(p1: tuple, p2: tuple) -> float:
    return math.hypot(p1[0] - p2[0], p1[1] - p2[1])


def _point_to_segment_dist(p: tuple, a: tuple, b: tuple) -> float:
    px, py = p
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return _dist(p, a)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    return _dist(p, (ax + t * dx, ay + t * dy))


# Radius sanity filter for donut-style dots -- rejects anything wildly
# too small/large to plausibly be a bar-mark anchor, so a random 2-vertex
# closed polyline elsewhere on the drawing (a bolt symbol, a different
# detail) doesn't get treated as a dot just because it happens to be a
# closed 2-point bulge shape too.
MIN_DOT_RADIUS_MM = 3
MAX_DOT_RADIUS_MM = 100


def find_solid_dots(doc, layer: str = DOT_LAYER) -> list[tuple]:
    """Find solid-filled circular markers (HATCH with solid_fill=1) on the
    given layer, returned as a list of (x, y) centers."""
    msp = doc.modelspace()
    dots = []
    for h in msp.query(f'HATCH[layer=="{layer}"]'):
        if not h.dxf.solid_fill:
            continue
        for path in h.paths:
            for edge in getattr(path, "edges", []):
                center = getattr(edge, "center", None)
                if center is not None:
                    dots.append((float(center[0]), float(center[1])))
    return dots


def find_donut_dots(doc, layer: Optional[str] = None) -> list[tuple]:
    """Find AutoCAD DONUT-style dots: a closed LWPOLYLINE with exactly 2
    vertices, both with bulge ~= 1 (two semicircular arcs forming a full
    circle). Center = midpoint of the two vertices; radius = half the
    distance between them. This is AutoCAD's actual DXF representation
    for its DONUT command -- verified against a synthetic test file, not
    guessed (see module test suite).

    layer=None scans every layer, since a drafter using DONUT may not use
    the same layer convention as HATCH-based dots.
    """
    msp = doc.modelspace()
    query = "LWPOLYLINE" if layer is None else f'LWPOLYLINE[layer=="{layer}"]'
    dots = []
    for p in msp.query(query):
        if not p.closed:
            continue
        points = list(p.get_points())
        if len(points) != 2:
            continue
        (x1, y1, _, _, b1), (x2, y2, _, _, b2) = points
        if abs(abs(b1) - 1.0) > 0.05 or abs(abs(b2) - 1.0) > 0.05:
            continue
        radius = math.hypot(x2 - x1, y2 - y1) / 2
        if not (MIN_DOT_RADIUS_MM <= radius <= MAX_DOT_RADIUS_MM):
            continue
        dots.append(((x1 + x2) / 2, (y1 + y2) / 2))
    return dots


def find_all_dots(doc, hatch_layer: str = DOT_LAYER, donut_layer: Optional[str] = None) -> list[tuple]:
    """Merge both known dot conventions (HATCH solid-fill and DONUT
    LWPOLYLINE) into one list, de-duplicating any that coincide (in case
    both conventions somehow mark the same physical point)."""
    dots = find_solid_dots(doc, layer=hatch_layer) + find_donut_dots(doc, layer=donut_layer)
    deduped = []
    for d in dots:
        if not any(_dist(d, existing) < 1.0 for existing in deduped):
            deduped.append(d)
    return deduped


def find_bar_lines(doc, layer: str = BAR_LAYER) -> list[tuple]:
    """Find candidate bar LINE entities on the given layer, returned as a
    list of (start_xy, end_xy, length). LWPOLYLINE is deliberately excluded
    -- on this drawing those turned out to be small rectangular symbols,
    not bars, and summing their segment lengths gives a wrong "length"."""
    msp = doc.modelspace()
    lines = []
    for l in msp.query(f'LINE[layer=="{layer}"]'):
        s = (float(l.dxf.start[0]), float(l.dxf.start[1]))
        e = (float(l.dxf.end[0]), float(l.dxf.end[1]))
        lines.append((s, e, _dist(s, e)))
    return lines


# ---------------------------------------------------------------------------
# 4c. BAR CHAINS -- grouping connected LINE segments into bent-bar paths,
#     so shape codes (not just straight-bar length) can be detected.
# ---------------------------------------------------------------------------
# A single LINE entity is definitely shape 20 (straight) -- one segment,
# no bend, nothing to detect. A bent bar (L-shape, crank, stirrup leg,
# etc.) is drawn as several separate LINE entities end-to-end. To detect
# its shape code we first need to recognise which LINEs actually belong
# together as one continuous bar.
#
# This is done conservatively: two LINEs are only chained together through
# a shared endpoint if EXACTLY two LINEs meet there. If three or more
# LINE-ends meet at the same point (a T-junction, or two unrelated bars
# happening to cross/touch), that point is treated as a junction and
# chaining stops there -- each LINE is instead treated as its own
# standalone segment. Guessing which of several crossing bars continues
# into which is exactly the kind of silent wrong-answer this tool avoids
# elsewhere, so the same caution applies here.

CHAIN_SNAP_TOL_MM = 2  # endpoints within this distance are treated as "the same point"


@dataclass
class BarChain:
    segments: list       # [(start_xy, end_xy, length), ...] in path order
    segment_lengths: list  # just the lengths, in path order (candidate A, B, C, ...)
    total_length: float   # sum of segment lengths (raw geometry, NOT the calculated BBS length)
    num_segments: int
    turn_angles: list     # degrees of direction change at each internal vertex


def _chain_snap(p: tuple, tol: float = CHAIN_SNAP_TOL_MM) -> tuple:
    return (round(p[0] / tol), round(p[1] / tol))


def _turn_angle(seg1: tuple, seg2: tuple) -> float:
    """Angle (degrees, 0-180) the direction changes between two ordered,
    connected segments. 0 = continues straight, 90 = right-angle bend."""
    v1 = (seg1[1][0] - seg1[0][0], seg1[1][1] - seg1[0][1])
    v2 = (seg2[1][0] - seg2[0][0], seg2[1][1] - seg2[0][1])
    a1 = math.atan2(v1[1], v1[0])
    a2 = math.atan2(v2[1], v2[0])
    diff = math.degrees(a2 - a1)
    diff = (diff + 180) % 360 - 180
    return abs(diff)


def _order_chain(segs: list) -> Optional[list]:
    """Order a set of connected (p1, p2) segments into a single continuous
    path, flipping each segment's direction as needed. Returns None if the
    segments don't form one simple open path (e.g. they form a closed
    loop) -- caller should not guess at an ordering in that case."""
    from collections import Counter
    counts = Counter()
    for p1, p2 in segs:
        counts[_chain_snap(p1)] += 1
        counts[_chain_snap(p2)] += 1
    free_ends = [pt for pt, c in counts.items() if c == 1]
    if len(free_ends) != 2:
        return None  # not a simple open path

    remaining = list(segs)
    used = [False] * len(remaining)
    ordered = []
    current_pt = None
    for i, (p1, p2) in enumerate(remaining):
        if _chain_snap(p1) == free_ends[0]:
            ordered.append((p1, p2)); used[i] = True; current_pt = p2; break
        if _chain_snap(p2) == free_ends[0]:
            ordered.append((p2, p1)); used[i] = True; current_pt = p1; break

    for _ in range(len(remaining) - 1):
        found = False
        for i, (p1, p2) in enumerate(remaining):
            if used[i]:
                continue
            if _chain_snap(p1) == _chain_snap(current_pt):
                ordered.append((p1, p2)); used[i] = True; current_pt = p2; found = True; break
            if _chain_snap(p2) == _chain_snap(current_pt):
                ordered.append((p2, p1)); used[i] = True; current_pt = p1; found = True; break
        if not found:
            return None
    return ordered


def build_bar_chains(doc, layer: str = BAR_LAYER) -> list:
    """Group connected LINE entities on `layer` into BarChains. A LINE
    that doesn't touch any other LINE becomes its own 1-segment chain."""
    msp = doc.modelspace()
    raw_lines = []
    for l in msp.query(f'LINE[layer=="{layer}"]'):
        s = (float(l.dxf.start[0]), float(l.dxf.start[1]))
        e = (float(l.dxf.end[0]), float(l.dxf.end[1]))
        raw_lines.append((s, e))

    point_to_lines: dict = defaultdict(list)
    for idx, (s, e) in enumerate(raw_lines):
        point_to_lines[_chain_snap(s)].append(idx)
        point_to_lines[_chain_snap(e)].append(idx)

    parent = list(range(len(raw_lines)))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    for pt, idxs in point_to_lines.items():
        unique_idxs = set(idxs)
        if len(unique_idxs) == 2:  # exactly two lines meet -> safe to chain
            a, b = list(unique_idxs)
            union(a, b)
        # len > 2 -> junction/branch: deliberately do NOT chain through it

    groups: dict = defaultdict(list)
    for idx in range(len(raw_lines)):
        groups[find(idx)].append(idx)

    chains = []
    for idxs in groups.values():
        segs = [raw_lines[i] for i in idxs]
        if len(segs) == 1:
            s, e = segs[0]
            length = _dist(s, e)
            chains.append(BarChain(segments=[(s, e, length)], segment_lengths=[round(length, 1)],
                                    total_length=length, num_segments=1, turn_angles=[]))
            continue

        ordered = _order_chain(segs)
        if ordered is None:
            # Couldn't resolve a clean single path (e.g. a closed loop) --
            # fall back to standalone segments rather than guess an order.
            for s, e in segs:
                length = _dist(s, e)
                chains.append(BarChain(segments=[(s, e, length)], segment_lengths=[round(length, 1)],
                                        total_length=length, num_segments=1, turn_angles=[]))
            continue

        seg_tuples = [(s, e, _dist(s, e)) for s, e in ordered]
        angles = [round(_turn_angle(ordered[i], ordered[i + 1]), 1) for i in range(len(ordered) - 1)]
        chains.append(BarChain(
            segments=seg_tuples,
            segment_lengths=[round(t[2], 1) for t in seg_tuples],
            total_length=sum(t[2] for t in seg_tuples),
            num_segments=len(seg_tuples),
            turn_angles=angles,
        ))
    return chains


def _chain_point_dist(dot: tuple, chain: BarChain) -> float:
    """Shortest distance from a point to any segment in the chain."""
    return min(_point_to_segment_dist(dot, s, e) for s, e, _ in chain.segments)


def suggest_shape_from_geometry(num_segments: int, turn_angles: list) -> tuple:
    """Ask sans282_shape_codes for shape-code candidates matching this
    geometry. Returns (auto_code, candidates) where auto_code is only set
    if EXACTLY ONE confirmed candidate exists -- otherwise None, since
    picking among several plausible confirmed codes (or accepting an
    unverified one) would be guessing, not detecting.

    A single segment is unambiguously shape 20 (straight bar) -- that's
    the library's own definition, not a guess, so it's always returned
    directly rather than run through the coarser multi-candidate matcher.
    """
    if num_segments == 1:
        return "20", ["20"]
    if shape_codes is None:
        return None, []
    try:
        candidates = shape_codes.match_shape_by_geometry(num_segments, turn_angles)
    except Exception:
        return None, []
    confirmed = [c for c in candidates if shape_codes.SHAPE_CODES[c].confidence == "confirmed" and c != "99"]
    auto_code = confirmed[0] if len(confirmed) == 1 else None
    return auto_code, candidates


def _axis_aligned_dist(p1: tuple, p2: tuple) -> float:
    """Distance along whichever single axis (X or Y) the two points are
    aligned on. Near zero for two points that share a coordinate (as a
    dot and its true label's defpoint do, on an orthogonally-drafted
    leader) -- large otherwise, even if the raw straight-line distance
    between them happens to be small. See module notes above."""
    return min(abs(p1[0] - p2[0]), abs(p1[1] - p2[1]))


def match_dot_geometry(
    labels: list[BarLabel],
    dots: list[tuple],
    chains: list,
    dot_label_tol: float = DOT_LABEL_TOL_MM,
    bar_match_tol: float = BAR_MATCH_TOL_MM,
) -> tuple[dict, list[dict]]:
    """For each dot: find the bar-mark label whose defpoint is
    axis-aligned with it (see _axis_aligned_dist) and the closest bar
    CHAIN (by perpendicular distance to any of its segments). If both are
    within tolerance, record that chain's full geometry against that mark
    -- not just its length, but its segment count, per-segment lengths,
    and turn angles, so a shape code can be suggested.

    Returns:
        geometry_by_mark: {mark: [ {length, num_segments, turn_angles,
            segment_lengths, auto_shape_code, shape_candidates}, ... ]}
        match_records: one dict per dot, for a full traceability report
    """
    # Only non-link, parsed labels are candidates -- links are handled
    # separately and are not expected to have a leader-dot at all.
    candidates = [
        l for l in labels
        if l.parsed_ok and l.defpoint is not None
        and suggest_shape_code(l.suffix) != "LINK"
    ]

    geometry_by_mark: dict = defaultdict(list)
    match_records = []

    for dot in dots:
        record = {"Dot X": round(dot[0], 1), "Dot Y": round(dot[1], 1)}

        if candidates:
            # Stage 1: axis-alignment is a HARD FILTER -- only candidates
            # that genuinely share a coordinate with the dot (a real
            # leader) qualify at all.
            aligned = [l for l in candidates if _axis_aligned_dist(dot, l.defpoint) <= dot_label_tol]
            if aligned:
                # Stage 2: two different bars can occasionally sit on the
                # very same gridline (e.g. two footings at the same level),
                # so more than one candidate can pass stage 1. Among just
                # those that genuinely qualify, the raw straight-line
                # distance is now a reliable tie-breaker -- the true match
                # is the nearer one (shorter, more plausible leader length).
                best_label = min(aligned, key=lambda l: _dist(dot, l.defpoint))
                label_dist = _axis_aligned_dist(dot, best_label.defpoint)
            else:
                best_label, label_dist = None, None
        else:
            best_label, label_dist = None, None

        if chains:
            best_chain_dist, best_chain = min(
                ((_chain_point_dist(dot, c), c) for c in chains),
                key=lambda x: x[0],
            )
        else:
            best_chain_dist, best_chain = None, None

        record["Matched Mark"] = best_label.mark if best_label else None
        record["Label Axis-Aligned Offset (mm)"] = round(label_dist, 1) if label_dist is not None else None
        record["Bar Line Distance (mm)"] = round(best_chain_dist, 1) if best_chain_dist is not None else None
        record["Bar Length (mm)"] = round(best_chain.total_length, 1) if best_chain else None
        record["Num Segments"] = best_chain.num_segments if best_chain else None
        record["Turn Angles"] = str(best_chain.turn_angles) if best_chain else None

        label_ok = best_label is not None and label_dist <= dot_label_tol
        bar_ok = best_chain is not None and best_chain_dist <= bar_match_tol
        record["Accepted"] = bool(label_ok and bar_ok)

        auto_code, shape_candidates = (None, [])
        if label_ok and bar_ok:
            auto_code, shape_candidates = suggest_shape_from_geometry(best_chain.num_segments, best_chain.turn_angles)
            geometry_by_mark[best_label.mark].append({
                "length": round(best_chain.total_length, 1),
                "num_segments": best_chain.num_segments,
                "turn_angles": best_chain.turn_angles,
                "segment_lengths": best_chain.segment_lengths,
                "auto_shape_code": auto_code,
                "shape_candidates": shape_candidates,
            })
        record["Shape Candidates"] = ",".join(shape_candidates) if shape_candidates else None
        record["Auto Shape Code"] = auto_code

        match_records.append(record)

    return dict(geometry_by_mark), match_records





# ---------------------------------------------------------------------------
# 5. REPORTING / AGGREGATION
# ---------------------------------------------------------------------------

def build_extraction_report(labels: list[BarLabel]) -> pd.DataFrame:
    rows = []
    for l in labels:
        rows.append({
            "Layer": l.layer,
            "Entity": l.entity_type,
            "Handle": l.handle,
            "Raw Label": l.raw_text,
            "Parsed OK": l.parsed_ok,
            "Count": l.count,
            "Type": l.bar_type,
            "Diameter (mm)": l.diameter,
            "Bar Mark": l.mark,
            "Spacing (mm c/c)": l.spacing,
            "Notes": l.suffix,
            "DXF Measured Length (mm)": round(l.measured_length, 1) if l.measured_length else None,
            "X": l.x,
            "Y": l.y,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["Parsed OK", "Bar Mark"], ascending=[False, True], na_position="last")
    return df



# A mark's length occurrences that differ by more than this fraction get
# flagged for manual review rather than silently resolved by "longest
# wins" -- e.g. 4 dots on one mark with genuinely different member
# depths, not 4 repeats of the same length.
LENGTH_DIVERGENCE_THRESHOLD = 0.10


def build_summary(labels: list[BarLabel], geometry_by_mark: Optional[dict] = None) -> pd.DataFrame:
    """geometry_by_mark: {mark: [ {length, num_segments, turn_angles,
    segment_lengths, auto_shape_code, shape_candidates}, ... ]}, from
    match_dot_geometry(). If not supplied, no lengths/shapes are
    populated (columns come back empty rather than falling back to the
    old, unreliable DIMENSION-measured value)."""
    geometry_by_mark = geometry_by_mark or {}

    groups: dict[tuple, list[BarLabel]] = defaultdict(list)
    for l in labels:
        if not l.parsed_ok:
            continue
        groups[l.mark].append(l)

    rows = []
    conflicts = []
    for mark, items in groups.items():
        types = {i.bar_type for i in items}
        dias = {i.diameter for i in items}
        spacings = {i.spacing for i in items if i.spacing is not None}
        suffixes = sorted({i.suffix for i in items if i.suffix})
        total_count = sum(i.count for i in items)

        conflict = None
        if len(types) > 1 or len(dias) > 1:
            conflict = f"CONFLICT: mark {mark} used with type(s) {types}, diameter(s) {dias} -- check drawing"
            conflicts.append(conflict)

        # LINKS: excluded from length/shape entirely (see module notes) --
        # a link's real length is its own bend perimeter, not anything we
        # currently extract, so length stays blank rather than guessed.
        #
        # STRAIGHT/BENT BARS: length and shape code now come from the
        # leader-dot -> bar CHAIN match (match_dot_geometry()), NOT from
        # the DIMENSION's own measured value -- that measured value was
        # found to frequently be an unrelated leftover distance. If a
        # mark has no accepted dot match, length/shape stay blank rather
        # than falling back to that unreliable number.
        is_link = suggest_shape_code(" ".join(suffixes)) == "LINK"

        auto_shape_code = None
        segment_lengths_for_row = []
        length_flag = ""

        if is_link:
            longest_length = None
            valid_lengths = []
        else:
            geoms = geometry_by_mark.get(mark, [])
            valid_lengths = sorted(g["length"] for g in geoms)
            longest_length = max(valid_lengths) if valid_lengths else None

            if len(valid_lengths) > 1:
                spread = (valid_lengths[-1] - valid_lengths[0]) / valid_lengths[-1]
                if spread > LENGTH_DIVERGENCE_THRESHOLD:
                    length_flag = (
                        f"LENGTHS DIVERGE across {len(valid_lengths)} occurrences of mark {mark} "
                        f"({valid_lengths[0]}-{valid_lengths[-1]}mm, {spread:.0%} spread) -- "
                        f"these may be genuinely different members (e.g. a stepped/tapered layout), "
                        f"not repeats of the same bar. Longest shown here; check All DXF Lengths before trusting one row for all of them."
                    )

            # Shape code: only auto-assign if every occurrence of this
            # mark agrees on both segment count and the same single
            # confirmed shape candidate. If occurrences disagree (e.g.
            # one dot found a straight 1-segment bar and another found a
            # bent 2-segment one for the "same" mark), that's exactly the
            # kind of thing worth a human looking at, not resolving
            # silently -- so no shape code is auto-filled in that case.
            all_candidates = []
            if geoms:
                seg_counts = {g["num_segments"] for g in geoms}
                auto_codes = {g["auto_shape_code"] for g in geoms if g["auto_shape_code"]}
                # Union of every shape-code candidate the geometry matcher saw
                # across all occurrences of this mark -- surfaced even when no
                # single code could be confidently auto-assigned, so the
                # engineer sees what to check FOR instead of seeing nothing.
                all_candidates = sorted({c for g in geoms for c in g.get("shape_candidates", [])})
                if len(seg_counts) == 1 and len(auto_codes) == 1:
                    auto_shape_code = next(iter(auto_codes))
                    # use the segment lengths from the LONGEST occurrence,
                    # consistent with the "longest governs" length rule
                    longest_geom = max(geoms, key=lambda g: g["length"])
                    segment_lengths_for_row = longest_geom["segment_lengths"]
                elif len(seg_counts) > 1:
                    seg_flag = (
                        f"GEOMETRY DIVERGES across occurrences of mark {mark} "
                        f"(segment counts found: {sorted(seg_counts)}) -- some occurrences look straight, "
                        f"others look bent. Shape code not auto-assigned; check the drawing."
                    )
                    length_flag = (length_flag + " " + seg_flag).strip() if length_flag else seg_flag
                elif all_candidates:
                    # Segment count/angles agree across occurrences, but the
                    # library couldn't narrow it to exactly one confirmed
                    # code (either >1 confirmed candidate fits this geometry,
                    # or the best match is still unverified in
                    # sans282_shape_codes.py). This is NOT the same as "no
                    # data" -- real geometry was matched, it's just not a
                    # single, solid answer -- so it must not default to any
                    # numeric code (including 20). List the real candidates
                    # instead and let the engineer pick from the drawing.
                    seg_flag = (
                        f"SHAPE CODE NOT CONFIRMED for mark {mark} ({next(iter(seg_counts))} segment(s)) -- "
                        f"candidate code(s) from geometry: {', '.join(all_candidates)}. More than one "
                        f"candidate fits, or the best match's formula isn't confirmed yet in "
                        f"sans282_shape_codes.py. No code is auto-filled; compare against the drawing / "
                        f"SANS 282:2011 Annex A and enter the correct one by hand."
                    )
                    length_flag = (length_flag + " " + seg_flag).strip() if length_flag else seg_flag

        rows.append({
            "Bar Mark": mark,
            "Type": next(iter(types)) if len(types) == 1 else "/".join(sorted(types)),
            "Diameter (mm)": next(iter(dias)) if len(dias) == 1 else "/".join(str(d) for d in sorted(dias)),
            "Occurrences on Drawing": len(items),
            "Total No. Off": total_count,
            "Spacing (mm c/c)": next(iter(spacings)) if len(spacings) == 1 else ("/".join(str(s) for s in sorted(spacings)) if spacings else None),
            "Notes": "; ".join(suffixes),
            "Is Link": is_link,
            "Suggested Shape Code (verify)": "LINK" if is_link else None,
            "Auto-Detected Shape Code": auto_shape_code,
            "Shape Candidates (if not confirmed)": ", ".join(all_candidates) if (not is_link and not auto_shape_code and all_candidates) else None,
            "Auto-Detected Segment Lengths (mm)": ", ".join(str(x) for x in segment_lengths_for_row) if segment_lengths_for_row else None,
            "Longest DXF Length (mm)": longest_length,
            "All DXF Lengths (mm)": ", ".join(str(x) for x in valid_lengths) if valid_lengths else None,
            "Length/Geometry Flag": length_flag,
            "Flag": conflict or "",
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["_sort"] = df["Bar Mark"].apply(lambda m: (0, int(m)) if str(m).isdigit() else (1, str(m)))
        df = df.sort_values("_sort").drop(columns="_sort").reset_index(drop=True)

    if conflicts:
        print(f"[!] {len(conflicts)} conflict(s) found -- see 'Flag' column in bar_mark_summary.csv:")
        for c in conflicts:
            print("    -", c)

    return df


def find_missing_marks(summary_df: pd.DataFrame) -> list[int]:
    """Flag gaps in the bar mark numbering sequence.

    Only considers numeric, non-link bar marks (links are often numbered
    in their own separate short sequence, e.g. 01/02, so mixing them into
    a straight-bar gap check produces false positives). If marks 03 and
    06 exist but 04, 05 don't, this returns [4, 5].
    """
    if summary_df.empty:
        return []
    non_link = summary_df[~summary_df["Is Link"]]
    numeric_marks = sorted(int(m) for m in non_link["Bar Mark"] if str(m).isdigit())
    if len(numeric_marks) < 2:
        return []
    full_range = set(range(numeric_marks[0], numeric_marks[-1] + 1))
    missing = sorted(full_range - set(numeric_marks))
    return missing


# ---------------------------------------------------------------------------
# 6. BBS WORKBOOK GENERATION (matches the SABS-style template layout)
# ---------------------------------------------------------------------------

WEIGHT_FACTORS = {8: 0.395, 10: 0.616, 12: 0.888, 16: 1.579, 20: 2.466, 25: 3.854, 32: 6.313}

REVIEW_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")  # pale yellow -- needs your input
DXF_LENGTH_FILL = PatternFill("solid", start_color="D9EAD3", end_color="D9EAD3")  # pale green -- taken from DXF, worth spot-checking
HEADER_FONT = Font(name="Arial", bold=True, size=9)
BODY_FONT = Font(name="Arial", size=9)
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _copy_header_block(src_ws: Worksheet, dst_ws: Worksheet, rows: int = 11):
    """Clone the title/header rows (client, project, column headings) from
    an existing BBS template so the generated sheet matches house style."""
    for row in src_ws.iter_rows(min_row=1, max_row=rows):
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = cell.font.copy()
                new_cell.fill = cell.fill.copy()
                new_cell.border = cell.border.copy()
                new_cell.alignment = cell.alignment.copy()
                new_cell.number_format = cell.number_format
    for merged_range in src_ws.merged_cells.ranges:
        if merged_range.max_row <= rows:
            dst_ws.merge_cells(str(merged_range))
    for col, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = dim.width


# ---------------------------------------------------------------------------
# 6b. SHAPE-CODE LENGTH FORMULA -- built from sans282_shape_codes.py ONLY
# ---------------------------------------------------------------------------
# Every formula written into the spreadsheet here was numerically verified
# against shape_codes.calculate_length() for multiple test cases before
# being trusted (see the test workbook this was built and checked against).
# Only shape codes with confidence == "confirmed" in the library get a
# real formula. Everything else (including any code not in the library at
# all) computes to blank, not a guessed number -- if you need a shape
# code that isn't confirmed yet, that's a sign to go verify it against
# your own copy of SANS 282:2011 and update sans282_shape_codes.py, not a
# sign to hardcode a formula here.

REF_TABLE_START_COL = 19  # column S -- reference table for standard r/h/n lookups


def _write_bend_allowance_table(ws: Worksheet, start_row: int = 1) -> int:
    """Write STANDARD_BEND_ALLOWANCES as a flat lookup table (Key, Type,
    Diameter, h, n, r) starting at REF_TABLE_START_COL. Returns the last
    row written, so the caller knows the lookup range."""
    headers = ["Key", "Type", "Dia (mm)", "h (mm)", "n (mm)", "r (mm)"]
    for j, h in enumerate(headers):
        c = ws.cell(row=start_row, column=REF_TABLE_START_COL + j, value=h)
        c.font = Font(name="Arial", bold=True, size=8, italic=True)

    row = start_row + 1
    if shape_codes is not None:
        for steel_type, table in shape_codes.STANDARD_BEND_ALLOWANCES.items():
            for dia, vals in table.items():
                key = f"{steel_type}{dia}"
                values = [key, steel_type, dia, vals["h"], vals["n"], vals["r"]]
                for j, v in enumerate(values):
                    ws.cell(row=row, column=REF_TABLE_START_COL + j, value=v).font = Font(name="Arial", size=8)
                row += 1
    last_row = row - 1

    note_col = get_column_letter(REF_TABLE_START_COL)
    ws.cell(row=start_row - 1 if start_row > 1 else start_row,
            column=REF_TABLE_START_COL,
            value="Standard bend allowances (SANS 282:2011 Fig 2/3) -- used to auto-look-up r for shape codes 37/38/45. "
                  "If the drawing calls for a NON-STANDARD radius (shape code suffix 'S'), this lookup is wrong for that "
                  "row -- check manually.").font = Font(italic=True, size=8)
    return last_row


def _r_lookup_formula(row: int, ref_last_row: int) -> str:
    """Excel formula: look up standard r for this row's Type (col C) and
    Diameter (col D) against the reference table. Blank if no match."""
    key_col = get_column_letter(REF_TABLE_START_COL)      # Key
    r_col = get_column_letter(REF_TABLE_START_COL + 5)     # r
    return (
        f'IFERROR(INDEX(${r_col}$2:${r_col}${ref_last_row},'
        f'MATCH(C{row}&D{row},${key_col}$2:${key_col}${ref_last_row},0)),"")'
    )


def _shape_length_formula(row: int, ref_last_row: int) -> str:
    """Build the Length (column H) formula for a given row, covering ONLY
    the shape codes confirmed in sans282_shape_codes.py:
        20: A
        41: A+B+C
        42: A+B+C+n
        37: A+B-r/2-d
        38: A+B+C-r-2d
        45: A+B+C-r/2-d
        39: A+0.57*B+C-1.57*d
    Column mapping: J=A, K=B, L=C, N=n, D(sheet col 4)=d (bar diameter).
    Any other shape code (including anything not in this list) -> blank.
    """
    r = _r_lookup_formula(row, ref_last_row)
    return (
        f'=IFERROR('
        f'IF(I{row}=20,J{row},'
        f'IF(I{row}=41,J{row}+K{row}+L{row},'
        f'IF(I{row}=42,J{row}+K{row}+L{row}+N{row},'
        f'IF(I{row}=37,J{row}+K{row}-({r})/2-D{row},'
        f'IF(I{row}=38,J{row}+K{row}+L{row}-({r})-2*D{row},'
        f'IF(I{row}=45,J{row}+K{row}+L{row}-({r})/2-D{row},'
        f'IF(I{row}=39,J{row}+0.57*K{row}+L{row}-1.57*D{row},'
        f'"")))))))'
        f',"")'
    )


def write_bbs_workbook(
    summary_df: pd.DataFrame,
    out_path: str,
    template_path: Optional[str] = None,
    member_name: str = "MEMBER",
    item_description: str = "EXTRACTED FROM DXF -- VERIFY LENGTH / SHAPE CODE / BEND DIMENSIONS",
    missing_marks: Optional[list[int]] = None,
):
    """Write a BBS-formatted workbook. Bar mark / type / diameter / count
    are filled in from the drawing. Length, shape code and bend dimensions
    (H, I, J-N) are left blank with a highlighted fill as a to-do for the
    engineer -- see module docstring for why these are not guessed.
    """
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "BBS from DXF"

    if template_path and Path(template_path).exists():
        src_wb = load_workbook(template_path)
        src_ws = src_wb.active
        _copy_header_block(src_ws, ws, rows=11)
        ws["B8"] = item_description
    else:
        # Minimal header if no template supplied
        headers = ["Member", "Mark", "Type", "", "No.\nMbrs", "No.\neach", "Total\nNo. off",
                   "Length\nmm", "Shape\nCode", "A\nmm", "B\nmm", "C\nmm", "D\nmm", "n\nmm", "TOTAL WEIGHT\n(TON)"]
        for col, h in enumerate(headers, start=1):
            c = ws.cell(row=10, column=col, value=h)
            c.font = HEADER_FONT

    start_row = 12
    if missing_marks:
        banner_text = (
            f"GAP DETECTED IN BAR MARK NUMBERING: mark(s) {', '.join(str(m) for m in missing_marks)} "
            f"do not appear anywhere on the drawing but fall between marks that do -- "
            f"check whether they were omitted by mistake before relying on this schedule."
        )
        ws.cell(row=start_row, column=1, value=banner_text)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=15)
        banner_cell = ws.cell(row=start_row, column=1)
        banner_cell.font = Font(name="Arial", bold=True, size=9, color="9C0006")
        banner_cell.fill = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
        banner_cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[start_row].height = 30
        start_row += 1

    first_data_row = start_row

    # Reference table for the shape-code length formulas (r lookups) --
    # written once, off to the side (columns S onward), sourced directly
    # from sans282_shape_codes.STANDARD_BEND_ALLOWANCES.
    ref_last_row = _write_bend_allowance_table(ws, start_row=1)
    if shape_codes is None:
        print("[!] sans282_shape_codes.py not found -- shape-code length formulas will be blank. "
              "Place it in the same folder as this script.")

    for offset, rec in enumerate(summary_df.to_dict("records")):
        r = start_row + offset
        ws.cell(row=r, column=1, value=member_name if offset == 0 else None)
        ws.cell(row=r, column=2, value=rec["Bar Mark"])
        ws.cell(row=r, column=3, value=rec["Type"])
        ws.cell(row=r, column=4, value=rec["Diameter (mm)"] if isinstance(rec["Diameter (mm)"], int) else None)
        ws.cell(row=r, column=5, value=1)  # No. in Mbrs -- default 1, confirm per member grouping
        ws.cell(row=r, column=6, value=rec["Total No. Off"])  # No. each
        ws.cell(row=r, column=7, value=f"=E{r}*F{r}")  # Total No. off

        # --- Length (H): use the longest DXF-measured length for this bar
        # mark when one was found (e.g. multiple occurrences of mark 01
        # across the drawing -> take the longest span, per your request).
        # If no usable measured length exists on the drawing for this mark,
        # fall back to the shape-code-driven formula so it's still ready
        # to compute once you fill in the shape code and bend dimensions.
        longest_length = rec.get("Longest DXF Length (mm)")
        length_formula = _shape_length_formula(r, ref_last_row)
        is_link = bool(rec.get("Is Link"))
        auto_shape_code = rec.get("Auto-Detected Shape Code")
        if isinstance(auto_shape_code, float) and pd.isna(auto_shape_code):
            auto_shape_code = None
        auto_seg_lengths_raw = rec.get("Auto-Detected Segment Lengths (mm)")
        if auto_seg_lengths_raw is None or (isinstance(auto_seg_lengths_raw, float) and pd.isna(auto_seg_lengths_raw)):
            auto_seg_lengths = []
        elif isinstance(auto_seg_lengths_raw, str):
            auto_seg_lengths = [float(x) for x in auto_seg_lengths_raw.split(",")]
        else:
            # pandas may have inferred a numeric dtype for this column when
            # every value in the file happens to be a single number (no
            # comma) -- e.g. a drawing with no multi-segment bent bars.
            auto_seg_lengths = [float(auto_seg_lengths_raw)]

        if is_link:
            # We know the current DXF geometry for links measures the
            # spaced RUN LENGTH, not the individual link length -- so we
            # deliberately show nothing here rather than a number that
            # would look like a real length. For the shape code cell,
            # write the literal text "LINK" rather than leaving it blank
            # (an empty cell reads as "not processed"; "LINK" reads as
            # "processed, but this is a special case needing your input")
            # and rather than guessing a numeric code -- 60 on THIS
            # drawing's own legend is a specific, different bent shape,
            # not a generic stand-in for "link", so it must not be
            # auto-written here. Length and shape code are left for
            # manual entry until link geometry is handled separately.
            ws.cell(row=r, column=8, value=None)  # H: length -- not available
            ws.cell(row=r, column=8).fill = REVIEW_FILL
            ws.cell(row=r, column=9, value="LINK")  # I: flag, not a shape code
            ws.cell(row=r, column=9).fill = REVIEW_FILL
        else:
            is_straight = str(auto_shape_code) in ("20", "20.0")

            if auto_shape_code is not None:
                # Detected from the actual bar geometry (segment count +
                # angles), not assumed -- e.g. a single LINE is definitely
                # shape 20; a 2-segment 90-degree bend is confidently
                # shape 37 with no other confirmed candidate. Distinct
                # fill from the "needs your input" yellow, since this is
                # a real detection, not a placeholder -- but still worth
                # a quick glance, especially the A/B/C letter assignment.
                ws.cell(row=r, column=9, value=int(float(auto_shape_code)))
                ws.cell(row=r, column=9).fill = DXF_LENGTH_FILL
                for i, seg_len in enumerate(auto_seg_lengths[:3]):  # J, K, L = A, B, C
                    ws.cell(row=r, column=10 + i, value=seg_len)
                    ws.cell(row=r, column=10 + i).fill = DXF_LENGTH_FILL
            else:
                # No numeric code is assumed here -- not 20, not anything
                # else. If geometry couldn't confirm a single shape code
                # (no dot matched at all, occurrences disagreed, or the
                # geometry fit more than one candidate / an unverified
                # library entry) the cell is left BLANK, flagged yellow for
                # manual identification. Column Q lists any real candidates
                # that were found, so this is never a bare unexplained gap.
                ws.cell(row=r, column=9, value=None)  # I: Shape Code -- not confirmed, verify
                ws.cell(row=r, column=9).fill = REVIEW_FILL

            if is_straight and longest_length is not None and str(longest_length).lower() != "nan":
                # A straight bar's length IS just its measured span -- no
                # bend correction applies, so the raw geometric value is
                # the correct answer, not an approximation of one.
                ws.cell(row=r, column=8, value=float(longest_length))  # H: from DXF, plain value
                ws.cell(row=r, column=8).fill = DXF_LENGTH_FILL
            else:
                # ANY bend changes the calculated length vs. the raw
                # straight-line sum of segments (SANS 282 formulas
                # subtract/add bend allowances like r/2+d) -- so for a
                # detected bent shape, or anything not confidently
                # detected as straight, the formula must run using the
                # shape code + A/B/C dimensions, never the raw DXF sum.
                ws.cell(row=r, column=8, value=length_formula)
                ws.cell(row=r, column=8).fill = DXF_LENGTH_FILL if auto_shape_code is not None else REVIEW_FILL

        for col in (10, 11, 12, 13, 14):  # J..N -- ensure any untouched cells still get the review fill
            if ws.cell(row=r, column=col).value is None:
                ws.cell(row=r, column=col).fill = REVIEW_FILL

        weight_formula = (
            f'=IFERROR(IF(D{r}=8,G{r}*H{r}*0.395,IF(D{r}=10,G{r}*H{r}*0.616,'
            f'IF(D{r}=12,G{r}*H{r}*0.888,IF(D{r}=16,G{r}*H{r}*1.579,'
            f'IF(D{r}=20,G{r}*H{r}*2.466,IF(D{r}=25,G{r}*H{r}*3.854,'
            f'IF(D{r}=32,G{r}*H{r}*6.313,"")))))))/1000/1000,"")'
        )
        ws.cell(row=r, column=15, value=weight_formula)  # O: TOTAL WEIGHT

        # Reference-only notes column (Q) -- spacing / suffix / DXF length,
        # so the engineer has the source context right next to the row.
        def _is_set(v):
            return v is not None and str(v) != "" and str(v).lower() != "nan"

        ref_bits = []
        if _is_set(rec.get("Spacing (mm c/c)")):
            ref_bits.append(f"spacing {rec['Spacing (mm c/c)']} c/c")
        if _is_set(rec.get("Notes")):
            ref_bits.append(str(rec["Notes"]))
        if is_link:
            ref_bits.append("LINK -- length not extracted (DXF measures run length, not link length); shape code 60 is confirmed (2A+2B+2n-1.5r-3d) but A/B leg dims aren't traced for closed loops yet -- assign A/B and confirm code manually")
        elif auto_shape_code is not None:
            if str(auto_shape_code) in ("20", "20.0"):
                ref_bits.append("shape code 20 confirmed by geometry (single straight bar line)")
            else:
                ref_bits.append(
                    f"shape code {auto_shape_code} auto-detected from bar geometry ({rec.get('Occurrences on Drawing')} occurrence(s)) -- "
                    f"segment lengths filled in PATH ORDER (first segment=A, second=B, ...); this order is NOT verified against "
                    f"the standard's own sketch/hook orientation for this code -- confirm the A/B/C assignment matches the drawing before trusting it."
                )
        elif _is_set(rec.get("Shape Candidates (if not confirmed)")):
            ref_bits.append(
                f"shape code NOT confirmed -- candidate code(s) from geometry: "
                f"{rec['Shape Candidates (if not confirmed)']}. No code is assumed; compare against "
                f"the drawing / SANS 282:2011 Annex A and enter the correct one by hand."
            )
        else:
            ref_bits.append("shape code not detected from geometry (no dot/bar match found) -- no code is assumed; identify and enter it manually from the drawing")
        if _is_set(rec.get("All DXF Lengths (mm)")):
            ref_bits.append(f"all DXF lengths found: {rec['All DXF Lengths (mm)']} mm (longest used in column H)")
        if _is_set(rec.get("Length/Geometry Flag")):
            ref_bits.append(str(rec["Length/Geometry Flag"]))
        if _is_set(rec.get("Flag")):
            ref_bits.append(str(rec["Flag"]))
        ws.cell(row=r, column=17, value="; ".join(ref_bits))

        for col in range(1, 16):
            ws.cell(row=r, column=col).font = BODY_FONT
            ws.cell(row=r, column=col).border = BOX

    last_row = start_row + len(summary_df) - 1
    total_row = last_row + 2
    ws.cell(row=total_row, column=9, value="TOTAL WEIGHT (TON)")
    ws.cell(row=total_row, column=15, value=f"=SUM(O{first_data_row}:O{last_row})")
    ws.cell(row=total_row, column=9).font = HEADER_FONT
    ws.cell(row=total_row, column=15).font = HEADER_FONT

    ws.cell(row=first_data_row - 1, column=17, value="Reference notes (spacing / drawing notes / DXF length -- not part of BBS format)").font = Font(italic=True, size=8)
    ws.column_dimensions["Q"].width = 55

    wb.save(out_path)


# ---------------------------------------------------------------------------
# 7. REUSABLE ENTRY POINT (used by both the CLI below and Streamlit)
# ---------------------------------------------------------------------------

def process_dxf(
    dxf_path,
    outdir,
    template_path=None,
    member_name="MEMBER",
    layers=None,
):
    """Run the full DXF -> BBS extraction pipeline and write outputs.

    This is the same logic that used to live inline inside main() -- it's
    been pulled out here so it can be called directly (e.g. from a
    Streamlit app) without going through argparse/sys.exit().

    Returns a dict with the resulting dataframes and the paths of every
    file written to `outdir`:
        {
            "report_df": ...,
            "summary_df": ...,
            "report_path": ...,
            "summary_path": ...,
            "dot_match_path": ...,
            "bbs_path": ...,
        }
    """
    doc = ezdxf.readfile(dxf_path)
    labels = extract_labels(doc, layers=layers)

    parsed = [l for l in labels if l.parsed_ok]
    unparsed = [l for l in labels if not l.parsed_ok]

    # --- leader-dot -> bar geometry matching (straight + bent bars) ------------
    hatch_dots = find_solid_dots(doc, layer=DOT_LAYER)
    donut_dots = find_donut_dots(doc, layer=None)
    dots = find_all_dots(doc, hatch_layer=DOT_LAYER, donut_layer=None)
    chains = build_bar_chains(doc, layer=BAR_LAYER)
    bent_chains = [c for c in chains if c.num_segments > 1]
    geometry_by_mark, dot_match_records = match_dot_geometry(labels, dots, chains)
    accepted = sum(1 for r in dot_match_records if r["Accepted"])
    unmatched = [r for r in dot_match_records if not r["Accepted"]]

    report_df = build_extraction_report(labels)
    summary_df = build_summary(labels, geometry_by_mark=geometry_by_mark)

    missing_marks = find_missing_marks(summary_df)

    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    report_path = outdir / "extraction_report.csv"
    summary_path = outdir / "bar_mark_summary.csv"
    dot_match_path = outdir / "dot_bar_matches.csv"
    bbs_path = outdir / "BBS_from_DXF.xlsx"

    report_df.to_csv(report_path, index=False)
    summary_df.to_csv(summary_path, index=False)
    pd.DataFrame(dot_match_records).to_csv(dot_match_path, index=False)
    write_bbs_workbook(summary_df, str(bbs_path), template_path=template_path, member_name=member_name,
                        missing_marks=missing_marks)

    return {
        "report_df": report_df,
        "summary_df": summary_df,
        "report_path": str(report_path),
        "summary_path": str(summary_path),
        "dot_match_path": str(dot_match_path),
        "bbs_path": str(bbs_path),
        # extra info, handy for a Streamlit UI to show progress/warnings
        "num_labels": len(labels),
        "num_parsed": len(parsed),
        "num_unparsed": len(unparsed),
        "num_dots": len(dots),
        "num_chains": len(chains),
        "num_bent_chains": len(bent_chains),
        "num_dots_matched": accepted,
        "num_dots_unmatched": len(unmatched),
        "missing_marks": missing_marks,
    }


# ---------------------------------------------------------------------------
# 8. CLI
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dxf_path", nargs="?", default=None,
                     help="Path to the DXF file to extract from. "
                          "If omitted, auto-detects a .dxf file inside FOLDER (set at the top of this script).")
    ap.add_argument("--layers", help="Comma-separated list of layers to scan (default: all layers)", default=None)
    ap.add_argument("--template", help="Existing BBS .xlsx to clone header/style from. "
                                        "If omitted, auto-detects an .xlsx file inside FOLDER.", default=None)
    ap.add_argument("--member", help="Member/element name for column A of the generated BBS", default="MEMBER")
    ap.add_argument("--outdir", help="Output directory (default: FOLDER)", default=None)
    args = ap.parse_args()

    # --- resolve file locations -------------------------------------------------
    dxf_path = args.dxf_path
    if dxf_path is None:
        dxf_path = find_file(FOLDER, ".dxf", "DXF drawing")
        if dxf_path is None:
            print(f"[!] No .dxf file found in:\n    {FOLDER}\n"
                  f"    Drop your drawing's DXF file into that folder, or pass a path directly:\n"
                  f"    python dxf_bbs_extractor.py path\\to\\your.dxf")
            sys.exit(1)

    template_path = args.template
    if template_path is None:
        template_path = find_file(FOLDER, ".xlsx", "BBS template")
        if template_path:
            print(f"(No --template given -- auto-using the .xlsx found in FOLDER. "
                  f"Pass --template none if you don't want to clone a template's style.)")

    outdir = args.outdir or FOLDER

    layers = [l.strip() for l in args.layers.split(",")] if args.layers else None

    print(f"Reading {dxf_path} ...")
    result = process_dxf(
        dxf_path=dxf_path,
        outdir=outdir,
        template_path=template_path,
        member_name=args.member,
        layers=layers,
    )

    print(f"Found {result['num_labels']} text-bearing entities on the target layer(s).")
    print(f"  {result['num_parsed']} matched the bar-label pattern.")
    if result["num_unparsed"]:
        print(f"  {result['num_unparsed']} did not match -- listed in extraction_report.csv for manual review.")

    print(f"  Leader-dot geometry: {result['num_dots']} dot(s) total, "
          f"{result['num_chains']} bar chain(s) ({result['num_bent_chains']} multi-segment), "
          f"{result['num_dots_matched']}/{result['num_dots']} dot(s) matched to a mark + bar geometry within tolerance.")
    if result["num_dots_unmatched"]:
        print(f"  {result['num_dots_unmatched']} dot(s) could not be confidently matched -- see dot_bar_matches.csv.")

    if result["missing_marks"]:
        print(f"[!] Gap in bar mark numbering -- mark(s) {', '.join(str(m) for m in result['missing_marks'])} "
              f"never appear on the drawing but sit between marks that do. "
              f"Check whether these were skipped by mistake.")

    print(
        f"\nWrote:\n"
        f"  {result['report_path']}\n"
        f"  {result['summary_path']}\n"
        f"  {result['dot_match_path']}\n"
        f"  {result['bbs_path']}"
    )


if __name__ == "__main__":
    main()
