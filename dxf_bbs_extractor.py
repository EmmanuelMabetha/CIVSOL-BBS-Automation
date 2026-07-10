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
from collections import defaultdict, Counter
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

# Fallback matching for label styles with no true defpoint (plain TEXT/MTEXT
# callouts, as opposed to DIMENSION text overrides). These don't sit
# axis-aligned with their dot, so we fall back to "nearest label by raw
# distance" -- but ONLY when that nearest label is clearly closer than the
# next-closest candidate. If two labels are roughly equidistant from a dot,
# picking either one risks silently attributing a real bar length to the
# wrong mark, so those are left unmatched and flagged instead of guessed.
PROXIMITY_MATCH_MAX_MM = 4000   # a label further than this from a dot is not considered a candidate at all
PROXIMITY_AMBIGUITY_GAP_MM = 200  # min gap between 1st and 2nd nearest label to accept the 1st as unambiguous

# Leader-chain matching: some drafters connect a dot to its label via an
# actual drawn L-shaped leader (a DIMENSION-layer LINE, axis-aligned with
# the dot but not touching it, bending once to run to near the MTEXT)
# instead of either a true DIMENSION defpoint or raw proximity. Verified
# against real geometry (see project notes) before being added here --
# this is not a speculative feature.
LEADER_LINE_LAYER_ALIASES = ("DIMENSION", "DIMENSIONS")
LEADER_X_ALIGN_TOL_MM = 5     # how tightly the first leader segment must share a coordinate with the dot
LEADER_BEND_TOL_MM = 5        # how tightly two leader segments must touch to count as one continuous chain
LEADER_TIP_TOL_MM = 400       # how close the chain's final point must land to an MTEXT to accept it
LEADER_MAX_HOPS = 4           # longest chain (in line segments) that will be followed
LEADER_START_AMBIGUITY_GAP_MM = 50  # if two starting segments are this close in distance, refuse rather than guess


# Ray-cast fallback: when the chain dead-ends (no more connected segments)
# without landing near a label, continue the LAST segment's own direction
# as a ray and look for a label in a narrow corridor ahead of it -- this
# is for leaders that point at their label without quite touching it or
# closing an exact right-angle bend. Verified against real geometry
# (mark 11 case) before being added, not speculative.
LEADER_RAY_CORRIDOR_MM = 200   # how far off the ray's line a label may sit and still count
LEADER_RAY_AMBIGUITY_GAP_MM = 150  # min gap (along the ray) between 1st and 2nd hit to accept the 1st

# Company template names are strict -- these are NOT fuzzy/substring matches,
# just an explicit, closed whitelist of the exact names each engineer's
# drawing is allowed to use for these two layers. Matching is
# case-insensitive only (a drafter using "reinforcement" vs "REINFORCEMENT"
# is not an error worth flagging). Add a new alias here ONLY when it's a
# genuine, approved alternate name in use on the standard -- this list is
# meant to stay short and deliberate, not silently absorb typos.
DOT_LAYER_ALIASES = ("DIMENSION", "DIMENSIONS")
BAR_LAYER_ALIASES = ("REINFORCEMENT", "REINF")


def _layer_matches(entity_layer: str, aliases: "tuple[str, ...]") -> bool:
    """Case-insensitive match against a strict, explicit alias whitelist."""
    return entity_layer.strip().upper() in {a.upper() for a in aliases}


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
    defpoint: Optional[tuple] = None  # anchor point used for leader-dot axis-alignment matching
    anchor_source: Optional[str] = None  # "defpoint" (DIMENSION) or "insertion point" (TEXT/MTEXT) -- for traceability
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


def suggest_shape_code(suffix: str) -> Optional[int]:
    """Very light heuristic hint only -- an engineer must confirm this.

    60 = closed link (BS8666 shape code), 20 = straight bar. Anything
    else is left blank rather than guessed.
    """
    upper = suffix.upper()
    if any(k in upper for k in LINK_KEYWORDS):
        return 60
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

        # Two different drafting styles show up across engineers:
        #   - DIMENSION-based labels (text override on a real dimension
        #     object) carry a genuine "defpoint" -- the geometric anchor
        #     AutoCAD uses for that dimension.
        #   - Plain TEXT/MTEXT labels have no such concept at all. For
        #     these, the entity's own insertion point is the only anchor
        #     available -- and on an orthogonally-drafted leader it plays
        #     the same role (it sits axis-aligned with the dot it belongs
        #     to). Using it lets the same axis-alignment matching logic
        #     work for both drafting conventions without requiring either
        #     drawing to be redrafted.
        defpoint = None
        anchor_source = None
        if entity.dxftype() == "DIMENSION":
            try:
                defpoint = (float(entity.dxf.defpoint[0]), float(entity.dxf.defpoint[1]))
                anchor_source = "defpoint"
            except Exception:
                defpoint = None
        if defpoint is None and entity.dxftype() in ("TEXT", "MTEXT"):
            defpoint = (round(x, 1), round(y, 1))
            anchor_source = "insertion point"

        label = BarLabel(
            layer=layer,
            entity_type=entity.dxftype(),
            handle=entity.dxf.handle,
            raw_text=cleaned,
            x=round(x, 1),
            y=round(y, 1),
            measured_length=get_measured_length(entity),
            defpoint=defpoint,
            anchor_source=anchor_source,
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


def find_solid_dots(doc, layer_aliases: "tuple[str, ...]" = DOT_LAYER_ALIASES) -> list[tuple]:
    """Find solid-filled circular markers (HATCH with solid_fill=1) on any
    of the given (case-insensitive, strict-whitelist) layer names,
    returned as a list of (x, y) centers.

    Two different boundary representations are both supported, since
    different CAD setups/exports produce different ones for what is
    visually the same solid dot:
      - EdgePath boundary with a true arc/circle edge (has .center directly).
      - PolylinePath boundary with exactly 2 vertices, both bulge ~= 1.0
        (two semicircular arcs forming a full circle) -- the same
        representation AutoCAD's DONUT command uses, just wrapped inside
        a HATCH instead of standing alone as an LWPOLYLINE.
    """
    msp = doc.modelspace()
    dots = []
    for h in msp.query("HATCH"):
        if not _layer_matches(h.dxf.layer, layer_aliases):
            continue
        if not h.dxf.solid_fill:
            continue
        for path in h.paths:
            for edge in getattr(path, "edges", []):
                center = getattr(edge, "center", None)
                if center is not None:
                    dots.append((float(center[0]), float(center[1])))
            vertices = getattr(path, "vertices", None)
            if vertices and len(vertices) == 2:
                (x1, y1, b1), (x2, y2, b2) = ((v[0], v[1], v[-1]) for v in vertices)
                if abs(abs(b1) - 1.0) <= 0.05 and abs(abs(b2) - 1.0) <= 0.05:
                    dots.append(((x1 + x2) / 2.0, (y1 + y2) / 2.0))
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


def find_all_dots(doc, hatch_layer_aliases: "tuple[str, ...]" = DOT_LAYER_ALIASES, donut_layer: Optional[str] = None) -> list[tuple]:
    """Merge both known dot conventions (HATCH solid-fill and DONUT
    LWPOLYLINE) into one list, de-duplicating any that coincide (in case
    both conventions somehow mark the same physical point)."""
    dots = find_solid_dots(doc, layer_aliases=hatch_layer_aliases) + find_donut_dots(doc, layer=donut_layer)
    deduped = []
    for d in dots:
        if not any(_dist(d, existing) < 1.0 for existing in deduped):
            deduped.append(d)
    return deduped


def find_bar_lines(doc, layer_aliases: "tuple[str, ...]" = BAR_LAYER_ALIASES) -> list[tuple]:
    """Find candidate bar LINE entities on any of the given (case-insensitive,
    strict-whitelist) layer names, returned as a list of
    (start_xy, end_xy, length). LWPOLYLINE is deliberately excluded -- on
    these drawings those turned out to be small rectangular symbols, not
    bars, and summing their segment lengths gives a wrong "length"."""
    msp = doc.modelspace()
    lines = []
    for l in msp.query("LINE"):
        if not _layer_matches(l.dxf.layer, layer_aliases):
            continue
        s = (float(l.dxf.start[0]), float(l.dxf.start[1]))
        e = (float(l.dxf.end[0]), float(l.dxf.end[1]))
        lines.append((s, e, _dist(s, e)))
    return lines


def _axis_aligned_dist(p1: tuple, p2: tuple) -> float:
    """Distance along whichever single axis (X or Y) the two points are
    aligned on. Near zero for two points that share a coordinate (as a
    dot and its true label's defpoint do, on an orthogonally-drafted
    leader) -- large otherwise, even if the raw straight-line distance
    between them happens to be small. See module notes above."""
    return min(abs(p1[0] - p2[0]), abs(p1[1] - p2[1]))


def find_leader_lines(doc, layer_aliases: "tuple[str, ...]" = LEADER_LINE_LAYER_ALIASES) -> list[tuple]:
    """Find candidate leader-line segments (LINE entities on the dimension
    layer) that might chain a dot to its label. Returned as a plain list
    of (start_xy, end_xy) -- unlike bar lines, length isn't meaningful
    here, only connectivity."""
    msp = doc.modelspace()
    lines = []
    for l in msp.query("LINE"):
        if not _layer_matches(l.dxf.layer, layer_aliases):
            continue
        s = (float(l.dxf.start[0]), float(l.dxf.start[1]))
        e = (float(l.dxf.end[0]), float(l.dxf.end[1]))
        lines.append((s, e))
    return lines


def _ray_cast_label(near: tuple, point: tuple, label_points: list) -> Optional[tuple]:
    """From `point`, continue in the direction (near -> point) as a ray and
    look for a label sitting in a narrow corridor ahead of it. Handles
    leaders that point at their label without an exact touching bend --
    verified against real geometry (a leader ending ~200mm off-axis and
    ~2000mm short of its label, but clearly aimed at it) before being
    added. Refuses if two labels compete for the same ray."""
    if not label_points:
        return None
    dx, dy = point[0] - near[0], point[1] - near[1]
    length = math.hypot(dx, dy)
    if length == 0:
        return None
    ux, uy = dx / length, dy / length
    hits = []
    for pt, label in label_points:
        vx, vy = pt[0] - point[0], pt[1] - point[1]
        along = vx * ux + vy * uy       # distance ahead, along the ray
        perp = abs(vx * uy - vy * ux)   # distance off the ray's line
        if along > 0 and perp <= LEADER_RAY_CORRIDOR_MM:
            hits.append((along, label))
    if not hits:
        return None
    hits.sort(key=lambda h: h[0])
    if len(hits) > 1 and (hits[1][0] - hits[0][0]) < LEADER_RAY_AMBIGUITY_GAP_MM:
        return None  # two labels compete for this ray -- refuse rather than guess
    return (hits[0][1], round(hits[0][0], 1))


def walk_leader_chain(dot: tuple, leader_lines: list, label_points: list) -> Optional[tuple]:
    """Trace a real, drawn leader chain from a dot to its label, instead of
    guessing by proximity. The chain is: dot -(axis-aligned, gapped)->
    first LINE segment -(touching bend)-> possibly more segments. At
    every point along the way, and whenever the chain dead-ends, two
    checks are tried before giving up: is a label already close by, and
    does continuing this segment's direction as a ray point straight at
    one (see _ray_cast_label) -- covers leaders that end near, but not
    exactly touching, their label.

    Refuses (returns None) at any point more than one candidate could
    apply -- an ambiguous starting segment, an ambiguous bend, or a ray
    with two competing labels -- rather than picking one and risking a
    silently wrong match. Every stage here was tuned and verified against
    real chain geometry pulled from the actual drawing, not assumed.

    label_points: list of (x, y), BarLabel pairs -- ALL parsed labels,
    including links. Filtering out links before this search would let the
    chain wander onto an unrelated straight-bar label whenever a dot's
    true match is a link; the caller decides what to do with a link match
    (skip its length) only after the geometry has resolved it correctly.

    Returns (matched_label, tip_distance_mm) or None.
    """
    candidates = []
    for s, e in leader_lines:
        for near, far in ((s, e), (e, s)):
            if abs(near[0] - dot[0]) <= LEADER_X_ALIGN_TOL_MM or abs(near[1] - dot[1]) <= LEADER_X_ALIGN_TOL_MM:
                candidates.append((_dist(dot, near), near, far))
    if not candidates:
        return None
    candidates.sort(key=lambda c: c[0])
    if len(candidates) > 1 and (candidates[1][0] - candidates[0][0]) < LEADER_START_AMBIGUITY_GAP_MM:
        return None  # ambiguous which segment to start from

    _, near, point = candidates[0]
    visited = {near}
    for _hop in range(LEADER_MAX_HOPS):
        # Check right here first -- if this point is already close to a
        # label, take it. This matters when several different dots' own
        # leader lines converge at a shared junction near the text: each
        # dot's chain reaches the junction independently, and the
        # junction itself is effectively "at the label" even though
        # multiple other branches also touch that same point (which would
        # otherwise look ambiguous).
        if label_points:
            nearest_pt, nearest_label = min(label_points, key=lambda lp: _dist(point, lp[0]))
            tip_dist = _dist(point, nearest_pt)
            if tip_dist <= LEADER_TIP_TOL_MM:
                return (nearest_label, round(tip_dist, 1))

        touching = [
            b for s, e in leader_lines for a, b in ((s, e), (e, s))
            if _dist(a, point) <= LEADER_BEND_TOL_MM and b not in visited and _dist(b, point) > LEADER_BEND_TOL_MM
        ]
        if not touching:
            return _ray_cast_label(near, point, label_points)
        if len(touching) > 1:
            return None  # ambiguous bend -- more than one line touches here
        visited.add(point)
        near, point = point, touching[0]

    return _ray_cast_label(near, point, label_points)


def detect_suspicious_length_clusters(bar_lines: list, min_repeat_count: int = 15) -> set:
    """Flag length values that repeat suspiciously often (rounded to the
    nearest 0.5mm) among candidate bar lines. Real bar geometry varies
    with footing/member size and rarely repeats to sub-mm precision dozens
    of times across unrelated marks -- a length appearing far more often
    than everything else in the file is more likely a fixed reference or
    symbol length than genuine coincidence. This is independent of
    link/non-link status -- verified link lengths were found to span
    350-6200mm (not a tight cluster), so "is it link-length" is not a
    reliable signal on its own; unusual repetition frequency is.

    Returns the set of suspicious rounded lengths (empty if nothing in
    the file stands out from the rest)."""
    counts: dict = defaultdict(int)
    for _, _, length in bar_lines:
        counts[round(length * 2) / 2] += 1
    if not counts:
        return set()
    freqs = sorted(counts.values())
    median_freq = freqs[len(freqs) // 2]
    return {length for length, cnt in counts.items() if cnt >= min_repeat_count and cnt >= median_freq * 5}


def match_dot_lengths(
    labels: list[BarLabel],
    dots: list[tuple],
    bar_lines: list[tuple],
    leader_lines: Optional[list] = None,
    dot_label_tol: float = DOT_LABEL_TOL_MM,
    bar_match_tol: float = BAR_MATCH_TOL_MM,
) -> tuple[dict, list[dict]]:
    """For each dot: find its bar-mark label (three methods, tried in order
    of reliability -- see below) and the closest bar LINE (by perpendicular
    distance). If both are within tolerance, record that bar LINE's length
    against that mark.

    IMPORTANT: label matching always searches the FULL set of parsed
    labels, including links. Filtering links out before matching would let
    a dot whose true match is a link wander onto a nearby non-link label
    instead of correctly finding no usable length -- that was a real bug
    caught during testing, not a hypothetical one. Once a label is
    resolved, links are excluded from the length output (same as always),
    but only *after* the geometry has correctly identified them as links.

    Returns:
        lengths_by_mark: {mark: [(length_mm, match_method), ...]}
        match_records: one dict per dot, for a full traceability report
    """
    all_candidates = [l for l in labels if l.parsed_ok and l.defpoint is not None]
    label_points = [(l.defpoint, l) for l in all_candidates]
    leader_lines = leader_lines or []
    suspicious_lengths = detect_suspicious_length_clusters(bar_lines)

    lengths_by_mark: dict = defaultdict(list)
    match_records = []

    for dot in dots:
        record = {"Dot X": round(dot[0], 1), "Dot Y": round(dot[1], 1)}

        match_method = None
        ambiguous = False
        best_label, label_dist = None, None

        if all_candidates:
            # Stage 1: axis-alignment is a HARD FILTER -- only candidates
            # that genuinely share a coordinate with the dot (a real
            # leader) qualify at all. This is the reliable path for
            # DIMENSION-style labels (true defpoint).
            aligned = [l for l in all_candidates if _axis_aligned_dist(dot, l.defpoint) <= dot_label_tol]
            if aligned:
                best_label = min(aligned, key=lambda l: _dist(dot, l.defpoint))
                label_dist = _axis_aligned_dist(dot, best_label.defpoint)
                match_method = "axis-aligned"
            else:
                # Stage 2: geometric leader-chain -- a real, drawn L-shaped
                # connector (dot -> axis-aligned but gapped -> LINE -> bend
                # -> LINE -> near a label). Verified against real geometry;
                # refuses on any ambiguous branch rather than guessing.
                chain_result = walk_leader_chain(dot, leader_lines, label_points)
                if chain_result:
                    best_label, label_dist = chain_result
                    match_method = "leader-chain"
                else:
                    # Stage 3: last resort -- nearest label by raw
                    # distance, accepted ONLY if unambiguously closer than
                    # the runner-up. See PROXIMITY_* constants above.
                    by_dist = sorted(all_candidates, key=lambda l: _dist(dot, l.defpoint))
                    nearest = by_dist[0]
                    nearest_dist = _dist(dot, nearest.defpoint)
                    runner_up_dist = _dist(dot, by_dist[1].defpoint) if len(by_dist) > 1 else float("inf")
                    if nearest_dist <= PROXIMITY_MATCH_MAX_MM and (runner_up_dist - nearest_dist) >= PROXIMITY_AMBIGUITY_GAP_MM:
                        best_label = nearest
                        label_dist = nearest_dist
                        match_method = "nearest-proximity"
                    elif nearest_dist <= PROXIMITY_MATCH_MAX_MM:
                        ambiguous = True

        is_link = best_label is not None and suggest_shape_code(best_label.suffix) == 60

        if bar_lines:
            best_bar_dist, best_bar_len = min(
                ((_point_to_segment_dist(dot, s, e), length) for s, e, length in bar_lines),
                key=lambda x: x[0],
            )
        else:
            best_bar_dist, best_bar_len = None, None

        record["Matched Mark"] = best_label.mark if best_label else None
        record["Is Link"] = is_link
        record["Match Method"] = match_method if best_label else ("ambiguous -- not matched" if ambiguous else None)
        record["Label Distance (mm)"] = round(label_dist, 1) if label_dist is not None else None
        record["Bar Line Distance (mm)"] = round(best_bar_dist, 1) if best_bar_dist is not None else None
        record["Bar Length (mm)"] = round(best_bar_len, 1) if best_bar_len is not None else None
        is_suspicious_length = best_bar_len is not None and any(
            abs(best_bar_len - s) <= 1 for s in suspicious_lengths
        )
        record["Suspicious Length"] = is_suspicious_length

        if match_method in ("axis-aligned", "leader-chain"):
            label_ok = best_label is not None and not is_link
        elif match_method == "nearest-proximity":
            label_ok = best_label is not None and not is_link
        else:
            label_ok = False
        bar_ok = best_bar_len is not None and best_bar_dist <= bar_match_tol
        record["Accepted"] = bool(label_ok and bar_ok)

        if label_ok and bar_ok:
            # Axis-aligned and leader-chain are both geometrically
            # verified (a real defpoint, or a real traced, ambiguity-
            # guarded leader chain) and are trusted directly.
            # Nearest-proximity has no such verification -- still
            # recorded here for the audit trail, but NOT fed into the
            # BBS length column by build_summary() (see Length
            # Confidence there). is_suspicious_length flags a length
            # that repeats suspiciously often across the drawing --
            # still trusted and used, but surfaced for a human check.
            lengths_by_mark[best_label.mark].append((round(best_bar_len, 1), match_method, is_suspicious_length))

        match_records.append(record)

    return dict(lengths_by_mark), match_records


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
            "Anchor Source": l.anchor_source,
            "X": l.x,
            "Y": l.y,
        })
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(["Parsed OK", "Bar Mark"], ascending=[False, True], na_position="last")
    return df


def build_summary(labels: list[BarLabel], dot_lengths: Optional[dict] = None) -> pd.DataFrame:
    """dot_lengths: {mark: [length_mm, ...]}, from match_dot_lengths(). If
    not supplied, no lengths are populated (length columns come back empty
    rather than falling back to the old, unreliable DIMENSION-measured
    value)."""
    dot_lengths = dot_lengths or {}

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

        # LINKS: excluded from length entirely (see module notes) -- a
        # link's real length is its own bend perimeter, not anything we
        # currently extract, so length stays blank rather than guessed.
        #
        # STRAIGHT BARS: length now comes from the leader-dot -> bar LINE
        # match (match_dot_lengths()), NOT from the DIMENSION's own
        # measured value -- that measured value was found to frequently
        # be an unrelated leftover distance (see module notes). If a mark
        # has no accepted dot match, length is left blank rather than
        # falling back to that unreliable number.
        is_link = suggest_shape_code(" ".join(suffixes)) == 60

        LENGTH_VARIANCE_FLAG_MM = 100  # spread beyond this within one mark's trusted lengths gets flagged

        if is_link:
            longest_length = None
            valid_lengths = []
            proximity_lengths = []
            suspicious_length_values = []
        else:
            raw_matches = dot_lengths.get(mark, [])
            # Only axis-aligned matches (true DIMENSION defpoint geometry)
            # are trusted for the actual length column -- proximity
            # matches proved unreliable (decoy tick-mark lines can be
            # closer to a dot than the real bar). Proximity matches are
            # still surfaced separately, clearly marked as NOT used, so
            # you have the raw data to check by hand if you want it.
            trusted_matches = [x for x in raw_matches if x[1] in ("axis-aligned", "leader-chain")]
            valid_lengths = sorted(x[0] for x in trusted_matches)
            proximity_lengths = sorted(x[0] for x in raw_matches if x[1] == "nearest-proximity")
            suspicious_length_values = sorted({x[0] for x in trusted_matches if x[2]})
            longest_length = max(valid_lengths) if valid_lengths else None

            if len(valid_lengths) > 1 and (valid_lengths[-1] - valid_lengths[0]) > LENGTH_VARIANCE_FLAG_MM:
                spread_note = (
                    f"LENGTH VARIES: mark {mark} has trusted lengths from {valid_lengths[0]} to "
                    f"{valid_lengths[-1]}mm (spread {valid_lengths[-1]-valid_lengths[0]:.1f}mm across "
                    f"{len(valid_lengths)} verified instance(s)) -- confirm with engineer whether this "
                    f"is genuine footing-to-footing variation or a mismatched length"
                )
                conflict = (conflict + "; " + spread_note) if conflict else spread_note
                conflicts.append(spread_note)
            if suspicious_length_values:
                suspicious_note = (
                    f"POSSIBLE DECOY LENGTH: mark {mark} includes length(s) "
                    f"{', '.join(str(v) for v in suspicious_length_values)}mm that repeat unusually often "
                    f"across the drawing (see extraction notes) -- verify before trusting"
                )
                conflict = (conflict + "; " + suspicious_note) if conflict else suspicious_note
                conflicts.append(suspicious_note)

        rows.append({
            "Bar Mark": mark,
            "Type": next(iter(types)) if len(types) == 1 else "/".join(sorted(types)),
            "Diameter (mm)": next(iter(dias)) if len(dias) == 1 else "/".join(str(d) for d in sorted(dias)),
            "Occurrences on Drawing": len(items),
            "Total No. Off": total_count,
            "Spacing (mm c/c)": next(iter(spacings)) if len(spacings) == 1 else ("/".join(str(s) for s in sorted(spacings)) if spacings else None),
            "Notes": "; ".join(suffixes),
            "Is Link": is_link,
            "Suggested Shape Code (verify)": 60 if is_link else None,
            "Longest DXF Length (mm)": longest_length,
            "All DXF Lengths (mm)": ", ".join(str(x) for x in valid_lengths) if valid_lengths else None,
            "Length Confidence": ("geometrically verified (axis-aligned/leader-chain)" if valid_lengths else None),
            "Proximity-Matched Lengths (NOT used -- unverified)": ", ".join(str(x) for x in proximity_lengths) if proximity_lengths else None,
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
PROXIMITY_LENGTH_FILL = PatternFill("solid", start_color="FCE5CD", end_color="FCE5CD")  # pale orange -- DXF length, but matched by nearest-label proximity, not axis-aligned geometry -- spot-check this one
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

        if is_link:
            # We know the current DXF geometry for links measures the
            # spaced RUN LENGTH, not the individual link length -- so we
            # deliberately show nothing here rather than a number that
            # would look like a real length. Length and shape code (60)
            # are left for manual entry until link geometry is handled
            # separately.
            ws.cell(row=r, column=8, value=None)  # H: length -- not available
            ws.cell(row=r, column=8).fill = REVIEW_FILL
            ws.cell(row=r, column=9, value=None)  # I: shape code -- not assumed for links
            ws.cell(row=r, column=9).fill = REVIEW_FILL
        else:
            if longest_length is not None and str(longest_length).lower() != "nan":
                ws.cell(row=r, column=8, value=float(longest_length))  # H: from DXF, axis-aligned match only
                ws.cell(row=r, column=8).fill = DXF_LENGTH_FILL
            else:
                ws.cell(row=r, column=8, value=length_formula)  # H: needs J (length) filled in
                ws.cell(row=r, column=8).fill = REVIEW_FILL
            # Per your instruction: assume shape code 20 (straight bar) for
            # every non-link mark for now, until told otherwise. This is an
            # ASSUMPTION, not something read off the drawing -- flagged with
            # the same "needs verifying" fill as the rest of the row, and
            # called out explicitly in the notes column.
            ws.cell(row=r, column=9, value=20)  # I: Shape Code -- assumed, verify
            ws.cell(row=r, column=9).fill = REVIEW_FILL

        for col in (10, 11, 12, 13, 14):  # J..N -- bend dimensions, needs engineering input
            ws.cell(row=r, column=col).fill = REVIEW_FILL

        weight_formula = (
            f"=IF(D{r}=8,G{r}*H{r}*0.395,IF(D{r}=10,G{r}*H{r}*0.616,"
            f"IF(D{r}=12,G{r}*H{r}*0.888,IF(D{r}=16,G{r}*H{r}*1.579,"
            f"IF(D{r}=20,G{r}*H{r}*2.466,IF(D{r}=25,G{r}*H{r}*3.854,"
            f"IF(D{r}=32,G{r}*H{r}*6.313))))))/1000/1000"
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
            ref_bits.append("LINK -- length not extracted (DXF measures run length, not link length); shape code + dims need manual entry")
        else:
            ref_bits.append("shape code 20 (straight bar) assumed for now -- verify")
            if _is_set(rec.get("All DXF Lengths (mm)")):
                ref_bits.append(f"all DXF lengths found: {rec['All DXF Lengths (mm)']} mm (longest used in column H)")
            if _is_set(rec.get("Proximity-Matched Lengths (NOT used -- unverified)")):
                ref_bits.append(f"NOT USED -- proximity-matched candidate length(s) found but not trusted (see dot_bar_matches.csv): {rec['Proximity-Matched Lengths (NOT used -- unverified)']} mm; length needs manual entry from drawing")
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
# 7. REUSABLE PIPELINE
# ---------------------------------------------------------------------------

def process_dxf(
    dxf_path: str,
    template_path: Optional[str] = None,
    member_name: str = "MEMBER",
    layers: Optional[list[str]] = None,
    outdir: Optional[str] = None,
) -> dict:
    """
    Core extraction pipeline. Reads a DXF, parses bar labels, matches
    leader-dot geometry, and writes BBS outputs.

    Returns
    -------
    dict
        report_df, summary_df, dot_matches_df, bbs_path, report_path,
        summary_path, dot_match_path, missing_marks, parsed_count,
        unparsed_count, total_dots, accepted_dots, method_counts
    """
    outdir = Path(outdir) if outdir else Path(FOLDER)
    outdir.mkdir(parents=True, exist_ok=True)

    doc = ezdxf.readfile(dxf_path)
    labels = extract_labels(doc, layers=layers)

    parsed = [l for l in labels if l.parsed_ok]
    unparsed = [l for l in labels if not l.parsed_ok]

    # --- leader-dot -> bar length matching ---------------------------------
    dots = find_all_dots(doc, hatch_layer_aliases=DOT_LAYER_ALIASES, donut_layer=None)
    bar_lines = find_bar_lines(doc, layer_aliases=BAR_LAYER_ALIASES)
    leader_lines = find_leader_lines(doc, layer_aliases=LEADER_LINE_LAYER_ALIASES)
    dot_lengths, dot_match_records = match_dot_lengths(
        labels, dots, bar_lines, leader_lines=leader_lines
    )
    accepted = sum(1 for r in dot_match_records if r["Accepted"])

    # --- reporting ---------------------------------------------------------
    report_df = build_extraction_report(labels)
    summary_df = build_summary(labels, dot_lengths=dot_lengths)
    missing_marks = find_missing_marks(summary_df)

    # --- write outputs -----------------------------------------------------
    report_path = outdir / "extraction_report.csv"
    summary_path = outdir / "bar_mark_summary.csv"
    dot_match_path = outdir / "dot_bar_matches.csv"
    bbs_path = outdir / "BBS_from_DXF.xlsx"

    report_df.to_csv(report_path, index=False)
    summary_df.to_csv(summary_path, index=False)
    pd.DataFrame(dot_match_records).to_csv(dot_match_path, index=False)
    write_bbs_workbook(
        summary_df,
        str(bbs_path),
        template_path=template_path,
        member_name=member_name,
        missing_marks=missing_marks,
    )

    return {
        "report_df": report_df,
        "summary_df": summary_df,
        "dot_matches_df": pd.DataFrame(dot_match_records),
        "bbs_path": str(bbs_path),
        "report_path": str(report_path),
        "summary_path": str(summary_path),
        "dot_match_path": str(dot_match_path),
        "missing_marks": missing_marks,
        "parsed_count": len(parsed),
        "unparsed_count": len(unparsed),
        "total_dots": len(dots),
        "accepted_dots": accepted,
        "method_counts": Counter(
            r["Match Method"] for r in dot_match_records if r["Accepted"]
        ),
    }


# ---------------------------------------------------------------------------
# 8. CLI
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("dxf_path", nargs="?", default=None,
                    help="Path to the DXF file to extract from. "
                         "If omitted, auto-detects a .dxf file inside FOLDER.")
    ap.add_argument("--layers", help="Comma-separated list of layers to scan (default: all layers)", default=None)
    ap.add_argument("--template", help="Existing BBS .xlsx to clone header/style from. "
                                       "If omitted, auto-detects an .xlsx file inside FOLDER.", default=None)
    ap.add_argument("--member", help="Member/element name for column A of the generated BBS", default="MEMBER")
    ap.add_argument("--outdir", help="Output directory (default: FOLDER)", default=None)
    args = ap.parse_args()

    # --- resolve file locations --------------------------------------------
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
        template_path=template_path,
        member_name=args.member,
        layers=layers,
        outdir=outdir,
    )

    print(f"Found {len(result['report_df'])} text-bearing entities on the target layer(s).")
    print(f"  {result['parsed_count']} matched the bar-label pattern.")
    if result['unparsed_count']:
        print(f"  {result['unparsed_count']} did not match -- listed in extraction_report.csv for manual review.")

    print(f"  Leader-dot geometry: {result['total_dots']} total dot(s), "
          f"{result['accepted_dots']}/{result['total_dots']} dot(s) matched to a mark + bar length within tolerance.")
    if result['method_counts']:
        breakdown = ", ".join(f"{v} {k}" for k, v in result['method_counts'].most_common())
        print(f"    ({breakdown})")
    unmatched = result['total_dots'] - result['accepted_dots']
    if unmatched:
        print(f"  {unmatched} dot(s) could not be confidently matched -- see dot_bar_matches.csv.")

    if result['missing_marks']:
        print(f"[!] Gap in bar mark numbering -- mark(s) {', '.join(str(m) for m in result['missing_marks'])} "
              f"never appear on the drawing but sit between marks that do. "
              f"Check whether these were skipped by mistake.")

    print(f"\nWrote:\n  {result['report_path']}\n  {result['summary_path']}\n  "
          f"{result['dot_match_path']}\n  {result['bbs_path']}")


if __name__ == "__main__":
    main()
