"""
dxf_bbs_extractor.py
=====================

General-purpose tool that reads bar-mark callouts off a DXF drawing
(e.g. "Extractor.dxf") and turns them into a clean, review-ready Bar
Bending Schedule (BBS) starting point in Excel -- or, via process_dxf(),
into a Streamlit app.

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
  4. Traces every solid "dot" marker back to the bar-mark label it
     belongs to (three methods, tried in order of reliability -- see
     section 4b), and from there to the actual bar LINE geometry, so
     length comes from the drawing rather than guesswork.
  5. Writes a BBS workbook in the same layout/formula style as a
     standard SABS/SANS 10100 schedule (matching the uploaded
     BBS01 template), with the objective fields (mark, type, diameter,
     count, length) filled in automatically -- including an explicit,
     confirmed shape code 60 for links/stirrups (see LINK_SHAPE_CODE
     below).

WHAT IT DELIBERATELY DOES **NOT** GUESS
-----------------------------------------
A link/stirrup's individual leg dimensions (A, B) and the general
straight-bar SHAPE CODE for anything bent are left blank for the
engineer to complete. A text label like "20Y8-01-300 LINKS" tells you
the bar type, diameter, mark and spacing -- it does NOT tell you the
leg lengths of the link, and this tool doesn't currently trace closed
loop geometry to measure them. Shape code 60 itself IS written in
automatically for links (see below) since that classification is
reliable from the label text alone; the dimensions that go with it are
not.

USAGE
-----
    python dxf_bbs_extractor.py Extractor.dxf \\
        --template BBS01_-_FOUNDATION_REINFORCEMENT.xlsx \\
        --outdir /mnt/user-data/outputs \\
        --member "FOOTING"

    # restrict to specific layers (default: every layer in the file)
    python dxf_bbs_extractor.py Extractor.dxf --layers DIMS,REINFORCEMENT

    # from Streamlit / any other Python caller, use process_dxf() directly
    # (see section 7 below) instead of going through argparse/CLI.

Outputs (written to --outdir):
    extraction_report.csv   raw, one row per label found on the drawing
    bar_mark_summary.csv    aggregated per bar mark, with conflict flags
    dot_bar_matches.csv     one row per dot, full match traceability
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
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import ezdxf
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# sans282_shape_codes.py is kept as a SEPARATE file, imported here, rather
# than pasted into this script -- see module notes in that file for why.
try:
    import sans282_shape_codes as shape_codes
except ImportError:
    shape_codes = None


# ---------------------------------------------------------------------------
# 0. FILE LOCATION -- only relevant to the CLI / FOLDER auto-detect flow.
#    Streamlit callers should use process_dxf() directly (section 7) and
#    can ignore this entirely.
# ---------------------------------------------------------------------------

FOLDER = r"C:\Users\T14s\OneDrive - University of Cape Town\Desktop\CIVSOL\Automation"


def find_file(folder: str, extension: str, label: str) -> Optional[str]:
    """Look for exactly one file of the given extension in `folder`.

    If there's more than one, asks you to pick. If there's none, returns
    None (the caller decides whether that's fatal)."""
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

# Words that indicate a link/stirrup rather than a straight bar.
LINK_KEYWORDS = ("LINK", "STIRRUP", "TIE")

# BS8666 shape code for a closed rectangular link/stirrup. Written
# directly into the BBS workbook and summary for every mark identified as
# a link -- this classification is reliable from the label text alone
# ("...LINKS"/"STIRRUP"/"TIE" in the suffix). What is NOT auto-filled is
# the A/B leg dimensions that go with it -- those still need the
# engineer's input, since this tool doesn't currently trace closed-loop
# geometry to measure them.
LINK_SHAPE_CODE = 60

# DXF dimensions with near-zero measured length usually mean the label was
# placed as a leader/callout with no real geometry behind it. Anything
# below this is treated as "no usable length" rather than a real one.
MIN_VALID_LENGTH_MM = 50

# ---------------------------------------------------------------------------
# 1b. BAR LENGTH FROM LEADER-DOT GEOMETRY
# ---------------------------------------------------------------------------
# Every straight-bar callout has a solid filled circle ("dot") sitting
# almost exactly on top of the bar it refers to. So instead of trusting a
# DIMENSION's own measured value (frequently an unrelated leftover
# distance from something else on the drawing), we:
#   1. Find every dot on DOT_LAYER_ALIASES (HATCH solid-fill or DONUT).
#   2. Find every candidate bar LINE on BAR_LAYER_ALIASES.
#   3. For each dot, find the closest bar LINE -- if within
#      BAR_MATCH_TOL_MM, that LINE's length is the bar length.
#   4. Work out which bar-mark label that dot belongs to -- see the
#      3-stage matcher in match_dot_lengths() below.
#
# Label matching, in order of reliability:
#   Stage 1 -- axis-alignment: dot and label anchor share an X or Y
#              coordinate almost exactly (orthogonal drafting practice).
#              Works for both a true DIMENSION defpoint AND a plain
#              TEXT/MTEXT label's own insertion point (see extract_labels
#              -- the insertion point is used as a defpoint stand-in for
#              labels with no true DIMENSION geometry behind them).
#   Stage 2 -- leader-chain: a real, drawn connector (one or more
#              DIMENSION-layer LINEs) traced from the dot to near a
#              label, for leaders that don't sit exactly axis-aligned
#              with their own label but do connect to it via drawn
#              line-work. Refuses on any ambiguous branch.
#   Stage 3 -- nearest-proximity: last resort, accepted only when
#              unambiguously closer than the runner-up candidate.
#
# Links are matched the same way as straight bars (never filtered out
# beforehand -- see match_dot_lengths() docstring for why), but are
# excluded from the LENGTH output once identified.

DOT_LABEL_TOL_MM = 5          # max axis-aligned offset between a dot and its label's anchor
BAR_MATCH_TOL_MM = 10         # max distance from a dot to the bar LINE it sits on

PROXIMITY_MATCH_MAX_MM = 4000     # a label further than this from a dot isn't a candidate at all
PROXIMITY_AMBIGUITY_GAP_MM = 200  # min gap between 1st/2nd nearest label to accept the 1st unambiguously

LEADER_LINE_LAYER_ALIASES = ("DIMENSION", "DIMENSIONS")
LEADER_X_ALIGN_TOL_MM = 5           # how tightly a leader segment's end must share a coordinate with the dot
LEADER_BEND_TOL_MM = 5              # how tightly two leader segments must touch to count as one chain
LEADER_TIP_TOL_MM = 400             # how close the chain's arrival point must land to a label to accept it
LEADER_MAX_HOPS = 4                 # longest chain (in line segments) that will be followed
LEADER_START_AMBIGUITY_GAP_MM = 50  # if two starting segments are this close in distance, refuse rather than guess

LEADER_RAY_CORRIDOR_MM = 200        # how far off a ray's line a label may sit and still count
LEADER_RAY_AMBIGUITY_GAP_MM = 150   # min gap (along the ray) between 1st/2nd hit to accept the 1st

# Strict, explicit layer-name whitelists (case-insensitive only -- not
# fuzzy/substring). Add a new alias here only when it's a genuine,
# approved alternate name in use on the standard.
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
    measured_length: Optional[float]  # DXF dimension geometric value -- reference only, unreliable
    defpoint: Optional[tuple] = None  # anchor point used for leader-dot axis-alignment matching
    anchor_source: Optional[str] = None  # "defpoint" (DIMENSION) or "insertion point" (TEXT/MTEXT)
    count: Optional[int] = None
    bar_type: Optional[str] = None
    diameter: Optional[int] = None
    mark: Optional[str] = None
    spacing: Optional[int] = None
    suffix: str = ""
    parsed_ok: bool = False
    is_link: bool = False   # set once, from suffix, right after parsing -- see parse_bar_label()


# ---------------------------------------------------------------------------
# 3. TEXT CLEANUP + PARSING (links / stirrups / other labels)
# ---------------------------------------------------------------------------

def clean_dxf_text(raw: str) -> str:
    """Strip AutoCAD MTEXT/DIMENSION formatting control codes.

    Dimension text overrides and MTEXT can carry inline formatting such
    as \\P (paragraph break), \\W (width factor), \\H (height), font/colour
    codes, and {..} grouping braces. We only want the visible text."""
    if raw is None:
        return ""
    text = raw
    text = re.sub(r"\\P", " ", text)              # paragraph break -> space
    text = re.sub(r"\\[A-Za-z][^;]*;", "", text)   # \Xvalue; formatting codes
    text = text.replace("{", "").replace("}", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def parse_bar_label(raw_text: str) -> dict:
    """Parse a cleaned label string into its structured fields.

    Returns a dict with parsed_ok=False if the text doesn't match the
    expected bar-mark pattern (e.g. it's an unrelated dimension/note)."""
    text = clean_dxf_text(raw_text)
    m = BAR_LABEL_PATTERN.match(text)
    if not m:
        return {"parsed_ok": False}

    gd = m.groupdict()
    suffix = gd["suffix"].strip()
    return {
        "parsed_ok": True,
        "count": int(gd["count"]),
        "bar_type": gd["type"].upper(),
        "diameter": int(gd["diameter"]),
        "mark": gd["mark"].upper(),
        "spacing": int(gd["spacing"]) if gd["spacing"] else None,
        "suffix": suffix,
        "is_link": is_link_suffix(suffix),
    }


def is_link_suffix(suffix: str) -> bool:
    """True if this label's suffix text identifies it as a link/stirrup/
    tie (LINK_KEYWORDS), rather than a straight/bent bar. This is the
    single source of truth for "is this mark a link" -- every other place
    in the script (matching, summary, workbook) calls this, or reads the
    is_link flag already set on the BarLabel by parse_bar_label(), rather
    than re-checking suffix text independently. Having two slightly
    different link-detection code paths drift apart was a real source of
    inconsistent LINK flagging in an earlier version of this script."""
    return any(k in suffix.upper() for k in LINK_KEYWORDS)


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
        elif etype in ("TEXT", "MTEXT"):
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
        #   - DIMENSION-based labels carry a genuine "defpoint".
        #   - Plain TEXT/MTEXT labels have no such concept -- their own
        #     insertion point is the only anchor available, and on an
        #     orthogonally-drafted leader it plays the same role (it
        #     sits axis-aligned with the dot it belongs to). Using it
        #     lets the same axis-alignment matching logic work for both
        #     drafting conventions.
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
            label.is_link = parsed["is_link"]

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


MIN_DOT_RADIUS_MM = 3
MAX_DOT_RADIUS_MM = 100


def find_solid_dots(doc, layer_aliases: "tuple[str, ...]" = DOT_LAYER_ALIASES) -> list[tuple]:
    """Find solid-filled circular markers (HATCH with solid_fill=1) on any
    of the given (case-insensitive, strict-whitelist) layer names.

    Two different boundary representations are both supported:
      - EdgePath boundary with a true arc/circle edge (has .center directly).
      - PolylinePath boundary with exactly 2 vertices, both bulge ~= 1.0
        (the same representation AutoCAD's DONUT command uses, just
        wrapped inside a HATCH instead of standing alone)."""
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


def find_circle_dots(doc, layer_aliases: "tuple[str, ...]" = DOT_LAYER_ALIASES) -> list[tuple]:
    """Find dots drawn as a plain CIRCLE entity (typically paired with a
    HATCH that fills it solid) -- a different DXF representation of the
    same "solid dot marker" convention, seen on drawings where the HATCH
    boundary is stored in a form find_solid_dots() can't read a center
    from directly. Reads the CIRCLE's own center, so it works regardless
    of how the paired HATCH's boundary happens to be stored."""
    msp = doc.modelspace()
    dots = []
    for c in msp.query("CIRCLE"):
        if not _layer_matches(c.dxf.layer, layer_aliases):
            continue
        radius = float(c.dxf.radius)
        if not (MIN_DOT_RADIUS_MM <= radius <= MAX_DOT_RADIUS_MM):
            continue
        center = c.dxf.center
        dots.append((float(center[0]), float(center[1])))
    return dots


def find_donut_dots(doc, layer: Optional[str] = None) -> list[tuple]:
    """Find AutoCAD DONUT-style dots: a closed LWPOLYLINE with exactly 2
    vertices, both with bulge ~= 1 (two semicircular arcs forming a full
    circle). layer=None scans every layer, since a drafter using DONUT
    may not use the same layer convention as HATCH-based dots."""
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


def find_all_dots(doc, hatch_layer_aliases: "tuple[str, ...]" = DOT_LAYER_ALIASES,
                   circle_layer_aliases: "tuple[str, ...]" = DOT_LAYER_ALIASES,
                   donut_layer: Optional[str] = None) -> list[tuple]:
    """Merge every known dot convention (HATCH solid-fill, plain CIRCLE,
    DONUT LWPOLYLINE) into one list, de-duplicating any that coincide."""
    dots = (
        find_solid_dots(doc, layer_aliases=hatch_layer_aliases)
        + find_circle_dots(doc, layer_aliases=circle_layer_aliases)
        + find_donut_dots(doc, layer=donut_layer)
    )
    deduped = []
    for d in dots:
        if not any(_dist(d, existing) < 1.0 for existing in deduped):
            deduped.append(d)
    return deduped


def find_bar_lines(doc, layer_aliases: "tuple[str, ...]" = BAR_LAYER_ALIASES) -> list[tuple]:
    """Find candidate bar LINE entities on any of the given layer names,
    returned as a list of (start_xy, end_xy, length)."""
    msp = doc.modelspace()
    lines = []
    for l in msp.query("LINE"):
        if not _layer_matches(l.dxf.layer, layer_aliases):
            continue
        s = (float(l.dxf.start[0]), float(l.dxf.start[1]))
        e = (float(l.dxf.end[0]), float(l.dxf.end[1]))
        lines.append((s, e, _dist(s, e)))
    return lines


def find_leader_lines(doc, layer_aliases: "tuple[str, ...]" = LEADER_LINE_LAYER_ALIASES) -> list[tuple]:
    """Find candidate leader-line segments that might chain a dot to its
    label. Returned as (start_xy, end_xy) -- length isn't meaningful here,
    only connectivity."""
    msp = doc.modelspace()
    lines = []
    for l in msp.query("LINE"):
        if not _layer_matches(l.dxf.layer, layer_aliases):
            continue
        s = (float(l.dxf.start[0]), float(l.dxf.start[1]))
        e = (float(l.dxf.end[0]), float(l.dxf.end[1]))
        lines.append((s, e))
    return lines


def _axis_aligned_dist(p1: tuple, p2: tuple) -> float:
    """Distance along whichever single axis (X or Y) the two points are
    aligned on. Near zero for two points that share a coordinate."""
    return min(abs(p1[0] - p2[0]), abs(p1[1] - p2[1]))


def _ray_cast_label(near: tuple, point: tuple, label_points: list) -> Optional[tuple]:
    """From `point`, continue in the direction (near -> point) as a ray and
    look for a label sitting in a narrow corridor ahead of it. Handles
    leaders that point at their label without an exact touching bend.
    Refuses if two labels compete for the same ray."""
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
        return None
    return (hits[0][1], round(hits[0][0], 1))


def walk_leader_chain(dot: tuple, leader_lines: list, label_points: list) -> Optional[tuple]:
    """Trace a real, drawn leader chain from a dot to its label, instead of
    guessing by proximity.

    The starting segment is found by axis-alignment (some endpoint of a
    DIMENSION-layer LINE shares an X or Y coordinate with the dot, even if
    not touching it -- an orthogonally-drafted, gapped leader). Refuses
    outright if two starting segments are close enough in distance to be
    ambiguous.

    IMPORTANT: once that starting segment is found, BOTH of its ends are
    checked for a nearby label before picking a walking direction -- not
    just the end furthest from the dot. On drawings where several dots
    along one row share a single long baseline/extension line (e.g. a
    row of 4+ identical bars all dotted along the same dimension line),
    "furthest from the dot" is only the label-ward direction for SOME of
    those dots -- for a dot sitting close to the label-side end of that
    shared line, the label-ward end is actually the *near* one. Checking
    only the far end silently dropped that dot's match in an earlier
    version of this script (verified against a real 4-dot row where the
    3 dots further from the label matched fine and the 4th, closest to
    the label, did not -- fixed here, all 4 now resolve correctly).

    If neither end of the starting segment is close enough, the chain
    continues hopping through touching segments, checking for a nearby
    label at each stop, until it dead-ends -- at which point a ray-cast
    along the last segment's own direction is tried as a final fallback.

    Refuses (returns None) at any ambiguous branch rather than guessing.

    label_points: list of (x, y), BarLabel pairs -- ALL parsed labels,
    including links (see match_dot_lengths() docstring for why).

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

    # Check BOTH ends of the first segment before committing to a walking
    # direction -- see docstring above for why this matters.
    if label_points:
        for candidate_point in (point, near):
            nearest_pt, nearest_label = min(label_points, key=lambda lp: _dist(candidate_point, lp[0]))
            tip_dist = _dist(candidate_point, nearest_pt)
            if tip_dist <= LEADER_TIP_TOL_MM:
                return (nearest_label, round(tip_dist, 1))

    visited = {near}
    for _hop in range(LEADER_MAX_HOPS):
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
    nearest 0.5mm) among candidate bar lines -- more likely a fixed
    reference/symbol length than genuine coincidence.

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
    """For each dot: find its bar-mark label (3-stage matcher, see module
    notes) and the closest bar LINE (by perpendicular distance). If both
    are within tolerance, record that bar LINE's length against that
    mark. EVERY dot that resolves is counted -- a bar mark drawn with
    several separate dots along a row (multiple parallel identical bars
    in cross-section) gets one entry per dot, not capped at any fixed
    number, so build_summary() sees every occurrence.

    IMPORTANT: label matching always searches the FULL set of parsed
    labels, including links. Filtering links out beforehand would let a
    dot whose true match is a link wander onto a nearby non-link label
    instead of correctly finding no usable length. Once a label is
    resolved, links are excluded from the length output (never their
    length, only their length), using each label's own is_link flag --
    the single source of truth set by parse_bar_label().

    Returns:
        lengths_by_mark: {mark: [(length_mm, match_method, is_suspicious), ...]}
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
            # that genuinely share a coordinate with the dot qualify.
            aligned = [l for l in all_candidates if _axis_aligned_dist(dot, l.defpoint) <= dot_label_tol]
            if aligned:
                best_label = min(aligned, key=lambda l: _dist(dot, l.defpoint))
                label_dist = _axis_aligned_dist(dot, best_label.defpoint)
                match_method = "axis-aligned"
            else:
                # Stage 2: geometric leader-chain.
                chain_result = walk_leader_chain(dot, leader_lines, label_points)
                if chain_result:
                    best_label, label_dist = chain_result
                    match_method = "leader-chain"
                else:
                    # Stage 3: last resort -- nearest label by raw
                    # distance, accepted only if unambiguously closer
                    # than the runner-up.
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

        is_link = best_label is not None and best_label.is_link

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

        label_ok = best_label is not None and not is_link
        bar_ok = best_bar_len is not None and best_bar_dist <= bar_match_tol
        record["Accepted"] = bool(label_ok and bar_ok)

        if label_ok and bar_ok:
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
            "Is Link": l.is_link,
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


LENGTH_VARIANCE_FLAG_MM = 100  # spread beyond this within one mark's trusted lengths gets flagged


def build_summary(labels: list[BarLabel], dot_lengths: Optional[dict] = None) -> pd.DataFrame:
    """dot_lengths: {mark: [(length_mm, match_method, is_suspicious), ...]},
    from match_dot_lengths(). If not supplied, no lengths are populated."""
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

        # is_link is read from the labels themselves (set once, in
        # parse_bar_label(), via is_link_suffix()) -- not re-derived here,
        # so this can never disagree with what match_dot_lengths() used.
        is_link = any(i.is_link for i in items)

        if is_link:
            longest_length = None
            valid_lengths = []
            proximity_lengths = []
            suspicious_length_values = []
        else:
            raw_matches = dot_lengths.get(mark, [])
            # Only axis-aligned / leader-chain matches (real geometric
            # verification) are trusted for the length column.
            # Nearest-proximity matches are surfaced separately, clearly
            # marked as NOT used.
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
                    f"across the drawing -- verify before trusting"
                )
                conflict = (conflict + "; " + suspicious_note) if conflict else suspicious_note
                conflicts.append(suspicious_note)

        rows.append({
            "Bar Mark": mark,
            "Type": next(iter(types)) if len(types) == 1 else "/".join(sorted(types)),
            "Diameter (mm)": next(iter(dias)) if len(dias) == 1 else "/".join(str(d) for d in sorted(dias)),
            "Occurrences on Drawing": len(items),
            "Dot Occurrences Matched": len(dot_lengths.get(mark, [])) if not is_link else None,
            "Total No. Off": total_count,
            "Spacing (mm c/c)": next(iter(spacings)) if len(spacings) == 1 else ("/".join(str(s) for s in sorted(spacings)) if spacings else None),
            "Notes": "; ".join(suffixes),
            "Is Link": is_link,
            "Suggested Shape Code (verify)": LINK_SHAPE_CODE if is_link else None,
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
    """Flag gaps in the bar mark numbering sequence. Only considers
    numeric, non-link bar marks (links are often numbered in their own
    separate short sequence)."""
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

REVIEW_FILL = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")       # pale yellow -- needs your input
DXF_LENGTH_FILL = PatternFill("solid", start_color="D9EAD3", end_color="D9EAD3")   # pale green -- taken from DXF / confirmed classification
LINK_SHAPE_FILL = PatternFill("solid", start_color="CFE2F3", end_color="CFE2F3")   # pale blue -- shape 60 confirmed, leg dims still needed
HEADER_FONT = Font(name="Arial", bold=True, size=9)
BODY_FONT = Font(name="Arial", size=9)
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _copy_header_block(src_ws: Worksheet, dst_ws: Worksheet, rows: int = 11):
    """Clone the title/header rows from an existing BBS template so the
    generated sheet matches house style."""
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


REF_TABLE_START_COL = 19  # column S -- reference table for standard r/h/n lookups


def _write_bend_allowance_table(ws: Worksheet, start_row: int = 1) -> int:
    """Write STANDARD_BEND_ALLOWANCES as a flat lookup table starting at
    REF_TABLE_START_COL. Returns the last row written."""
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

    ws.cell(row=start_row - 1 if start_row > 1 else start_row,
            column=REF_TABLE_START_COL,
            value="Standard bend allowances (SANS 282:2011 Fig 2/3) -- used to auto-look-up r for shape codes 37/38/45. "
                  "If the drawing calls for a NON-STANDARD radius (shape code suffix 'S'), check manually.").font = Font(italic=True, size=8)
    return last_row


def _r_lookup_formula(row: int, ref_last_row: int) -> str:
    key_col = get_column_letter(REF_TABLE_START_COL)
    r_col = get_column_letter(REF_TABLE_START_COL + 5)
    return (
        f'IFERROR(INDEX(${r_col}$2:${r_col}${ref_last_row},'
        f'MATCH(C{row}&D{row},${key_col}$2:${key_col}${ref_last_row},0)),"")'
    )


def _shape_length_formula(row: int, ref_last_row: int) -> str:
    """Length (column H) formula, covering ONLY the shape codes confirmed
    in sans282_shape_codes.py: 20, 41, 42, 37, 38, 45, 39.
    Column mapping: J=A, K=B, L=C, N=n, D=bar diameter."""
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
    are filled in from the drawing. For links/stirrups, shape code
    LINK_SHAPE_CODE (60) is written directly -- that classification is
    reliable from the label text alone -- but the A/B leg dimensions are
    still left blank for the engineer, since this tool doesn't currently
    trace closed-loop geometry to measure them. For straight bars, length
    comes from the DXF where a trusted dot match was found; shape code is
    assumed 20 (straight) pending confirmation, same as before."""
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

        longest_length = rec.get("Longest DXF Length (mm)")
        length_formula = _shape_length_formula(r, ref_last_row)
        is_link = bool(rec.get("Is Link"))

        if is_link:
            # Shape code IS written in -- 60 is a reliable classification
            # from the label text alone. Length stays blank (current DXF
            # geometry for links measures the spaced RUN LENGTH, not the
            # individual link length, so showing a number here would look
            # like real data when it isn't). A/B leg dims also stay blank
            # -- not traced yet -- but with the distinct "confirmed shape,
            # dims still needed" fill rather than the plain review fill,
            # so it reads as "60 is right, just fill in the legs" rather
            # than "nothing has been done here".
            ws.cell(row=r, column=8, value=None)  # H: length -- not available for links
            ws.cell(row=r, column=8).fill = REVIEW_FILL
            ws.cell(row=r, column=9, value=LINK_SHAPE_CODE)  # I: shape code -- confirmed
            ws.cell(row=r, column=9).fill = LINK_SHAPE_FILL
            for col in (10, 11):  # J, K -- A, B leg dimensions
                ws.cell(row=r, column=col).fill = LINK_SHAPE_FILL
            for col in (12, 13, 14):  # L, M, N -- not used by shape 60, leave plain review fill
                ws.cell(row=r, column=col).fill = REVIEW_FILL
        else:
            if longest_length is not None and str(longest_length).lower() != "nan":
                ws.cell(row=r, column=8, value=float(longest_length))  # H: from DXF
                ws.cell(row=r, column=8).fill = DXF_LENGTH_FILL
            else:
                ws.cell(row=r, column=8, value=length_formula)
                ws.cell(row=r, column=8).fill = REVIEW_FILL
            ws.cell(row=r, column=9, value=20)  # I: Shape Code -- assumed straight, verify
            ws.cell(row=r, column=9).fill = REVIEW_FILL
            for col in (10, 11, 12, 13, 14):
                ws.cell(row=r, column=col).fill = REVIEW_FILL

        weight_formula = (
            f'=IFERROR(IF(D{r}=8,G{r}*H{r}*0.395,IF(D{r}=10,G{r}*H{r}*0.616,'
            f'IF(D{r}=12,G{r}*H{r}*0.888,IF(D{r}=16,G{r}*H{r}*1.579,'
            f'IF(D{r}=20,G{r}*H{r}*2.466,IF(D{r}=25,G{r}*H{r}*3.854,'
            f'IF(D{r}=32,G{r}*H{r}*6.313,"")))))))/1000/1000,"")'
        )
        ws.cell(row=r, column=15, value=weight_formula)

        def _is_set(v):
            return v is not None and str(v) != "" and str(v).lower() != "nan"

        ref_bits = []
        if _is_set(rec.get("Spacing (mm c/c)")):
            ref_bits.append(f"spacing {rec['Spacing (mm c/c)']} c/c")
        if _is_set(rec.get("Notes")):
            ref_bits.append(str(rec["Notes"]))
        if is_link:
            ref_bits.append(
                f"LINK -- shape code {LINK_SHAPE_CODE} confirmed (closed rectangular link/stirrup); "
                f"length not extracted (DXF measures run length, not individual link length); "
                f"A/B leg dimensions need manual entry from the drawing"
            )
        else:
            ref_bits.append("shape code 20 (straight bar) assumed for now -- verify")
            if _is_set(rec.get("Dot Occurrences Matched")):
                ref_bits.append(f"{rec['Dot Occurrences Matched']} dot(s) traced to this mark")
            if _is_set(rec.get("All DXF Lengths (mm)")):
                ref_bits.append(f"all DXF lengths found: {rec['All DXF Lengths (mm)']} mm (longest used in column H)")
            if _is_set(rec.get("Proximity-Matched Lengths (NOT used -- unverified)")):
                ref_bits.append(f"NOT USED -- proximity-matched candidate length(s) found but not trusted (see dot_bar_matches.csv): {rec['Proximity-Matched Lengths (NOT used -- unverified)']} mm")
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

    This is the function a Streamlit app (or any other Python caller)
    should call directly -- it does no argparse/sys.exit()/interactive
    prompting, just takes plain arguments and returns a plain dict.

    Returns:
        {
            "report_df": pd.DataFrame,
            "summary_df": pd.DataFrame,
            "report_path": str, "summary_path": str,
            "dot_match_path": str, "bbs_path": str,
            "num_labels": int, "num_parsed": int, "num_unparsed": int,
            "num_dots": int, "num_bar_lines": int, "num_leader_lines": int,
            "num_dots_matched": int, "num_dots_unmatched": int,
            "match_method_counts": {method: count, ...},
            "missing_marks": [int, ...],
        }
    """
    doc = ezdxf.readfile(dxf_path)
    labels = extract_labels(doc, layers=layers)

    parsed = [l for l in labels if l.parsed_ok]
    unparsed = [l for l in labels if not l.parsed_ok]

    dots = find_all_dots(doc, hatch_layer_aliases=DOT_LAYER_ALIASES,
                          circle_layer_aliases=DOT_LAYER_ALIASES, donut_layer=None)
    bar_lines = find_bar_lines(doc, layer_aliases=BAR_LAYER_ALIASES)
    leader_lines = find_leader_lines(doc, layer_aliases=LEADER_LINE_LAYER_ALIASES)

    dot_lengths, dot_match_records = match_dot_lengths(labels, dots, bar_lines, leader_lines=leader_lines)
    accepted = sum(1 for r in dot_match_records if r["Accepted"])
    unmatched = [r for r in dot_match_records if not r["Accepted"]]
    method_counts = dict(Counter(r["Match Method"] for r in dot_match_records if r["Accepted"]))

    report_df = build_extraction_report(labels)
    summary_df = build_summary(labels, dot_lengths=dot_lengths)
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
        "num_labels": len(labels),
        "num_parsed": len(parsed),
        "num_unparsed": len(unparsed),
        "num_dots": len(dots),
        "num_bar_lines": len(bar_lines),
        "num_leader_lines": len(leader_lines),
        "num_dots_matched": accepted,
        "num_dots_unmatched": len(unmatched),
        "match_method_counts": method_counts,
        "missing_marks": missing_marks,
    }


# ---------------------------------------------------------------------------
# 8. CLI (thin wrapper around process_dxf())
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
    result = process_dxf(dxf_path=dxf_path, outdir=outdir, template_path=template_path,
                          member_name=args.member, layers=layers)

    print(f"Found {result['num_labels']} text-bearing entities on the target layer(s).")
    print(f"  {result['num_parsed']} matched the bar-label pattern.")
    if result["num_unparsed"]:
        print(f"  {result['num_unparsed']} did not match -- listed in extraction_report.csv for manual review.")

    print(f"  Leader-dot geometry: {result['num_dots']} dot(s), {result['num_bar_lines']} candidate bar line(s), "
          f"{result['num_leader_lines']} candidate leader line(s), "
          f"{result['num_dots_matched']}/{result['num_dots']} dot(s) matched to a mark + bar length within tolerance.")
    if result["match_method_counts"]:
        breakdown = ", ".join(f"{v} {k}" for k, v in Counter(result["match_method_counts"]).most_common())
        print(f"    ({breakdown})")
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
    main()n()
