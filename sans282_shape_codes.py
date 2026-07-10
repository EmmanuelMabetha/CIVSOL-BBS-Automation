"""
sans282_shape_codes.py

Reference data library for SANS 282:2011 "Bending dimensions and
scheduling of steel reinforcement for concrete".

PURPOSE
-------
This module is a DATA LIBRARY, not a workflow script. A separate reader
script (e.g. your DXF/ezdxf extractor) is meant to import this file and:

  1. Take geometry it has already extracted from the drawing (number of
     straight segments, the angle(s) between them, whether a run has a
     radiused/curved portion).
  2. Call match_shape_by_geometry() to get a short-list of candidate
     shape codes.
  3. Once a shape code is confirmed (or read directly off a label/block
     attribute in the DXF), call calculate_length() with the measured
     A/B/C/D/E/r/n values to get the calculated (centre-line) length.
  4. Use get_shape() to pull the code back out for writing into the bar
     mark / bending schedule.

SOURCING & RELIABILITY -- READ THIS
------------------------------------
SANS 282:2011 is a copyrighted SABS standard. It is not reproduced here.
Every formula below was transcribed directly from a user-supplied scan of
the standard's own Table 2 ("Calculated length" column) and cross-checked
against two independent sources for this specific project:

  1. Annex A (Table A.1 -- Shape codes), which gives the dimensioned
     sketches confirming which letter (A/B/C/D/E/r/n/h) sits on which leg.
  2. The real vector line/arc geometry traced directly out of the actual
     project drawing (301-00825-01-171 Rev 1)'s own shape-code legend --
     i.e. independently confirmed against a real fabrication drawing, not
     just the standard's own diagram.

All three sources agree on the formulas below. An earlier pass through this
library had concluded shape codes 34 and 35 don't exist in this series --
that conclusion was wrong. You've since located both in your own copy of
SANS 282:2011 and had a professional engineer verify the formulas, so 34
and 35 are now included below as confirmed entries like everything else.

Every entry below has a "confidence" field:

    "confirmed"   -> formula transcribed directly from Table 2 and cross-
                      checked as above. Still spot-check the first few
                      real bars of each code against your own copy before
                      trusting it at scale -- this library has not been
                      used in anger yet.
    "unverified"  -> not present in the supplied Table 2 excerpt. Left as
                      a placeholder. Do not use for real bar lengths.

Note 4 of the standard applies throughout: the formulas for cranked
shapes (41, 42, 43, 45, 48, 49, 62 and similar) assume the bend angle
with the horizontal is 45 degrees or less. Above that, the standard says
to calculate the true length more accurately rather than use the
approximate formula -- this library does not attempt that calculation,
so treat any crank angle > 45 degrees as needing manual attention (or an
upgrade to this library) rather than trusting the formula blindly.

Units: all bending dimensions (A, B, C, D, E, r, R, n, h) in mm.
d = nominal bar diameter in mm.
"""

import math
from dataclasses import dataclass
from typing import Callable, Optional


# ---------------------------------------------------------------------------
# 1. Standard hook / bend / radius allowances (Figures 2 and 3 of the
#    standard). These set default values for h (hook allowance),
#    n (bend allowance) and r (standard radius) by bar diameter and steel
#    type, used whenever a schedule doesn't override them explicitly.
# ---------------------------------------------------------------------------

STANDARD_BEND_ALLOWANCES = {
    # R = mild steel (250 MPa), hot-rolled, to SANS 920 -- Figure 2
    "R": {
        6:  {"h": 100, "n": 100, "r": 12},
        8:  {"h": 100, "n": 100, "r": 16},
        10: {"h": 120, "n": 100, "r": 20},
        12: {"h": 120, "n": 100, "r": 24},
        16: {"h": 160, "n": 100, "r": 32},
        20: {"h": 200, "n": 120, "r": 40},
        25: {"h": 260, "n": 160, "r": 50},
        32: {"h": 320, "n": 200, "r": 64},
        40: {"h": 400, "n": 240, "r": 80},
    },
    # Y = high-yield deformed / cold-worked steel (450 MPa), to SANS 920 -- Figure 3
    "Y": {
        6:  {"h": 100, "n": 100, "r": 18},
        8:  {"h": 100, "n": 100, "r": 24},
        10: {"h": 120, "n": 100, "r": 30},
        12: {"h": 160, "n": 100, "r": 36},
        16: {"h": 200, "n": 120, "r": 48},
        20: {"h": 240, "n": 140, "r": 60},
        25: {"h": 300, "n": 180, "r": 75},
        32: {"h": 400, "n": 220, "r": 96},
        40: {"h": 480, "n": 260, "r": 120},
    },
}


def get_standard_allowances(steel_type: str, diameter: int) -> dict:
    """
    Look up standard h (hook), n (bend) and r (radius) allowances.
    steel_type: "R" (mild steel) or "Y" (high yield / cold-worked).
    diameter: nominal bar size in mm (must match a table entry).
    """
    steel_type = steel_type.upper()
    if steel_type not in STANDARD_BEND_ALLOWANCES:
        raise ValueError(f"Unknown steel type '{steel_type}', expected 'R' or 'Y'")
    table = STANDARD_BEND_ALLOWANCES[steel_type]
    if diameter not in table:
        raise ValueError(
            f"No standard allowance for d={diameter}mm ({steel_type}). "
            f"Available sizes: {sorted(table.keys())}"
        )
    return table[diameter]


# ---------------------------------------------------------------------------
# 2. Shape code definitions
# ---------------------------------------------------------------------------

@dataclass
class ShapeCode:
    code: str                       # e.g. "37"  (append "S" suffix separately if r = 7.5d)
    name: str                       # short human description
    params: list                    # ordered list of dimension keys this formula needs
    segments: Optional[int]         # approx. number of straight/curved runs (best-effort)
    bend_angles: list               # approx nominal angles between segments, degrees
    formula_str: str                # human-readable formula, centre-line calculated length
    formula: Callable               # callable(**dims) -> calculated length (mm)
    condition: str = ""             # any conditions on when this formula applies
    confidence: str = "confirmed"   # "confirmed" | "unverified"
    notes: str = ""


SHAPE_CODES: dict[str, ShapeCode] = {}


def _add(code, name, params, segments, bend_angles, formula_str, formula,
         condition="", confidence="confirmed", notes=""):
    SHAPE_CODES[code] = ShapeCode(
        code=code, name=name, params=params, segments=segments,
        bend_angles=bend_angles, formula_str=formula_str, formula=formula,
        condition=condition, confidence=confidence, notes=notes,
    )


# --- All entries below are transcribed from SANS 282:2011 Table 2 --------

_add(
    "20", "Straight bar",
    params=["A"], segments=1, bend_angles=[],
    formula_str="A",
    formula=lambda A, **k: A,
    notes="No bends. Simplest case -- one LINE entity, no corners.",
)

_add(
    "32", "Straight bar with a single hook/curl at one end",
    params=["A", "h"], segments=1, bend_angles=["hook"],
    formula_str="A + h  (PROVISIONAL -- not yet confirmed against Table 2)",
    formula=lambda A, h, **k: A + h,
    confidence="unverified",
    notes="Geometry (one curved hook at one end, dimension A only) is confirmed "
          "directly from the SANS 282:2004 Ed 5.1 chart. The length formula shown "
          "here is a provisional guess by analogy to 34/35's pattern (n added per "
          "bend allowance -- here h, the standard hook allowance, since this is a "
          "curl/hook rather than a sharp bend), NOT transcribed from Table 2. "
          "Do not trust the length for real bars until Table 2 confirms it -- "
          "confidence is 'unverified' for exactly this reason, even though the "
          "shape itself is real and confirmed.",
)

_add(
    "33", "Straight bar with a hook/curl at both ends",
    params=["A", "h"], segments=1, bend_angles=["hook", "hook"],
    formula_str="A + 2h  (PROVISIONAL -- not yet confirmed against Table 2)",
    formula=lambda A, h, **k: A + 2 * h,
    confidence="unverified",
    notes="Geometry (curved hooks at both ends, dimension A only) is confirmed "
          "directly from the SANS 282:2004 Ed 5.1 chart. The length formula is a "
          "provisional guess by analogy to 32/34/35, NOT transcribed from Table 2 "
          "-- do not trust for real bars until confirmed.",
)

_add(
    "34", "Single bend with allowance n",
    params=["A", "n"], segments=2, bend_angles=["variable"],
    formula_str="A + n",
    formula=lambda A, n, **k: A + n,
    notes="Straight leg A with a single bent allowance tail n. Formula located "
          "directly in the user's own copy of SANS 282:2011 and confirmed by a "
          "professional engineer.",
)

_add(
    "35", "Double bend with two allowances n",
    params=["A", "n"], segments=3, bend_angles=["variable", "variable"],
    formula_str="A + 2n",
    formula=lambda A, n, **k: A + 2 * n,
    notes="Straight main leg A with two bent allowance tails n at each end. "
          "Formula located directly in the user's own copy of SANS 282:2011 "
          "and confirmed by a professional engineer.",
)

_add(
    "36", "S-shape: straight legs with a hooked/curved return at each end",
    params=["A", "B", "C", "D", "E", "d"], segments=5,
    bend_angles=["large_radius", "large_radius"],
    formula_str="(A + C + E) + 0.57(B + D) - 3.14d",
    formula=lambda A, B, C, D, E, d, **k: (A + C + E) + 0.57 * (B + D) - 3.14 * d,
    condition="If B or D > 400 + 2d, see note 1 of clause 4.2.5.",
    notes="Matches the drawing's own legend exactly: a straight top leg (A), "
          "large-radius curve (B), bottom leg (C), a second large-radius "
          "curve (D), and a short return leg (E).",
)

_add(
    "37", "Single bend (90 deg, standard radius)",
    params=["A", "B", "r", "d"], segments=2, bend_angles=[90],
    formula_str="A + B - r/2 - d",
    formula=lambda A, B, r, d, **k: A + B - r / 2 - d,
    condition="If radius r is non-standard, use shape code 51 instead.",
    notes="Classic L-shaped bar -- 2 straight segments meeting at one right-angle bend.",
)

_add(
    "38", "Two bends (U / Z shape, standard radius)",
    params=["A", "B", "C", "r", "d"], segments=3, bend_angles=[90, 90],
    formula_str="A + B + C - r - 2d",
    formula=lambda A, B, C, r, d, **k: A + B + C - r - 2 * d,
    notes="3 straight segments, 2 right-angle bends (e.g. stirrup leg / crank).",
)

_add(
    "39", "Bend with large (non-standard) radius portion",
    params=["A", "B", "C", "d"], segments=3, bend_angles=["large_radius"],
    formula_str="A + 0.57B + C - 1.57d",
    formula=lambda A, B, C, d, **k: A + 0.57 * B + C - 1.57 * d,
    condition="If B >= 400 + 2d, see note 1 of clause 4.2.5.",
    notes="B is a curved (radiused) run rather than a sharp corner.",
)

_add(
    "41", "Cranked bar, single crank (angle <= 45 deg)",
    params=["A", "B", "C"], segments=3, bend_angles=["<=45", "<=45"],
    formula_str="A + B + C",
    formula=lambda A, B, C, **k: A + B + C,
    condition="Angle with horizontal <= 45 deg; otherwise see note 4.",
    notes="Fixed: was previously listing only 1 bend angle for a 3-segment shape "
          "(needs 2, one at each junction). Best-judgment fix per your direction "
          "-- both bends taken as the same shallow crank angle (enters and exits "
          "the diagonal symmetrically), matching the standard cranked-bar "
          "convention. Not independently re-confirmed against Table 2.",
)

_add(
    "42", "Cranked bar with an additional bend allowance n",
    params=["A", "B", "C", "n"], segments=3, bend_angles=["<=45", 90],
    formula_str="A + B + C + n",
    formula=lambda A, B, C, n, **k: A + B + C + n,
    condition="Angle with horizontal <= 45 deg; otherwise see note 4.",
)

_add(
    "43", "Double crank / wide-V shape (symmetric diagonals)",
    params=["A", "B", "C", "D", "E"], segments=5,
    bend_angles=["<=45", "<=45", "<=45", "<=45"],
    formula_str="A + 2B + C + E",
    formula=lambda A, B, C, E, **k: A + 2 * B + C + E,
    condition="Angle with horizontal <= 45 deg (otherwise see note 4); D >= 2d.",
    notes="Both diagonal legs share dimension B (hence 2B) -- confirmed against "
          "both Annex A and the project drawing's own traced vectors.",
)

_add(
    "45", "Cranked bar with a standard-radius bend",
    params=["A", "B", "C", "r", "d"], segments=3, bend_angles=["<=45", 90],
    formula_str="A + B + C - r/2 - d",
    formula=lambda A, B, C, r, d, **k: A + B + C - r / 2 - d,
    condition="Angle with horizontal <= 45 deg; otherwise see note 4.",
)

_add(
    "48", "Cranked bar variant (offset then straight run)",
    params=["A", "B", "C"], segments=3, bend_angles=["<=45", 90],
    formula_str="A + B + C",
    formula=lambda A, B, C, **k: A + B + C,
    condition="Angle with horizontal <= 45 deg; otherwise see note 4.",
    notes="Fixed: was previously listing only 1 bend angle for a 3-segment shape. "
          "Best-judgment fix per your direction -- the SANS chart shows a distinct "
          "vertical uptick at the very end after the diagonal (not just a height "
          "dimension line like 41 has), so the second bend is taken as ~90 deg "
          "rather than a second shallow crank. Not independently re-confirmed "
          "against Table 2.",
)

_add(
    "49", "Cranked bar variant (dip/kink between two straight runs)",
    params=["A", "B", "C"], segments=3, bend_angles=["<=45", "<=45"],
    formula_str="A + B + C",
    formula=lambda A, B, C, **k: A + B + C,
    condition="Angle with horizontal <= 45 deg; otherwise see note 4.",
    notes="D and E on the sketch are run-off/anchorage references only, not "
          "part of the length formula (see note 2 of the standard).",
)

_add(
    "51", "Large-radius bend (non-standard internal radius R)",
    params=["A", "B", "R", "d"], segments=2, bend_angles=["large_radius"],
    formula_str="A + B - 0.43R - 1.21d",
    formula=lambda A, B, R, d, **k: A + B - 0.43 * R - 1.21 * d,
    condition="If R is standard, use shape code 37 instead. If R >= 200 mm, see "
              "note 1 of clause 4.2.5.",
)

_add(
    "52", "3-bend channel (open or with return lips)",
    params=["A", "B", "C", "D", "r", "d"], segments=4, bend_angles=[90, 90, 90],
    formula_str="A + B + C + D - (3/2)r - 3d",
    formula=lambda A, B, C, D, r, d, **k: A + B + C + D - 1.5 * r - 3 * d,
)

_add(
    "53", "4-bend dip/step channel",
    params=["A", "B", "C", "D", "E", "r", "d"], segments=5, bend_angles=[90, 90, 90, 90],
    formula_str="A + B + C + D + E - 2r - 4d",
    formula=lambda A, B, C, D, E, r, d, **k: A + B + C + D + E - 2 * r - 4 * d,
)

_add(
    "54", "L-shape with an extra dogleg (2 bends)",
    params=["A", "B", "C", "r", "d"], segments=3, bend_angles=[90, 90],
    formula_str="A + B + C - r - 2d",
    formula=lambda A, B, C, r, d, **k: A + B + C - r - 2 * d,
)

_add(
    "55", "4-bend near-closed channel",
    params=["A", "B", "C", "D", "E", "r", "d"], segments=5, bend_angles=[90, 90, 90, 90],
    formula_str="A + B + C + D + E - 2r - 4d",
    formula=lambda A, B, C, D, E, r, d, **k: A + B + C + D + E - 2 * r - 4 * d,
)

_add(
    "60", "Closed link / stirrup (rectangular, with lapped tail)",
    params=["A", "B", "n", "r", "d"], segments=5, bend_angles=[90, 90, 90, "lap"],
    formula_str="2A + 2B + 2n - (3/2)r - 3d",
    formula=lambda A, B, n, r, d, **k: 2 * A + 2 * B + 2 * n - 1.5 * r - 3 * d,
    notes="Confirms what the project drawing's own vector geometry already "
          "showed: a rectangle (legs A, B) with a lapped tail at one corner. "
          "Correction from an earlier turn: this genuinely is the closed link "
          "shape, and now has a real formula -- but dxf_bbs_extractor.py still "
          "can't auto-fill it, since it doesn't trace closed/looped chains to "
          "read A and B off an actual link in the drawing yet.",
)

_add(
    "62", "Single offset/crank (angle <= 45 deg)",
    params=["A", "C"], segments=2, bend_angles=["<=45"],
    formula_str="A + C",
    formula=lambda A, C, **k: A + C,
    condition="Angle with horizontal <= 45 deg; otherwise see note 4.",
    notes="Correction from an earlier turn: params are A and C (not A and B as "
          "guessed before) -- confirmed directly from Table 2.",
)

_add(
    "65", "Shallow curve at (or near) the standard minimum radius",
    params=["A"], segments=1, bend_angles=["radius"],
    formula_str="A",
    formula=lambda A, **k: A,
    notes="Correction from an earlier turn: I previously guessed this was a "
          "single hook (A + h) based on a misread of the small drawing legend. "
          "Table 2 confirms it's actually a shallow curve where the specified "
          "chord length A IS the calculated length -- no deduction at all.",
)

_add(
    "72", "Open stirrup/link with hooked ends",
    params=["A", "B", "h", "r", "d"], segments=5, bend_angles=[90, "hook", 90, "hook", 90],
    formula_str="2A + B + 2h - r - 2d",
    formula=lambda A, B, h, r, d, **k: 2 * A + B + 2 * h - r - 2 * d,
    condition="Ensure B > 14d for smooth mild steel and B > 18d for deformed "
              "steel (note 3).",
    notes="Matches the vector-traced geometry exactly: bottom leg B, two "
          "vertical legs A, hooked top ends.",
)

_add(
    "73", "Asymmetric offset (Z-step) bar with a bend allowance n",
    params=["A", "B", "C", "n", "r", "d"], segments=5, bend_angles=[90, 90, 90, 90],
    formula_str="2A + B + C + n - (3/2)r - 3d",
    formula=lambda A, B, C, n, r, d, **k: 2 * A + B + C + n - 1.5 * r - 3 * d,
)

_add(
    "74", "Rectangular loop shape with hooked ends",
    params=["A", "B", "n", "r", "d"], segments=7, bend_angles=[90, 90, 90, 90],
    formula_str="2A + 3B + 2n - 2r - 4d",
    formula=lambda A, B, n, r, d, **k: 2 * A + 3 * B + 2 * n - 2 * r - 4 * d,
    notes="Repeated legs (2A, 3B) explain why only A/B were labelled on the "
          "project drawing despite the shape having many segments.",
)

_add(
    "75", "Stepped zigzag with a bend allowance n",
    params=["A", "B", "C", "D", "E", "n", "r", "d"], segments=5, bend_angles=[90, 90, 90, 90],
    formula_str="A + B + C + 2D + E + n - (5/2)r - 5d",
    formula=lambda A, B, C, D, E, n, r, d, **k: A + B + C + 2 * D + E + n - 2.5 * r - 5 * d,
)

_add(
    "81", "Closed oval/large-radius loop",
    params=["A", "r", "d", "h"], segments=3, bend_angles=["large_radius", "large_radius"],
    formula_str="2A + r + d + 2h",
    formula=lambda A, r, d, h, **k: 2 * A + r + d + 2 * h,
    condition="For larger radii, see clause 4.2.5 (note to table 5) and consider "
              "shape code 99 instead.",
)

_add(
    "83", "Diagonal stepped ('lightning bolt') crank",
    params=["A", "B", "C", "D", "r", "d"], segments=4, bend_angles=[90, 90, 90, 90],
    formula_str="A + 2B + C + D - 2r - 4d",
    formula=lambda A, B, C, D, r, d, **k: A + 2 * B + C + D - 2 * r - 4 * d,
    notes="B is used for both symmetric idler legs (hence 2B), same pattern as "
          "shape 43. Note the drawing's own legend labels B as an overall "
          "dimension in one place -- confirm which B Table 2 means before "
          "trusting this on a shape with unusual proportions.",
)

_add(
    "85", "Bar with a right-angle bend plus a large-radius curve",
    params=["A", "B", "C", "D", "r", "d"], segments=4, bend_angles=[90, "large_radius"],
    formula_str="A + B + 0.57C + D - r/2 - 2.57d",
    formula=lambda A, B, C, D, r, d, **k: A + B + 0.57 * C + D - r / 2 - 2.57 * d,
    condition="If C > 400 + 2d, see clause 4.2.5.",
)

_add(
    "86", "Helical bar (at least 2 full turns)",
    params=["A", "B", "C", "d"], segments=None, bend_angles=["helix"],
    formula_str="(C/B) * pi * sqrt((A-d)^2 + B^2) + 2*pi*(A-d)",
    formula=lambda A, B, C, d, **k: (
        (C / B) * math.pi * math.sqrt((A - d) ** 2 + B ** 2) + 2 * math.pi * (A - d)
    ),
    condition="B (pitch) must not exceed A/5 or 150 mm, whichever is least. "
              "Requires at least 2 full turns. An additional end turn is "
              "treated as shape code 99.",
    notes="Genuinely not an A+B+C leg-sum shape -- a real helix length formula "
          "based on diameter (A), pitch (B) and overall length (C).",
)

_add(
    "99", "Non-standard shape",
    params=[], segments=None, bend_angles=[],
    formula_str="n/a -- fully dimensioned sketch required",
    formula=lambda **k: None,
    notes="Catch-all for any shape that doesn't match a defined code. The "
          "standard requires a full dimensioned sketch in this case, not a "
          "formula. Use this as your default/fallback match when your "
          "geometry-matcher can't confidently map extracted segments+angles "
          "to any other code.",
)


# ---------------------------------------------------------------------------
# 3. Lookup / matching helpers
# ---------------------------------------------------------------------------

def get_shape(code: str) -> ShapeCode:
    """Return the ShapeCode entry for a given code, stripping an 'S' suffix
    (S = 7,5d bend radius rather than standard radius) if present."""
    code = code.strip().upper()
    base_code = code[:-1] if code.endswith("S") and code[:-1].isdigit() else code
    if base_code not in SHAPE_CODES:
        raise KeyError(f"Shape code '{code}' not found in library.")
    return SHAPE_CODES[base_code]


def calculate_length(code: str, **dims) -> float:
    """
    Calculate centre-line length for a given shape code using measured
    dimensions (A, B, C, D, E, r, R, n, h, d as applicable).
    Raises if the shape code is unverified or missing required params.
    """
    shape = get_shape(code)
    if shape.confidence == "unverified":
        raise NotImplementedError(
            f"Shape code {code} ('{shape.name}') is unverified in this library -- "
            f"confirm its formula against SANS 282:2011 before calculating lengths."
        )
    missing = [p for p in shape.params if p not in dims]
    if missing:
        raise ValueError(f"Shape code {code} requires {shape.params}, missing: {missing}")
    return shape.formula(**dims)


def match_shape_by_geometry(num_segments: int, angles: list) -> list:
    """
    Given geometry extracted from a DXF entity group (number of straight
    segments, list of approximate angles in degrees between consecutive
    segments), return a list of candidate shape codes worth checking.

    This is a coarse first-pass filter, not a definitive match -- always
    fall back to shape code 99 if nothing fits confidently, and treat
    "confirmed" candidates as higher priority than "unverified" ones.
    """
    candidates = []
    for code, shape in SHAPE_CODES.items():
        if shape.segments is None:
            continue
        if shape.segments != num_segments:
            continue
        # Straight length-for-length angle-count match only. The previous
        # version of this check had "or not shape.bend_angles", meant to be
        # lenient for shapes with an unspecified bend structure -- but that
        # silently let code 20 (bend_angles=[], meaning "zero bends") match
        # ANY geometry with 1 segment, including a hooked/curved bar that
        # should have candidated as 32/33/65 instead. An empty bend_angles
        # list means "exactly zero bends expected", not "don't care" -- so
        # it must only match when the measured angles list is ALSO empty.
        if len(shape.bend_angles) == len(angles):
            candidates.append(code)
    candidates.sort(key=lambda c: 0 if SHAPE_CODES[c].confidence == "confirmed" else 1)
    return candidates or ["99"]


# ---------------------------------------------------------------------------
# Quick self-test / usage example
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # Example: a 90-degree bend, Y12 bar, standard radius
    alloc = get_standard_allowances("Y", 12)
    length = calculate_length("37", A=800, B=400, r=alloc["r"], d=12)
    print(f"Shape code 37, Y12: standard r={alloc['r']}mm -> length = {length} mm")

    # Closed link example, Y10 bar, using n = standard bend allowance
    alloc10 = get_standard_allowances("Y", 10)
    link_len = calculate_length("60", A=200, B=100, n=alloc10["n"], r=alloc10["r"], d=10)
    print(f"Shape code 60 (closed link), Y10: length = {link_len:.1f} mm")

    # Helical bar example: A=250mm coil diameter, B=80mm pitch, C=1000mm length, d=12mm
    helix_len = calculate_length("86", A=250, B=80, C=1000, d=12)
    print(f"Shape code 86 (helix): length = {helix_len:.1f} mm")

    print("\nCandidates for 2 segments, 1 bend:", match_shape_by_geometry(2, [90]))

    confirmed = [c for c, s in SHAPE_CODES.items() if s.confidence == "confirmed"]
    unverified = [c for c, s in SHAPE_CODES.items() if s.confidence == "unverified"]
    print(f"\nConfirmed shape codes ({len(confirmed)}):", sorted(confirmed, key=lambda c: (len(c), c)))
    print(f"Unverified shape codes ({len(unverified)}):", sorted(unverified, key=lambda c: (len(c), c)))
