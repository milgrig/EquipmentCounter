"""
pdf_cables_by_method.py — Detect cable routes on PDF lighting plans and
                         break down total length by laying method.

Inputs : a PDF page with red/blue cable trasse drawings.
Outputs: a CableMethodReport with:
  * total length per (cable color × laying method × cross-section)
  * the scale used (mm per PDF point)
  * per-route polyline data (for visualization / debugging)

This module is intentionally written from scratch (it does not modify
existing modules) and complements:
  * pdf_count_geometry.py  — measures simple total length, no laying-method
                             attribution, uses a strict MIN_SEGMENT length
                             that loses ~80% of micro-segments.
  * pdf_count_cables.py    — extracts text annotations (L=…м) but does not
                             measure geometry.

Laying-method detection uses geometric markers because PDF dash pattern
is not present on these drawings (diagnostic confirmed all cable lines
are stroked solid):
  * "в гофре" / "в трубе"  — small circles drawn along the cable line
                             (legend item: "Проводка прокладываемая в трубе").
  * "в лотке"               — rectangular tray glyph straddling the cable.
  * "по конструкциям"       — bare line, no marker.
  * "в земле"               — marker outside the building footprint
                             (currently classified as "по конструкциям"
                             unless detected via legend-driven keyword).

Optional cross-check: text labels of the form "L=NNм" near the route
(parsed by pdf_count_cables.py) can be aggregated and compared.

Usage:
    python pdf_cables_by_method.py <path.pdf> [--page N] [--scale-mm-per-pt 26.47]
                                              [--vor <path.xlsx>]
                                              [--save-overlay overlay.png]
"""

from __future__ import annotations

import io
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Iterable, Optional

import pdfplumber

from pdf_legend_parser import parse_legend, LegendItem, LegendResult

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

PT_PER_INCH = 72.0
MM_PER_INCH = 25.4
MM_PER_PT = MM_PER_INCH / PT_PER_INCH  # ≈ 0.3528 — paper mm per PDF point

# Cable colors (exact hits seen on KPP-007: (1,0,0) and (0,0,1))
COLOR_RED = (1.0, 0.0, 0.0)
COLOR_BLUE = (0.0, 0.0, 1.0)
COLOR_TOL = 0.10  # generous to catch slightly off-channel reds/blues

# Marker detection
CIRCLE_MAX_RADIUS_PT = 4.0     # small circles glued to cable runs
CIRCLE_MIN_RADIUS_PT = 0.5
MARKER_TO_LINE_RADIUS = 5.0    # how close a marker must lie to a cable line
TRAY_RECT_MIN_W = 8.0
TRAY_RECT_MAX_W = 35.0
TRAY_RECT_MAX_H = 10.0

# Tray T-tick markers: short perpendicular ticks crossing the cable line.
# Per legend on GPK3 plans: tray glyph = a horizontal segment with two
# 3.5 pt vertical ticks on opposite sides of the cable midline.
# We treat any short black/colored segment of length [TICK_LEN_MIN, TICK_LEN_MAX]
# whose direction differs from a nearby cable line by ~90° as a tray tick.
TICK_LEN_MIN_PT = 1.5
TICK_LEN_MAX_PT = 6.0
TICK_TO_LINE_PT = 4.0          # tick midpoint must be within this of a cable
TICK_PERP_TOL_DEG = 25.0       # 90° ± this is "perpendicular"
TICK_HIT_RADIUS = 6.0          # tick→polyline-endpoint distance for "hit"
TRAY_TICK_MIN_HITS = 2         # need ≥N ticks on a polyline for "lotok"
TRAY_TICK_MIN_DENSITY = 0.5    # ticks per 100 pt of cable
TRAY_TICK_HARD_HITS = 5        # OR ≥N absolute ticks regardless of density
TRAY_TICK_SHORT_LEN_PT = 100.0 # < this length → tighten the hit threshold
TRAY_TICK_SHORT_HITS = 4       # short polylines need ≥N hits

# Polyline assembly
ENDPOINT_TOL = 0.6  # pt: endpoints within this distance are the same node

# Parallel-line dedup (centerline merging)
# Cable trays are physically drawn as TWO parallel edges 3-6 pt apart on the
# PDF (≈ 100-200 mm at 1:100 scale, matching real lotok widths). Without
# dedup, the detector double-counts every metre of tray run. We collapse
# such pairs into a single centerline.
PARALLEL_MIN_LEN_PT = 30.0      # only consider long segments (>≈1 m at 1:100)
PARALLEL_ANGLE_TOL_DEG = 3.0    # max angular difference for "parallel"
PARALLEL_PERP_MIN_PT = 0.3      # below this they're co-linear, not a pair
PARALLEL_PERP_MAX_PT = 6.0      # above this they're independent runs
PARALLEL_OVERLAP_MIN = 0.5      # min fraction of shorter projected onto longer

# Exact duplicate dedup: PDFs from CAD exporters routinely contain
# multiple stacked copies of the same line (some up to 7× on GPK3 trays).
# We collapse any segments whose endpoints match within DUP_ENDPOINT_TOL pt.
DUP_ENDPOINT_TOL = 0.4          # pt: endpoints within this distance are equal
DUP_MIN_LEN_PT = 1.0            # only dedup segments at least this long

# Default exclusion margins
GRID_AXIS_MARGIN = 60
TITLE_BLOCK_MIN_LINES = 6

# Scale detection
KNOWN_DIMS_MM = [
    1500, 1800, 2400, 3000, 3600, 4200, 4500, 4800, 5400, 6000,
    7200, 9000, 12000,
]
DIM_TEXT_RE = re.compile(r"^(\d{3,5})$")


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class Polyline:
    """A connected chain of cable segments of one color."""
    color: str = ""                    # 'red' / 'blue'
    points: list[tuple[float, float]] = field(default_factory=list)
    length_pt: float = 0.0
    laying_method: str = ""            # 'gofra/truba' | 'lotok' | 'po_konstrukciyam'
    nearby_label_m: Optional[float] = None  # length in meters from L=… label

    @property
    def length_m(self) -> float:
        return self.length_pt * 0.0  # placeholder; report fills via scale


@dataclass
class CableMethodReport:
    pdf_path: str = ""
    page_index: int = 0
    page_size: tuple[float, float] = (0.0, 0.0)
    scale_mm_per_pt: float = 0.0
    scale_source: str = ""
    polylines: list[Polyline] = field(default_factory=list)
    # totals[(color, method)] = total length in metres
    totals_m: dict[tuple[str, str], float] = field(default_factory=dict)
    # Diagnostics
    raw_red_lines: int = 0
    raw_blue_lines: int = 0
    circle_markers: int = 0
    tray_rects: int = 0
    label_total_m: float = 0.0
    # Parallel-line dedup diagnostics (per color)
    dedup_red: dict = field(default_factory=dict)
    dedup_blue: dict = field(default_factory=dict)
    # Sheet-level default laying method, parsed from general notes
    sheet_default_method: str = "po_konstrukciyam"
    sheet_default_info: dict = field(default_factory=dict)
    # Tray-tick markers found
    tray_ticks: int = 0
    # Sheet-level cable marks (most common to least)
    sheet_cable_marks: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Color / geometry helpers
# ---------------------------------------------------------------------------

def _is_color(c, target: tuple, tol: float = COLOR_TOL) -> bool:
    if not isinstance(c, tuple) or len(c) != len(target):
        return False
    return all(abs(a - b) < tol for a, b in zip(c, target))


def _classify_color(c) -> str:
    if _is_color(c, COLOR_RED):
        return "red"
    if _is_color(c, COLOR_BLUE):
        return "blue"
    return ""


def _seg_len(ln: dict) -> float:
    return math.hypot(ln["x1"] - ln["x0"], ln["bottom"] - ln["top"])


def _midpoint(ln: dict) -> tuple[float, float]:
    return ((ln["x0"] + ln["x1"]) / 2, (ln["top"] + ln["bottom"]) / 2)


def _in_bbox(p: tuple[float, float], bbox: tuple[float, float, float, float]) -> bool:
    x, y = p
    return bbox[0] <= x <= bbox[2] and bbox[1] <= y <= bbox[3]


# ---------------------------------------------------------------------------
# Exclusion zones
# ---------------------------------------------------------------------------

def _detect_title_block(page, lines):
    pw, ph = page.width, page.height
    h_lines = [
        ln for ln in lines
        if abs(ln["top"] - ln["bottom"]) < 2
        and abs(ln["x1"] - ln["x0"]) > 50
        and ln["x0"] > pw * 0.4
        and ln["top"] > ph * 0.6
    ]
    if len(h_lines) < TITLE_BLOCK_MIN_LINES:
        return None
    tb_x0 = min(ln["x0"] for ln in h_lines)
    tb_x1 = max(ln["x1"] for ln in h_lines)
    tb_y0 = min(ln["top"] for ln in h_lines)
    tb_y1 = max(ln["top"] for ln in h_lines)
    return (tb_x0, tb_y0, tb_x1, tb_y1)


def _build_zones(page, lines, legend_bbox):
    pw, ph = page.width, page.height
    m = GRID_AXIS_MARGIN
    zones = []
    if legend_bbox and legend_bbox != (0, 0, 0, 0):
        zones.append((
            "legend",
            (legend_bbox[0] - 10, legend_bbox[1] - 30,
             legend_bbox[2] + 10, legend_bbox[3] + 10),
        ))
    tb = _detect_title_block(page, lines)
    if tb:
        zones.append(("title_block", tb))
    zones.extend([
        ("grid_left", (0, 0, m, ph)),
        ("grid_right", (pw - m, 0, pw, ph)),
        ("grid_top", (0, 0, pw, m)),
        ("grid_bottom", (0, ph - m, pw, ph)),
    ])
    return zones


def _excluded(p, zones):
    for _, z in zones:
        if _in_bbox(p, z):
            return True
    return False


# ---------------------------------------------------------------------------
# Scale detection (text dim + axis fallback)
# ---------------------------------------------------------------------------

STANDARD_SCALES = [50, 75, 100, 150, 200, 250, 400, 500]


def _round_scale_score(mm_per_pt: float) -> float:
    """Lower is better: distance to the nearest STANDARD drawing scale.

    Russian electrical drawings almost always use 1:50, 1:100, 1:200 or
    1:500. We score by the absolute difference to the nearest entry in
    STANDARD_SCALES; perfect match → 0, off by 5 → penalty 5.
    """
    real_N = mm_per_pt / MM_PER_PT
    return min(abs(real_N - s) for s in STANDARD_SCALES)


def _detect_scale_from_dims(words, lines) -> Optional[tuple[float, str, dict]]:
    """Try to derive scale from KNOWN_DIMS_MM text values that repeat at
    same Y row (dimension chain) or X column.

    The spacing between adjacent dimension texts on a chain ≈ the
    dimension value in PDF points only when each label sits in the
    middle of its own segment. We therefore *also* require that the
    derived mm_per_pt produces a "round" 1:N drawing scale (multiple
    of 25), otherwise the chain is rejected and we fall back to axis
    detection.
    """
    rows: dict[tuple[int, float], list[dict]] = defaultdict(list)
    cols: dict[tuple[int, float], list[dict]] = defaultdict(list)
    for w in words:
        m = DIM_TEXT_RE.match(w["text"])
        if not m:
            continue
        v = int(m.group(1))
        if v not in KNOWN_DIMS_MM:
            continue
        y_key = round(w["top"] / 8) * 8
        x_key = round(w["x0"] / 8) * 8
        rows[(v, y_key)].append(w)
        cols[(v, x_key)].append(w)

    # Pick best repeat in either rows or cols
    best = None  # (round_score, mm_per_pt, label, info)
    for group_dict, axis in ((rows, "row"), (cols, "col")):
        for (v, k), items in group_dict.items():
            if len(items) < 2:
                continue
            if axis == "row":
                items.sort(key=lambda w: w["x0"])
                spacings = [items[i + 1]["x0"] - items[i]["x0"]
                            for i in range(len(items) - 1)]
            else:
                items.sort(key=lambda w: w["top"])
                spacings = [items[i + 1]["top"] - items[i]["top"]
                            for i in range(len(items) - 1)]
            spacings = [s for s in spacings if s > 10]
            if not spacings:
                continue
            spacings.sort()
            med = spacings[len(spacings) // 2]
            mm_per_pt = v / med
            score = _round_scale_score(mm_per_pt)
            cand = (score, mm_per_pt,
                    f"dim_text_{v}mm_{axis}",
                    dict(value_mm=v, span_pt=round(med, 2), axis=axis,
                         repeats=len(items),
                         round_score=round(score, 2)))
            if best is None or score < best[0]:
                best = cand

    if best is None:
        return None
    # Reject if not close to a round scale (within 5 of nearest 1:N step)
    if best[0] > 5.0:
        return None
    return (best[1], best[2], best[3])


def _detect_scale_from_axes(words,
                             page_size: Optional[tuple[float, float]] = None,
                             ) -> Optional[tuple[float, str, dict]]:
    """Fallback: use axis letter pairs (А, Б, В, Г, …) seen on the same
    column or row. Spacing between adjacent axes in industrial buildings
    is typically 6000 mm. We use whichever value is in KNOWN_DIMS_MM and
    matches the observed pixel-spacing closest.

    Title-block heuristic: letters falling in the bottom-right 25% of
    the page are dropped — they are usually section codes inside the
    штамп (e.g. "ИКЛГ", "НОПР"), not building axes.

    Returns (mm_per_pt, source, info) or None.
    """
    axis_letters = "АБВГДЕЖИКЛМНОПРСТУФХЦЧШЩЭЮЯ"
    pos_by_letter: dict[str, list[tuple[float, float]]] = defaultdict(list)
    pw, ph = (page_size or (0, 0))
    tb_x_thr = pw * 0.55 if pw else None  # right-half cutoff
    tb_y_thr = ph * 0.80 if ph else None  # bottom 20% cutoff
    for w in words:
        t = w["text"]
        if len(t) == 1 and t in axis_letters:
            cx = (w["x0"] + w.get("x1", w["x0"])) / 2
            cy = (w["top"] + w.get("bottom", w["top"])) / 2
            # Drop title-block letters (bottom-right corner)
            if (tb_x_thr is not None and tb_y_thr is not None
                    and cx > tb_x_thr and cy > tb_y_thr):
                continue
            pos_by_letter[t].append((cx, cy))

    if len(pos_by_letter) < 2:
        return None

    # Find adjacent axis letter pairs that share an x-column (vertical axes)
    # or a y-row (horizontal axes). Letters in Russian alphabet order are
    # adjacent if their indices differ by 1 in the LETTER_ORDER string.
    # We tighten the alignment threshold to 8 pt (was 30) so that stray
    # legend / description occurrences of the same letter do not pollute
    # the axis-pair list.
    LETTER_ORDER = axis_letters
    ALIGN_TOL = 8.0
    distances = []
    for i in range(len(LETTER_ORDER) - 1):
        a, b = LETTER_ORDER[i], LETTER_ORDER[i + 1]
        if a not in pos_by_letter or b not in pos_by_letter:
            continue
        for ax, ay in pos_by_letter[a]:
            for bx, by in pos_by_letter[b]:
                # vertical axis pair: nearly identical x → measure dy
                if abs(ax - bx) < ALIGN_TOL and abs(ay - by) > 30:
                    distances.append(("vertical", abs(ay - by)))
                # horizontal axis pair: nearly identical y → measure dx
                if abs(ay - by) < ALIGN_TOL and abs(ax - bx) > 30:
                    distances.append(("horizontal", abs(ax - bx)))

    if not distances:
        return None

    # Building axes are physically long; spurious axis-letter occurrences
    # in legends / title blocks tend to be close together (< 100 pt).
    # Evaluate every observed span individually and pick the
    # (span, KNOWN_DIM) combination that lands closest to a STANDARD
    # 1:N drawing scale. When the round-score is tied (which happens
    # often because 1:50 and 1:75 and 1:100 are all standard), prefer
    # column-grid values that are more typical for industrial buildings:
    # 6000 > 7200 > 4800 > 4500 > 5400 > 3600 > 9000 > 12000 > others.
    VALUE_PRIOR_RANK = {
        6000: 0, 7200: 1, 4800: 2, 4500: 3, 5400: 4, 3600: 5,
        9000: 6, 12000: 7, 4200: 8, 3000: 9, 2400: 10, 1800: 11,
        1500: 12,
    }
    candidates = []
    for _kind, span in distances:
        if span < 100:
            continue
        for v in KNOWN_DIMS_MM:
            mm_per_pt = v / span
            if not (5.0 <= mm_per_pt <= 80.0):
                continue
            score = _round_scale_score(mm_per_pt)
            # Bucket scores within 1.0 of each other so tie-break by
            # building-grid prior actually fires. Without bucketing,
            # 3000/170 (score=0.005) would always beat 6000/170 (score=0.04).
            score_bucket = round(score)
            prior = VALUE_PRIOR_RANK.get(v, 99)
            candidates.append((score_bucket, prior, -span, score,
                               mm_per_pt, v, span))
    if not candidates:
        return None
    candidates.sort()
    (best_bucket, best_prior, _neg_span, best_score,
     best_mm_per_pt, best_v, best_span) = candidates[0]
    if best_bucket > 1:
        return None
    return (best_mm_per_pt,
            f"axis_pair_{best_v}mm",
            dict(value_mm=best_v, span_pt=round(best_span, 2),
                 pairs_used=len(distances),
                 round_score=round(best_score, 3),
                 value_prior=best_prior))


def _detect_scale(words, lines, page) -> tuple[float, str, dict]:
    """Run scale detectors in order: building axes first (most reliable on
    plan sheets), then dim-text repeats, then a page-size heuristic.
    """
    cand = _detect_scale_from_axes(words, page_size=(page.width, page.height))
    if cand:
        return cand
    cand = _detect_scale_from_dims(words, lines)
    if cand:
        return cand
    # Heuristic: A1 paper landscape, 1:100
    pw = max(page.width, page.height)
    paper_mm = pw / 72 * 25.4
    mm_per_pt = paper_mm / pw * 100
    return (mm_per_pt, "page_heuristic_1:100",
            dict(paper_mm=round(paper_mm, 1)))


# ---------------------------------------------------------------------------
# Polyline assembly
# ---------------------------------------------------------------------------

class _UnionFind:
    def __init__(self):
        self.parent: dict = {}

    def find(self, x):
        while self.parent.get(x, x) != x:
            self.parent[x] = self.parent.get(self.parent.get(x, x), x)
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[ra] = rb


def _key(p: tuple[float, float], tol: float = ENDPOINT_TOL) -> tuple[int, int]:
    return (int(round(p[0] / tol)), int(round(p[1] / tol)))


def _dedupe_exact_duplicates(segments: list[dict]) -> tuple[list[dict], dict]:
    """Remove segments that are identical (or near-identical) duplicates
    of another segment in the input.

    A segment is a duplicate of another if both endpoints match within
    DUP_ENDPOINT_TOL pt (in either order). PDFs from many CAD exporters
    write the same line several times; on GPK3 we observed up to 7 copies
    of one tray line on the same coordinates — a 7× overcount of length.

    Returns (kept_segments, stats) with how many duplicates were removed.
    """
    seen: dict[tuple, int] = {}
    kept: list[dict] = []
    removed = 0
    removed_pt = 0.0
    for ln in segments:
        L = _seg_len(ln)
        if L < DUP_MIN_LEN_PT:
            kept.append(ln)
            continue
        # Canonical key: snap each endpoint to a grid of DUP_ENDPOINT_TOL,
        # then sort the two endpoints so direction doesn't matter.
        ax = round(ln["x0"] / DUP_ENDPOINT_TOL)
        ay = round(ln["top"] / DUP_ENDPOINT_TOL)
        bx = round(ln["x1"] / DUP_ENDPOINT_TOL)
        by = round(ln["bottom"] / DUP_ENDPOINT_TOL)
        a, b = (ax, ay), (bx, by)
        key = (a, b) if a <= b else (b, a)
        if key in seen:
            removed += 1
            removed_pt += L
            continue
        seen[key] = 1
        kept.append(ln)
    stats = dict(input_count=len(segments), kept_count=len(kept),
                 dup_removed=removed, removed_pt=round(removed_pt, 1))
    return kept, stats


def _dedupe_parallel_segments(segments: list[dict]) -> tuple[list[dict], dict]:
    """Collapse close-parallel pairs of long segments into a single
    centerline segment. This is essential for cable-tray (лоток) drawings
    where each tray is rendered as two parallel edge lines 3-6 pt apart.

    Returns (kept_segments, stats). ``stats`` reports how many segments
    were merged and how much pt-length was removed — useful for sanity.

    Algorithm:
      1. Split inputs into LONG (>= PARALLEL_MIN_LEN_PT) and SHORT.
      2. For each LONG segment, compute (angle_mod_pi, midpoint, length).
      3. Greedily pair LONG segments where:
         - angular difference ≤ PARALLEL_ANGLE_TOL_DEG,
         - perpendicular distance between midpoints in
           (PARALLEL_PERP_MIN_PT, PARALLEL_PERP_MAX_PT),
         - projected overlap fraction ≥ PARALLEL_OVERLAP_MIN.
      4. For each pair: drop one segment, replace the other with its
         centerline (mid of the two parallel lines).
      5. Unpaired LONG segments and all SHORT segments pass through
         unchanged.
    """
    long_idx, short_idx = [], []
    for i, ln in enumerate(segments):
        if _seg_len(ln) >= PARALLEL_MIN_LEN_PT:
            long_idx.append(i)
        else:
            short_idx.append(i)

    # Pre-compute geometry for long segments
    geom: dict[int, tuple] = {}
    for i in long_idx:
        ln = segments[i]
        x1, y1 = ln["x0"], ln["top"]
        x2, y2 = ln["x1"], ln["bottom"]
        L = math.hypot(x2 - x1, y2 - y1)
        if L < 0.5:
            continue
        a = math.atan2(y2 - y1, x2 - x1) % math.pi
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        geom[i] = (a, mx, my, L, x1, y1, x2, y2)

    angle_tol = math.radians(PARALLEL_ANGLE_TOL_DEG)
    used: set[int] = set()
    kept: list[dict] = []
    merged_pairs = 0
    removed_pt = 0.0

    long_sorted = sorted(geom.keys(), key=lambda i: -geom[i][3])
    for i in long_sorted:
        if i in used:
            continue
        a1, mx1, my1, L1, x1, y1, x2, y2 = geom[i]
        dx, dy = x2 - x1, y2 - y1
        partner = None
        for j in long_sorted:
            if j == i or j in used:
                continue
            a2, mx2, my2, L2, _, _, _, _ = geom[j]
            da = abs(a1 - a2)
            if da > math.pi / 2:
                da = math.pi - da
            if da > angle_tol:
                continue
            # perpendicular distance from (mx2,my2) to line of segment i
            perp = abs(dy * (mx2 - x1) - dx * (my2 - y1)) / L1
            if not (PARALLEL_PERP_MIN_PT < perp < PARALLEL_PERP_MAX_PT):
                continue
            # projected overlap of segment j onto direction of segment i
            ux, uy = dx / L1, dy / L1
            j_x1, j_y1 = segments[j]["x0"], segments[j]["top"]
            j_x2, j_y2 = segments[j]["x1"], segments[j]["bottom"]
            t1 = (j_x1 - x1) * ux + (j_y1 - y1) * uy
            t2 = (j_x2 - x1) * ux + (j_y2 - y1) * uy
            lo, hi = (t1, t2) if t1 < t2 else (t2, t1)
            overlap = max(0.0, min(L1, hi) - max(0.0, lo))
            if overlap < PARALLEL_OVERLAP_MIN * min(L1, L2):
                continue
            partner = j
            break

        if partner is None:
            kept.append(segments[i])
            continue

        # Build centerline: average endpoints of i and j (with j's
        # endpoints projected onto i's direction so we keep i's length).
        a2, mx2, my2, L2, jx1, jy1, jx2, jy2 = geom[partner]
        # nx,ny = unit normal to segment i
        nx, ny = -dy / L1, dx / L1
        # signed offset of partner mid relative to segment i
        off = (mx2 - mx1) * nx + (my2 - my1) * ny
        cx_shift, cy_shift = nx * off / 2, ny * off / 2
        center = dict(segments[i])  # shallow copy preserves color/width
        center["x0"] = x1 + cx_shift
        center["top"] = y1 + cy_shift
        center["x1"] = x2 + cx_shift
        center["bottom"] = y2 + cy_shift
        kept.append(center)
        used.add(i)
        used.add(partner)
        merged_pairs += 1
        # We dropped one full partner segment + adjusted i (length unchanged).
        removed_pt += L2

    # Pass through unpaired long segments and ALL short segments
    for i in long_idx:
        if i in geom and i not in used and segments[i] not in kept:
            # Already added unpaired ones above; this guards the L<0.5 case.
            kept.append(segments[i])
    for i in short_idx:
        kept.append(segments[i])

    stats = dict(
        input_count=len(segments),
        kept_count=len(kept),
        merged_pairs=merged_pairs,
        removed_pt=round(removed_pt, 1),
    )
    return kept, stats


def _assemble_polylines(segments: list[dict], color: str) -> list[Polyline]:
    """Connect segments by shared endpoints into polylines.

    Each output Polyline carries length = sum of its segment lengths.
    Geometry order is not reconstructed; we only need length & a points
    bag to find marker-on-route hits.
    """
    uf = _UnionFind()
    seg_records = []
    for ln in segments:
        a = (ln["x0"], ln["top"])
        b = (ln["x1"], ln["bottom"])
        ka, kb = _key(a), _key(b)
        uf.union(ka, kb)
        seg_records.append((ka, kb, a, b, _seg_len(ln)))

    groups: dict = defaultdict(list)
    for ka, kb, a, b, L in seg_records:
        groups[uf.find(ka)].append((a, b, L))

    polylines: list[Polyline] = []
    for _, recs in groups.items():
        pts = []
        total = 0.0
        for a, b, L in recs:
            pts.append(a)
            pts.append(b)
            total += L
        polylines.append(Polyline(color=color, points=pts, length_pt=total))
    return polylines


# ---------------------------------------------------------------------------
# Marker collection
# ---------------------------------------------------------------------------

def _collect_circle_markers(page, zones) -> list[tuple[float, float]]:
    """Find small circular shapes on the page.

    pdfplumber exposes circles as either:
      * ``page.curves`` items (Bezier segments of an arc)
      * ``page.rects`` with near-equal w/h and a stroking color

    Best signal on real plans: PyMuPDF ``get_drawings()`` reports closed
    paths; pdfplumber in many PDFs records small circles as 4-segment
    Bezier curves. We treat any curve with bbox width and height ≤
    2*CIRCLE_MAX_RADIUS_PT as a circle marker.
    """
    out: list[tuple[float, float]] = []
    for cv in (page.curves or []):
        x0 = cv.get("x0"); x1 = cv.get("x1")
        y0 = cv.get("top"); y1 = cv.get("bottom")
        if None in (x0, x1, y0, y1):
            continue
        w = abs(x1 - x0); h = abs(y1 - y0)
        r = max(w, h) / 2
        if not (CIRCLE_MIN_RADIUS_PT <= r <= CIRCLE_MAX_RADIUS_PT):
            continue
        cx = (x0 + x1) / 2; cy = (y0 + y1) / 2
        if _excluded((cx, cy), zones):
            continue
        out.append((cx, cy))

    # Tiny rects with equal sides also act as circle markers
    for r in (page.rects or []):
        w = abs(r["x1"] - r["x0"]); h = abs(r["bottom"] - r["top"])
        if not (CIRCLE_MIN_RADIUS_PT * 2 <= max(w, h) <= CIRCLE_MAX_RADIUS_PT * 2):
            continue
        if abs(w - h) > 1.5:
            continue
        cx = (r["x0"] + r["x1"]) / 2; cy = (r["top"] + r["bottom"]) / 2
        if _excluded((cx, cy), zones):
            continue
        out.append((cx, cy))

    return out


def _collect_tray_rects(page, zones) -> list[tuple[float, float, float, float]]:
    """Find rectangles that look like a cable-tray glyph straddling a
    cable line. Heuristic: width 8–35 pt, height ≤ 10 pt, stroked.
    """
    out: list[tuple[float, float, float, float]] = []
    for r in (page.rects or []):
        w = abs(r["x1"] - r["x0"]); h = abs(r["bottom"] - r["top"])
        if not (TRAY_RECT_MIN_W <= w <= TRAY_RECT_MAX_W and h <= TRAY_RECT_MAX_H):
            continue
        cx = (r["x0"] + r["x1"]) / 2; cy = (r["top"] + r["bottom"]) / 2
        if _excluded((cx, cy), zones):
            continue
        out.append((r["x0"], r["top"], r["x1"], r["bottom"]))
    return out


def _collect_tray_ticks(page, zones,
                         cable_segs: list[dict],
                         ) -> list[tuple[float, float]]:
    """Find short black/dark line segments that act as TRAY ticks: short
    perpendicular strokes crossing a cable line.

    Per legend on Russian electrical drawings, the tray ("на лотке")
    glyph consists of two short perpendicular ticks (~3.5 pt) on either
    side of the cable line midpoint. We look for any line segment of
    length in [TICK_LEN_MIN_PT, TICK_LEN_MAX_PT] whose midpoint lies
    within TICK_TO_LINE_PT of a cable segment AND whose direction is
    perpendicular (±TICK_PERP_TOL_DEG) to that cable segment.

    Returns a list of (cx, cy) midpoints. Each tick is a single hint
    that the nearest cable segment is on a tray.
    """
    if not cable_segs:
        return []
    # Pre-index cable segments by coarse grid (32-pt cells) for fast lookup
    grid: dict[tuple[int, int], list[tuple]] = defaultdict(list)
    cell = 32
    for cs in cable_segs:
        cx1, cy1 = cs["x0"], cs["top"]
        cx2, cy2 = cs["x1"], cs["bottom"]
        L = math.hypot(cx2 - cx1, cy2 - cy1)
        if L < 0.5:
            continue
        ang = math.atan2(cy2 - cy1, cx2 - cx1) % math.pi
        # Bucket every cell that the segment passes near
        n = max(1, int(L / cell) + 1)
        for k in range(n + 1):
            t = k / n
            x = cx1 + (cx2 - cx1) * t
            y = cy1 + (cy2 - cy1) * t
            grid[(int(x // cell), int(y // cell))].append((cx1, cy1, cx2, cy2, ang, L))

    perp_tol = math.radians(TICK_PERP_TOL_DEG)
    out: list[tuple[float, float]] = []
    for ln in (page.lines or []):
        x1, y1 = ln["x0"], ln["top"]
        x2, y2 = ln["x1"], ln["bottom"]
        L = math.hypot(x2 - x1, y2 - y1)
        if not (TICK_LEN_MIN_PT <= L <= TICK_LEN_MAX_PT):
            continue
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        if _excluded((mx, my), zones):
            continue
        # Only consider colored ticks (red/blue): legend ticks on cables
        # of the same color. Black ticks would dominate from text glyphs.
        sc = ln.get("stroking_color")
        if not _classify_color(sc):
            continue
        ang = math.atan2(y2 - y1, x2 - x1) % math.pi
        # Look for nearby cable segment with perpendicular direction
        cx_cell, cy_cell = int(mx // cell), int(my // cell)
        found = False
        for dx_c in (-1, 0, 1):
            if found:
                break
            for dy_c in (-1, 0, 1):
                bucket = grid.get((cx_cell + dx_c, cy_cell + dy_c), [])
                for cx1, cy1, cx2, cy2, c_ang, c_L in bucket:
                    # Skip self
                    if abs(cx1 - x1) < 0.1 and abs(cy1 - y1) < 0.1:
                        continue
                    # Perpendicular distance from tick midpoint to cable line
                    cdx, cdy = cx2 - cx1, cy2 - cy1
                    perp = abs(cdy * (mx - cx1) - cdx * (my - cy1)) / c_L
                    if perp > TICK_TO_LINE_PT:
                        continue
                    # Perpendicular angle?
                    da = abs(ang - c_ang)
                    if da > math.pi / 2:
                        da = math.pi - da
                    diff_from_perp = abs(da - math.pi / 2)
                    if diff_from_perp > perp_tol:
                        continue
                    out.append((mx, my))
                    found = True
                    break
        # done for this tick
    return out


def _attribute_method(pl: Polyline,
                      circle_markers: list[tuple[float, float]],
                      tray_rects: list[tuple[float, float, float, float]],
                      tray_ticks: Optional[list[tuple[float, float]]] = None,
                      sheet_default: str = "po_konstrukciyam",
                      ) -> str:
    """Decide laying method for a single polyline.

    Priority:
      1. Tray rect overlapping any segment endpoint → "lotok"
      2. ≥1 tray tick (perpendicular short stroke crossing a cable
         segment) within ~6 pt of a polyline endpoint → "lotok"
      3. ≥1 circle marker within MARKER_TO_LINE_RADIUS pt of a polyline
         endpoint → "gofra/truba"
      4. otherwise → ``sheet_default`` (parsed from sheet's general notes
         when available, else "po_konstrukciyam")
    """
    pts = pl.points
    # Bounding box of polyline for cheap pre-filter
    if not pts:
        return sheet_default
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    bx0, by0, bx1, by1 = min(xs) - 8, min(ys) - 8, max(xs) + 8, max(ys) + 8

    # Tray check (rect-based). Only fire if the rect is genuinely
    # ON the polyline (within ~3 pt of one of its vertices), not just
    # inside its bbox. This prevents random light-fixture rectangles
    # near (but not on) a cable line from mis-classifying the polyline.
    for tx0, ty0, tx1, ty1 in tray_rects:
        cx = (tx0 + tx1) / 2; cy = (ty0 + ty1) / 2
        if not (bx0 <= cx <= bx1 and by0 <= cy <= by1):
            continue
        for p in pts:
            if math.hypot(p[0] - cx, p[1] - cy) < 3.0:
                return "lotok"

    # Tray check (tick-based: short perpendicular strokes).
    # Require ≥ TRAY_TICK_MIN_HITS ticks lying within TICK_HIT_RADIUS of
    # any polyline endpoint AND a per-pt-length density above
    # TRAY_TICK_MIN_DENSITY (ticks per 100 pt of cable). For SHORT
    # polylines (< TRAY_TICK_SHORT_LEN_PT) we additionally require
    # TRAY_TICK_SHORT_HITS hits to avoid false positives from light-
    # fixture / outlet glyphs that contain perpendicular micro-strokes.
    if tray_ticks and pl.length_pt > 0:
        hits = 0
        for mx, my in tray_ticks:
            if not (bx0 <= mx <= bx1 and by0 <= my <= by1):
                continue
            for p in pts:
                if math.hypot(p[0] - mx, p[1] - my) <= TICK_HIT_RADIUS:
                    hits += 1
                    break
        if hits >= TRAY_TICK_MIN_HITS:
            density = hits / max(pl.length_pt, 1) * 100.0
            short_pl = pl.length_pt < TRAY_TICK_SHORT_LEN_PT
            min_hits_required = TRAY_TICK_SHORT_HITS if short_pl else TRAY_TICK_MIN_HITS
            # If the sheet itself says cables go in a conduit, raise
            # the bar — only believe lotok promotion with HARD evidence.
            if sheet_default == "gofra/truba":
                min_hits_required = max(min_hits_required, TRAY_TICK_HARD_HITS)
            if hits >= min_hits_required and (
                density >= TRAY_TICK_MIN_DENSITY
                or hits >= TRAY_TICK_HARD_HITS
            ):
                return "lotok"

    # Circle marker check
    near = 0
    for mx, my in circle_markers:
        if not (bx0 <= mx <= bx1 and by0 <= my <= by1):
            continue
        for p in pts:
            if math.hypot(p[0] - mx, p[1] - my) <= MARKER_TO_LINE_RADIUS:
                near += 1
                break
        if near >= 1:
            return "gofra/truba"
    return sheet_default


# ---------------------------------------------------------------------------
# Sheet-level default method (parsed from general notes)
# ---------------------------------------------------------------------------

# Each rule: (regex, method, weight). The first regex with the highest
# accumulated weight wins. Patterns are designed to match phrases found
# in real Russian electrical drawings' "Общие указания" / "Примечания"
# blocks and in legend rows.
SHEET_METHOD_RULES = [
    # "на лотке" / "по лоткам" / "в металлических лотках"
    (re.compile(r"\bна\s+лотк[еа]\b", re.IGNORECASE), "lotok", 3),
    (re.compile(r"\bпо\s+лоткам\b", re.IGNORECASE), "lotok", 3),
    (re.compile(r"в\s+(металл\w*\s+)?(перфорированн\w*\s+)?лотк\w+",
                re.IGNORECASE), "lotok", 2),
    # "в трубе" / "в трубах ПВХ/ПНД" / "в гофре"
    (re.compile(r"\bв\s+труб[еах]\w*\s*(пвх|пнд)?\b", re.IGNORECASE),
        "gofra/truba", 2),
    (re.compile(r"\bв\s+гофр[еа]\b", re.IGNORECASE), "gofra/truba", 2),
    # "по строительным конструкциям" / "открыто по конструкциям"
    (re.compile(r"по\s+(строительн\w+\s+)?конструкц\w+", re.IGNORECASE),
        "po_konstrukciyam", 2),
    (re.compile(r"\bоткрыт[оа]\s+по\s+", re.IGNORECASE),
        "po_konstrukciyam", 2),
]


# Cable mark patterns. We look for canonical Russian power-cable marks
# either with explicit cross-section ("ВВГнг(А)-LS 5х2,5") or by mark
# prefix on its own; the cross-section is then matched by proximity.
CABLE_MARK_PREFIX_RE = re.compile(
    r"\b("
    r"ВВГ\w*|АВВГ\w*|ВБ[бБ]?Шв?\w*|ПвБбШв\w*|"
    r"ППГ\w*|NYM\w*|КГ\w*|КВВГ\w*|КГВВ\w*|"
    r"АПвВнг\w*|ПвВнг\w*"
    r")"
    r"(?:\s*\(?[А-Яа-я]{1,2}\)?)?"
    r"(?:\s*[-]\s*(?:LS|FRLS|HF|FRHF|FR|нг|FR\w*))?",
    re.IGNORECASE,
)
CABLE_SECTION_RE = re.compile(
    r"(\d+)\s*[xх×]\s*(\d+(?:[.,]\d+)?)"
)


def _detect_sheet_cable_marks(words: list[dict]) -> list[dict]:
    """Find unique cable mark mentions on the sheet. Each mark is
    returned with its prefix (e.g. "ВВГнг(А)-LS"), an associated
    cross-section if found nearby, and a count of mentions.

    The aim is sheet-level identification — most Russian electrical
    plans state one or two cable marks in the general notes and use
    those everywhere on the sheet.
    """
    found: dict[str, dict] = {}
    # Group words into lines for context
    lines_by_y: dict[int, list[dict]] = defaultdict(list)
    for w in words:
        lines_by_y[round(w["top"] / 3) * 3].append(w)

    for ws in lines_by_y.values():
        ws.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws)
        for pm in CABLE_MARK_PREFIX_RE.finditer(text):
            mark = pm.group(0).strip()
            # Look for cross-section within 60 chars after the mark
            tail = text[pm.end():pm.end() + 60]
            sm = CABLE_SECTION_RE.search(tail)
            section = None
            if sm:
                cores, area = sm.group(1), sm.group(2)
                section = f"{cores}x{area.replace(',', '.')}"
            key = (mark.upper(), section)
            if key not in found:
                found[key] = dict(mark=mark, section=section, mentions=0,
                                  example=text[:120])
            found[key]["mentions"] += 1
    # Sort by mentions desc
    return sorted(found.values(), key=lambda d: -d["mentions"])


def _detect_sheet_default_method(words: list[dict]) -> tuple[str, dict]:
    """Scan sheet text and pick the most frequently mentioned laying
    method. Returns (method, info). If no rule fires we default to
    "po_konstrukciyam" (which is the most common method on Russian
    plans anyway).

    Strategy: join words into lines (by ``top`` proximity), run regex
    rules against the joined text, accumulate weighted hit counts per
    method, then return the winner. Title-block / штамп rows are not
    excluded because their text rarely matches our patterns.
    """
    # Group words into lines
    lines_by_y: dict[int, list[dict]] = defaultdict(list)
    for w in words:
        lines_by_y[round(w["top"] / 3) * 3].append(w)

    score: dict[str, int] = defaultdict(int)
    hits: dict[str, list[str]] = defaultdict(list)
    for ws in lines_by_y.values():
        ws.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in ws)
        for pat, method, weight in SHEET_METHOD_RULES:
            if pat.search(text):
                score[method] += weight
                if len(hits[method]) < 3:
                    hits[method].append(text[:80])

    if not score:
        return "po_konstrukciyam", dict(reason="no_rule_fired", scores={})

    # Winner = highest score
    method = max(score.items(), key=lambda kv: kv[1])[0]
    return method, dict(scores=dict(score), examples=dict(hits))


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def measure_cables(
    pdf_path: str,
    page_index: int = 0,
    legend_result: Optional[LegendResult] = None,
    scale_mm_per_pt: Optional[float] = None,
) -> CableMethodReport:
    if legend_result is None:
        try:
            legend_result = parse_legend(pdf_path)
        except Exception:
            legend_result = None

    legend_bbox = (0, 0, 0, 0)
    if legend_result and legend_result.legend_bbox != (0, 0, 0, 0):
        legend_bbox = legend_result.legend_bbox

    rep = CableMethodReport(pdf_path=pdf_path, page_index=page_index)

    with pdfplumber.open(pdf_path) as pdf:
        if page_index >= len(pdf.pages):
            return rep
        page = pdf.pages[page_index]
        rep.page_size = (page.width, page.height)
        words = page.extract_words(x_tolerance=2, y_tolerance=2) or []
        lines = page.lines or []

        # Scale
        if scale_mm_per_pt:
            rep.scale_mm_per_pt = scale_mm_per_pt
            rep.scale_source = "user_supplied"
        else:
            mm_per_pt, source, _info = _detect_scale(words, lines, page)
            rep.scale_mm_per_pt = round(mm_per_pt, 4)
            rep.scale_source = source

        zones = _build_zones(page, lines, legend_bbox)

        # Filter colored cable lines
        red_segs, blue_segs = [], []
        for ln in lines:
            sc = ln.get("stroking_color")
            cls = _classify_color(sc)
            if not cls:
                continue
            mp = _midpoint(ln)
            if _excluded(mp, zones):
                continue
            (red_segs if cls == "red" else blue_segs).append(ln)
        rep.raw_red_lines = len(red_segs)
        rep.raw_blue_lines = len(blue_segs)

        # Markers
        circle_markers = _collect_circle_markers(page, zones)
        tray_rects = _collect_tray_rects(page, zones)
        # Tray ticks (perpendicular short strokes crossing cable lines)
        tray_ticks = _collect_tray_ticks(page, zones, red_segs + blue_segs)
        rep.circle_markers = len(circle_markers)
        rep.tray_rects = len(tray_rects)
        rep.tray_ticks = len(tray_ticks)

        # Sheet-level default laying method (from general notes / legend)
        sheet_default, sheet_info = _detect_sheet_default_method(words)
        rep.sheet_default_method = sheet_default
        rep.sheet_default_info = sheet_info
        # Sheet-level cable mark(s)
        rep.sheet_cable_marks = _detect_sheet_cable_marks(words)

        # Dedupe exact duplicates first (CAD exporters often emit a
        # cable line 2-7× on the same coordinates), then close-parallel
        # pairs (tray = two parallel edges).
        red_segs, dup_red = _dedupe_exact_duplicates(red_segs)
        blue_segs, dup_blue = _dedupe_exact_duplicates(blue_segs)
        red_segs, rep.dedup_red = _dedupe_parallel_segments(red_segs)
        blue_segs, rep.dedup_blue = _dedupe_parallel_segments(blue_segs)
        # Merge stats: keep both passes visible
        rep.dedup_red["dup_removed"] = dup_red.get("dup_removed", 0)
        rep.dedup_red["dup_removed_pt"] = dup_red.get("removed_pt", 0.0)
        rep.dedup_blue["dup_removed"] = dup_blue.get("dup_removed", 0)
        rep.dedup_blue["dup_removed_pt"] = dup_blue.get("removed_pt", 0.0)

        # Polylines per color. Pick the fallback as follows:
        #   * if the sheet's general notes mention "в гофре/трубе" much
        #     more strongly than other methods → use "gofra/truba" as
        #     fallback (KPP-30 case: notes say all cables go in PVC
        #     conduits → lotok would be a false positive everywhere).
        #   * else default to "po_konstrukciyam" — safe guess.
        # We deliberately do NOT auto-promote to "lotok" because lighting
        # plans have many short wall-drops that genuinely are
        # po_konstrukciyam (to switches/sockets), and tray promotion via
        # ticks already covers the lotok case.
        scores = sheet_info.get("scores", {})
        gofra_score = scores.get("gofra/truba", 0)
        if gofra_score >= 4 and gofra_score >= 2 * scores.get("lotok", 0):
            fallback_method = "gofra/truba"
        else:
            fallback_method = "po_konstrukciyam"
        rep.sheet_default_method = fallback_method
        rep.sheet_default_info = sheet_info
        for color, segs in (("red", red_segs), ("blue", blue_segs)):
            polylines = _assemble_polylines(segs, color)
            for pl in polylines:
                pl.laying_method = _attribute_method(
                    pl, circle_markers, tray_rects,
                    tray_ticks=tray_ticks,
                    sheet_default=fallback_method,
                )
                rep.polylines.append(pl)

        # Aggregate totals
        for pl in rep.polylines:
            length_m = pl.length_pt * rep.scale_mm_per_pt / 1000.0
            key = (pl.color, pl.laying_method)
            rep.totals_m[key] = rep.totals_m.get(key, 0.0) + length_m

        # Optional: aggregate L=…м labels for cross-check
        for w in words:
            t = w["text"]
            m = re.match(r"^L=(\d+(?:[,\.]\d+)?)м$", t)
            if m:
                try:
                    rep.label_total_m += float(m.group(1).replace(",", "."))
                except ValueError:
                    pass

    return rep


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _print_report(rep: CableMethodReport) -> None:
    print(f"PDF: {rep.pdf_path}")
    print(f"Page: {rep.page_index}  size={rep.page_size}")
    print(f"Scale: {rep.scale_mm_per_pt} mm/pt   ({rep.scale_source})")
    real_scale = rep.scale_mm_per_pt / MM_PER_PT
    print(f"  ≈ 1:{real_scale:.0f} drawing scale")
    print()
    print(f"Raw cable lines: red={rep.raw_red_lines}, blue={rep.raw_blue_lines}")
    if rep.dedup_red or rep.dedup_blue:
        for color, st in (("red", rep.dedup_red), ("blue", rep.dedup_blue)):
            if not st:
                continue
            dr = st.get("dup_removed", 0)
            mp = st.get("merged_pairs", 0)
            if dr or mp:
                print(f"  dedup-{color}: exact dups removed={dr} "
                      f"({st.get('dup_removed_pt', 0)} pt); "
                      f"parallel pairs merged={mp} "
                      f"({st.get('removed_pt', 0)} pt)")
    print(f"Markers: circles={rep.circle_markers}, "
          f"tray rects={rep.tray_rects}, tray ticks={rep.tray_ticks}")
    print(f"Sheet default method: {rep.sheet_default_method} "
          f"(scores={rep.sheet_default_info.get('scores', {})})")
    if rep.sheet_cable_marks:
        print(f"Sheet cable marks (top 5):")
        for cm in rep.sheet_cable_marks[:5]:
            sec = cm['section'] or '?'
            print(f"  {cm['mark']:<25} section={sec:<10} x{cm['mentions']}")
    print(f"Polylines: {len(rep.polylines)}")
    print()
    print("=== Length by (color × laying method) ===")
    print(f"{'color':<6} {'method':<22} {'length, m':>12}")
    print("-" * 45)
    by_method = defaultdict(float)
    for (c, m), L in sorted(rep.totals_m.items()):
        print(f"{c:<6} {m:<22} {L:>12.1f}")
        by_method[m] += L
    print()
    print("=== Length by laying method (sum of red+blue) ===")
    for m, L in sorted(by_method.items(), key=lambda x: -x[1]):
        print(f"  {m:<22} {L:>10.1f} м")
    print()
    print(f"Sum total (all colors, all methods): "
          f"{sum(rep.totals_m.values()):.1f} м")
    if rep.label_total_m > 0:
        print(f"Cross-check (sum of L=…м labels):    "
              f"{rep.label_total_m:.1f} м")


def main() -> None:
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                      errors="replace")
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8",
                                      errors="replace")

    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Usage: python pdf_cables_by_method.py <pdf> [--page N] "
              "[--scale-mm-per-pt 26.47]")
        sys.exit(1)

    pdf_path = args[0]
    page_index = 0
    user_scale: Optional[float] = None

    i = 1
    while i < len(args):
        a = args[i]
        if a == "--page" and i + 1 < len(args):
            page_index = int(args[i + 1]); i += 2
        elif a == "--scale-mm-per-pt" and i + 1 < len(args):
            user_scale = float(args[i + 1]); i += 2
        else:
            i += 1

    rep = measure_cables(pdf_path, page_index=page_index,
                         scale_mm_per_pt=user_scale)
    _print_report(rep)


if __name__ == "__main__":
    main()
