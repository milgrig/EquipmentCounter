"""
pdf_count_geometry.py — Measure cable routes by analyzing PDF line geometry.

On ЭО and ГПК electrical plans, cable routes are drawn as colored lines:
  - RED  (1,0,0) = аварийное (emergency) lighting
  - BLUE (0,0,1) = рабочее  (working) lighting

The drawing scale is derived from dimension text (e.g. "6000" mm between
grid axis lines whose PDF spacing is ~170 pt → scale ≈ 35.29 mm/pt).

This module:
  1. Detects drawing scale from dimension text + nearby line segments
  2. Classifies lines by stroking_color
  3. Filters out exclusion zones (title block, legend, grid margins)
  4. Measures total cable route length per color / linewidth group
  5. Reconstructs connected polyline routes (graph-based)

Usage:
    python pdf_count_geometry.py <path.pdf> [--page N] [--all-pages]
"""

from __future__ import annotations

import io
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

from pdf_legend_parser import parse_legend, LegendResult

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class CableRoute:
    """A group of cable segments sharing color + linewidth."""
    color: str = ""                # 'red' / 'blue'
    linewidth: float = 0.0         # PDF linewidth value
    total_length_pt: float = 0.0   # total route length in PDF points
    total_length_m: float = 0.0    # total route length in meters
    segment_count: int = 0         # number of line segments
    route_count: int = 0           # number of connected polylines
    route_paths: list = field(default_factory=list)  # connected polylines
    page_index: int = 0


@dataclass
class ScaleInfo:
    """Drawing scale information."""
    mm_per_pt: float = 0.0        # millimeters per PDF point
    source: str = ""               # how the scale was determined
    dim_value_mm: int = 0          # the dimension value used (e.g. 6000)
    dim_span_pt: float = 0.0       # the measured span in PDF points
    confidence: str = ""           # 'high', 'medium', 'low'


@dataclass
class GeometryResult:
    """Result of geometric cable measurement."""
    routes: list[CableRoute] = field(default_factory=list)
    scale: Optional[ScaleInfo] = None
    # Summary
    total_red_length_m: float = 0.0
    total_blue_length_m: float = 0.0
    total_red_segments: int = 0
    total_blue_segments: int = 0
    # By linewidth
    by_linewidth: dict[str, dict] = field(default_factory=dict)
    # Metadata
    pages_scanned: list[int] = field(default_factory=list)
    total_lines_on_page: int = 0
    cable_lines_count: int = 0
    exclusion_zones: list[tuple[str, tuple[float, float, float, float]]] = \
        field(default_factory=list)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Cable line colors (RGB tuples)
COLOR_RED = (1.0, 0.0, 0.0)
COLOR_BLUE = (0.0, 0.0, 1.0)
COLOR_TOLERANCE = 0.05

# Known dimension values to search for (mm)
KNOWN_DIMS_MM = [6000, 4800, 4200, 3600, 3000, 12000, 9000, 7200]

# Dimension text regex: 3-5 digit integers
DIM_TEXT_RE = re.compile(r"^(\d{3,5})$")

# Scale detection: max distance from dim text to dim line (pt)
DIM_LINE_SEARCH_RADIUS = 30

# Minimum segment length to count (pt) — filter tiny symbol fragments
MIN_SEGMENT_LENGTH_PT = 3.0

# Polyline endpoint merge tolerance (pt)
ENDPOINT_TOLERANCE = 2.0

# Grid axis margin
GRID_AXIS_MARGIN = 60

# Title block detection
TITLE_BLOCK_MIN_LINES = 6


# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------

def _color_matches(color, target: tuple, tol: float = COLOR_TOLERANCE) -> bool:
    """Check if a color tuple matches target within tolerance."""
    if not isinstance(color, tuple) or len(color) != len(target):
        return False
    return all(abs(c - t) < tol for c, t in zip(color, target))


def _classify_line_color(stroking_color) -> str:
    """Classify a line's stroking_color as 'red', 'blue', or 'other'."""
    if stroking_color is None:
        return "other"
    if isinstance(stroking_color, (int, float)):
        return "other"  # grayscale — structural/grid
    if not isinstance(stroking_color, tuple):
        return "other"
    if _color_matches(stroking_color, COLOR_RED):
        return "red"
    if _color_matches(stroking_color, COLOR_BLUE):
        return "blue"
    return "other"


# ---------------------------------------------------------------------------
# Exclusion zone detection
# ---------------------------------------------------------------------------

def _detect_title_block(
    page,
    lines: list[dict],
) -> Optional[tuple[float, float, float, float]]:
    """Detect the title block (штамп) from PDF lines."""
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

    x0_counter: dict[float, int] = {}
    for ln in h_lines:
        rx = round(ln["x0"], 0)
        x0_counter[rx] = x0_counter.get(rx, 0) + 1

    best_x0 = max(x0_counter, key=x0_counter.get)  # type: ignore[arg-type]
    all_x0s = sorted(x0_counter.keys())

    tb_x0 = pw
    for x in all_x0s:
        if x0_counter[x] >= 3 and x < tb_x0:
            tb_x0 = x

    tb_x1 = max(ln["x1"] for ln in h_lines)
    tb_y0 = min(
        ln["top"] for ln in h_lines
        if abs(round(ln["x0"], 0) - tb_x0) < 10
        or abs(round(ln["x0"], 0) - best_x0) < 10
    )
    tb_y1 = max(ln["top"] for ln in h_lines)

    return (tb_x0, tb_y0, tb_x1, tb_y1)


def _build_exclusion_zones(
    page,
    lines: list[dict],
    legend_bbox: Optional[tuple[float, float, float, float]] = None,
) -> list[tuple[str, tuple[float, float, float, float]]]:
    """Build exclusion zones: title block, grid margins, legend."""
    pw, ph = page.width, page.height
    m = GRID_AXIS_MARGIN
    zones: list[tuple[str, tuple[float, float, float, float]]] = []

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


def _line_in_zone(
    ln: dict,
    zone: tuple[float, float, float, float],
    margin: float = 0,
) -> bool:
    """Check if a line segment is fully inside a bounding box zone."""
    x0 = min(ln["x0"], ln["x1"])
    x1 = max(ln["x0"], ln["x1"])
    y0 = min(ln["top"], ln["bottom"])
    y1 = max(ln["top"], ln["bottom"])

    return (
        x0 >= zone[0] - margin
        and x1 <= zone[2] + margin
        and y0 >= zone[1] - margin
        and y1 <= zone[3] + margin
    )


def _line_intersects_zone(
    ln: dict,
    zone: tuple[float, float, float, float],
) -> bool:
    """Check if line midpoint is inside zone (simpler, less strict)."""
    mx = (ln["x0"] + ln["x1"]) / 2
    my = (ln["top"] + ln["bottom"]) / 2
    return (
        zone[0] <= mx <= zone[2]
        and zone[1] <= my <= zone[3]
    )


def _line_excluded(
    ln: dict,
    zones: list[tuple[str, tuple[float, float, float, float]]],
) -> bool:
    """Check if a line's midpoint falls in any exclusion zone."""
    for _, zb in zones:
        if _line_intersects_zone(ln, zb):
            return True
    return False


# ---------------------------------------------------------------------------
# Scale detection
# ---------------------------------------------------------------------------

def _detect_scale(
    page,
    words: list[dict],
    lines: list[dict],
) -> ScaleInfo:
    """
    Detect drawing scale from dimension text and nearby line segments.

    Strategy:
    1. Find dimension text (e.g. '6000') that appears multiple times at same Y
    2. Find horizontal dim lines at that Y coordinate
    3. Match dim lines to consistent spacing → mm_per_pt

    Fallback: use page size heuristic (A1 = 841×594mm, 1:100 common).
    """
    pw, ph = page.width, page.height

    # Collect dimension words grouped by Y position (row)
    dim_words: list[tuple[int, dict]] = []
    for w in words:
        m = DIM_TEXT_RE.match(w["text"])
        if not m:
            continue
        val = int(m.group(1))
        if val in KNOWN_DIMS_MM:
            dim_words.append((val, w))

    # Group dim words by (value, y_row)
    # Multiple identical dim values at same Y → dimension line row
    row_groups: dict[tuple[int, float], list[dict]] = defaultdict(list)
    for val, w in dim_words:
        y_key = round(w["top"] / 10) * 10  # bucket by 10pt
        row_groups[(val, y_key)].append(w)

    # Find the best row: most repetitions of same dimension value
    best_row = None
    best_count = 0
    best_val = 0
    for (val, y_key), row_words in row_groups.items():
        if len(row_words) > best_count:
            best_count = len(row_words)
            best_row = row_words
            best_val = val

    if best_row and best_count >= 2:
        # Sort words by X position
        sorted_words = sorted(best_row, key=lambda w: w["x0"])

        # Calculate average spacing between consecutive dim texts
        spacings = []
        for i in range(len(sorted_words) - 1):
            dx = sorted_words[i + 1]["x0"] - sorted_words[i]["x0"]
            spacings.append(dx)

        if spacings:
            # Use median spacing for robustness
            spacings.sort()
            median_spacing = spacings[len(spacings) // 2]

            if median_spacing > 10:  # sanity check
                mm_per_pt = best_val / median_spacing

                # Also try to verify with horizontal dim lines
                ref_y = sorted_words[0]["top"]
                dim_lines = [
                    ln for ln in lines
                    if abs(ln["top"] - ref_y) < DIM_LINE_SEARCH_RADIUS
                    and abs(ln["y1"] - ln["y0"]) < 2  # horizontal
                    and abs(ln["x1"] - ln["x0"]) > 30  # not too short
                ]

                # Check if dim lines have consistent spacing matching text
                if dim_lines:
                    line_widths = [abs(ln["x1"] - ln["x0"]) for ln in dim_lines]
                    line_widths.sort()
                    median_line_w = line_widths[len(line_widths) // 2]

                    # If dim lines exist at consistent width, use that
                    if abs(median_line_w - median_spacing) < 20:
                        mm_per_pt = best_val / median_line_w
                        return ScaleInfo(
                            mm_per_pt=round(mm_per_pt, 4),
                            source=f"dim_line_{best_val}mm",
                            dim_value_mm=best_val,
                            dim_span_pt=round(median_line_w, 1),
                            confidence="high",
                        )

                return ScaleInfo(
                    mm_per_pt=round(mm_per_pt, 4),
                    source=f"dim_text_{best_val}mm",
                    dim_value_mm=best_val,
                    dim_span_pt=round(median_spacing, 1),
                    confidence="medium",
                )

    # Fallback: page size heuristic
    # Common architectural formats: A1 (841×594mm), A0 (1189×841mm)
    # At 1:100 scale, 1mm real = 0.01mm drawing = 0.72/25.4*0.01 pt
    # But we need real_mm per PDF_pt
    # A1 paper: 841mm wide → pw points. So 1 pt = 841/pw mm (paper mm)
    # At 1:100 scale: 1 pt = 841/pw * 100 real mm
    paper_w_mm = max(pw, ph) / 72 * 25.4  # page width in mm (at 72 DPI)
    # Guess scale from paper size
    if paper_w_mm > 1100:  # A0-ish
        scale_factor = 100
    elif paper_w_mm > 780:  # A1-ish
        scale_factor = 100
    elif paper_w_mm > 550:  # A2-ish
        scale_factor = 50
    else:
        scale_factor = 100  # default

    mm_per_pt = paper_w_mm / max(pw, ph) * scale_factor
    return ScaleInfo(
        mm_per_pt=round(mm_per_pt, 4),
        source=f"page_size_heuristic_1:{scale_factor}",
        dim_value_mm=0,
        dim_span_pt=0,
        confidence="low",
    )


# ---------------------------------------------------------------------------
# Line segment measurement
# ---------------------------------------------------------------------------

def _segment_length(ln: dict) -> float:
    """Calculate length of a line segment in PDF points."""
    dx = ln["x1"] - ln["x0"]
    dy = ln["y1"] - ln["y0"]
    return math.sqrt(dx * dx + dy * dy)


# ---------------------------------------------------------------------------
# Polyline route reconstruction
# ---------------------------------------------------------------------------

def _build_routes(segments: list[dict], tolerance: float = ENDPOINT_TOLERANCE) -> list[list[dict]]:
    """
    Build connected polyline routes from individual line segments.

    Uses a union-find (disjoint set) approach: segments sharing endpoints
    (within tolerance) belong to the same route.

    Returns list of connected segment groups (routes).
    """
    if not segments:
        return []

    n = len(segments)

    # Union-Find
    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    # Build endpoint index for fast neighbor search
    # Bucket endpoints into grid cells
    cell_size = tolerance * 2
    endpoint_grid: dict[tuple[int, int], list[tuple[int, int]]] = defaultdict(list)
    # Each segment has 2 endpoints: (seg_idx, 0=start, 1=end)

    endpoints = []
    for i, seg in enumerate(segments):
        p0 = (seg["x0"], seg["y0"])
        p1 = (seg["x1"], seg["y1"])
        endpoints.append((i, p0, p1))

        for px, py in [p0, p1]:
            gx = int(px / cell_size)
            gy = int(py / cell_size)
            endpoint_grid[(gx, gy)].append(i)

    # For each segment, find other segments with nearby endpoints
    for i, seg_i in enumerate(segments):
        for px, py in [(seg_i["x0"], seg_i["y0"]), (seg_i["x1"], seg_i["y1"])]:
            gx = int(px / cell_size)
            gy = int(py / cell_size)

            # Check neighboring cells
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for j in endpoint_grid.get((gx + dx, gy + dy), []):
                        if j <= i:
                            continue
                        seg_j = segments[j]
                        # Check if any endpoint of j is near (px, py)
                        for qx, qy in [(seg_j["x0"], seg_j["y0"]),
                                        (seg_j["x1"], seg_j["y1"])]:
                            if abs(px - qx) < tolerance and abs(py - qy) < tolerance:
                                union(i, j)
                                break

    # Group segments by root
    groups: dict[int, list[dict]] = defaultdict(list)
    for i, seg in enumerate(segments):
        root = find(i)
        groups[root].append(seg)

    return list(groups.values())


# ---------------------------------------------------------------------------
# Core extraction logic
# ---------------------------------------------------------------------------

def measure_cables(
    pdf_path: str,
    legend_result: Optional[LegendResult] = None,
    pages: Optional[list[int]] = None,
) -> GeometryResult:
    """
    Measure cable routes by analyzing PDF line geometry.

    Args:
        pdf_path: Path to the PDF file.
        legend_result: Pre-parsed legend result (for exclusion zones).
        pages: Specific page indices to scan. If None, scans all pages.

    Returns:
        GeometryResult with route measurements, scale, and summaries.
    """
    # Parse legend for exclusion zones
    if legend_result is None:
        try:
            legend_result = parse_legend(pdf_path)
        except Exception:
            legend_result = None

    legend_bbox = None
    legend_page = -1
    if legend_result and legend_result.items:
        legend_bbox = legend_result.legend_bbox
        legend_page = legend_result.page_index

    all_routes: list[CableRoute] = []
    result_scale: Optional[ScaleInfo] = None
    all_zones: list[tuple[str, tuple[float, float, float, float]]] = []
    pages_scanned: list[int] = []
    total_lines_count = 0
    cable_lines_total = 0

    total_red_m = 0.0
    total_blue_m = 0.0
    total_red_segs = 0
    total_blue_segs = 0
    by_lw: dict[str, dict] = {}

    # Collect segments per page first, detect best scale, then compute lengths
    # This is a two-phase approach so scale from any page applies globally.
    page_data: list[tuple[int, dict[float, list[dict]], dict[float, list[dict]]]] = []

    with pdfplumber.open(pdf_path) as pdf:
        scan_pages = pages if pages is not None else list(range(len(pdf.pages)))

        for page_idx in scan_pages:
            if page_idx >= len(pdf.pages):
                continue

            page = pdf.pages[page_idx]
            pdf_lines = page.lines or []
            words = page.extract_words(x_tolerance=3, y_tolerance=3) or []

            if not pdf_lines:
                continue

            pages_scanned.append(page_idx)
            total_lines_count += len(pdf_lines)

            # Detect scale — keep best across all pages
            page_scale = _detect_scale(page, words, pdf_lines)
            if result_scale is None:
                result_scale = page_scale
            elif (result_scale.confidence != "high"
                  and page_scale.confidence == "high"):
                result_scale = page_scale
            elif (result_scale.confidence == "low"
                  and page_scale.confidence == "medium"):
                result_scale = page_scale

            # Build exclusion zones
            lb = legend_bbox if page_idx == legend_page else None
            zones = _build_exclusion_zones(page, pdf_lines, lb)

            if page_idx == legend_page or not all_zones:
                all_zones = zones

            # Classify and filter lines
            red_segments: dict[float, list[dict]] = defaultdict(list)
            blue_segments: dict[float, list[dict]] = defaultdict(list)

            for ln in pdf_lines:
                if _line_excluded(ln, zones):
                    continue

                seg_len = _segment_length(ln)
                if seg_len < MIN_SEGMENT_LENGTH_PT:
                    continue

                color = _classify_line_color(ln.get("stroking_color"))
                if color == "other":
                    continue

                lw = round(ln.get("linewidth", 0), 3)
                cable_lines_total += 1

                if color == "red":
                    red_segments[lw].append(ln)
                else:
                    blue_segments[lw].append(ln)

            page_data.append((page_idx, dict(red_segments), dict(blue_segments)))

    # Phase 2: compute lengths using the best scale found
    mm_per_pt = result_scale.mm_per_pt if result_scale else 35.0

    for page_idx, red_segments, blue_segments in page_data:
        for lw, segs in sorted(red_segments.items()):
            total_pt = sum(_segment_length(s) for s in segs)
            total_m = total_pt * mm_per_pt / 1000.0
            routes = _build_routes(segs)

            route = CableRoute(
                color="red",
                linewidth=lw,
                total_length_pt=round(total_pt, 1),
                total_length_m=round(total_m, 1),
                segment_count=len(segs),
                route_count=len(routes),
                route_paths=routes,
                page_index=page_idx,
            )
            all_routes.append(route)
            total_red_m += total_m
            total_red_segs += len(segs)

            lw_key = f"red/{lw}"
            if lw_key not in by_lw:
                by_lw[lw_key] = {
                    "color": "red", "linewidth": lw,
                    "length_m": 0.0, "segments": 0, "routes": 0,
                }
            by_lw[lw_key]["length_m"] = round(
                by_lw[lw_key]["length_m"] + total_m, 1)
            by_lw[lw_key]["segments"] += len(segs)
            by_lw[lw_key]["routes"] += len(routes)

        for lw, segs in sorted(blue_segments.items()):
            total_pt = sum(_segment_length(s) for s in segs)
            total_m = total_pt * mm_per_pt / 1000.0
            routes = _build_routes(segs)

            route = CableRoute(
                color="blue",
                linewidth=lw,
                total_length_pt=round(total_pt, 1),
                total_length_m=round(total_m, 1),
                segment_count=len(segs),
                route_count=len(routes),
                route_paths=routes,
                page_index=page_idx,
            )
            all_routes.append(route)
            total_blue_m += total_m
            total_blue_segs += len(segs)

            lw_key = f"blue/{lw}"
            if lw_key not in by_lw:
                by_lw[lw_key] = {
                    "color": "blue", "linewidth": lw,
                    "length_m": 0.0, "segments": 0, "routes": 0,
                }
            by_lw[lw_key]["length_m"] = round(
                by_lw[lw_key]["length_m"] + total_m, 1)
            by_lw[lw_key]["segments"] += len(segs)
            by_lw[lw_key]["routes"] += len(routes)

    return GeometryResult(
        routes=all_routes,
        scale=result_scale,
        total_red_length_m=round(total_red_m, 1),
        total_blue_length_m=round(total_blue_m, 1),
        total_red_segments=total_red_segs,
        total_blue_segments=total_blue_segs,
        by_linewidth=by_lw,
        pages_scanned=pages_scanned,
        total_lines_on_page=total_lines_count,
        cable_lines_count=cable_lines_total,
        exclusion_zones=all_zones,
    )


# ---------------------------------------------------------------------------
# CLI interface
# ---------------------------------------------------------------------------

def main() -> None:
    """CLI entry point for testing."""
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8", errors="replace"
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8", errors="replace"
        )

    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print("Usage: python pdf_count_geometry.py <path.pdf> [--page N] [--all-pages]")
        print()
        print("Options:")
        print("  --page N      Scan specific page (0-based)")
        print("  --all-pages   Scan all pages (default: legend page only)")
        print("  --routes      Show individual route details")
        sys.exit(1)

    pdf_path = args[0]
    scan_pages = None
    show_routes = "--routes" in args

    # Parse --page N
    for i, arg in enumerate(args[1:], 1):
        if arg == "--page" and i + 1 < len(args):
            try:
                scan_pages = [int(args[i + 1])]
            except ValueError:
                print(f"Invalid page number: {args[i + 1]}")
                sys.exit(1)

    all_pages_mode = "--all-pages" in args

    print(f"Geometry measurement: {pdf_path}")
    print()

    # Parse legend
    legend = None
    try:
        legend = parse_legend(pdf_path)
        if legend.items:
            print(f"Legend: {len(legend.items)} items on page {legend.page_index + 1}")
        else:
            print("No legend found")
    except Exception as e:
        print(f"Legend parse failed: {e}")

    print()

    # Determine pages
    if scan_pages is None:
        if all_pages_mode:
            scan_pages = None
        elif legend and legend.items:
            scan_pages = [legend.page_index]
        else:
            scan_pages = None

    # Measure
    result = measure_cables(pdf_path, legend, scan_pages)

    # Scale info
    if result.scale:
        s = result.scale
        print(f"=== Scale ===")
        print(f"  {s.mm_per_pt} mm/pt ({s.source})")
        if s.dim_value_mm:
            print(f"  {s.dim_value_mm} mm = {s.dim_span_pt} pt")
        print(f"  Confidence: {s.confidence}")
        # Show real-world scale ratio
        # 1 pt = mm_per_pt mm in reality
        # Standard notation: 1:N where N = mm_per_pt / (25.4/72)
        pt_mm = 25.4 / 72  # 1 pt = 0.3528 mm on paper
        scale_ratio = s.mm_per_pt / pt_mm
        print(f"  Drawing scale ≈ 1:{scale_ratio:.0f}")
    print()

    # Line stats
    print(f"Pages scanned: {[p + 1 for p in result.pages_scanned]}")
    print(f"Total lines: {result.total_lines_on_page}")
    print(f"Cable lines (red+blue): {result.cable_lines_count}")
    print()

    if result.exclusion_zones:
        print("Exclusion zones:")
        for name, bbox in result.exclusion_zones:
            print(f"  {name}: ({bbox[0]:.0f}, {bbox[1]:.0f})"
                  f" — ({bbox[2]:.0f}, {bbox[3]:.0f})")
        print()

    # By linewidth table
    if result.by_linewidth:
        print("=== Cable Routes by Color/Linewidth ===")
        print(f"{'Color':<8s} {'LW':>6s} {'Segments':>10s} "
              f"{'Routes':>8s} {'Length(m)':>12s}")
        print("-" * 50)
        for lw_key, info in sorted(result.by_linewidth.items()):
            print(f"  {info['color']:<6s} {info['linewidth']:>6.3f} "
                  f"{info['segments']:>10d} {info['routes']:>8d} "
                  f"{info['length_m']:>12.1f}")
        print()

    # Summary
    color_labels = {"red": "аварийное", "blue": "рабочее"}
    print("=== Summary ===")
    print(f"  RED  ({color_labels['red']}):  "
          f"{result.total_red_segments} segments, "
          f"{result.total_red_length_m:.1f} m")
    print(f"  BLUE ({color_labels['blue']}): "
          f"{result.total_blue_segments} segments, "
          f"{result.total_blue_length_m:.1f} m")
    total_m = result.total_red_length_m + result.total_blue_length_m
    print(f"  TOTAL: {total_m:.1f} m")

    # Route details
    if show_routes and result.routes:
        print()
        print("=== Route Details ===")
        for route in result.routes:
            print(f"\n  {route.color} lw={route.linewidth}: "
                  f"{route.segment_count} segments, "
                  f"{route.route_count} routes, "
                  f"{route.total_length_m:.1f}m (p.{route.page_index + 1})")
            # Show top 5 longest routes
            if route.route_paths:
                route_lengths = []
                for rp in route.route_paths:
                    rl = sum(_segment_length(s) for s in rp)
                    route_lengths.append((rl, len(rp)))
                route_lengths.sort(key=lambda x: -x[0])
                mm_per_pt = result.scale.mm_per_pt if result.scale else 35.0
                print(f"    Top routes:")
                for rl, rc in route_lengths[:5]:
                    print(f"      {rl * mm_per_pt / 1000:.1f}m "
                          f"({rc} segments, {rl:.0f}pt)")


if __name__ == "__main__":
    main()
