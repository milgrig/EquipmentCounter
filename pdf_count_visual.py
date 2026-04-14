"""
pdf_count_visual.py — Count equipment by visual (graphical) symbol matching.

On ЭМ (mechanical) and some ЭО plans, equipment is represented by
graphical symbols (drawings), not text markers.  The legend table shows
each symbol as a small graphic in the "Обозначение" column.

This module:
  1. Extracts symbol images from the legend table using PyMuPDF (fitz)
  2. Renders the full drawing page at high DPI
  3. Matches each symbol template on full grayscale page using OpenCV
     template matching with pyramid (coarse-to-fine) acceleration (T142)
  4. Single scale (1.0x) with rotation (0, 90) — pyramid speed-up
  5. Pre-NMS exclusion zone filtering + non-maximum suppression
  6. Post-hoc color validation (rejects matches with wrong color)

Dependencies:
  - fitz (PyMuPDF) for PDF rendering
  - cv2 (OpenCV) for template matching
  - numpy for array operations

Usage:
    python pdf_count_visual.py <path.pdf> [--page N] [--threshold 0.7]
"""

from __future__ import annotations

import io
import logging
import math
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Optional

logger = logging.getLogger(__name__)

import cv2
import fitz  # PyMuPDF
import numpy as np
import pdfplumber

from pdf_legend_parser import parse_legend, LegendResult, LegendItem
from pdf_color_layers import extract_raster_layers

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class VisualMatch:
    """A single visual match of a legend symbol on the drawing."""
    symbol_index: int            # index in legend items list
    description: str = ""        # legend item description
    x: float = 0.0              # center x on page (PDF pt)
    y: float = 0.0              # center y on page (PDF pt)
    confidence: float = 0.0      # match confidence (0..1)
    scale: float = 1.0           # matched scale factor
    rotation: int = 0            # matched rotation angle (degrees)
    color: str = ""              # detected color ('red', 'blue', '')
    page_index: int = 0


@dataclass
class VisualResult:
    """Result of visual symbol matching."""
    matches: list[VisualMatch] = field(default_factory=list)
    counts: dict[int, int] = field(default_factory=dict)   # symbol_index → count
    descriptions: dict[int, str] = field(default_factory=dict)  # symbol_index → desc
    # Metadata
    page_index: int = 0
    legend_page: int = 0
    symbols_extracted: int = 0
    render_dpi: int = 200
    threshold: float = 0.7
    scales_used: list[float] = field(default_factory=list)
    rotations_used: list[int] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Rendering DPI
RENDER_DPI = 200            # full-page rendering (balance speed vs accuracy)
SYMBOL_DPI = 300            # symbol cell rendering (higher for template clarity)

# Template matching
DEFAULT_THRESHOLD = 0.75     # minimum match confidence
SCALES = [1.0]                        # single scale (T142: 3→1 cuts 3× match calls)
ROTATIONS = [0, 90]                   # 0° + 90° (symbols may be placed vertically)

# Pyramid (coarse-to-fine) matching (T142)
PYRAMID_DOWNSAMPLE = 0.5    # coarse pass at 50% resolution
PYRAMID_COARSE_THRESH = 0.55 # lower threshold for coarse pass (catch all candidates)
PYRAMID_VERIFY_PAD = 1.5     # expand candidate ROI by this factor for fine verification

# Non-maximum suppression
NMS_OVERLAP_THRESH = 0.3    # IoU threshold for NMS
NMS_ADAPTIVE_RATIO = 0.6    # suppress radius = max(template_w, template_h) * ratio

# Symbol cell extraction: minimum symbol size (pixels at SYMBOL_DPI)
MIN_SYMBOL_SIZE_PX = 8       # skip empty/too-small symbols
MIN_SYMBOL_AREA_PX = 100     # minimum non-white pixel area

# Color detection
COLOR_RED = (1.0, 0.0, 0.0)
COLOR_BLUE = (0.0, 0.0, 1.0)
COLOR_TOLERANCE = 0.05

# Exclusion zone margin
GRID_AXIS_MARGIN = 60
TITLE_BLOCK_MIN_LINES = 6

# Context-aware false positive filter (T133)
FP_LINE_DENSITY_THRESH = 0.45   # max ratio of edge pixels in ROI (hatching)
FP_TEXT_OVERLAP_THRESH = 0.30   # max ratio of ROI area covered by text chars
FP_ROI_EXPAND = 1.5             # expand ROI by this factor for context analysis

# Non-equipment template skip patterns (T142)
# Templates whose description matches these patterns are typically wiring/cable
# path descriptions that produce massive raw detections and no useful counts.
NON_EQUIPMENT_KEYWORDS = [
    "прокладыв",      # Проводка прокладываемая (wiring being laid)
    "кабельн",         # Кабельная трасса (cable route)
    "трасс",           # cable route / path
    "провод скрыт",    # hidden wiring
    "линия связи",     # communication line
]

# Legend table column structure (ЭМ-style drawings)
# These are used when the legend parser provides item bboxes
# The symbol column is typically between the first and second vertical separators
SYMBOL_COL_WIDTH_MIN_PT = 30   # minimum width for symbol column (pt)
SYMBOL_COL_WIDTH_MAX_PT = 120  # maximum width for symbol column (pt)


# ---------------------------------------------------------------------------
# Exclusion zone detection (shared logic)
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
            (legend_bbox[0] - 5, legend_bbox[1] - 5,
             legend_bbox[2] + 5, legend_bbox[3] + 5),
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


def _pt_in_zone(
    x: float,
    y: float,
    zone: tuple[float, float, float, float],
    margin: float = 0,
) -> bool:
    """Check if a point falls within a bounding box zone."""
    return (
        x >= zone[0] - margin
        and x <= zone[2] + margin
        and y >= zone[1] - margin
        and y <= zone[3] + margin
    )


def _pt_excluded(
    x: float,
    y: float,
    zones: list[tuple[str, tuple[float, float, float, float]]],
    return_zone_name: bool = False,
) -> bool | str:
    """Check if a point is in any exclusion zone.

    Args:
        return_zone_name: If True, return the zone name (str) that matched,
                          or "" if no zone matched.  If False (default),
                          return bool.
    """
    for name, zb in zones:
        if _pt_in_zone(x, y, zb):
            return name if return_zone_name else True
    return "" if return_zone_name else False


# ---------------------------------------------------------------------------
# Color-cluster equipment zone detection
# ---------------------------------------------------------------------------

# Morphological kernel lengths for cable line removal (pixels at RENDER_DPI)
_LINE_KERNEL_LEN = 40   # lines shorter than this are kept as equipment
_EQUIP_DILATE_ITER = 3  # dilate iterations to connect nearby equipment parts
_EQUIP_DILATE_K = 5     # dilate kernel size (px)
_EQUIP_ZONE_PAD_PX = 20 # padding around each equipment cluster (pixels)
_MIN_EQUIP_AREA_PX = 30 # minimum contour area to count as equipment cluster


def _build_equipment_zone_mask(
    color_mask: np.ndarray,
    line_kernel_len: int = _LINE_KERNEL_LEN,
) -> np.ndarray:
    """Build a binary mask of equipment zones from a color layer mask.

    Process:
      1. Remove long straight cable lines via morphological opening
      2. Dilate remaining pixels to connect nearby symbol parts
      3. Return binary mask where 255 = equipment zone

    Args:
        color_mask: Binary mask (H×W, uint8) of a single color layer
                    (e.g. blue mask from HSV).
        line_kernel_len: Minimum length for a feature to be considered a
                         cable line (pixels). Longer → more aggressive removal.

    Returns:
        Binary mask (H×W, uint8): 255 = equipment zone, 0 = background/cable.
    """
    if color_mask is None or color_mask.size == 0:
        return np.zeros((1, 1), dtype=np.uint8)

    # Step 1: Detect long horizontal and vertical lines
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (line_kernel_len, 1))
    h_lines = cv2.morphologyEx(color_mask, cv2.MORPH_OPEN, h_kernel)

    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, line_kernel_len))
    v_lines = cv2.morphologyEx(color_mask, cv2.MORPH_OPEN, v_kernel)

    lines_mask = cv2.bitwise_or(h_lines, v_lines)

    # Step 2: Subtract cable lines → equipment-only pixels
    equip_mask = cv2.bitwise_and(color_mask, cv2.bitwise_not(lines_mask))

    # Step 3: Dilate to connect nearby symbol fragments into clusters
    dilate_k = cv2.getStructuringElement(
        cv2.MORPH_RECT, (_EQUIP_DILATE_K, _EQUIP_DILATE_K)
    )
    equip_dilated = cv2.dilate(equip_mask, dilate_k, iterations=_EQUIP_DILATE_ITER)

    # Step 4: Find contours and build padded zone mask
    contours, _ = cv2.findContours(
        equip_dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )

    zone_mask = np.zeros_like(color_mask)
    for cnt in contours:
        area = cv2.contourArea(cnt)
        if area < _MIN_EQUIP_AREA_PX:
            continue
        x, y, w, h = cv2.boundingRect(cnt)
        # Add padding
        x0 = max(0, x - _EQUIP_ZONE_PAD_PX)
        y0 = max(0, y - _EQUIP_ZONE_PAD_PX)
        x1 = min(zone_mask.shape[1], x + w + _EQUIP_ZONE_PAD_PX)
        y1 = min(zone_mask.shape[0], y + h + _EQUIP_ZONE_PAD_PX)
        zone_mask[y0:y1, x0:x1] = 255

    return zone_mask


def _point_in_zone_mask(
    cx_px: float,
    cy_px: float,
    zone_mask: np.ndarray,
) -> bool:
    """Check if a pixel coordinate falls within equipment zone mask."""
    ix = int(round(cx_px))
    iy = int(round(cy_px))
    if iy < 0 or iy >= zone_mask.shape[0] or ix < 0 or ix >= zone_mask.shape[1]:
        return False
    return zone_mask[iy, ix] > 0


def build_equipment_cluster_bboxes(
    pdf_path: str,
    page_index: int = 0,
    render_dpi: int = RENDER_DPI,
    page_bgr: Optional[np.ndarray] = None,
    proximity_pt: float = 25.0,
) -> dict[str, list[tuple[float, float, float, float]]]:
    """Build equipment cluster bounding boxes per color in PDF pt coords.

    Uses color layer masks → cable line removal → contour detection
    to find where equipment clusters are on the drawing.

    Args:
        pdf_path:  Path to PDF.
        page_index:  Page to analyse.
        render_dpi:  DPI for rasterisation.
        page_bgr:  Pre-rendered page (optional).
        proximity_pt: Extra padding around each cluster in PDF pt.

    Returns:
        ``{"red": [(x0,y0,x1,y1), ...], "blue": [...]}`` in PDF pt coords.
    """
    if page_bgr is None:
        page_bgr = _render_page(pdf_path, page_index, dpi=render_dpi)

    raster_layers, _ = extract_raster_layers(
        pdf_path, page_index, render_dpi, page_bgr=page_bgr,
    )

    px_per_pt = render_dpi / 72.0
    result: dict[str, list[tuple[float, float, float, float]]] = {}

    for cname in ("red", "blue"):
        layer = raster_layers.get(cname)
        if layer is None or layer.mask is None:
            result[cname] = []
            continue

        mask = layer.mask
        # Remove cable lines
        h_k = cv2.getStructuringElement(cv2.MORPH_RECT, (_LINE_KERNEL_LEN, 1))
        h_lines = cv2.morphologyEx(mask, cv2.MORPH_OPEN, h_k)
        v_k = cv2.getStructuringElement(cv2.MORPH_RECT, (1, _LINE_KERNEL_LEN))
        v_lines = cv2.morphologyEx(mask, cv2.MORPH_OPEN, v_k)
        lines = cv2.bitwise_or(h_lines, v_lines)
        equip = cv2.bitwise_and(mask, cv2.bitwise_not(lines))

        # Dilate to connect fragments
        dk = cv2.getStructuringElement(
            cv2.MORPH_RECT, (_EQUIP_DILATE_K, _EQUIP_DILATE_K)
        )
        equip_d = cv2.dilate(equip, dk, iterations=_EQUIP_DILATE_ITER)

        contours, _ = cv2.findContours(
            equip_d, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
        )

        bboxes: list[tuple[float, float, float, float]] = []
        pad_px = proximity_pt * px_per_pt
        for cnt in contours:
            if cv2.contourArea(cnt) < _MIN_EQUIP_AREA_PX:
                continue
            x, y, w, h = cv2.boundingRect(cnt)
            # Convert to PDF pt with padding
            x0 = max(0.0, (x - pad_px) / px_per_pt)
            y0 = max(0.0, (y - pad_px) / px_per_pt)
            x1 = (x + w + pad_px) / px_per_pt
            y1 = (y + h + pad_px) / px_per_pt
            bboxes.append((x0, y0, x1, y1))

        result[cname] = bboxes

    return result


# ---------------------------------------------------------------------------
# Context-aware false positive filter (T133)
# ---------------------------------------------------------------------------

def _is_false_positive(
    page_gray: np.ndarray,
    cx_px: float,
    cy_px: float,
    w_px: float,
    h_px: float,
    chars_in_drawing: list[dict],
    px_per_pt: float,
) -> bool:
    """Check if a template match is a false positive by analyzing its context.

    Three checks:
    1. Line density: hatching/dense line areas produce edge-heavy ROIs
    2. Text overlap: matches on top of text characters are false positives
    3. Template size sanity: extremely small matched regions are suspect

    Args:
        page_gray: Grayscale page image (full or color-layer).
        cx_px, cy_px: Match center in pixel coordinates.
        w_px, h_px: Match dimensions in pixels.
        chars_in_drawing: pdfplumber char dicts in the drawing area (pre-filtered).
        px_per_pt: Pixels per PDF point conversion factor.

    Returns:
        True if the match should be rejected as a false positive.
    """
    img_h, img_w = page_gray.shape[:2]

    # Expand ROI for context analysis
    ew = w_px * FP_ROI_EXPAND / 2
    eh = h_px * FP_ROI_EXPAND / 2
    x1 = max(0, int(cx_px - ew))
    y1 = max(0, int(cy_px - eh))
    x2 = min(img_w, int(cx_px + ew))
    y2 = min(img_h, int(cy_px + eh))

    if x2 <= x1 or y2 <= y1:
        return True  # degenerate ROI

    roi = page_gray[y1:y2, x1:x2]
    roi_area = roi.shape[0] * roi.shape[1]
    if roi_area < 25:
        return True

    # --- Check 1: Line/edge density (hatching detection) ---
    # Dense hatching produces many Canny edges; real symbols have moderate edges
    edges = cv2.Canny(roi, 50, 150)
    edge_ratio = np.count_nonzero(edges) / roi_area
    if edge_ratio > FP_LINE_DENSITY_THRESH:
        return True

    # --- Check 2: Text overlap ---
    # Convert match center to PDF pt and check overlap with text chars
    cx_pt = cx_px / px_per_pt
    cy_pt = cy_px / px_per_pt
    half_w_pt = (w_px / px_per_pt) / 2
    half_h_pt = (h_px / px_per_pt) / 2

    # Count how much of the match area is covered by text characters
    text_area = 0.0
    match_area_pt = (2 * half_w_pt) * (2 * half_h_pt)
    if match_area_pt > 0:
        for ch in chars_in_drawing:
            ch_x0 = float(ch.get("x0", 0))
            ch_y0 = float(ch.get("top", 0))
            ch_x1 = float(ch.get("x1", 0))
            ch_y1 = float(ch.get("bottom", 0))

            # Compute intersection
            ix0 = max(cx_pt - half_w_pt, ch_x0)
            iy0 = max(cy_pt - half_h_pt, ch_y0)
            ix1 = min(cx_pt + half_w_pt, ch_x1)
            iy1 = min(cy_pt + half_h_pt, ch_y1)

            if ix1 > ix0 and iy1 > iy0:
                text_area += (ix1 - ix0) * (iy1 - iy0)

        text_overlap_ratio = text_area / match_area_pt
        if text_overlap_ratio > FP_TEXT_OVERLAP_THRESH:
            return True

    return False


# ---------------------------------------------------------------------------
# Symbol extraction from legend
# ---------------------------------------------------------------------------

def _find_legend_symbol_column(
    page_lines: list[dict],
    legend_bbox: tuple[float, float, float, float],
) -> Optional[tuple[float, float]]:
    """
    Find the symbol column boundaries within the legend table.

    The legend table has vertical separator lines. The symbol column
    is typically the first narrow column (after the left edge of the table).

    Returns (x_left, x_right) of the symbol column, or None.
    """
    lb_x0, lb_y0, lb_x1, lb_y1 = legend_bbox

    # Find vertical lines within the legend bbox
    v_lines_x: list[float] = []
    for ln in page_lines:
        # Vertical line: small horizontal span, large vertical span
        if abs(ln["x1"] - ln["x0"]) > 3:
            continue
        if abs(ln["top"] - ln["bottom"]) < 20:
            continue
        x = (ln["x0"] + ln["x1"]) / 2
        # Must be within the legend's horizontal range
        if lb_x0 - 5 <= x <= lb_x1 + 5:
            # Must be within the legend's vertical range
            mid_y = (ln["top"] + ln["bottom"]) / 2
            if lb_y0 - 5 <= mid_y <= lb_y1 + 5:
                v_lines_x.append(round(x, 1))

    if not v_lines_x:
        return None

    # Deduplicate close vertical lines
    v_lines_x.sort()
    unique_x: list[float] = []
    for x in v_lines_x:
        if not unique_x or abs(x - unique_x[-1]) > 3:
            unique_x.append(x)

    # The symbol column is between the first two vertical lines after the table left edge
    # (or between left edge and first interior vertical line)
    if len(unique_x) < 2:
        return None

    # Find the symbol column — narrow column at the start
    for i in range(len(unique_x) - 1):
        width = unique_x[i + 1] - unique_x[i]
        if SYMBOL_COL_WIDTH_MIN_PT <= width <= SYMBOL_COL_WIDTH_MAX_PT:
            return (unique_x[i], unique_x[i + 1])

    return None


def _find_legend_row_boundaries(
    page_lines: list[dict],
    legend_bbox: tuple[float, float, float, float],
    sym_col: tuple[float, float],
) -> list[tuple[float, float]]:
    """
    Find horizontal row boundaries within the legend table.

    Returns list of (y_top, y_bottom) for each data row.
    """
    lb_x0, lb_y0, lb_x1, lb_y1 = legend_bbox
    sx0, sx1 = sym_col

    # Find horizontal lines spanning at least the symbol column
    h_lines_y: list[float] = []
    for ln in page_lines:
        if abs(ln["top"] - ln["bottom"]) > 3:
            continue
        if abs(ln["x1"] - ln["x0"]) < 20:
            continue
        y = (ln["top"] + ln["bottom"]) / 2
        # Must be in legend vertical range
        if lb_y0 - 5 <= y <= lb_y1 + 5:
            # Must span at least part of the symbol column
            if ln["x0"] <= sx1 and ln["x1"] >= sx0:
                h_lines_y.append(round(y, 1))

    if not h_lines_y:
        return []

    # Deduplicate
    h_lines_y.sort()
    unique_y: list[float] = []
    for y in h_lines_y:
        if not unique_y or abs(y - unique_y[-1]) > 3:
            unique_y.append(y)

    # Build row pairs
    rows: list[tuple[float, float]] = []
    for i in range(len(unique_y) - 1):
        height = unique_y[i + 1] - unique_y[i]
        if height > 8:  # skip header separator lines that are too close
            rows.append((unique_y[i], unique_y[i + 1]))

    return rows


def _extract_symbol_images(
    pdf_path: str,
    legend_result: LegendResult,
    dpi: int = SYMBOL_DPI,
) -> list[tuple[int, LegendItem, Optional[np.ndarray]]]:
    """
    Extract symbol images from the legend table.

    Uses PyMuPDF to render each symbol cell at high DPI.

    Returns list of (item_index, legend_item, image_array_or_None).
    """
    results: list[tuple[int, LegendItem, Optional[np.ndarray]]] = []

    if not legend_result.items:
        return results

    page_idx = legend_result.page_index
    legend_bbox = legend_result.legend_bbox

    if legend_bbox == (0, 0, 0, 0):
        return results

    # Open with pdfplumber to find table structure (symbol column X range)
    sym_col = None

    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return results
        page = pdf.pages[page_idx]
        page_lines = page.lines or []

        sym_col = _find_legend_symbol_column(page_lines, legend_bbox)
        if sym_col is None:
            return results

    # Compute per-item Y boundaries from item bbox centres.
    # Row boundary = midpoint between consecutive item Y-centres.
    # This is more robust than searching for horizontal table lines,
    # which may be missing in merged-cell tables.
    items = legend_result.items
    y_centres = []
    for item in items:
        b = item.bbox
        if b == (0, 0, 0, 0):
            y_centres.append(None)
        else:
            y_centres.append((b[1] + b[3]) / 2)

    item_row_bounds: list[tuple[float, float] | None] = []
    for i, yc in enumerate(y_centres):
        if yc is None:
            item_row_bounds.append(None)
            continue
        # top boundary: midpoint to previous item, or legend top
        prev_yc = None
        for j in range(i - 1, -1, -1):
            if y_centres[j] is not None:
                prev_yc = y_centres[j]
                break
        if prev_yc is not None:
            row_top = (prev_yc + yc) / 2
        else:
            row_top = legend_bbox[1]
        # bottom boundary: midpoint to next item, or legend bottom
        next_yc = None
        for j in range(i + 1, len(y_centres)):
            if y_centres[j] is not None:
                next_yc = y_centres[j]
                break
        if next_yc is not None:
            row_bot = (yc + next_yc) / 2
        else:
            row_bot = legend_bbox[3]
        item_row_bounds.append((row_top, row_bot))

    # Open with fitz to render symbol cells
    doc = fitz.open(pdf_path)
    fitz_page = doc[page_idx]
    zoom = dpi / 72.0

    sx0, sx1 = sym_col
    # Padding inside the cell to exclude table border lines (typically 1-2pt thick)
    pad = 4  # pt

    for item_idx, item in enumerate(items):
        bounds = item_row_bounds[item_idx]
        if bounds is None:
            results.append((item_idx, item, None))
            continue

        row_top, row_bot = bounds

        # Define clip rect for this symbol cell
        clip = fitz.Rect(sx0 + pad, row_top + pad, sx1 - pad, row_bot - pad)

        # Render at high DPI
        mat = fitz.Matrix(zoom, zoom)
        pix = fitz_page.get_pixmap(matrix=mat, clip=clip, alpha=False)

        # Convert to numpy array
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
            pix.height, pix.width, 3
        )
        # fitz outputs RGB, OpenCV uses BGR
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

        # Check if the symbol has any meaningful content
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        # Count non-white pixels (anything darker than 240)
        non_white = np.count_nonzero(gray < 240)

        if non_white < MIN_SYMBOL_AREA_PX:
            # Too few non-white pixels → empty cell
            results.append((item_idx, item, None))
            continue

        results.append((item_idx, item, img))

    doc.close()
    return results


# ---------------------------------------------------------------------------
# Page rendering
# ---------------------------------------------------------------------------

def _render_page(
    pdf_path: str,
    page_idx: int,
    dpi: int = RENDER_DPI,
) -> np.ndarray:
    """
    Render a PDF page at specified DPI using PyMuPDF.

    Returns BGR numpy array.
    """
    doc = fitz.open(pdf_path)
    page = doc[page_idx]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)

    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
        pix.height, pix.width, 3
    )
    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    doc.close()
    return img


# ---------------------------------------------------------------------------
# Template preprocessing
# ---------------------------------------------------------------------------

def _preprocess_template(
    img: np.ndarray,
    keep_color: str = "",
) -> np.ndarray:
    """
    Preprocess a symbol template for matching.

    Steps:
    0. (optional) If *keep_color* is "red" or "blue", mask out pixels of
       the OTHER colour so multi-colour legend cells become single-colour.
    1. Convert to grayscale
    2. Apply binary threshold (white background → white, symbol → black)
    3. Remove text labels (small isolated components) — keep only the
       largest connected component group (the graphical symbol itself)
    4. Crop to bounding box of remaining content
    5. Add small padding
    """
    work = img.copy()

    # --- colour isolation ---------------------------------------------------
    if keep_color in ("red", "blue"):
        hsv = cv2.cvtColor(work, cv2.COLOR_BGR2HSV)
        sat = hsv[:, :, 1]
        hue = hsv[:, :, 0]
        val = hsv[:, :, 2]
        coloured = (sat > 40) & (val > 40)

        if keep_color == "red":
            # keep red (hue 0-10 | 170-180), blank blue pixels
            other_mask = coloured & (hue > 10) & (hue < 170)
        else:
            # keep blue (hue 100-130), blank red pixels
            other_mask = coloured & ((hue < 100) | (hue > 130))

        # Paint "other colour" pixels white
        work[other_mask] = (255, 255, 255)

    gray = cv2.cvtColor(work, cv2.COLOR_BGR2GRAY)

    # Binary threshold: dark pixels = symbol content
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

    # Find bounding box of content
    coords = cv2.findNonZero(binary)
    if coords is None:
        return gray  # no content found

    # --- Remove text labels and duplicate symbol variants ---
    # Legend cells may contain:
    #   a) Text labels ("2", "7АЭ") — small isolated components
    #   b) Multiple symbol variants (e.g. blue+red versions stacked)
    # Strategy: find connected components, keep ONLY those that
    # physically overlap the largest component's bounding box.
    # This isolates a single symbol shape for template matching.
    num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(binary)

    if num_labels > 2:  # more than 1 foreground component
        # Find the largest foreground component (skip label 0 = background)
        areas = stats[1:, cv2.CC_STAT_AREA]
        largest_label = int(np.argmax(areas)) + 1  # +1 because we skipped bg

        # Bounding box of the largest component
        lx = stats[largest_label, cv2.CC_STAT_LEFT]
        ly = stats[largest_label, cv2.CC_STAT_TOP]
        lw = stats[largest_label, cv2.CC_STAT_WIDTH]
        lh = stats[largest_label, cv2.CC_STAT_HEIGHT]

        # Keep only components whose centroid falls within the
        # largest component's bbox (+ small margin). This removes
        # both text labels AND duplicate symbol variants that are
        # vertically/horizontally separated.
        margin = 3
        keep_labels = {largest_label}
        for lbl in range(1, num_labels):
            if lbl == largest_label:
                continue
            cx_comp = centroids[lbl][0]
            cy_comp = centroids[lbl][1]

            # Component centroid must be inside the largest's bbox
            overlapping = (lx - margin <= cx_comp <= lx + lw + margin and
                           ly - margin <= cy_comp <= ly + lh + margin)

            if overlapping:
                keep_labels.add(lbl)

        # Erase components that are NOT kept (set to white in gray)
        for lbl in range(1, num_labels):
            if lbl not in keep_labels:
                gray[labels == lbl] = 255
                binary[labels == lbl] = 0

    # Re-find bounding box after text removal
    coords = cv2.findNonZero(binary)
    if coords is None:
        return gray

    x, y, w, h = cv2.boundingRect(coords)

    # Crop with small padding
    pad = 3
    y0 = max(0, y - pad)
    y1 = min(gray.shape[0], y + h + pad)
    x0 = max(0, x - pad)
    x1 = min(gray.shape[1], x + w + pad)

    cropped = gray[y0:y1, x0:x1]

    if cropped.shape[0] < MIN_SYMBOL_SIZE_PX or cropped.shape[1] < MIN_SYMBOL_SIZE_PX:
        return gray  # too small after crop

    return cropped


def _preprocess_color_template(img: np.ndarray) -> tuple[np.ndarray, str]:
    """
    Preprocess template preserving color information.

    Returns (cropped_bgr, dominant_color) where dominant_color is
    'red', 'blue', or '' for grayscale/black symbols.
    """
    # Detect dominant color from non-white, non-black pixels
    hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)

    # Mask for colored pixels (saturation > 50, value > 50)
    # Note: pure red (255,0,0) and blue (0,0,255) have V=255, so no upper V limit
    mask = (hsv[:, :, 1] > 50) & (hsv[:, :, 2] > 50)

    color = ""
    colored_count = np.count_nonzero(mask)

    if colored_count > 20:
        # Get average hue of colored pixels
        colored_hues = hsv[:, :, 0][mask]
        avg_hue = np.median(colored_hues)

        # Red: hue 0-10 or 170-180; Blue: hue 100-130
        if avg_hue < 10 or avg_hue > 170:
            color = "red"
        elif 100 <= avg_hue <= 130:
            color = "blue"

    # Crop to content
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)
    coords = cv2.findNonZero(binary)

    if coords is not None:
        x, y, w, h = cv2.boundingRect(coords)
        pad = 3
        y0 = max(0, y - pad)
        y1 = min(img.shape[0], y + h + pad)
        x0 = max(0, x - pad)
        x1 = min(img.shape[1], x + w + pad)
        cropped = img[y0:y1, x0:x1]
    else:
        cropped = img

    return cropped, color


def _rotate_image(img: np.ndarray, angle: int) -> np.ndarray:
    """Rotate image by 0, 90, 180, or 270 degrees."""
    if angle == 0:
        return img
    elif angle == 90:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    elif angle == 180:
        return cv2.rotate(img, cv2.ROTATE_180)
    elif angle == 270:
        return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return img


def _scale_image(img: np.ndarray, scale: float) -> np.ndarray:
    """Scale image by a factor."""
    if abs(scale - 1.0) < 0.01:
        return img
    h, w = img.shape[:2]
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))
    return cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_LINEAR)


# ---------------------------------------------------------------------------
# Template matching
# ---------------------------------------------------------------------------

def _match_template_multi(
    page_gray: np.ndarray,
    template_gray: np.ndarray,
    threshold: float,
    scales: list[float],
    rotations: list[int],
    dpi_ratio: float,
    page_gray_coarse: np.ndarray | None = None,
) -> list[tuple[float, float, float, float, int, float, float]]:
    """
    Match a template at multiple scales and rotations using pyramid
    (coarse-to-fine) strategy for speed (T142).

    Pass 1 (coarse): match at PYRAMID_DOWNSAMPLE resolution with a lower
    threshold to find candidate regions quickly.
    Pass 2 (fine): for each candidate, verify at full resolution in a
    local ROI around the candidate location.

    If the template is small (< 20 px either dimension at coarse scale),
    falls back to direct full-resolution matching to avoid aliasing.

    Args:
        page_gray: Grayscale page image (full resolution)
        template_gray: Grayscale template image
        threshold: Minimum confidence
        scales: Scale factors to try
        rotations: Rotation angles to try
        dpi_ratio: SYMBOL_DPI / RENDER_DPI for template rescaling
        page_gray_coarse: Pre-computed coarse page image (optional, for reuse)

    Returns:
        List of (x_center, y_center, width, height, rotation, scale, confidence)
        in pixel coordinates of the full-resolution page image.
    """
    ds = PYRAMID_DOWNSAMPLE
    matches: list[tuple[float, float, float, float, int, float, float]] = []

    for rot in rotations:
        rotated = _rotate_image(template_gray, rot)

        for scale in scales:
            effective_scale = scale / dpi_ratio
            scaled = _scale_image(rotated, effective_scale)

            th, tw = scaled.shape[:2]

            # Skip if template is larger than page
            if th >= page_gray.shape[0] or tw >= page_gray.shape[1]:
                continue
            # Skip if template is too small
            if th < 5 or tw < 5:
                continue

            # Decide: use pyramid or direct matching
            coarse_th = int(th * ds)
            coarse_tw = int(tw * ds)
            use_pyramid = (
                page_gray_coarse is not None
                and coarse_th >= 10 and coarse_tw >= 10
            )

            if use_pyramid:
                # --- PASS 1: coarse matching ---
                tpl_coarse = cv2.resize(
                    scaled, (coarse_tw, coarse_th),
                    interpolation=cv2.INTER_AREA,
                )
                result_c = cv2.matchTemplate(
                    page_gray_coarse, tpl_coarse, cv2.TM_CCOEFF_NORMED,
                )
                locs_c = np.where(result_c >= PYRAMID_COARSE_THRESH)

                if locs_c[0].size == 0:
                    continue

                # Cluster coarse candidates to reduce verification count:
                # use quick NMS-like grid suppression (keep best per cell).
                # Cell size = template size at coarse resolution.
                cell_h = max(coarse_th, 1)
                cell_w = max(coarse_tw, 1)
                best_per_cell: dict[tuple[int, int], tuple[int, int, float]] = {}
                for yc, xc in zip(locs_c[0], locs_c[1]):
                    gi, gj = int(yc / cell_h), int(xc / cell_w)
                    conf_c = float(result_c[yc, xc])
                    key = (gi, gj)
                    if key not in best_per_cell or conf_c > best_per_cell[key][2]:
                        best_per_cell[key] = (int(yc), int(xc), conf_c)

                # --- PASS 2: fine verification at full res ---
                pad = PYRAMID_VERIFY_PAD
                ph_full, pw_full = page_gray.shape[:2]
                for (yc, xc, _conf_c) in best_per_cell.values():
                    # Map coarse coords → full-res coords
                    fx = int(xc / ds)
                    fy = int(yc / ds)

                    # ROI in full-res page around candidate
                    roi_x1 = max(0, fx - int(tw * pad / 2))
                    roi_y1 = max(0, fy - int(th * pad / 2))
                    roi_x2 = min(pw_full, fx + tw + int(tw * pad / 2))
                    roi_y2 = min(ph_full, fy + th + int(th * pad / 2))

                    roi = page_gray[roi_y1:roi_y2, roi_x1:roi_x2]
                    if roi.shape[0] < th or roi.shape[1] < tw:
                        continue

                    result_f = cv2.matchTemplate(
                        roi, scaled, cv2.TM_CCOEFF_NORMED,
                    )
                    locs_f = np.where(result_f >= threshold)
                    for yf, xf in zip(locs_f[0], locs_f[1]):
                        conf = float(result_f[yf, xf])
                        # Convert ROI-local coords to full-page coords
                        cx = (roi_x1 + xf) + tw / 2.0
                        cy = (roi_y1 + yf) + th / 2.0
                        matches.append((cx, cy, tw, th, rot, scale, conf))
            else:
                # --- Direct full-resolution matching (small templates) ---
                result = cv2.matchTemplate(
                    page_gray, scaled, cv2.TM_CCOEFF_NORMED,
                )
                locations = np.where(result >= threshold)
                for y, x in zip(locations[0], locations[1]):
                    conf = float(result[y, x])
                    cx = x + tw / 2.0
                    cy = y + th / 2.0
                    matches.append((cx, cy, tw, th, rot, scale, conf))

    return matches


# ---------------------------------------------------------------------------
# Non-maximum suppression
# ---------------------------------------------------------------------------

def _nms(
    detections: list[tuple[float, float, float, float, int, float, float]],
    overlap_thresh: float = NMS_OVERLAP_THRESH,
    adaptive_ratio: float = NMS_ADAPTIVE_RATIO,
) -> list[tuple[float, float, float, float, int, float, float]]:
    """
    Non-maximum suppression using both IoU and adaptive distance criteria.

    Suppress radius is adaptive per detection: max(w, h) * adaptive_ratio.
    Small templates get a small radius (avoids over-suppression of real
    clusters), large templates get a large radius (eliminates duplicates).

    Each detection is (cx, cy, w, h, rotation, scale, confidence).
    """
    if not detections:
        return []

    # Sort by confidence (descending)
    sorted_dets = sorted(detections, key=lambda d: d[6], reverse=True)

    keep: list[tuple[float, float, float, float, int, float, float]] = []

    while sorted_dets:
        best = sorted_dets.pop(0)
        keep.append(best)

        bcx, bcy, bw, bh = best[0], best[1], best[2], best[3]

        # Adaptive suppress radius: proportional to template size
        suppress_radius = max(bw, bh) * adaptive_ratio

        remaining = []
        for det in sorted_dets:
            dcx, dcy, dw, dh = det[0], det[1], det[2], det[3]

            # Distance-based suppression
            dist = math.sqrt((bcx - dcx) ** 2 + (bcy - dcy) ** 2)
            if dist < suppress_radius:
                continue  # suppress — too close to a better detection

            # Also check IoU for larger templates
            bx1, by1 = bcx - bw / 2, bcy - bh / 2
            bx2, by2 = bcx + bw / 2, bcy + bh / 2
            ox1, oy1 = dcx - dw / 2, dcy - dh / 2
            ox2, oy2 = dcx + dw / 2, dcy + dh / 2

            ix1 = max(bx1, ox1)
            iy1 = max(by1, oy1)
            ix2 = min(bx2, ox2)
            iy2 = min(by2, oy2)

            inter_w = max(0, ix2 - ix1)
            inter_h = max(0, iy2 - iy1)
            inter_area = inter_w * inter_h

            area_a = (bx2 - bx1) * (by2 - by1)
            area_b = (ox2 - ox1) * (oy2 - oy1)
            union_area = area_a + area_b - inter_area

            iou = inter_area / union_area if union_area > 0 else 0
            if iou >= overlap_thresh:
                continue  # suppress — significant overlap

            remaining.append(det)

        sorted_dets = remaining

    return keep


# ---------------------------------------------------------------------------
# Color detection from page region
# ---------------------------------------------------------------------------

def _detect_match_color(
    page_bgr: np.ndarray,
    cx: float,
    cy: float,
    w: float,
    h: float,
) -> str:
    """
    Detect the dominant color of a matched region on the page.

    Returns 'red', 'blue', or ''.
    """
    # Extract region
    x1 = max(0, int(cx - w / 2))
    y1 = max(0, int(cy - h / 2))
    x2 = min(page_bgr.shape[1], int(cx + w / 2))
    y2 = min(page_bgr.shape[0], int(cy + h / 2))

    if x2 <= x1 or y2 <= y1:
        return ""

    region = page_bgr[y1:y2, x1:x2]
    hsv = cv2.cvtColor(region, cv2.COLOR_BGR2HSV)

    # Mask colored pixels
    mask = (hsv[:, :, 1] > 50) & (hsv[:, :, 2] > 50) & (hsv[:, :, 2] < 250)
    colored_count = np.count_nonzero(mask)

    if colored_count < 10:
        return ""

    colored_hues = hsv[:, :, 0][mask]
    avg_hue = np.median(colored_hues)

    if avg_hue < 10 or avg_hue > 170:
        return "red"
    elif 100 <= avg_hue <= 130:
        return "blue"

    return ""


# ---------------------------------------------------------------------------
# Core matching logic
# ---------------------------------------------------------------------------

def match_symbols(
    pdf_path: str,
    legend_result: Optional[LegendResult] = None,
    page: Optional[int] = None,
    threshold: float = DEFAULT_THRESHOLD,
    scales: Optional[list[float]] = None,
    rotations: Optional[list[int]] = None,
    render_dpi: int = RENDER_DPI,
    detect_colors: bool = True,
) -> VisualResult:
    """
    Find and count equipment by matching legend symbols on drawing pages.

    Args:
        pdf_path: Path to the PDF file.
        legend_result: Pre-parsed legend. If None, parsed automatically.
        page: Page index to scan. If None, uses legend page.
        threshold: Minimum match confidence (0..1).
        scales: Scale factors to try. Default: [0.8, 0.9, 1.0, 1.1, 1.2].
        rotations: Rotation angles to try. Default: [0, 90, 180, 270].
        render_dpi: DPI for full page rendering. Default: 200.
        detect_colors: Whether to detect match colors.

    Returns:
        VisualResult with matches, counts, and metadata.
    """
    if scales is None:
        scales = list(SCALES)
    if rotations is None:
        rotations = list(ROTATIONS)

    # Parse legend
    if legend_result is None:
        legend_result = parse_legend(pdf_path)

    if not legend_result.items:
        return VisualResult(threshold=threshold, scales_used=scales,
                            rotations_used=rotations)

    legend_page = legend_result.page_index
    scan_page = page if page is not None else legend_page

    # Extract symbol images from legend (visual matching for ALL items,
    # not just graphical-only — text symbols also have graphical shapes)
    symbol_data = _extract_symbol_images(pdf_path, legend_result)
    symbols_with_images = [
        (idx, item, img) for idx, item, img in symbol_data
        if img is not None
    ]

    if not symbols_with_images:
        return VisualResult(
            page_index=scan_page,
            legend_page=legend_page,
            symbols_extracted=0,
            threshold=threshold,
            scales_used=scales,
            rotations_used=rotations,
        )

    # Auto-reduce DPI for very large pages to keep matching tractable
    # Template matching is O(page_pixels * template_pixels) per scale/rotation
    with pdfplumber.open(pdf_path) as pdf:
        if scan_page < len(pdf.pages):
            sp = pdf.pages[scan_page]
            page_area_pt2 = sp.width * sp.height
            # Reduce DPI so the rendered page stays under ~10000px max dimension
            # (increased from 5000 to preserve template quality on large drawings)
            max_dim_px = max(sp.width, sp.height) * render_dpi / 72.0
            if max_dim_px > 10000:
                render_dpi = int(10000 / max(sp.width, sp.height) * 72)
                render_dpi = max(render_dpi, 100)  # floor at 100 DPI

    # DPI ratio for template rescaling
    dpi_ratio = SYMBOL_DPI / render_dpi

    # Preprocess templates (with per-template threshold adjustment)
    # tuple: (idx, item, gray_template, color, adjusted_threshold)
    templates: list[tuple[int, LegendItem, np.ndarray, str, float]] = []
    for idx, item, img in symbols_with_images:
        cropped, detected_color = _preprocess_color_template(img)
        # Use legend parser's color as primary for colour isolation
        iso_color = getattr(item, "color", "") or detected_color
        gray_template = _preprocess_template(img, keep_color=iso_color)

        th, tw = gray_template.shape[:2]

        if th < MIN_SYMBOL_SIZE_PX or tw < MIN_SYMBOL_SIZE_PX:
            continue

        # Skip extremely elongated templates (aspect ratio > 8:1).
        # These are typically cable/wiring path symbols, not equipment.
        # Equipment symbols (e.g. row of 3 rectangles) can have aspect ~6:1.
        aspect = max(th, tw) / max(min(th, tw), 1)
        if aspect > 8.0:
            continue

        # Skip non-equipment templates by description keyword (T142).
        # Wiring/cable path descriptions produce massive raw detections
        # (tens of thousands) and are never real countable equipment.
        desc_lower = item.description.lower()
        if any(kw in desc_lower for kw in NON_EQUIPMENT_KEYWORDS):
            logger.debug(
                "skip_non_equip[%d] %s", idx, item.description[:40],
            )
            continue

        # Use legend parser's color as primary (it comes from PDF stroking_color
        # analysis and is more reliable than detecting color in the extracted
        # symbol image, which is typically black-on-white).
        sym_color = getattr(item, "color", "") or detected_color

        # Compute template complexity at page DPI scale
        # (templates are rendered at SYMBOL_DPI but matched at RENDER_DPI)
        page_scale_template = _scale_image(gray_template, 1.0 / dpi_ratio)
        _, tpl_binary = cv2.threshold(
            page_scale_template, 200, 255, cv2.THRESH_BINARY_INV
        )
        content_px_at_page = np.count_nonzero(tpl_binary)

        # All items are matched on full grayscale page (not colour layers),
        # so raise threshold for small/simple templates to reduce FPs.
        # Items with a known color get a lower boost because post-hoc
        # color validation already filters out wrong-color false positives.
        has_color = bool(sym_color) and detect_colors
        adj_threshold = threshold
        if has_color:
            # Color-validated items: post-hoc color check provides extra
            # FP protection, so we can afford a lower confidence threshold.
            if content_px_at_page < 100:
                adj_threshold = max(threshold, 0.83)
            # >=100 px with color: use base threshold as-is (color
            # validation provides sufficient FP protection)
        else:
            # No color info: only template confidence protects against FP,
            # so use higher thresholds.
            if content_px_at_page < 100:
                adj_threshold = max(threshold, 0.88)
            elif content_px_at_page < 300:
                adj_threshold = max(threshold, 0.85)
            elif content_px_at_page < 500:
                adj_threshold = max(threshold, 0.82)

        templates.append((idx, item, gray_template, sym_color, adj_threshold))

    # Render the target page
    page_bgr = _render_page(pdf_path, scan_page, dpi=render_dpi)
    page_gray = cv2.cvtColor(page_bgr, cv2.COLOR_BGR2GRAY)

    # Pre-compute coarse (downsampled) page for pyramid matching (T142).
    # This is done once and reused across all templates.
    ds = PYRAMID_DOWNSAMPLE
    page_gray_coarse = cv2.resize(
        page_gray,
        (int(page_gray.shape[1] * ds), int(page_gray.shape[0] * ds)),
        interpolation=cv2.INTER_AREA,
    )

    # NOTE: Color layer extraction removed from matching pipeline (T140).
    # Legend templates are black-on-white drawings that don't correlate
    # with color-separated raster layers. Matching is now ALWAYS done on
    # full grayscale; color info from the legend parser is used for
    # post-hoc validation only (via _detect_match_color on page_bgr).

    # Build exclusion zones (in PDF pt coordinates) and extract chars for FP filter
    exclusion_zones: list[tuple[str, tuple[float, float, float, float]]] = []
    drawing_chars: list[dict] = []  # chars in drawing area for FP text overlap check
    with pdfplumber.open(pdf_path) as pdf:
        if scan_page < len(pdf.pages):
            pp_page = pdf.pages[scan_page]
            pp_lines = pp_page.lines or []
            lb = legend_result.legend_bbox if scan_page == legend_page else None
            exclusion_zones = _build_exclusion_zones(pp_page, pp_lines, lb)
            page_width_pt = pp_page.width
            page_height_pt = pp_page.height

            # Extract chars in drawing area (outside exclusion zones) for T133 FP filter
            for ch in (pp_page.chars or []):
                ch_x = float(ch.get("x0", 0))
                ch_y = float(ch.get("top", 0))
                if not _pt_excluded(ch_x, ch_y, exclusion_zones):
                    drawing_chars.append(ch)

    # Conversion factor: pixel → PDF pt
    px_per_pt = render_dpi / 72.0

    all_matches: list[VisualMatch] = []
    counts: dict[int, int] = {}
    descriptions: dict[int, str] = {}

    for idx, item, template, sym_color, adj_threshold in templates:
        descriptions[idx] = item.description

        # Determine expected color for this item.
        # Use the legend parser's color (from PDF stroking_color analysis)
        # as primary, falling back to template image color detection.
        expected_color = getattr(item, "color", "") or sym_color

        # ALWAYS match on full grayscale page. Legend templates are
        # black-on-white line drawings which don't correlate with
        # color-separated layers (where only colored ink is present).
        # Matching on color layers yields max ~0.66 confidence vs ~0.96
        # on the full page for the same template. Color information from
        # the legend parser is used for POST-HOC validation only.
        match_target = page_gray

        # Match template on page (using per-template adjusted threshold)
        raw_detections = _match_template_multi(
            match_target, template, adj_threshold, scales, rotations, dpi_ratio,
            page_gray_coarse=page_gray_coarse,
        )

        # Pre-filter: remove raw detections inside exclusion zones BEFORE
        # NMS (T141). Without this, high-confidence legend-zone matches
        # dominate NMS selection and suppress real drawing-area detections.
        drawing_detections = []
        n_pre_excluded = 0
        for det in raw_detections:
            cx_px, cy_px = det[0], det[1]
            x_pt = cx_px / px_per_pt
            y_pt = cy_px / px_per_pt
            if _pt_excluded(x_pt, y_pt, exclusion_zones):
                n_pre_excluded += 1
            else:
                drawing_detections.append(det)

        # Apply NMS only on drawing-area detections
        filtered = _nms(drawing_detections, NMS_OVERLAP_THRESH)

        # --- Diagnostic counters (T141) ---
        n_raw = len(raw_detections)
        n_drawing_raw = len(drawing_detections)
        n_nms = len(filtered)
        n_excluded = 0
        n_fp = 0
        n_color_reject = 0
        excluded_zone_names: dict[str, int] = {}

        # Convert pixel coords to PDF pt and filter exclusion zones
        valid_matches: list[VisualMatch] = []
        for cx_px, cy_px, w_px, h_px, rot, scale, conf in filtered:
            # Convert to PDF pt
            x_pt = cx_px / px_per_pt
            y_pt = cy_px / px_per_pt

            # Check exclusion zones
            zone_hit = _pt_excluded(
                x_pt, y_pt, exclusion_zones, return_zone_name=True,
            )
            if zone_hit:
                n_excluded += 1
                excluded_zone_names[zone_hit] = (
                    excluded_zone_names.get(zone_hit, 0) + 1
                )
                continue

            # Context-aware false positive filter (T133)
            if _is_false_positive(
                match_target, cx_px, cy_px, w_px, h_px,
                drawing_chars, px_per_pt,
            ):
                n_fp += 1
                continue

            # Post-hoc color validation: since we match on full grayscale,
            # use color info from legend parser for validation only.
            # After finding a match, check if the match region has the
            # expected color using _detect_match_color() on page_bgr.
            # Only reject if a DIFFERENT color is confidently detected.
            match_color = ""
            if detect_colors:
                detected_color = _detect_match_color(
                    page_bgr, cx_px, cy_px, w_px, h_px
                )
                if expected_color:
                    # Template has a known color from legend parser.
                    # Reject ONLY if a different color is confidently
                    # detected (e.g. template is blue but match region
                    # is red). Accept when no color detected (the
                    # symbol may be drawn in thin lines with few
                    # saturated pixels) or when colors match.
                    if detected_color and detected_color != expected_color:
                        n_color_reject += 1
                        continue
                    match_color = expected_color
                else:
                    # No expected color — use whatever was detected
                    match_color = detected_color or sym_color

            vm = VisualMatch(
                symbol_index=idx,
                description=item.description,
                x=round(x_pt, 1),
                y=round(y_pt, 1),
                confidence=round(conf, 3),
                scale=scale,
                rotation=rot,
                color=match_color,
                page_index=scan_page,
            )
            valid_matches.append(vm)

        counts[idx] = len(valid_matches)
        all_matches.extend(valid_matches)

        # --- Emit diagnostic log (T141) ---
        n_valid = len(valid_matches)
        zones_str = (
            " zones=" + ",".join(f"{k}:{v}" for k, v in excluded_zone_names.items())
            if excluded_zone_names else ""
        )
        logger.debug(
            "match_symbols[%d] %s: raw=%d pre_excl=%d drawing=%d nms=%d "
            "excl=%d fp=%d color_rej=%d valid=%d thr=%.2f%s",
            idx, item.description[:40], n_raw, n_pre_excluded,
            n_drawing_raw, n_nms,
            n_excluded, n_fp, n_color_reject, n_valid,
            adj_threshold, zones_str,
        )

    return VisualResult(
        matches=all_matches,
        counts=counts,
        descriptions=descriptions,
        page_index=scan_page,
        legend_page=legend_page,
        symbols_extracted=len(templates),
        render_dpi=render_dpi,
        threshold=threshold,
        scales_used=scales,
        rotations_used=rotations,
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
        print("Usage: python pdf_count_visual.py <path.pdf> [options]")
        print()
        print("Options:")
        print("  --page N          Scan specific page (0-based)")
        print("  --threshold T     Match confidence threshold (default: 0.65)")
        print("  --dpi N           Render DPI (default: 200)")
        print("  --no-color        Disable color detection")
        print("  --save-debug DIR  Save debug images to directory")
        print("  --detail          Show individual match positions")
        print("  --verbose         Show per-template filter diagnostics")
        sys.exit(1)

    pdf_path = args[0]
    scan_page = None
    threshold = DEFAULT_THRESHOLD
    render_dpi = RENDER_DPI
    detect_colors = True
    debug_dir = None
    show_detail = "--detail" in args or "-d" in args
    verbose = "--verbose" in args or "-v" in args

    # Parse options
    i = 1
    while i < len(args):
        if args[i] == "--page" and i + 1 < len(args):
            try:
                scan_page = int(args[i + 1])
            except ValueError:
                print(f"Invalid page number: {args[i + 1]}")
                sys.exit(1)
            i += 2
        elif args[i] == "--threshold" and i + 1 < len(args):
            try:
                threshold = float(args[i + 1])
            except ValueError:
                print(f"Invalid threshold: {args[i + 1]}")
                sys.exit(1)
            i += 2
        elif args[i] == "--dpi" and i + 1 < len(args):
            try:
                render_dpi = int(args[i + 1])
            except ValueError:
                print(f"Invalid DPI: {args[i + 1]}")
                sys.exit(1)
            i += 2
        elif args[i] == "--no-color":
            detect_colors = False
            i += 1
        elif args[i] == "--save-debug" and i + 1 < len(args):
            debug_dir = args[i + 1]
            i += 2
        else:
            i += 1

    print(f"Visual symbol matching: {pdf_path}")
    print()

    # Parse legend
    legend = parse_legend(pdf_path)
    if not legend.items:
        print("No legend found — cannot match symbols.")
        sys.exit(0)

    print(f"Legend: {len(legend.items)} items on page {legend.page_index + 1}")
    for it in legend.items:
        sym_label = f"'{it.symbol}'" if it.symbol else "(graphic)"
        print(f"  {sym_label}: {it.description[:60]}")
    print()

    # Enable verbose diagnostics if requested
    if verbose:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(handler)
        logger.setLevel(logging.DEBUG)

    # Run matching
    print(f"Matching with threshold={threshold}, DPI={render_dpi}")
    print(f"Scales: {SCALES}")
    print(f"Rotations: {ROTATIONS}")
    print()

    result = match_symbols(
        pdf_path,
        legend_result=legend,
        page=scan_page,
        threshold=threshold,
        render_dpi=render_dpi,
        detect_colors=detect_colors,
    )

    print(f"Symbols extracted from legend: {result.symbols_extracted}")
    print(f"Scanning page: {result.page_index + 1}")
    print()

    # Results table
    if result.counts:
        print(f"{'#':<4s} {'Count':>6s} {'Color':>8s}  Description")
        print("-" * 70)
        total = 0
        for idx in sorted(result.counts.keys()):
            cnt = result.counts[idx]
            desc = result.descriptions.get(idx, "?")[:50]
            # Get dominant color for this symbol's matches
            sym_matches = [m for m in result.matches if m.symbol_index == idx]
            colors = sorted({m.color for m in sym_matches if m.color})
            color_str = "/".join(colors) if colors else "—"
            total += cnt
            print(f"  {idx:<3d} {cnt:>5d}  {color_str:>7s}  {desc}")

        print("-" * 70)
        print(f"  {'':3s} {total:>5d}  TOTAL")
    else:
        print("No symbols matched on the page.")

    # Detail view
    if show_detail and result.matches:
        print()
        print("=== Match Details ===")
        for i, m in enumerate(result.matches, 1):
            desc_short = m.description[:30]
            color_str = f" [{m.color}]" if m.color else ""
            print(f"  {i:3d}. #{m.symbol_index} ({m.confidence:.2f})"
                  f" at ({m.x:.0f}, {m.y:.0f})"
                  f" s={m.scale} r={m.rotation}{color_str}"
                  f"  {desc_short}")

    # Save debug image if requested
    if debug_dir and result.matches:
        os.makedirs(debug_dir, exist_ok=True)
        _save_debug_image(
            pdf_path, result, render_dpi, debug_dir,
        )
        print(f"\nDebug images saved to {debug_dir}/")


def _save_debug_image(
    pdf_path: str,
    result: VisualResult,
    dpi: int,
    output_dir: str,
) -> None:
    """Save a debug image showing matched positions on the page."""
    page_bgr = _render_page(pdf_path, result.page_index, dpi=dpi)
    px_per_pt = dpi / 72.0

    # Color map for symbols
    color_map = [
        (0, 0, 255),    # red
        (255, 0, 0),    # blue
        (0, 255, 0),    # green
        (0, 255, 255),  # yellow
        (255, 0, 255),  # magenta
        (255, 128, 0),  # orange
        (128, 0, 255),  # purple
        (0, 128, 255),  # light blue
    ]

    for m in result.matches:
        cx = int(m.x * px_per_pt)
        cy = int(m.y * px_per_pt)
        color_idx = m.symbol_index % len(color_map)
        color = color_map[color_idx]

        # Draw circle and label
        radius = 15
        cv2.circle(page_bgr, (cx, cy), radius, color, 2)
        label = f"#{m.symbol_index}:{m.confidence:.2f}"
        cv2.putText(
            page_bgr, label,
            (cx + radius + 2, cy + 5),
            cv2.FONT_HERSHEY_SIMPLEX, 0.4, color, 1,
        )

    out_path = os.path.join(
        output_dir,
        f"debug_visual_p{result.page_index + 1}.png"
    )
    cv2.imwrite(out_path, page_bgr)
    print(f"  Saved: {out_path}")

    # Also save extracted templates
    symbol_data = _extract_symbol_images(
        pdf_path,
        parse_legend(pdf_path),
    )
    for idx, item, img in symbol_data:
        if img is not None:
            tpl_path = os.path.join(output_dir, f"template_{idx}.png")
            cv2.imwrite(tpl_path, img)


if __name__ == "__main__":
    main()
