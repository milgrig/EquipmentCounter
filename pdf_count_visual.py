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
from typing import Callable, Optional

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
    disambiguated_indices: set[int] = field(default_factory=set)  # symbol indices resolved by text disambiguation (S021)
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

# Shape verification (visual anti-flood)
# Verifies that the matched ROI actually resembles the template's binary shape.
SHAPE_VERIFY_MIN_RECALL = 0.22
SHAPE_VERIFY_MIN_PRECISION = 0.08
# Hard cap for simple one-letter compound markers (e.g. 1А) to avoid
# catastrophic flood when a tiny template matches repetitive background.
SIMPLE_COMPOUND_MAX_MATCHES = 80
# Candidate-first detection for simple one-letter compounds.
SIMPLE_COMPOUND_USE_CANDIDATES = True

# Connected-component candidate proposal (pixel units at render DPI).
CANDIDATE_BIN_THRESH = 205
CANDIDATE_LINE_KERNEL = 35
CANDIDATE_MIN_AREA_PX = 28
CANDIDATE_MAX_AREA_PX = 6000
CANDIDATE_MIN_SIDE_PX = 6
CANDIDATE_MAX_SIDE_PX = 180
CANDIDATE_MAX_ASPECT = 4.0
CANDIDATE_PAD_PX = 2
CANDIDATE_MAX_COUNT = 6000

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
# Shape verification filter (anti-flood for repetitive backgrounds)
# ---------------------------------------------------------------------------

def _is_shape_mismatch(
    page_gray: np.ndarray,
    template_gray: np.ndarray,
    cx_px: float,
    cy_px: float,
    w_px: float,
    h_px: float,
    rot: int,
    scale: float,
    dpi_ratio: float,
    variant_cache: dict[tuple[int, float], np.ndarray],
    min_recall: float = SHAPE_VERIFY_MIN_RECALL,
    min_precision: float = SHAPE_VERIFY_MIN_PRECISION,
) -> bool:
    """Check whether a matched ROI resembles the template's binary shape.

    Returns True when the ROI shape is too dissimilar from the corresponding
    rotated/scaled template variant (likely a false-positive flood match).
    """
    key = (rot, round(scale, 3))
    tpl_bin = variant_cache.get(key)
    if tpl_bin is None:
        rotated = _rotate_image(template_gray, rot)
        effective_scale = scale / dpi_ratio
        scaled = _scale_image(rotated, effective_scale)
        _, tpl_bin = cv2.threshold(scaled, 200, 255, cv2.THRESH_BINARY_INV)
        variant_cache[key] = tpl_bin

    tpl_h, tpl_w = tpl_bin.shape[:2]
    if tpl_h < 4 or tpl_w < 4:
        return False

    x1 = max(0, int(cx_px - w_px / 2))
    y1 = max(0, int(cy_px - h_px / 2))
    x2 = min(page_gray.shape[1], int(cx_px + w_px / 2))
    y2 = min(page_gray.shape[0], int(cy_px + h_px / 2))
    if x2 <= x1 or y2 <= y1:
        return True

    roi = page_gray[y1:y2, x1:x2]
    if roi.size == 0:
        return True

    roi_norm = cv2.resize(roi, (tpl_w, tpl_h), interpolation=cv2.INTER_LINEAR)
    _, roi_bin = cv2.threshold(roi_norm, 200, 255, cv2.THRESH_BINARY_INV)

    # Mild dilation compensates small alignment offsets.
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    tpl_d = cv2.dilate(tpl_bin, k, iterations=1)
    roi_d = cv2.dilate(roi_bin, k, iterations=1)

    tpl_fg = np.count_nonzero(tpl_d)
    roi_fg = np.count_nonzero(roi_d)
    if tpl_fg < 12:
        return False

    overlap = np.count_nonzero(cv2.bitwise_and(tpl_d, roi_d))
    recall = overlap / tpl_fg
    precision = overlap / max(roi_fg, 1)

    return recall < min_recall or precision < min_precision


def _build_component_candidates(
    page_gray: np.ndarray,
    exclusion_zones: list[tuple[str, tuple[float, float, float, float]]],
    px_per_pt: float,
) -> list[tuple[int, int, int, int]]:
    """Propose symbol candidates from connected components on drawing pixels."""
    _, binary = cv2.threshold(page_gray, CANDIDATE_BIN_THRESH, 255, cv2.THRESH_BINARY_INV)

    # Remove long cable/grid lines so components are closer to symbol blobs.
    h_k = cv2.getStructuringElement(cv2.MORPH_RECT, (CANDIDATE_LINE_KERNEL, 1))
    v_k = cv2.getStructuringElement(cv2.MORPH_RECT, (1, CANDIDATE_LINE_KERNEL))
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_k)
    v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_k)
    lines = cv2.bitwise_or(h_lines, v_lines)
    fg = cv2.bitwise_and(binary, cv2.bitwise_not(lines))

    # Light cleanup of isolated noise pixels.
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    fg = cv2.morphologyEx(fg, cv2.MORPH_OPEN, k)

    num_labels, _labels, stats, _ = cv2.connectedComponentsWithStats(fg, connectivity=8)
    candidates: list[tuple[int, int, int, int, int]] = []
    img_h, img_w = page_gray.shape[:2]
    for lbl in range(1, num_labels):
        x = int(stats[lbl, cv2.CC_STAT_LEFT])
        y = int(stats[lbl, cv2.CC_STAT_TOP])
        w = int(stats[lbl, cv2.CC_STAT_WIDTH])
        h = int(stats[lbl, cv2.CC_STAT_HEIGHT])
        area = int(stats[lbl, cv2.CC_STAT_AREA])

        if area < CANDIDATE_MIN_AREA_PX or area > CANDIDATE_MAX_AREA_PX:
            continue
        if w < CANDIDATE_MIN_SIDE_PX or h < CANDIDATE_MIN_SIDE_PX:
            continue
        if w > CANDIDATE_MAX_SIDE_PX or h > CANDIDATE_MAX_SIDE_PX:
            continue
        if max(w, h) / max(min(w, h), 1) > CANDIDATE_MAX_ASPECT:
            continue

        cx_pt = (x + w / 2.0) / px_per_pt
        cy_pt = (y + h / 2.0) / px_per_pt
        if _pt_excluded(cx_pt, cy_pt, exclusion_zones):
            continue

        x0 = max(0, x - CANDIDATE_PAD_PX)
        y0 = max(0, y - CANDIDATE_PAD_PX)
        x1 = min(img_w, x + w + CANDIDATE_PAD_PX)
        y1 = min(img_h, y + h + CANDIDATE_PAD_PX)
        candidates.append((x0, y0, x1 - x0, y1 - y0, area))

    # Keep strongest (largest) candidates if sheet is extremely dense.
    if len(candidates) > CANDIDATE_MAX_COUNT:
        candidates.sort(key=lambda c: c[4], reverse=True)
        candidates = candidates[:CANDIDATE_MAX_COUNT]

    return [(x, y, w, h) for x, y, w, h, _a in candidates]


def _match_template_on_candidates(
    page_gray: np.ndarray,
    template_gray: np.ndarray,
    candidates: list[tuple[int, int, int, int]],
    threshold: float,
    scales: list[float],
    rotations: list[int],
    dpi_ratio: float,
) -> list[tuple[float, float, float, float, int, float, float]]:
    """Classify each candidate ROI by template similarity (best variant wins)."""
    if not candidates:
        return []

    best_by_candidate: dict[int, tuple[float, float, float, float, int, float, float]] = {}
    variant_cache: dict[tuple[int, float], np.ndarray] = {}

    for rot in rotations:
        for scale in scales:
            key = (rot, round(scale, 3))
            tpl = variant_cache.get(key)
            if tpl is None:
                rotated = _rotate_image(template_gray, rot)
                effective_scale = scale / dpi_ratio
                tpl = _scale_image(rotated, effective_scale)
                variant_cache[key] = tpl

            th, tw = tpl.shape[:2]
            if th < 5 or tw < 5:
                continue

            for ci, (x, y, w, h) in enumerate(candidates):
                roi = page_gray[y:y + h, x:x + w]
                if roi.size == 0:
                    continue

                roi_norm = cv2.resize(roi, (tw, th), interpolation=cv2.INTER_LINEAR)
                conf = float(cv2.matchTemplate(
                    roi_norm, tpl, cv2.TM_CCOEFF_NORMED
                )[0, 0])
                if conf < threshold:
                    continue

                cx = x + w / 2.0
                cy = y + h / 2.0
                det = (cx, cy, float(w), float(h), rot, scale, conf)
                prev = best_by_candidate.get(ci)
                if prev is None or conf > prev[6]:
                    best_by_candidate[ci] = det

    return list(best_by_candidate.values())


def _auto_candidate_threshold(
    candidate_detections: list[tuple[float, float, float, float, int, float, float]],
    base_threshold: float,
) -> float:
    """Pick adaptive threshold for candidate-classifier confidence.

    Candidate scores are usually lower than dense sliding-template scores.
    We keep only top-confidence tail (roughly top 3%) with clamps.
    """
    if not candidate_detections:
        return base_threshold

    confs = np.array([d[6] for d in candidate_detections], dtype=np.float32)
    if confs.size < 20:
        return base_threshold

    p97 = float(np.percentile(confs, 97))
    p99 = float(np.percentile(confs, 99))
    adaptive = max(base_threshold, p97, p99 - 0.015)
    return float(min(0.62, max(0.30, adaptive)))


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
            # For the first item: legend_bbox[1] is the top of the entire
            # table, which includes the header row ("Обозначение" etc.).
            # Use the item's own bbox top minus a small margin instead,
            # so the header text is excluded from the symbol cell crop.
            item_top = items[i].bbox[1] if items[i].bbox != (0, 0, 0, 0) else yc
            row_top = max(legend_bbox[1], item_top - 5)
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

        # Guard: if clip is degenerate (too small), skip
        if clip.is_empty or clip.width < 2 or clip.height < 2:
            results.append((item_idx, item, None))
            continue

        # Render at high DPI
        mat = fitz.Matrix(zoom, zoom)
        pix = fitz_page.get_pixmap(matrix=mat, clip=clip, alpha=False)

        # Guard: empty pixmap (can happen with tiny clips at high zoom)
        if pix.width < 1 or pix.height < 1 or len(pix.samples) == 0:
            results.append((item_idx, item, None))
            continue

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
    min_radius_px: float = 60.0,
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

        # Adaptive suppress radius: proportional to template size.
        # Floor at configurable minimum radius to prevent duplicate detections
        # of small templates at closely-spaced fixtures (S021).
        suppress_radius = max(max(bw, bh) * adaptive_ratio, min_radius_px)

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
# Pictogram detection (text-based)  — T149
# ---------------------------------------------------------------------------

# Pictograms are small sticker-like labels on the drawing (e.g. "ВЫХОД" / EXIT).
# They don't have a separate legend item or a visual template — they are pure
# text labels.  We detect them by searching for specific text words on the
# drawing page outside exclusion zones.
#
# "ДОХЫВ" is "ВЫХОД" read in reversed column order when the label is placed
# vertically (rotated 90°).  pdfplumber extracts the chars column-by-column,
# so the word appears reversed.

PICTOGRAM_KEYWORDS: list[tuple[str, str]] = [
    # (search_word, canonical_name)
    ("ВЫХОД", "Пиктограмма Выход"),
    ("ДОХЫВ", "Пиктограмма Выход"),   # vertical/reversed "ВЫХОД"
]


@dataclass
class PictogramResult:
    """Result of pictogram detection on a drawing page."""
    counts: dict[str, int] = field(default_factory=dict)   # canonical_name → count
    positions: list[tuple[str, float, float]] = field(default_factory=list)
    # (canonical_name, x_pt, y_pt)
    page_index: int = 0


def detect_pictograms(
    pdf_path: str,
    legend_result: Optional[LegendResult] = None,
    page: Optional[int] = None,
) -> PictogramResult:
    """
    Detect pictogram labels on a drawing page by searching for known
    text keywords (e.g. "ВЫХОД") outside exclusion zones.

    Args:
        pdf_path: Path to the PDF file.
        legend_result: Pre-parsed legend (for exclusion zone bbox).
        page: Page index to scan. If None, uses legend page.

    Returns:
        PictogramResult with counts and positions per canonical name.
    """
    if legend_result is None:
        legend_result = parse_legend(pdf_path)

    legend_page = legend_result.page_index
    scan_page = page if page is not None else legend_page

    counts: dict[str, int] = defaultdict(int)
    positions: list[tuple[str, float, float]] = []

    with pdfplumber.open(pdf_path) as pdf:
        if scan_page >= len(pdf.pages):
            return PictogramResult(page_index=scan_page)

        pp_page = pdf.pages[scan_page]
        pp_lines = pp_page.lines or []
        lb = legend_result.legend_bbox if scan_page == legend_page else None
        exclusion_zones = _build_exclusion_zones(pp_page, pp_lines, lb)

        # Extract words from the page
        words = pp_page.extract_words(
            keep_blank_chars=False, use_text_flow=False,
        )

        for w in words:
            w_text = w.get("text", "").strip().strip('"').strip("«»")
            for keyword, canonical in PICTOGRAM_KEYWORDS:
                if w_text != keyword:
                    continue

                # Use center of the word bbox
                x_pt = (float(w["x0"]) + float(w["x1"])) / 2
                y_pt = (float(w["top"]) + float(w["bottom"])) / 2

                # Check exclusion zones
                if _pt_excluded(x_pt, y_pt, exclusion_zones):
                    continue

                counts[canonical] += 1
                positions.append((canonical, round(x_pt, 1), round(y_pt, 1)))
                break  # don't double-count same word

    logger.debug(
        "detect_pictograms: page=%d  %s",
        scan_page,
        "  ".join(f"{k}={v}" for k, v in counts.items()) or "none found",
    )

    return PictogramResult(
        counts=dict(counts),
        positions=positions,
        page_index=scan_page,
    )


# ---------------------------------------------------------------------------
# Confusable template disambiguation (S021)
# ---------------------------------------------------------------------------

# Tokens that distinguish visually identical template variants.
# If two legend descriptions differ only by one of these tokens,
# they are "confusable" — template matching cannot separate them,
# but reading nearby text on the drawing can.
_DISAMBIG_TOKENS = {
    "TH", "EM", "EX", "IP65", "IP66", "AT", "FP", "DLW",
    # Russian operational variants for visually identical symbols/templates.
    "АВАРИЙНОГО", "АВАРИЙНОЕ", "АВАРИЙНЫЙ",
    "РАБОЧЕГО", "РАБОЧЕЕ", "РАБОЧИЙ",
}


def _find_confusable_groups(
    templates: list[tuple[int, object, object, str, float]],
) -> list[tuple[int, int, set[str], set[str]]]:
    """Identify pairs of templates whose descriptions differ only by
    a few distinguishing tokens (e.g. 'TH', 'Ex').

    Returns list of (idx_a, idx_b, tokens_a, tokens_b) where tokens_X
    are the words present in description X but not in the other.
    """
    import re as _re

    groups: list[tuple[int, int, set[str], set[str]]] = []
    n = len(templates)
    for i in range(n):
        for j in range(i + 1, n):
            idx_a = templates[i][0]
            idx_b = templates[j][0]
            desc_a = templates[i][1].description or ""
            desc_b = templates[j][1].description or ""
            # Tokenize: split into words
            words_a = set(
                tok.upper()
                for tok in _re.findall(r"[A-Za-zА-Яа-яЁё0-9.]+", desc_a)
            )
            words_b = set(
                tok.upper()
                for tok in _re.findall(r"[A-Za-zА-Яа-яЁё0-9.]+", desc_b)
            )
            only_a = words_a - words_b
            only_b = words_b - words_a
            # They are confusable if they share most words and differ
            # only by known disambiguation tokens
            diff_all = only_a | only_b
            if not diff_all:
                continue  # identical descriptions — not our problem
            common = words_a & words_b
            # For short labels (e.g. "Щит аварийного освещения" vs
            # "Щит рабочего освещения"), allow 2 shared words.
            if len(common) < 3:
                if not (len(common) >= 2 and max(len(words_a), len(words_b)) <= 4):
                    continue  # too different to be confusable
            # Require high word overlap. For long descriptions keep a strict
            # threshold; for short labels allow a slightly lower overlap.
            # This still prevents false grouping of different fixture
            # families (e.g. SLICK.PRS 50 vs ARCTIC.OPL 1200), while letting
            # short variants like "щит аварийного/рабочего освещения" pair.
            max_len = max(len(words_a), len(words_b))
            required_overlap = 0.80 if max_len > 4 else 0.60
            if len(common) / max_len < required_overlap:
                continue
            # All differing tokens must be known disambig tokens or
            # trivial (empty set on one side).  If the difference
            # includes unknown words (e.g. different model numbers
            # like '50' vs '30'), they are genuinely different items.
            non_disambig = diff_all - _DISAMBIG_TOKENS
            if non_disambig:
                continue  # different model/size — not confusable
            # Check that at least one side has a known disambig token
            if diff_all & _DISAMBIG_TOKENS:
                groups.append((idx_a, idx_b, only_a, only_b))
    return groups


def _collect_text_near_point(
    x_pt: float,
    y_pt: float,
    drawing_chars: list[dict],
    radius_pt: float = 40.0,
) -> str:
    """Collect all text characters within radius_pt of (x_pt, y_pt).

    Returns concatenated text (uppercase) for keyword search.
    """
    chars = []
    for ch in drawing_chars:
        cx = float(ch.get("x0", 0))
        cy = float(ch.get("top", 0))
        if abs(cx - x_pt) <= radius_pt and abs(cy - y_pt) <= radius_pt:
            chars.append(ch.get("text", ""))
    return "".join(chars).upper()


def _disambiguate_matches(
    matches: list,  # list[VisualMatch]
    confusable_groups: list[tuple[int, int, set[str], set[str]]],
    drawing_chars: list[dict],
) -> list:
    """For matches at locations where confusable templates overlap,
    assign each match to the correct template variant based on
    nearby text keywords.

    Returns updated list of matches with duplicates removed.
    """
    if not confusable_groups:
        return matches

    # Build lookup: symbol_index → list of (partner_index, my_tokens, partner_tokens)
    confusable_map: dict[int, list[tuple[int, set[str], set[str]]]] = {}
    for idx_a, idx_b, tok_a, tok_b in confusable_groups:
        confusable_map.setdefault(idx_a, []).append((idx_b, tok_a, tok_b))
        confusable_map.setdefault(idx_b, []).append((idx_a, tok_b, tok_a))

    CLUSTER_RADIUS = 25.0  # pt — matches within this radius are at "same location"

    # Group matches by physical location (cluster nearby matches)
    used = [False] * len(matches)
    clusters: list[list[int]] = []  # list of match-index lists
    for i, m in enumerate(matches):
        if used[i]:
            continue
        cluster = [i]
        used[i] = True
        for j in range(i + 1, len(matches)):
            if used[j]:
                continue
            dist = math.sqrt((m.x - matches[j].x) ** 2 +
                             (m.y - matches[j].y) ** 2)
            if dist < CLUSTER_RADIUS:
                cluster.append(j)
                used[j] = True
        clusters.append(cluster)

    result = []
    for cluster in clusters:
        cluster_matches = [matches[i] for i in cluster]
        # Check if cluster contains a confusable pair
        sym_indices = {m.symbol_index for m in cluster_matches}
        has_confusable = False
        for si in sym_indices:
            if si in confusable_map:
                for partner, _, _ in confusable_map[si]:
                    if partner in sym_indices:
                        has_confusable = True
                        break
            if has_confusable:
                break

        if not has_confusable or len(cluster_matches) == 1:
            # No disambiguation needed — keep all
            result.extend(cluster_matches)
            continue

        # Read text near this location
        cx = sum(m.x for m in cluster_matches) / len(cluster_matches)
        cy = sum(m.y for m in cluster_matches) / len(cluster_matches)
        nearby_text = _collect_text_near_point(cx, cy, drawing_chars)

        # For each confusable pair in this cluster, decide which template wins
        winners: dict[int, float] = {}  # symbol_index → best confidence
        for m in cluster_matches:
            if m.symbol_index not in confusable_map:
                # Not confusable — keep as-is
                winners[m.symbol_index] = max(
                    winners.get(m.symbol_index, 0), m.confidence
                )
                continue

            # Check which variant's distinguishing tokens appear in text
            for partner, my_tokens, partner_tokens in confusable_map[m.symbol_index]:
                if partner not in sym_indices:
                    continue
                # Count how many of MY distinguishing tokens appear nearby
                my_hits = sum(1 for t in my_tokens if t.upper() in nearby_text)
                partner_hits = sum(1 for t in partner_tokens if t.upper() in nearby_text)
                if my_hits > partner_hits:
                    # My tokens found — I'm the right variant
                    winners[m.symbol_index] = max(
                        winners.get(m.symbol_index, 0), m.confidence
                    )
                elif partner_hits > my_hits:
                    # Partner tokens found — partner wins, skip me
                    pass
                else:
                    # Tie — keep higher confidence
                    winners[m.symbol_index] = max(
                        winners.get(m.symbol_index, 0), m.confidence
                    )

        # Build output: one match per winning symbol at this location
        for m in cluster_matches:
            if m.symbol_index in winners:
                result.append(m)
                # Remove from winners so we don't add duplicates of same symbol
                del winners[m.symbol_index]

    logger.debug(
        "disambiguate: %d matches in, %d out, %d confusable groups",
        len(matches), len(result), len(confusable_groups),
    )
    return result


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
    progress_callback: Optional[Callable] = None,
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
    simple_compound_candidates: list[tuple[int, int, int, int]] = []
    if SIMPLE_COMPOUND_USE_CANDIDATES:
        simple_compound_candidates = _build_component_candidates(
            page_gray, exclusion_zones, px_per_pt,
        )
        logger.debug(
            "candidate_proposal: %d component candidates",
            len(simple_compound_candidates),
        )

    all_matches: list[VisualMatch] = []
    counts: dict[int, int] = {}
    descriptions: dict[int, str] = {}

    for idx, item, template, sym_color, adj_threshold in templates:
        descriptions[idx] = item.description
        sym_text = (item.symbol or "").strip()
        is_simple_compound = bool(re.fullmatch(r"\d{1,2}[А-Яа-яЁё]", sym_text))

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

        # Match template on page. For tiny one-letter compounds use
        # candidate-first path: propose components, then classify each ROI.
        if is_simple_compound and SIMPLE_COMPOUND_USE_CANDIDATES:
            # Candidate classifier scores are lower than dense sliding
            # template matching scores (ROI normalized to candidate box),
            # so we calibrate threshold from score distribution per symbol.
            all_candidate_detections = _match_template_on_candidates(
                match_target, template, simple_compound_candidates,
                0.0, scales, rotations, dpi_ratio,
            )
            candidate_thresh = _auto_candidate_threshold(
                all_candidate_detections, base_threshold=0.32,
            )
            raw_detections = [
                d for d in all_candidate_detections if d[6] >= candidate_thresh
            ]
            logger.debug(
                "candidate_calib[%d] %s: all=%d thr=%.3f kept=%d",
                idx, sym_text, len(all_candidate_detections),
                candidate_thresh, len(raw_detections),
            )
        else:
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
        nms_ratio = 1.2 if is_simple_compound else NMS_ADAPTIVE_RATIO
        nms_min_radius = 120.0 if is_simple_compound else 60.0
        filtered = _nms(
            drawing_detections, NMS_OVERLAP_THRESH,
            adaptive_ratio=nms_ratio, min_radius_px=nms_min_radius,
        )

        # --- Diagnostic counters (T141) ---
        n_raw = len(raw_detections)
        n_drawing_raw = len(drawing_detections)
        n_nms = len(filtered)
        n_excluded = 0
        n_fp = 0
        n_shape_reject = 0
        n_color_reject = 0
        excluded_zone_names: dict[str, int] = {}
        variant_cache: dict[tuple[int, float], np.ndarray] = {}

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

            # Shape verification: reject repetitive-background matches that
            # pass template score but don't resemble the actual symbol form.
            if _is_shape_mismatch(
                match_target, template, cx_px, cy_px, w_px, h_px,
                rot, scale, dpi_ratio, variant_cache,
                min_recall=(0.35 if is_simple_compound else SHAPE_VERIFY_MIN_RECALL),
                min_precision=(0.12 if is_simple_compound else SHAPE_VERIFY_MIN_PRECISION),
            ):
                n_shape_reject += 1
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

        # Spatial dedup for tiny one-letter compounds (e.g. 1А).
        # Keep only the strongest match per local grid cell so repetitive
        # background fragments cannot explode the count.
        if is_simple_compound and valid_matches:
            cell_pt = 120.0  # coarse spatial cell for noisy tiny templates
            best_by_cell: dict[tuple[int, int], VisualMatch] = {}
            for m in valid_matches:
                cell = (int(m.x // cell_pt), int(m.y // cell_pt))
                prev = best_by_cell.get(cell)
                if prev is None or m.confidence > prev.confidence:
                    best_by_cell[cell] = m
            if len(best_by_cell) < len(valid_matches):
                logger.debug(
                    "match_symbols[%d] spatial_dedup: %d -> %d",
                    idx, len(valid_matches), len(best_by_cell),
                )
            valid_matches = list(best_by_cell.values())

        # Flood guard for simple one-letter compounds (e.g. 1А, 2А):
        # these tiny templates are prone to massive accidental repeats.
        if (re.fullmatch(r"\d{1,2}[А-Яа-яЁё]", sym_text)
                and len(valid_matches) > SIMPLE_COMPOUND_MAX_MATCHES):
            valid_matches.sort(key=lambda m: m.confidence, reverse=True)
            dropped = len(valid_matches) - SIMPLE_COMPOUND_MAX_MATCHES
            valid_matches = valid_matches[:SIMPLE_COMPOUND_MAX_MATCHES]
            logger.warning(
                "match_symbols[%d] flood_guard: symbol=%s limited %d->%d",
                idx, sym_text, len(valid_matches) + dropped, len(valid_matches),
            )

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
            "excl=%d fp=%d shape_rej=%d color_rej=%d valid=%d thr=%.2f%s",
            idx, item.description[:40], n_raw, n_pre_excluded,
            n_drawing_raw, n_nms,
            n_excluded, n_fp, n_shape_reject, n_color_reject, n_valid,
            adj_threshold, zones_str,
        )

        # --- Progress callback (SSE streaming support) ---
        if progress_callback is not None:
            progress_callback(idx, item, n_valid)

    # --- Confusable group detection (S021) ------------------------------------
    # Identify template pairs that are visually identical but differ by a
    # text token in their description (e.g. "TH", "Ex").  These pairs
    # cannot be separated by confidence alone and need text disambiguation.
    confusable_groups = _find_confusable_groups(templates)
    confusable_pair_set: set[tuple[int, int]] = set()
    for idx_a, idx_b, _, _ in confusable_groups:
        confusable_pair_set.add((min(idx_a, idx_b), max(idx_a, idx_b)))
    if confusable_groups:
        logger.debug(
            "confusable_groups: %d pairs: %s",
            len(confusable_groups),
            [(a, b) for a, b, _, _ in confusable_groups],
        )

    # --- Cross-symbol NMS (T147, refined T148, S021) --------------------------
    # When two DIFFERENT templates match at the same physical location,
    # deduplicate.  But SKIP suppression for confusable pairs — they will
    # be resolved by text disambiguation instead.
    CROSS_NMS_RADIUS_PT = 25.0   # ~25pt ≈ 9mm — catches offset matches
    CROSS_NMS_CONF_GAP = 0.04   # min confidence advantage to suppress

    if len(all_matches) > 1:
        # Sort by confidence descending
        all_matches.sort(key=lambda m: m.confidence, reverse=True)

        deduped: list[VisualMatch] = []
        for m in all_matches:
            dominated = False
            for kept in deduped:
                if kept.symbol_index == m.symbol_index:
                    continue  # same symbol — already handled by per-symbol NMS
                # Skip suppression for confusable pairs (S021) — they need
                # text disambiguation, not confidence-based suppression.
                pair_key = (min(kept.symbol_index, m.symbol_index),
                            max(kept.symbol_index, m.symbol_index))
                if pair_key in confusable_pair_set:
                    continue
                dist = math.sqrt((m.x - kept.x) ** 2 + (m.y - kept.y) ** 2)
                if dist < CROSS_NMS_RADIUS_PT:
                    conf_gap = kept.confidence - m.confidence
                    if conf_gap >= CROSS_NMS_CONF_GAP:
                        # Clear winner — suppress the weaker match
                        dominated = True
                        break
                    # else: gap too small — both survive (ambiguous pair)
            if not dominated:
                deduped.append(m)

        n_cross_suppressed = len(all_matches) - len(deduped)
        if n_cross_suppressed > 0:
            logger.debug(
                "cross_symbol_nms: %d matches suppressed (before=%d after=%d)",
                n_cross_suppressed, len(all_matches), len(deduped),
            )
            all_matches = deduped

    # --- Text-aided disambiguation (S021) -------------------------------------
    # For confusable pairs that survived cross-NMS (both templates matched
    # at the same location), read nearby text to decide which variant wins.
    if confusable_groups and drawing_chars:
        all_matches = _disambiguate_matches(
            all_matches, confusable_groups, drawing_chars,
        )

    # Rebuild counts from final matches
    counts = {}
    for desc_idx in descriptions:
        counts[desc_idx] = 0
    for m in all_matches:
        counts[m.symbol_index] = counts.get(m.symbol_index, 0) + 1

    # Collect indices that were resolved by text disambiguation
    disambig_set: set[int] = set()
    for idx_a, idx_b, _, _ in confusable_groups:
        disambig_set.add(idx_a)
        disambig_set.add(idx_b)

    return VisualResult(
        matches=all_matches,
        counts=counts,
        descriptions=descriptions,
        disambiguated_indices=disambig_set,
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

    # Pictogram detection (T149)
    picto_result = detect_pictograms(pdf_path, legend, page=scan_page)
    if picto_result.counts:
        print()
        print("Pictograms (text-based):")
        for name, count in picto_result.counts.items():
            print(f"  {name}: {count}")

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
