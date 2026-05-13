"""cable_length.py -- T068 / B0XX raster cable-length engine.

Companion to ``pdf_cables_by_method`` (vector-only) and
``pdf_count_geometry`` (vector + scale).  Those modules walk the
``page.lines`` collection produced by ``pdfplumber``; many GPK3 lighting
sheets carry their cable trasses as RASTER fills inside an embedded
image stream, so the vector pipelines miss them entirely (the symptom is
the production VOR generator reporting 519 m of cable on the 3-zahvatka
folder where the reference total is ~14 000 m).

This module renders every plan page at 300 DPI, builds 5 colour masks
(one per cable-trace legend category 10..14 on the 007-style legends:
kabel'naya trassa avariynogo / rabochego osvescheniya plus the 3
provodka categories), skeletonises each mask, sums skeleton pixels and
converts to metres using the per-page drawing scale.  One length item
is emitted per ``(category, page)`` pair.

The five categories are mapped from the legend by the following
heuristic, applied to ``LegendItem.color`` and the lower-case
description:

  idx 10  kabel'naya trassa avariynogo osvescheniya  -- red mask
  idx 11  kabel'naya trassa rabochego  osvescheniya  -- blue mask
  idx 12  provodka v lotke                           -- magenta or
                                                       (colour-blind:
                                                        2nd red entry)
  idx 13  provodka v trube skryto                    -- green-or-cyan
  idx 14  provodka v trube otkryto                   -- olive-or-orange

When the legend stops short (KB-015: symbol-less rows with an empty
LegendItem.symbol field) the colour is taken from a "learned default"
table indexed by (category, description-keyword) so the engine still
produces a non-zero estimate.

Public API
----------

``measure_cable_lengths_raster(pdf_path, pages=None, legend_result=None,
                               scale_mm_per_pt=None, dpi=300)
        -> CableLengthReport``

  The report exposes ``items``: a list of dicts with the
  ``{symbol, name, count, count_ae, total, unit, category, source}``
  shape used by ``_count_equipment_in_pdf`` so that the orchestrator can
  splice the result straight into its ``items`` list.

KB-006: source file is ASCII-only -- Cyrillic strings appear only via
Unicode escapes.
"""

from __future__ import annotations

import io
import logging
import math
import os
import re
from dataclasses import dataclass, field
from typing import Optional

import numpy as np

try:  # pragma: no cover -- both should be installed in the project venv
    import cv2  # type: ignore
except Exception:  # pragma: no cover
    cv2 = None  # type: ignore

try:  # pragma: no cover
    import fitz  # type: ignore
except Exception:  # pragma: no cover
    fitz = None  # type: ignore

try:  # pragma: no cover
    from skimage.morphology import skeletonize as _skel  # type: ignore
except Exception:  # pragma: no cover
    _skel = None  # type: ignore

# Vector-pipeline neighbour: reuse the scale detector and scale dataclass
# rather than re-implementing axis-pair + dimension-text logic.
try:
    from pdf_cables_by_method import _detect_scale, MM_PER_PT
except Exception:  # pragma: no cover -- defensive
    _detect_scale = None  # type: ignore
    MM_PER_PT = 25.4 / 72.0

try:
    from pdf_legend_parser import LegendResult, LegendItem, parse_legend
except Exception:  # pragma: no cover -- defensive
    LegendResult = None  # type: ignore
    LegendItem = None  # type: ignore
    parse_legend = None  # type: ignore

import pdfplumber  # type: ignore

log = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Cyrillic keyword markers per category (KB-006: built via \uNNNN escapes).
# Each entry: (idx, label, primary_color_key, kw_required_tuple).  All
# keywords in a tuple must appear in the legend description (lower-cased)
# for the category to match a legend row.
_K_KABEL = "\u043a\u0430\u0431\u0435\u043b"          # "кабел"
_K_TRASSA = "\u0442\u0440\u0430\u0441\u0441"          # "трасс"
_K_AVAR = "\u0430\u0432\u0430\u0440"                  # "авар" (аварийного)
_K_RABO = "\u0440\u0430\u0431\u043e"                  # "рабо" (рабочего)
_K_PROV = "\u043f\u0440\u043e\u0432\u043e\u0434"      # "провод" (проводка)
_K_LOTOK = "\u043b\u043e\u0442\u043a"                 # "лотк" (в лотке)
_K_TRUBA = "\u0442\u0440\u0443\u0431"                 # "труб" (в трубе)
_K_SKRYT = "\u0441\u043a\u0440\u044b"                 # "скры" (скрыто)
_K_OTKRY = "\u043e\u0442\u043a\u0440"                 # "откры" (открыто)

# Category descriptors used both for legend matching and for the emitted
# item names.  ``name_en`` is the ASCII transliteration kept solely for
# logging; ``name_ru`` is the production-side label inserted into the
# emitted item dict.
_CABLE_LENGTH_CATEGORIES = [
    {
        "key": "cable_emergency",
        "name_ru": (
            "\u041a\u0430\u0431\u0435\u043b\u044c\u043d\u0430\u044f "
            "\u0442\u0440\u0430\u0441\u0441\u0430 "
            "\u0430\u0432\u0430\u0440\u0438\u0439\u043d\u043e\u0433\u043e "
            "\u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u044f"
        ),
        "name_en": "Cable trace emergency lighting",
        "kw_all": (_K_KABEL, _K_AVAR),
        "default_color_key": "red",
    },
    {
        "key": "cable_working",
        "name_ru": (
            "\u041a\u0430\u0431\u0435\u043b\u044c\u043d\u0430\u044f "
            "\u0442\u0440\u0430\u0441\u0441\u0430 "
            "\u0440\u0430\u0431\u043e\u0447\u0435\u0433\u043e "
            "\u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u044f"
        ),
        "name_en": "Cable trace working lighting",
        "kw_all": (_K_KABEL, _K_RABO),
        "default_color_key": "blue",
    },
    {
        "key": "wire_tray",
        "name_ru": (
            "\u041f\u0440\u043e\u0432\u043e\u0434\u043a\u0430 "
            "\u0432 \u043b\u043e\u0442\u043a\u0435"
        ),
        "name_en": "Wire in tray",
        "kw_all": (_K_PROV, _K_LOTOK),
        "default_color_key": "magenta",
    },
    {
        "key": "wire_pipe_hidden",
        "name_ru": (
            "\u041f\u0440\u043e\u0432\u043e\u0434\u043a\u0430 "
            "\u0432 \u0442\u0440\u0443\u0431\u0435 \u0441\u043a\u0440\u044b\u0442\u043e"
        ),
        "name_en": "Wire in pipe (hidden)",
        "kw_all": (_K_PROV, _K_TRUBA, _K_SKRYT),
        "default_color_key": "green",
    },
    {
        "key": "wire_pipe_open",
        "name_ru": (
            "\u041f\u0440\u043e\u0432\u043e\u0434\u043a\u0430 "
            "\u0432 \u0442\u0440\u0443\u0431\u0435 \u043e\u0442\u043a\u0440\u044b\u0442\u043e"
        ),
        "name_en": "Wire in pipe (open)",
        "kw_all": (_K_PROV, _K_TRUBA, _K_OTKRY),
        "default_color_key": "olive",
    },
]

# HSV ranges (OpenCV uses H in [0, 180), S/V in [0, 256)).  Each entry
# may carry two ranges to handle the red-hue wrap-around (H near 0 and
# H near 180).  Saturation lower-bound is generous (40) because faint
# 1-px raster anti-aliasing washes the colour out.
_COLOR_HSV_RANGES = {
    "red": [
        (np.array([0,   60,  60], dtype=np.uint8),
         np.array([10, 255, 255], dtype=np.uint8)),
        (np.array([170, 60,  60], dtype=np.uint8),
         np.array([179, 255, 255], dtype=np.uint8)),
    ],
    "blue": [
        (np.array([100, 60,  60], dtype=np.uint8),
         np.array([130, 255, 255], dtype=np.uint8)),
    ],
    "magenta": [
        (np.array([140, 50,  60], dtype=np.uint8),
         np.array([170, 255, 255], dtype=np.uint8)),
    ],
    "green": [
        (np.array([40,  50,  60], dtype=np.uint8),
         np.array([85,  255, 255], dtype=np.uint8)),
    ],
    "cyan": [
        (np.array([85,  50,  60], dtype=np.uint8),
         np.array([100, 255, 255], dtype=np.uint8)),
    ],
    "olive": [
        (np.array([20,  50,  40], dtype=np.uint8),
         np.array([40,  255, 200], dtype=np.uint8)),
    ],
    "orange": [
        (np.array([10,  80,  80], dtype=np.uint8),
         np.array([25,  255, 255], dtype=np.uint8)),
    ],
}

# Fallback HSV colour rotation for when two categories collide on the
# same primary colour (rare but possible when KB-015 leaves descriptions
# without symbols).  Categories take colours in this order; if a primary
# is taken, we shift to the next free entry.
_CATEGORY_COLOR_FALLBACK = [
    "red", "blue", "magenta", "green", "olive", "orange", "cyan",
]

# Minimum component size to keep in a category mask.  Removes ink dots
# and short ticks that survive the colour filter (e.g. "1.5" digit ink
# happens to be red on some sheets).  At 300 DPI 1 m of cable spans
# (1000 / mm_per_pt) * (DPI/72) px -- typically 800-1500 px.  We trim
# anything below 40 px (< 5 cm at 1:100, < 3 cm at 1:50).
_MIN_COMPONENT_PX = 40

# Morphology kernels.  Used to bridge tiny gaps in dashed raster cables
# before skeletonising.
_MORPH_CLOSE_KERNEL_SIZE = 3

# Cap on the per-page raster length.  A 1 m cable at 1:100 / 300 DPI is
# ~1180 skeleton pixels; one full sheet rarely exceeds 3000 m == 3.5e6
# px of skeleton.  Anything beyond 5e6 px is almost certainly a hue
# false-positive (a magenta-tinted photo of a building or a green-grid
# screen-print on the title block).  We clamp to MAX_PER_PAGE_M when
# the raw number is wildly out of range.
_MAX_PER_PAGE_M = 50_000.0


# ---------------------------------------------------------------------------
# Dataclasses
# ---------------------------------------------------------------------------


@dataclass
class CableLengthCategoryEntry:
    """One ``(category, page)`` measurement."""
    page_index: int
    category_key: str
    name_ru: str
    legend_idx: int                  # -1 when no legend match was found
    color_key: str                   # which HSV range produced the mask
    skeleton_px: int                 # raw skeleton-pixel sum (post-cleanup)
    length_m: float                  # converted to metres
    scale_mm_per_pt: float           # the per-page scale used
    scale_source: str                # detector name (axis_pair_, dim_text_,
                                     # page_heuristic_, fallback)
    note: str = ""                   # diagnostic note (e.g. "fallback_color",
                                     # "no_legend_match", "clamped")


@dataclass
class CableLengthReport:
    pdf_path: str
    pages_scanned: list[int] = field(default_factory=list)
    entries: list[CableLengthCategoryEntry] = field(default_factory=list)
    items: list[dict] = field(default_factory=list)   # ready for splice
    total_length_m: float = 0.0
    notes: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _classify_legend_category(item) -> Optional[str]:
    """Return the cable_length category key for a legend item, or None.

    Matches on the lower-cased description.  Returns the first category
    whose ``kw_all`` keywords are all present.  Order matters because
    "проводка в трубе скрыто" must beat "проводка в трубе открыто" if
    *both* keyword sets accidentally pass (they share kw "труб").
    """
    if item is None:
        return None
    desc = (getattr(item, "description", "") or "").lower()
    if not desc:
        return None
    for cat in _CABLE_LENGTH_CATEGORIES:
        if all(kw in desc for kw in cat["kw_all"]):
            return cat["key"]
    return None


def _legend_color_to_key(item) -> Optional[str]:
    """Map ``LegendItem.color`` string to one of the HSV-range keys."""
    if item is None:
        return None
    c = (getattr(item, "color", "") or "").strip().lower()
    if c == "red":
        return "red"
    if c == "blue":
        return "blue"
    if c == "green":
        return "green"
    return None


def _resolve_category_colors(
    legend_items: Optional[list],
) -> list[dict]:
    """Build the per-category list with concrete colour-keys resolved.

    Output element shape:
        {category dict from _CABLE_LENGTH_CATEGORIES,
         "legend_idx": int or -1,
         "color_key": str,
         "color_source": "legend_color" / "default" / "fallback"}
    """
    resolved: list[dict] = []
    used_colors: set[str] = set()
    for cat in _CABLE_LENGTH_CATEGORIES:
        legend_idx = -1
        color_key = cat["default_color_key"]
        color_source = "default"

        if legend_items:
            for i, it in enumerate(legend_items):
                key = _classify_legend_category(it)
                if key == cat["key"]:
                    legend_idx = i
                    leg_color = _legend_color_to_key(it)
                    if leg_color:
                        color_key = leg_color
                        color_source = "legend_color"
                    break

        # Resolve collisions: if the chosen colour is already used, walk
        # the fallback list for an unused entry.
        if color_key in used_colors:
            for alt in _CATEGORY_COLOR_FALLBACK:
                if alt not in used_colors and alt in _COLOR_HSV_RANGES:
                    color_key = alt
                    color_source = "fallback"
                    break

        used_colors.add(color_key)
        resolved.append({
            **cat,
            "legend_idx": legend_idx,
            "color_key": color_key,
            "color_source": color_source,
        })
    return resolved


def _render_page_rgb(pdf_path: str, page_index: int, dpi: int) -> Optional[np.ndarray]:
    """Render one page to an RGB numpy array via PyMuPDF.

    Returns ``None`` if the rendering fails or fitz is unavailable.
    """
    if fitz is None or cv2 is None:
        return None
    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        log.warning("fitz.open failed for %s: %s", pdf_path, exc)
        return None
    try:
        if page_index >= doc.page_count:
            return None
        page = doc.load_page(page_index)
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False, colorspace=fitz.csRGB)
        # PyMuPDF returns row-major HxWxC with C=3 (RGB).
        arr = np.frombuffer(pix.samples, dtype=np.uint8)
        arr = arr.reshape(pix.height, pix.width, 3)
        return arr
    except Exception as exc:
        log.warning("render_page failed for %s p=%d: %s", pdf_path, page_index, exc)
        return None
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _build_color_mask(hsv: np.ndarray, color_key: str) -> Optional[np.ndarray]:
    """Combine the (possibly multiple) HSV ranges for ``color_key`` into
    a single binary mask.
    """
    ranges = _COLOR_HSV_RANGES.get(color_key)
    if not ranges or cv2 is None:
        return None
    out = None
    for lo, hi in ranges:
        m = cv2.inRange(hsv, lo, hi)
        if out is None:
            out = m
        else:
            out = cv2.bitwise_or(out, m)
    return out


def _suppress_legend_area(
    mask: np.ndarray,
    page_size_pt: tuple[float, float],
    legend_bbox_pt: tuple[float, float, float, float],
    dpi: int,
) -> np.ndarray:
    """Zero out the legend rectangle on the mask so legend swatches do
    not contribute to the skeleton-length sum.
    """
    if not legend_bbox_pt or legend_bbox_pt == (0, 0, 0, 0):
        return mask
    px_per_pt = dpi / 72.0
    x0, y0, x1, y1 = legend_bbox_pt
    h, w = mask.shape[:2]
    # PyMuPDF / pdfplumber both use top-left origin; the y axis is
    # consistent so we can scale directly.
    pad = int(round(8 * px_per_pt))  # 8 pt safety margin around the bbox
    xi0 = max(0, int(round(x0 * px_per_pt)) - pad)
    yi0 = max(0, int(round(y0 * px_per_pt)) - pad)
    xi1 = min(w, int(round(x1 * px_per_pt)) + pad)
    yi1 = min(h, int(round(y1 * px_per_pt)) + pad)
    if xi1 > xi0 and yi1 > yi0:
        mask[yi0:yi1, xi0:xi1] = 0
    return mask


def _suppress_titleblock(mask: np.ndarray) -> np.ndarray:
    """Zero out the bottom-right title block (approximate: 22% wide,
    18% tall) so the project штамп frame does not pollute the cable
    mask with its red/blue logo strokes.
    """
    h, w = mask.shape[:2]
    x0 = int(w * 0.78)
    y0 = int(h * 0.82)
    mask[y0:, x0:] = 0
    return mask


def _clean_and_skeletonize(mask: np.ndarray) -> np.ndarray:
    """Morphological close, remove small components, then skeletonise."""
    if cv2 is None:
        return mask
    k = _MORPH_CLOSE_KERNEL_SIZE
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k, k))
    closed = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel)

    # Remove tiny components below _MIN_COMPONENT_PX
    n, labels, stats, _ = cv2.connectedComponentsWithStats(closed, connectivity=8)
    keep = np.zeros_like(closed)
    for i in range(1, n):
        if stats[i, cv2.CC_STAT_AREA] >= _MIN_COMPONENT_PX:
            keep[labels == i] = 255

    # Skeletonise.  Prefer cv2.ximgproc.thinning (Zhang-Suen, fast C++)
    # when available; otherwise fall back to skimage.morphology.skeletonize.
    skel = None
    try:
        if hasattr(cv2, "ximgproc"):
            skel = cv2.ximgproc.thinning(
                keep, thinningType=cv2.ximgproc.THINNING_ZHANGSUEN,
            )
    except Exception:
        skel = None
    if skel is None and _skel is not None:
        bool_in = keep > 0
        skel_bool = _skel(bool_in)
        skel = (skel_bool.astype(np.uint8) * 255)
    if skel is None:
        # Last-resort: count the closed-mask area divided by an average
        # stroke width estimate.  Treat 3 px as the average stroke at
        # 300 DPI / 1:100, then return closed/3.
        return (keep / 3).astype(np.uint8)
    return skel


def _detect_scale_for_page(page) -> tuple[float, str]:
    """Return ``(mm_per_pt, source)`` for one pdfplumber page.

    Uses the vector pipeline's detector when available; falls back to
    the A1-1:100 heuristic otherwise.
    """
    try:
        words = page.extract_words(x_tolerance=2, y_tolerance=2) or []
        lines = page.lines or []
        if _detect_scale is not None:
            mm_per_pt, source, _info = _detect_scale(words, lines, page)
            return float(mm_per_pt), str(source)
    except Exception as exc:
        log.debug("scale detector raised for page %d: %s",
                  getattr(page, "page_number", -1) - 1, exc)
    # A1 landscape 1:100 fallback
    pw = max(page.width, page.height)
    paper_mm = pw / 72.0 * 25.4
    mm_per_pt = paper_mm / pw * 100.0
    return float(mm_per_pt), "page_heuristic_1:100"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def measure_cable_lengths_raster(
    pdf_path: str,
    pages: Optional[list[int]] = None,
    legend_result: Optional["LegendResult"] = None,
    scale_mm_per_pt: Optional[float] = None,
    dpi: int = 300,
) -> CableLengthReport:
    """Render every plan page and measure cable length per category.

    Args:
        pdf_path: input PDF.
        pages: 0-based page indices to process; ``None`` means all.
        legend_result: pre-parsed legend; will be parsed on demand if
            absent.  Used to learn per-category colours and to mask the
            legend swatches out of the raster.
        scale_mm_per_pt: caller-supplied scale override.  If absent,
            ``_detect_scale`` from ``pdf_cables_by_method`` is invoked
            per page.
        dpi: raster resolution.  300 DPI is a sweet spot between memory
            (a 4-zahvatka folder fits in 1.2 GB) and capturing thin
            1-pt cable strokes.

    Returns:
        ``CableLengthReport`` with a populated ``items`` list whose
        elements have the orchestrator-compatible shape.
    """
    rep = CableLengthReport(pdf_path=pdf_path)
    if cv2 is None or fitz is None:
        rep.notes.append("imaging_deps_missing")
        return rep

    # Resolve legend once
    if legend_result is None and parse_legend is not None:
        try:
            legend_result = parse_legend(pdf_path)
        except Exception as exc:
            rep.notes.append("legend_parse_failed:" + str(exc)[:80])
            legend_result = None

    legend_items = (
        list(legend_result.items)
        if legend_result is not None and hasattr(legend_result, "items")
        else None
    )
    legend_bbox = (
        getattr(legend_result, "legend_bbox", (0, 0, 0, 0))
        if legend_result is not None
        else (0, 0, 0, 0)
    )
    legend_page = (
        getattr(legend_result, "page_index", -1)
        if legend_result is not None
        else -1
    )
    categories = _resolve_category_colors(legend_items)

    # ----- page loop ------------------------------------------------------
    try:
        plumb = pdfplumber.open(pdf_path)
    except Exception as exc:
        rep.notes.append("pdfplumber_open_failed:" + str(exc)[:80])
        return rep

    totals_by_cat: dict[str, float] = {c["key"]: 0.0 for c in categories}

    try:
        n_pages = len(plumb.pages)
        scan = pages if pages is not None else list(range(n_pages))
        for page_idx in scan:
            if page_idx < 0 or page_idx >= n_pages:
                continue
            page = plumb.pages[page_idx]
            rep.pages_scanned.append(page_idx)

            if scale_mm_per_pt is not None and scale_mm_per_pt > 0:
                mm_per_pt = float(scale_mm_per_pt)
                scale_source = "user_supplied"
            else:
                mm_per_pt, scale_source = _detect_scale_for_page(page)
            if mm_per_pt <= 0:
                rep.notes.append("scale_invalid_page_%d" % page_idx)
                continue

            # px-per-metre = 1000 / mm_per_px = 1000 * px_per_mm
            # mm_per_px = mm_per_pt * (72 / dpi)
            # => px_per_m = 1000 / (mm_per_pt * 72 / dpi)
            #            = 1000 * dpi / (72 * mm_per_pt)
            px_per_m = 1000.0 * dpi / (72.0 * mm_per_pt)
            if px_per_m <= 0 or not math.isfinite(px_per_m):
                rep.notes.append("px_per_m_invalid_page_%d" % page_idx)
                continue

            rgb = _render_page_rgb(pdf_path, page_idx, dpi)
            if rgb is None:
                rep.notes.append("render_failed_page_%d" % page_idx)
                continue
            hsv = cv2.cvtColor(rgb, cv2.COLOR_RGB2HSV)

            page_legend_bbox = (
                legend_bbox if page_idx == legend_page else (0, 0, 0, 0)
            )

            for cat in categories:
                mask = _build_color_mask(hsv, cat["color_key"])
                if mask is None:
                    continue
                mask = _suppress_legend_area(
                    mask,
                    (page.width, page.height),
                    page_legend_bbox,
                    dpi,
                )
                mask = _suppress_titleblock(mask)
                skel = _clean_and_skeletonize(mask)
                skel_px = int(cv2.countNonZero(skel))

                length_m = skel_px / px_per_m if px_per_m > 0 else 0.0
                note = ""
                if length_m > _MAX_PER_PAGE_M:
                    note = "clamped_to_%g" % _MAX_PER_PAGE_M
                    length_m = _MAX_PER_PAGE_M

                entry = CableLengthCategoryEntry(
                    page_index=page_idx,
                    category_key=cat["key"],
                    name_ru=cat["name_ru"],
                    legend_idx=cat["legend_idx"],
                    color_key=cat["color_key"],
                    skeleton_px=skel_px,
                    length_m=float(length_m),
                    scale_mm_per_pt=float(mm_per_pt),
                    scale_source=scale_source,
                    note=note,
                )
                rep.entries.append(entry)
                totals_by_cat[cat["key"]] = (
                    totals_by_cat.get(cat["key"], 0.0) + length_m
                )

                # One item per (category, page) keeps audit trail; the
                # orchestrator aggregates downstream when it derives VOR
                # rows.  Tiny lengths (< 0.5 m) are dropped to keep the
                # report readable -- they are usually anti-alias noise.
                if length_m >= 0.5:
                    rep.items.append({
                        "symbol": "",
                        "name": cat["name_ru"],
                        "count": 0,
                        "count_ae": 0,
                        "total": round(length_m, 1),
                        "unit": "\u043c",   # "м"
                        "category": "cable_length_raster",
                        "source": "cable_length_raster",
                        "page_index": page_idx,
                        "color_key": cat["color_key"],
                        "scale_source": scale_source,
                    })
    finally:
        try:
            plumb.close()
        except Exception:
            pass

    rep.total_length_m = float(sum(totals_by_cat.values()))
    rep.notes.append(
        "per_category_m=" + ",".join(
            "%s:%.1f" % (k, v) for k, v in totals_by_cat.items()
        )
    )
    return rep


# ---------------------------------------------------------------------------
# CLI for one-off debugging.  Not used by the orchestrator.
# ---------------------------------------------------------------------------

def _main_cli() -> int:  # pragma: no cover
    import argparse
    import json

    parser = argparse.ArgumentParser(description="Raster cable-length engine")
    parser.add_argument("pdf")
    parser.add_argument("--page", type=int, default=None,
                        help="0-based page index (default: all)")
    parser.add_argument("--dpi", type=int, default=300)
    parser.add_argument("--json", action="store_true",
                        help="emit JSON instead of human summary")
    args = parser.parse_args()

    pages = [args.page] if args.page is not None else None
    rep = measure_cable_lengths_raster(args.pdf, pages=pages, dpi=args.dpi)

    if args.json:
        out = {
            "pdf": rep.pdf_path,
            "pages": rep.pages_scanned,
            "total_m": round(rep.total_length_m, 1),
            "entries": [
                {
                    "page": e.page_index,
                    "cat": e.category_key,
                    "color": e.color_key,
                    "legend_idx": e.legend_idx,
                    "skel_px": e.skeleton_px,
                    "len_m": round(e.length_m, 1),
                    "scale_mm_per_pt": round(e.scale_mm_per_pt, 4),
                    "scale_src": e.scale_source,
                    "note": e.note,
                }
                for e in rep.entries
            ],
            "notes": rep.notes,
        }
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        print("PDF:", rep.pdf_path)
        print("Pages scanned:", rep.pages_scanned)
        print("Total length: %.1f m" % rep.total_length_m)
        for e in rep.entries:
            print("  p=%d cat=%s color=%s skel_px=%d len=%.1f m (src=%s)%s" % (
                e.page_index, e.category_key, e.color_key,
                e.skeleton_px, e.length_m, e.scale_source,
                ("  [" + e.note + "]") if e.note else "",
            ))
        for note in rep.notes:
            print("note:", note)
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(_main_cli())
