#!/usr/bin/env python3
"""
DXF Visualizer — render DXF drawing to PNG with equipment markers.

Renders the full DXF drawing and overlays colored circle markers on each
piece of found equipment.  A color legend maps markers to equipment names.

Usage:
    python dxf_visualizer.py drawing.dxf
    python dxf_visualizer.py drawing.dxf -o output.png --dpi 200
"""

import argparse
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

_READY = False
_IMPORT_ERR = ""

try:
    import ezdxf
    from ezdxf.addons.drawing import Frontend, RenderContext
    from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    _READY = True
except ImportError as exc:
    _IMPORT_ERR = str(exc)

from equipment_counter import (
    EquipmentItem,
    process_dxf,
    _clean_mtext,
    _find_all_legend_bboxes,
    _detect_grid_labels,
    SYMBOL_RE,
    CIRCUIT_VARIANT_RE,
    _SPECIAL_LABELS,
)

# ── 20 visually distinct marker colours ──────────────────────────────

MARKER_COLORS = [
    "#e6194b", "#3cb44b", "#4363d8", "#f58231", "#911eb4",
    "#42d4f4", "#f032e6", "#bfef45", "#fabed4", "#469990",
    "#dcbeff", "#9A6324", "#800000", "#aaffc3", "#808000",
    "#000075", "#a9a9a9", "#e6beff", "#fffac8", "#ffd8b1",
]
LEGEND_BOX_COLOR = "#FFD700"
_PLAN_ANN_RE = re.compile(r"^(\d{1,3})\s*[-–—]\s*(.+)", re.MULTILINE)

_EQUIP_LAYERS = frozenset({
    "эом", "спз",
})
_EQUIP_BLOCK_KW = (
    "выкл", "светильник", "slick", "arctic", "свет.",
)
_SKIP_BLOCK_KW = (
    "ось", "сетки", "помещен", "рамка", "формат", "разрез",
    "стрелка", "фахверк", "воронка", "вкладыш", "подвес",
    "лотки", "кабел",
)
_LIGHTING_LAYERS = frozenset({
    "рабочее освещение", "аварийное освещение", "аварийное освещени",
    "эом", "спз",
})


def _hidden_layers(doc) -> frozenset[str]:
    """Return set of layer names that are frozen or turned off."""
    hidden: set[str] = set()
    for layer in doc.layers:
        if layer.is_frozen() or not layer.is_on():
            hidden.add(layer.dxf.name)
    return frozenset(hidden)


def _collect_block_markers(
    msp,
    bboxes: list[tuple[float, float, float, float]],
    plan_x_range: tuple[float, float] | None = None,
    hidden: frozenset[str] = frozenset(),
    log=print,
) -> list[tuple[float, float, str, str]]:
    """Find INSERT block positions for electrical equipment on the plan.

    Returns list of (x, y, block_name, layer) for equipment blocks
    found outside legend bounding boxes and within plan_x_range.
    """
    results: list[tuple[float, float, str, str]] = []
    for e in msp.query("INSERT"):
        try:
            ix, iy = e.dxf.insert.x, e.dxf.insert.y
            bname = e.dxf.name
            layer = e.dxf.layer
        except Exception:
            continue
        if layer in hidden:
            continue
        if plan_x_range and not (plan_x_range[0] <= ix <= plan_x_range[1]):
            continue
        bl = bname.lower()
        ll = layer.lower()
        if bl.startswith("*u") or bl.startswith("a$c"):
            continue
        if any(kw in bl for kw in _SKIP_BLOCK_KW):
            continue
        if _in_any_bbox(ix, iy, bboxes):
            continue
        is_equip = (
            ll in _EQUIP_LAYERS
            or any(kw in bl for kw in _EQUIP_BLOCK_KW)
        )
        if is_equip:
            results.append((ix, iy, bname, layer))
    return results


# ── Position collector ───────────────────────────────────────────────

def _in_any_bbox(x: float, y: float,
                 bboxes: list[tuple[float, float, float, float]]) -> bool:
    return any(xmin < x < xmax and ymin < y < ymax
               for xmin, ymin, xmax, ymax in bboxes)


def _dedup_points(pts: list[tuple[float, float]],
                   tol: float = 1.0) -> list[tuple[float, float]]:
    """Remove duplicate (x, y) positions within *tol* DXF units."""
    if not pts:
        return pts
    seen: list[tuple[float, float]] = []
    for x, y in pts:
        if not any(abs(x - sx) < tol and abs(y - sy) < tol for sx, sy in seen):
            seen.append((x, y))
    return seen


def _fuzzy_name_match(a: str, b: str) -> bool:
    """Check if two equipment names match loosely (substring or token overlap)."""
    al, bl = a.lower().strip(), b.lower().strip()
    if al == bl:
        return True
    # One is a substring of the other
    if al in bl or bl in al:
        return True
    # Token overlap ≥ 60%
    at = set(re.findall(r"[a-zA-Zа-яА-ЯёЁ]+|\d+", al))
    bt = set(re.findall(r"[a-zA-Zа-яА-ЯёЁ]+|\d+", bl))
    if at and bt:
        overlap = len(at & bt) / max(len(at), len(bt))
        if overlap >= 0.6:
            return True
    return False


def _collect_markers(
    msp,
    entries: list[tuple[str, float, float]],
    all_bboxes: list[tuple[float, float, float, float]],
    items: list[EquipmentItem],
    grid_positions: set[tuple[int, int]] | None = None,
    entry_layers: dict[tuple[float, float], str] | None = None,
    log=print,
) -> dict[str, list[tuple[float, float]]]:
    """Find (x, y) positions of each equipment symbol on the plan.

    T019 fix: every parsed equipment item with coordinates MUST get markers.
    Detection heuristic based on symbol-prefix convention set by process_dxf():
      plain numeric/cyrillic → MTEXT scan
      B-prefix              → INSERT block scan
      P-prefix              → plan annotation MTEXT scan (fuzzy name match)
      D-prefix / ⚙ / Щ     → geometric circle scan (relaxed layer filter)
      ВЫХОД etc.            → special MTEXT scan

    A final fallback pass picks up any active items that still have zero
    markers after the primary passes, preventing silent drops.
    """
    positions: dict[str, list[tuple[float, float]]] = defaultdict(list)
    active = {it.symbol: it for it in items if it.count > 0 or it.count_ae > 0}
    if not active:
        return {}

    numbered = {s for s in active if SYMBOL_RE.match(s) and len(s) <= 4}
    specials = {s for s in active if s in _SPECIAL_LABELS}
    p_items  = {s: active[s] for s in active if s.startswith("P")}

    # ── Pre-filter entries outside legend/grid once ──
    plan_entries: list[tuple[str, float, float]] = []
    for text, x, y in entries:
        if _in_any_bbox(x, y, all_bboxes):
            continue
        if grid_positions and (round(x), round(y)) in grid_positions:
            continue
        plan_entries.append((text, x, y))

    # ── 1. MTEXT: numbered symbols, specials, plan annotations ──
    for text, x, y in plan_entries:
        fl = text.split("\n")[0].strip()

        if fl in specials:
            positions[fl].append((x, y))
            continue

        if SYMBOL_RE.match(fl) and len(fl) <= 4:
            if fl in numbered:
                # B019 fix: Short numeric symbols (1, 2, 3...) are confirmed
                # legend entries — accept them on any layer, not just _LIGHTING_LAYERS.
                positions[fl].append((x, y))
                continue
            # B020 fix: А/АЭ variants (e.g. "1А", "1АЭ", "2АЭ") should be
            # mapped to their base symbol so markers appear on the plan.
            m_var = CIRCUIT_VARIANT_RE.match(fl)
            if m_var and m_var.group(1) in numbered:
                positions[m_var.group(1)].append((x, y))
                continue
            continue

        # T019 fix: P-prefix matching — use fuzzy name comparison instead of
        # exact equality so that minor differences in spacing, abbreviation,
        # or trailing characters do not silently drop markers.
        if p_items:
            m = _PLAN_ANN_RE.match(fl)
            if m:
                desc = m.group(2).strip()
                for sym, it in p_items.items():
                    if _fuzzy_name_match(it.name, desc):
                        positions[sym].append((x, y))
                        break

    # ── 2. INSERT blocks (B-prefix items) ──
    b_items = sorted(
        (s for s in active if s.startswith("B")),
        key=lambda s: int(s[1:]) if s[1:].isdigit() else 0,
    )
    if b_items and all_bboxes:
        bbox0 = all_bboxes[0]
        xmin, ymin, xmax, ymax = bbox0
        legend_block_names: set[str] = set()
        plan_block_pts: dict[str, list[tuple[float, float]]] = defaultdict(list)

        for e in msp.query("INSERT"):
            ix, iy = e.dxf.insert.x, e.dxf.insert.y
            bname = e.dxf.name
            bl = bname.lower()
            if "ось" in bl or "сетки" in bl or "помещен" in bl:
                continue
            if xmin < ix < xmax and ymin < iy < ymax:
                legend_block_names.add(bname)
            elif not _in_any_bbox(ix, iy, all_bboxes):
                plan_block_pts[bname].append((ix, iy))

        sorted_blocks = sorted(legend_block_names)
        for bname, sym in zip(sorted_blocks, b_items):
            if bname in plan_block_pts:
                positions[sym].extend(plan_block_pts[bname])

    # ── 3. Geometric circles (D-prefix, ⚙, Щ) ──
    # T019 fix: relaxed layer filter — accept circles on ANY layer, not just
    # ANNO/СИМВ, to avoid silent drops when real drawings use different layers.
    geo_items = {s for s in active
                 if s.startswith("D") or s == "⚙" or s.startswith("Щ")}
    if geo_items:
        red_pts: list[tuple[float, float]] = []
        other_pts: list[tuple[float, float]] = []
        for c in msp.query("CIRCLE"):
            r = c.dxf.radius
            if abs(r - 50) > 15:          # T019: relaxed from ±5 to ±15
                continue
            cx, cy = c.dxf.center.x, c.dxf.center.y
            if _in_any_bbox(cx, cy, all_bboxes):
                continue
            color = c.dxf.get("color", 256)
            (red_pts if color == 1 else other_pts).append((cx, cy))

        for sym in geo_items:
            if sym.startswith("Щ"):
                positions[sym].extend(red_pts)
            else:
                positions[sym].extend(other_pts)

    # ── 4. T019 fallback: pick up any active items still missing markers ──
    # For items that weren't matched by any pass above, attempt a broad MTEXT
    # scan matching the item's name substring in plan text.  This prevents
    # the "0 отмечено" problem for unusual symbol types.
    missing_syms = {s for s in active if not positions.get(s)}
    if missing_syms:
        # Build a name→symbol lookup for fallback matching
        name_to_sym: dict[str, str] = {}
        for sym in missing_syms:
            it = active[sym]
            if it.name and it.name not in ("", "[?]"):
                name_to_sym[sym] = it.name

        for text, x, y in plan_entries:
            fl = text.split("\n")[0].strip()
            # Try matching MTEXT text against item names (fuzzy)
            for sym, name in name_to_sym.items():
                if _fuzzy_name_match(fl, name) or fl == sym:
                    positions[sym].append((x, y))
                    break

        # Report still-missing items for diagnostics
        still_missing = [s for s in missing_syms if not positions.get(s)]
        if still_missing:
            log(f"  [viz] WARNING: {len(still_missing)} items with 0 markers "
                f"after fallback: {still_missing[:10]}")

    # ── 5. Deduplicate positions per symbol ──
    result: dict[str, list[tuple[float, float]]] = {}
    for sym, pts in positions.items():
        deduped = _dedup_points(pts)
        if deduped:
            result[sym] = deduped

    return result


# ── Main visualisation function ──────────────────────────────────────

def _get_best_layout(doc) -> object | None:
    """Return the best Paper Space layout to render (if any).

    Prefers the first elevation-named layout (e.g. "+0,000"), falling back
    to any non-Model layout with viewports.  Returns None if only Model
    Space exists.
    """
    elev_layouts = []
    other_layouts = []
    for layout in doc.layouts:
        name = layout.name.strip()
        if name == "Model":
            continue
        if re.match(r"[+]?\d+[.,]\d+$", name):
            elev_layouts.append(layout)
        else:
            other_layouts.append(layout)

    # Prefer elevation-named layouts
    if elev_layouts:
        return elev_layouts[0]
    # Fall back to any non-Model layout that has viewports
    for layout in other_layouts:
        vps = list(layout.query("VIEWPORT"))
        if len(vps) > 1:      # viewport id=1 is always present (paper bg)
            return layout
    if other_layouts:
        return other_layouts[0]
    return None


def _compute_viewport_crop(doc) -> tuple[float, float, float, float] | None:
    """Compute Model-Space crop rectangle from the largest Paper Space viewport.

    Returns (xmin, ymin, xmax, ymax) in Model-Space coordinates, or None
    if no usable viewport is found.
    """
    best_area = 0.0
    best_bounds = None
    for layout in doc.layouts:
        if layout.name.strip() == "Model":
            continue
        for vp in layout.query("VIEWPORT"):
            if getattr(vp.dxf, "id", 0) == 1:
                continue
            try:
                vh = vp.dxf.view_height
                pw, ph = vp.dxf.width, vp.dxf.height
                vc = vp.dxf.view_center_point
                if vh <= 0 or ph <= 0:
                    continue
                vw = vh * (pw / ph)
                area = vw * vh
                if area > best_area:
                    best_area = area
                    best_bounds = (
                        vc.x - vw / 2, vc.y - vh / 2,
                        vc.x + vw / 2, vc.y + vh / 2,
                    )
            except Exception:
                continue
    return best_bounds


def visualize_dxf(
    dxf_path: str,
    png_path: str | None = None,
    items: list[EquipmentItem] | None = None,
    dpi: int = 200,
    log=print,
) -> str:
    """Render a DXF drawing to PNG and overlay coloured equipment markers.

    Parameters
    ----------
    dxf_path : str
        Path to the source DXF file.
    png_path : str | None
        Output PNG path.  Defaults to ``<dxf_path>.png``.
    items : list[EquipmentItem] | None
        Pre-computed equipment list.  If *None*, ``process_dxf()`` is called.
    dpi : int
        Output image resolution (default 200).
    log : callable
        Logging function (default ``print``).

    Returns
    -------
    str
        Absolute path of the saved PNG file.
    """
    if not _READY:
        raise ImportError(
            f"Visualization requires ezdxf + matplotlib: {_IMPORT_ERR}"
        )

    _MAX_DXF_SIZE_MB = 80
    fsize_mb = Path(dxf_path).stat().st_size / (1024 * 1024)
    if fsize_mb > _MAX_DXF_SIZE_MB:
        return visualize_combined_dxf(dxf_path, png_path, items=items,
                                      dpi=dpi, log=log)

    if png_path is None:
        png_path = str(Path(dxf_path).with_suffix(".png"))

    t0 = time.time()

    # ── Read DXF once ──
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    # ── Parse equipment if needed ──
    if items is None:
        log("  [viz] Parsing equipment…")
        items = process_dxf(dxf_path)

    active_items = [it for it in items if it.count > 0 or it.count_ae > 0]

    # ── Collect MTEXT entries from loaded doc ──
    hidden = _hidden_layers(doc)
    entries: list[tuple[str, float, float]] = []
    entry_layers: dict[tuple[float, float], str] = {}
    for e in msp.query("MTEXT"):
        if e.dxf.layer in hidden:
            continue
        clean = _clean_mtext(e.text)
        if clean:
            x, y = e.dxf.insert.x, e.dxf.insert.y
            entries.append((clean, x, y))
            entry_layers[(x, y)] = e.dxf.layer

    all_bboxes = _find_all_legend_bboxes(entries)
    grid_positions = _detect_grid_labels(msp)

    # ── Collect equipment positions ──
    positions = _collect_markers(msp, entries, all_bboxes, active_items,
                                 grid_positions, entry_layers=entry_layers,
                                 log=log)
    total_markers = sum(len(pts) for pts in positions.values())
    total_parsed = sum(it.count + it.count_ae for it in active_items)
    log(f"  [viz] {total_markers} markers across {len(positions)} equipment types"
        f" (parsed: {total_parsed} items across {len(active_items)} types)")
    # T019: per-item diagnostics for items with marker gap
    for it in active_items:
        found = len(positions.get(it.symbol, []))
        expected = it.count + it.count_ae
        if found == 0 and expected > 0:
            log(f"  [viz] ⚠ {it.symbol} ({it.name[:40]}): "
                f"0 markers / {expected} parsed")

    # ── Render Model Space + viewport crop (V-003) ──
    # Always render Model Space so equipment markers (in MS coords) draw correctly.
    # Use viewport crop from Paper Space to eliminate white-space / orphan entities.
    log("  [viz] Rendering DXF (Model Space)…")
    fig = plt.figure()
    ax = fig.add_axes([0, 0, 1, 1])
    ctx = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    config = Configuration.defaults().with_changes(
        background_policy=BackgroundPolicy.WHITE,
    )
    frontend = Frontend(ctx, backend, config=config)
    frontend.draw_layout(msp)

    ax.set_aspect("equal")
    ax.autoscale()

    # V-003: Viewport-aware cropping — use Paper Space viewport bounds
    vp_crop = _compute_viewport_crop(doc)
    if vp_crop:
        xmin, ymin, xmax, ymax = vp_crop
        pad_x = (xmax - xmin) * 0.03
        pad_y = (ymax - ymin) * 0.03
        ax.set_xlim(xmin - pad_x, xmax + pad_x)
        ax.set_ylim(ymin - pad_y, ymax + pad_y)
        log(f"  [viz] Viewport crop applied: "
            f"{xmax - xmin:.0f} x {ymax - ymin:.0f} DXF units")

    xlim = ax.get_xlim()
    ylim = ax.get_ylim()
    drawing_w = xlim[1] - xlim[0]
    drawing_h = ylim[1] - ylim[0]

    # ── V-002: Adaptive marker radius with clamp ──
    marker_r = max(drawing_w, drawing_h) * 0.003
    marker_r = max(50, min(500, marker_r))  # clamp to [50, 500] DXF units

    # ── Assign colours ──
    color_map: dict[str, str] = {}
    for i, it in enumerate(active_items):
        color_map[it.symbol] = MARKER_COLORS[i % len(MARKER_COLORS)]

    # ── Draw equipment markers ──
    for sym, pts in positions.items():
        clr = color_map.get(sym, "#FF0000")
        for x, y in pts:
            ax.add_patch(plt.Circle(
                (x, y), marker_r,
                fill=False, edgecolor=clr, linewidth=2.0,
                alpha=0.75, zorder=10,
            ))

    # ── Outline legend bounding boxes ──
    for bbox in all_bboxes:
        bx0, by0, bx1, by1 = bbox
        ax.add_patch(plt.Rectangle(
            (bx0, by0), bx1 - bx0, by1 - by0,
            fill=False, edgecolor=LEGEND_BOX_COLOR,
            linewidth=2, linestyle="--", alpha=0.7, zorder=9,
        ))

    # ── Colour legend panel ──
    handles: list[mpatches.Patch] = []
    for it in active_items:
        clr = color_map.get(it.symbol, "#FF0000")
        found = len(positions.get(it.symbol, []))
        total = it.count + it.count_ae
        label = f"{it.symbol}: {it.name[:50]}  [{found} отмечено / {total} подсчитано]"
        handles.append(mpatches.Patch(
            facecolor=clr, edgecolor="black", linewidth=0.5, label=label,
        ))

    if handles:
        leg = ax.legend(
            handles=handles, loc="upper left",
            fontsize=7, framealpha=0.92, fancybox=True, edgecolor="gray",
        )
        leg.set_zorder(20)

    # ── V-004: Improved figure sizing for higher resolution ──
    aspect = drawing_h / drawing_w if drawing_w > 0 else 1
    target_px_w = 8000  # increased from ~3200 for legible text
    fig_w = target_px_w / dpi
    fig_h = max(10, fig_w * aspect)
    fig.set_size_inches(fig_w, fig_h)

    # ── Save ──
    fig.savefig(png_path, dpi=dpi, bbox_inches="tight",
                facecolor="white", pad_inches=0.3)
    plt.close(fig)

    elapsed = time.time() - t0
    size_mb = Path(png_path).stat().st_size / 1024 / 1024
    log(f"  [viz] Saved {png_path} ({size_mb:.1f} MB, {elapsed:.1f}s)")
    return png_path


# ── Combined DXF per-sheet rendering ─────────────────────────────────

_SKIP_ENTITY_TYPES = frozenset(("MULTILEADER", "ACAD_TABLE", "ACAD_PROXY_ENTITY"))

_COUNT_ANN_RE = re.compile(r"^(\d+)\s*[-–]\s*(.+)", re.IGNORECASE)


def _entity_coord(entity, idx: int) -> float | None:
    """Extract the primary coordinate (0=X, 1=Y) of an entity."""
    dt = entity.dxftype()
    try:
        if dt in ("MTEXT", "TEXT", "INSERT", "CIRCLE", "HATCH", "WIPEOUT"):
            ins = entity.dxf.insert if hasattr(entity.dxf, "insert") else None
            return (ins.x if idx == 0 else ins.y) if ins else None
        if dt == "LINE":
            a = entity.dxf.start[idx]
            b = entity.dxf.end[idx]
            return (a + b) / 2
        if dt in ("LWPOLYLINE", "POLYLINE"):
            pts = list(entity.vertices()) if dt == "POLYLINE" else entity.get_points()
            if pts:
                return sum(p[idx] for p in pts) / len(pts)
        if dt == "DIMENSION":
            dp = entity.dxf.defpoint if hasattr(entity.dxf, "defpoint") else None
            return (dp.x if idx == 0 else dp.y) if dp else None
        if dt == "SPLINE":
            cps = entity.control_points
            if cps:
                return sum(p[idx] for p in cps) / len(cps)
        if dt == "ARC":
            c = entity.dxf.center
            return c.x if idx == 0 else c.y
    except Exception:
        pass
    return None


def _entity_x(entity) -> float | None:
    return _entity_coord(entity, 0)


def _entity_y(entity) -> float | None:
    return _entity_coord(entity, 1)


def _get_viewport_plan_size(doc) -> tuple[float, float]:
    """Get typical plan sheet size from paper-space viewport dimensions."""
    elev_layouts = []
    for layout in doc.layouts:
        name = layout.name.strip()
        if name == "Model":
            continue
        if re.match(r"[+]?\d+[.,]\d+$", name):
            elev_layouts.append(layout)

    sizes: list[tuple[float, float]] = []
    for layout in elev_layouts:
        best_area, best_w, best_h = 0.0, 0.0, 0.0
        for vp in layout.query("VIEWPORT"):
            if getattr(vp.dxf, "id", 0) == 1:
                continue
            try:
                vh = vp.dxf.view_height
                pw, ph = vp.dxf.width, vp.dxf.height
                if vh <= 0 or ph <= 0:
                    continue
                vw = vh * (pw / ph)
                if vw * vh > best_area:
                    best_area, best_w, best_h = vw * vh, vw, vh
            except Exception:
                continue
        if best_area > 0:
            sizes.append((best_w, best_h))

    if not sizes:
        return 0.0, 0.0
    sizes.sort(key=lambda s: s[0] * s[1])
    return sizes[len(sizes) // 2]


# Regex for annotation summaries ("P1 - 133 × description")
_ANN_SUMMARY_RE = re.compile(
    r"^(P?\d{1,3})\s*[-–—:]\s*(\d{1,4})\s*[×хxX*]\s*(.+)",
)


def visualize_combined_dxf(
    dxf_path: str,
    png_path: str | None = None,
    items: list[EquipmentItem] | None = None,
    dpi: int = 200,
    log=print,
) -> str:
    """Render a combined multi-sheet DXF file as separate PNGs per sheet.

    Detects sheet boundaries from legend positions in model space, using
    viewport dimensions for proper sizing.  Overlays equipment markers
    from individual symbol labels found on each plan.
    """
    if not _READY:
        raise ImportError(f"Visualization requires ezdxf + matplotlib: {_IMPORT_ERR}")

    t0 = time.time()
    log("  [viz] Loading combined DXF…")
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    # ── Step 1: Detect sheet boundaries from legend positions ──
    legend_positions: list[tuple[float, float]] = []
    for e in msp:
        if e.dxftype() == "MTEXT":
            plain = e.plain_text().replace("\n", " ").strip()
            if "Условн" in plain and "обозначен" in plain.lower():
                legend_positions.append((e.dxf.insert.x, e.dxf.insert.y))

    if len(legend_positions) < 2:
        raise ValueError("Not a combined DXF: fewer than 2 legends found")

    legend_positions.sort(key=lambda p: p[0])
    legend_xs = [lx for lx, _ in legend_positions]
    legend_y = min(ly for _, ly in legend_positions)

    # Get viewport dimensions for proper edge sheet sizing
    vp_w, vp_h = _get_viewport_plan_size(doc)

    # Typical gap between adjacent legends
    gaps = [legend_xs[i + 1] - legend_xs[i] for i in range(len(legend_xs) - 1)]
    typical_gap = sorted(gaps)[len(gaps) // 2] if gaps else (vp_w if vp_w > 0 else 60000)

    # Elevation names from layouts
    layout_elevs: list[float] = []
    for layout in doc.layouts:
        name = layout.name.strip()
        if name != "Model" and re.match(r"[+]?\d+[.,]\d+$", name):
            layout_elevs.append(float(re.match(r"[+]?(\d+[.,]\d+)", name).group(1).replace(",", ".")))
    layout_elevs.sort()

    # Build sheet list — constrain width using viewport dimensions
    half_vp = vp_w / 2 if vp_w > 0 else typical_gap * 0.45
    sheets: list[dict] = []
    for i, lx in enumerate(legend_xs):
        mid_left = (legend_xs[i - 1] + lx) / 2 if i > 0 else lx - half_vp
        mid_right = (lx + legend_xs[i + 1]) / 2 if i < len(legend_xs) - 1 else lx + half_vp
        x_left = max(mid_left, lx - half_vp)
        x_right = min(mid_right, lx + half_vp)
        elev = layout_elevs[i] if i < len(layout_elevs) else None
        elev_str = f"+{elev:.3f}".replace(".", "_") if elev is not None else f"sheet{i + 1}"
        name = f"{elev}" if elev is not None else elev_str
        sheets.append({
            "x_left": x_left, "x_right": x_right,
            "legend_x": lx, "elev_str": elev_str, "name": name,
        })

    log(f"  [viz] {len(sheets)} sheets (gap={typical_gap:.0f},"
        f" vp={vp_w:.0f}×{vp_h:.0f})")

    # ── Step 2: Collect MTEXT entries for equipment detection ──
    hidden = _hidden_layers(doc)
    entries: list[tuple[str, float, float]] = []
    entry_layers: dict[tuple[float, float], str] = {}
    for e in msp.query("MTEXT"):
        if e.dxf.layer in hidden:
            continue
        clean = _clean_mtext(e.text)
        if clean:
            x, y = e.dxf.insert.x, e.dxf.insert.y
            entries.append((clean, x, y))
            entry_layers[(x, y)] = e.dxf.layer

    # ── Step 3: Compute Y bounds per sheet from entity distribution ──
    _POINT_TYPES = frozenset(("MTEXT", "TEXT", "INSERT", "CIRCLE"))
    sheet_entity_ys: dict[int, list[float]] = defaultdict(list)
    for e in [e for e in msp if e.dxftype() in _POINT_TYPES]:
        x = _entity_x(e)
        y = _entity_y(e)
        if x is None or y is None:
            continue
        for si, sh in enumerate(sheets):
            if sh["x_left"] <= x <= sh["x_right"]:
                if y >= legend_y - 3000:
                    sheet_entity_ys[si].append(y)
                break

    # Median Y range across sheets
    y_ranges: list[float] = []
    for ys in sheet_entity_ys.values():
        if len(ys) >= 4:
            ys_s = sorted(ys)
            n = len(ys_s)
            y_ranges.append(ys_s[min(n - 1, int(n * 0.99))] - ys_s[max(0, int(n * 0.01))])
    median_yr = sorted(y_ranges)[len(y_ranges) // 2] if y_ranges else (vp_h if vp_h > 0 else 80000)

    # Set Y bounds per sheet
    sheet_bounds: dict[int, tuple[float, float, float, float]] = {}
    for si, sh in enumerate(sheets):
        ys = sorted(sheet_entity_ys.get(si, []))
        if len(ys) < 4:
            continue
        n = len(ys)
        y_lo = ys[max(0, int(n * 0.01))]
        y_hi = ys[min(n - 1, int(n * 0.99))]
        y_rng = y_hi - y_lo

        # Trim outlier sheets
        if median_yr > 0 and y_rng > median_yr * 2.5:
            cluster = list(ys)
            for _ in range(3):
                if len(cluster) < 6:
                    break
                rng = cluster[-1] - cluster[0]
                if rng <= 0:
                    break
                bg, bi = 0.0, -1
                for j in range(len(cluster) - 1):
                    g = cluster[j + 1] - cluster[j]
                    if g > bg:
                        bg, bi = g, j
                if bg < rng * 0.25:
                    break
                lo, hi = cluster[:bi + 1], cluster[bi + 1:]
                cluster = lo if len(lo) >= len(hi) else hi
            y_lo, y_hi = cluster[0], cluster[-1]
            log(f"  [viz] Sheet {si}: Y trimmed {y_rng:.0f} → {y_hi - y_lo:.0f}")

        # Extend down to include legend + room schedule
        y_lo = min(y_lo, legend_y - 5000)
        pad_x = max((sh["x_right"] - sh["x_left"]) * 0.02, 300)
        pad_y = max((y_hi - y_lo) * 0.02, 300)
        sheet_bounds[si] = (
            sh["x_left"] - pad_x, y_lo - pad_y,
            sh["x_right"] + pad_x, y_hi + pad_y,
        )

    # ── Step 4: Full equipment detection and per-sheet filtering ──
    log("  [viz] Detecting equipment positions…")
    all_bboxes = _find_all_legend_bboxes(entries)
    grid_positions = _detect_grid_labels(msp)

    if items is None:
        log("  [viz] No pre-computed items — running process_dxf (may be slow)…")
        items = process_dxf(dxf_path)

    active_items = [it for it in items if it.count > 0 or it.count_ae > 0]
    sym_name_map = {it.symbol: it.name for it in items}

    # MTEXT-based symbol markers
    all_positions = _collect_markers(msp, entries, all_bboxes, active_items,
                                     grid_positions, entry_layers=entry_layers,
                                     log=log)
    total_text = sum(len(pts) for pts in all_positions.values())

    # INSERT block-based equipment markers (switches, panels, etc.)
    plan_x_lo = min(sh["x_left"] for sh in sheets) - 5000
    plan_x_hi = max(sh["x_right"] for sh in sheets) + 5000
    block_markers = _collect_block_markers(
        msp, all_bboxes, plan_x_range=(plan_x_lo, plan_x_hi),
        hidden=hidden, log=log)
    block_by_layer: dict[str, list[tuple[float, float]]] = defaultdict(list)
    for bx, by, bname, layer in block_markers:
        block_by_layer[layer].append((bx, by))

    # MTEXT on lighting layers — individual lamp/group labels on the plan
    light_mtext: list[tuple[float, float]] = []
    for e in msp.query("MTEXT"):
        if e.dxf.layer in hidden:
            continue
        ll = e.dxf.layer.lower()
        if ll not in _LIGHTING_LAYERS:
            continue
        try:
            mx, my = e.dxf.insert.x, e.dxf.insert.y
        except Exception:
            continue
        if not (plan_x_lo <= mx <= plan_x_hi):
            continue
        if _in_any_bbox(mx, my, all_bboxes):
            continue
        if grid_positions and (round(mx), round(my)) in grid_positions:
            continue
        light_mtext.append((mx, my))

    total_extra = len(block_markers) + len(light_mtext)
    log(f"  [viz] {total_text} text markers + {len(block_markers)} block markers"
        f" + {len(light_mtext)} light-layer MTEXT")

    # Merge into all_positions
    for layer, pts in block_by_layer.items():
        key = f"⬤ {layer}"
        all_positions[key] = pts
        sym_name_map[key] = layer
    if light_mtext:
        all_positions["⬤ Светильники"] = light_mtext
        sym_name_map["⬤ Светильники"] = "Рабочее/аварийное освещение"

    sheet_data: list[dict] = []
    for si, sh in enumerate(sheets):
        if si not in sheet_bounds:
            continue
        xl, yb, xr, yt = sheet_bounds[si]

        sheet_positions: dict[str, list[tuple[float, float]]] = {}
        for sym, pts in all_positions.items():
            filt = [(x, y) for x, y in pts if xl <= x <= xr and yb <= y <= yt]
            if filt:
                sheet_positions[sym] = filt

        n_sh = sum(len(v) for v in sheet_positions.values())
        log(f"  [viz] Sheet {sh['name']}: {n_sh} markers / "
            f"{len(sheet_positions)} types")

        sheet_data.append({
            **sh, "bounds": sheet_bounds[si],
            "sym_to_name": sym_name_map,
            "sym_positions": sheet_positions,
        })

    # ── Step 5: Render model space once ──
    log("  [viz] Rendering model space…")
    total_msp = sum(1 for _ in msp)
    safe_entities = [e for e in msp if e.dxftype() not in _SKIP_ENTITY_TYPES]
    log(f"  [viz] {len(safe_entities)} safe entities"
        f" (skipped {total_msp - len(safe_entities)})")

    fig = plt.figure()
    ax = fig.add_axes([0, 0, 1, 1])
    ctx = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    config = Configuration.defaults().with_changes(
        background_policy=BackgroundPolicy.WHITE,
    )
    frontend = Frontend(ctx, backend, config=config)

    skipped = 0
    for entity in safe_entities:
        try:
            props = ctx.resolve_all(entity)
            if props.is_visible:
                frontend.draw_entity(entity, props)
        except Exception:
            skipped += 1
    if skipped:
        log(f"  [viz] {skipped} entities skipped (render errors)")

    ax.set_aspect("equal")
    ax.autoscale()

    # ── Step 6: Save cropped PNG per sheet with markers ──
    base = Path(dxf_path).with_suffix("")
    saved_paths: list[str] = []

    for si, sd in enumerate(sheet_data):
        out_path = str(base) + f"__{sd['elev_str']}.png"
        xl, yb, xr, yt = sd["bounds"]
        dw, dh = xr - xl, yt - yb
        if dw <= 0 or dh <= 0:
            continue

        ax.set_aspect("equal")
        ax.set_xlim(xl, xr)
        ax.set_ylim(yb, yt)

        # ── V-002: Draw equipment markers with clamped radius ──
        marker_r = min(dw, dh) * 0.004
        marker_r = max(50, min(500, marker_r))  # clamp to [50, 500] DXF units
        marker_patches: list = []
        color_idx = 0
        handles: list[mpatches.Patch] = []
        sym_to_name = sd["sym_to_name"]
        sym_positions = sd["sym_positions"]

        for sym, pts in sorted(sym_positions.items()):
            name = sym_to_name.get(sym, sym)
            clr = MARKER_COLORS[color_idx % len(MARKER_COLORS)]
            color_idx += 1
            for px, py in pts:
                circ = plt.Circle(
                    (px, py), marker_r,
                    fill=True, facecolor=clr, edgecolor="black",
                    linewidth=1.8, alpha=0.45, zorder=10,
                )
                ax.add_patch(circ)
                marker_patches.append(circ)
            label = (f"{sym}: {name[:60]} [{len(pts)}]"
                     if name != sym else f"{sym} [{len(pts)}]")
            handles.append(mpatches.Patch(
                facecolor=clr, edgecolor="black", linewidth=0.5,
                label=label,
            ))

        legend_obj = None
        if handles:
            legend_obj = ax.legend(
                handles=handles, loc="upper left",
                fontsize=7, framealpha=0.92, fancybox=True,
                edgecolor="gray", borderpad=0.5,
                handlelength=1.2, handletextpad=0.5,
            )
            legend_obj.set_zorder(20)

        target_px_w = 8000  # V-004: increased for legible text
        fig_w_in = target_px_w / dpi
        fig_h_in = fig_w_in * (dh / dw) if dw > 0 else fig_w_in
        fig.set_size_inches(fig_w_in, fig_h_in)

        fig.savefig(out_path, dpi=dpi, bbox_inches="tight",
                    facecolor="white", pad_inches=0.1)

        for p in marker_patches:
            p.remove()
        if legend_obj:
            legend_obj.remove()

        n_markers = sum(len(pts) for pts in sym_positions.values())
        size_mb = Path(out_path).stat().st_size / 1024 / 1024
        log(f"  [viz] {sd['name']} ({sd['elev_str']}): {size_mb:.1f} MB,"
            f" {n_markers} markers / {len(sym_positions)} types")
        saved_paths.append(out_path)

    plt.close(fig)

    elapsed = time.time() - t0
    log(f"  [viz] Done: {len(saved_paths)} PNGs in {elapsed:.1f}s")
    return ";".join(saved_paths) if saved_paths else ""


# ── CLI ──────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Render DXF to PNG with equipment markers"
    )
    parser.add_argument("dxf", help="Path to DXF file")
    parser.add_argument("-o", "--output", default=None, help="Output PNG path")
    parser.add_argument("--dpi", type=int, default=200, help="Image DPI (default 200)")
    args = parser.parse_args()

    dxf_path = Path(args.dxf)
    if not dxf_path.exists():
        sys.exit(f"File not found: {dxf_path}")
    if dxf_path.suffix.lower() != ".dxf":
        sys.exit(f"Not a DXF file: {dxf_path}")

    png = visualize_dxf(str(dxf_path), args.output, dpi=args.dpi)
    print(f"\nDone: {png}")


if __name__ == "__main__":
    main()
