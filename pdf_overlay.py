#!/usr/bin/env python3
"""
PDF Overlay — render PDF pages with equipment markers from DXF analysis.

Takes the original PDF (high-quality rendering) and overlays coloured
circle markers at equipment positions detected from the DXF file.

Coordinate mapping is calibrated by matching "Условные обозначения"
(legend header) text positions between DXF model-space and PDF pages.
The scale is derived from the paper-space layout limits (LIMMIN/LIMMAX)
and the physical paper size.

Usage:
    python pdf_overlay.py drawing.pdf drawing.dxf
    python pdf_overlay.py drawing.pdf drawing.dxf --dpi 200
"""

import argparse
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

import ezdxf
import fitz  # PyMuPDF
import pdfplumber
from PIL import Image, ImageDraw, ImageFont

from equipment_counter import (
    EquipmentItem,
    process_dxf,
    _clean_mtext,
    _find_all_legend_bboxes,
    _detect_grid_labels,
)
from dxf_visualizer import (
    _collect_markers,
    _collect_block_markers,
    _in_any_bbox,
    _hidden_layers,
    _LIGHTING_LAYERS,
    MARKER_COLORS,
)


def _hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    for name in ("arial.ttf", "segoeui.ttf", "tahoma.ttf", "DejaVuSans.ttf"):
        try:
            return ImageFont.truetype(name, size=size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def _collect_all_markers(
    msp,
    doc,
    entries: list[tuple[str, float, float]],
    entry_layers: dict[tuple[float, float], str],
    items: list[EquipmentItem],
    all_bboxes: list[tuple[float, float, float, float]],
    grid_positions: set[tuple[int, int]] | None,
    plan_x_range: tuple[float, float] | None,
    log=print,
) -> tuple[dict[str, list[tuple[float, float]]], dict[str, str]]:
    """Collect all equipment markers from DXF model space."""
    active_items = [it for it in items if it.count > 0 or it.count_ae > 0]
    sym_name_map = {it.symbol: it.name for it in items}
    hidden = _hidden_layers(doc)

    positions = _collect_markers(msp, entries, all_bboxes, active_items,
                                 grid_positions, entry_layers=entry_layers)

    block_markers = _collect_block_markers(
        msp, all_bboxes, plan_x_range=plan_x_range,
        hidden=hidden, log=log)
    block_by_layer: dict[str, list[tuple[float, float]]] = defaultdict(list)
    for bx, by, _bname, layer in block_markers:
        block_by_layer[layer].append((bx, by))

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
        if plan_x_range and not (plan_x_range[0] <= mx <= plan_x_range[1]):
            continue
        if _in_any_bbox(mx, my, all_bboxes):
            continue
        if grid_positions and (round(mx), round(my)) in grid_positions:
            continue
        light_mtext.append((mx, my))

    for layer, pts in block_by_layer.items():
        key = f"⬤ {layer}"
        positions[key] = pts
        sym_name_map[key] = layer
    if light_mtext:
        positions["⬤ Светильники"] = light_mtext
        sym_name_map["⬤ Светильники"] = "Рабочее/аварийное освещение"

    total = sum(len(v) for v in positions.values())
    log(f"  [pdf] {total} markers across {len(positions)} equipment types")
    return positions, sym_name_map


def _draw_legend_panel(
    draw: ImageDraw.ImageDraw,
    entries: list[tuple[str, str, int]],
    img_w: int,
    img_h: int,
) -> None:
    """Draw a semi-transparent legend in the top-left corner."""
    font = _load_font(22)
    line_h = 34
    pad = 18
    swatch = 24

    max_tw = 0
    for _clr, label, count in entries:
        text = f"{label} [{count}]"
        bb = draw.textbbox((0, 0), text, font=font)
        max_tw = max(max_tw, bb[2] - bb[0])

    pw = pad * 3 + swatch + max_tw + pad
    ph = pad * 2 + len(entries) * line_h
    x0, y0 = 30, 30

    try:
        draw.rounded_rectangle(
            [x0, y0, x0 + pw, y0 + ph],
            radius=10, fill=(255, 255, 255, 215),
            outline=(120, 120, 120, 220), width=2,
        )
    except AttributeError:
        draw.rectangle(
            [x0, y0, x0 + pw, y0 + ph],
            fill=(255, 255, 255, 215),
            outline=(120, 120, 120, 220), width=2,
        )

    for i, (clr_hex, label, count) in enumerate(entries):
        rgb = _hex_to_rgb(clr_hex)
        ey = y0 + pad + i * line_h
        draw.rectangle(
            [x0 + pad, ey + 2, x0 + pad + swatch, ey + 2 + swatch],
            fill=(*rgb, 210), outline=(0, 0, 0, 200),
        )
        draw.text(
            (x0 + pad * 2 + swatch, ey),
            f"{label} [{count}]",
            font=font, fill=(0, 0, 0, 255),
        )


# ── Layout data extraction ───────────────────────────────────────────

def _get_layout_info(doc) -> list[dict]:
    """Extract paper-space layout data sorted by elevation.

    Includes the main viewport's paper-space bounds so we can compute
    which area of the PDF page corresponds to the plan drawing.
    """
    infos: list[dict] = []
    for layout in doc.layouts:
        name = layout.name.strip()
        if name == "Model":
            continue
        if not re.match(r"[+]?\d+[.,]\d+$", name):
            continue
        elev = float(
            re.match(r"[+]?(\d+[.,]\d+)", name)
            .group(1).replace(",", ".")
        )
        try:
            limmin = layout.dxf.limmin
            limmax = layout.dxf.limmax
        except AttributeError:
            continue

        paper_range_x = limmax.x - limmin.x
        paper_range_y = limmax.y - limmin.y
        if paper_range_x <= 0 or paper_range_y <= 0:
            continue

        # Find the main (largest) viewport
        best_area, vp_data = 0.0, None
        for vp in layout.query("VIEWPORT"):
            if getattr(vp.dxf, "id", 0) == 1:
                continue
            try:
                pw, ph = vp.dxf.width, vp.dxf.height
                area = pw * ph
                if area > best_area:
                    best_area = area
                    center = vp.dxf.center
                    vp_data = {
                        "ps_cx": center.x, "ps_cy": center.y,
                        "ps_w": pw, "ps_h": ph,
                    }
            except Exception:
                continue

        infos.append({
            "elev": elev,
            "name": name,
            "limmin_x": limmin.x, "limmin_y": limmin.y,
            "paper_range_x": paper_range_x,
            "paper_range_y": paper_range_y,
            "vp": vp_data,
        })
    infos.sort(key=lambda d: d["elev"])
    return infos


# ── Main overlay function ────────────────────────────────────────────

def visualize_on_pdf(
    pdf_path: str,
    dxf_path: str,
    items: list[EquipmentItem] | None = None,
    dpi: int = 200,
    log=print,
) -> list[str]:
    """Overlay equipment markers on rendered PDF pages.

    1. Loads DXF, extracts layout paper-space limits for scale computation.
    2. Detects all equipment positions in model space.
    3. Opens PDF, finds large plan pages with legend text for calibration.
    4. For each matched sheet: computes transform, renders page, overlays
       markers, saves PNG.

    Returns list of saved PNG paths.
    """
    t0 = time.time()

    # ── 1. Load DXF ──
    log("  [pdf] Loading DXF…")
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    # Legend positions for calibration
    dxf_legends: list[tuple[float, float]] = []
    for e in msp:
        if e.dxftype() == "MTEXT":
            plain = e.plain_text().replace("\n", " ").strip()
            if "Условн" in plain and "обозначен" in plain.lower():
                dxf_legends.append((e.dxf.insert.x, e.dxf.insert.y))
    dxf_legends.sort(key=lambda p: p[0])
    log(f"  [pdf] {len(dxf_legends)} DXF legends")

    # Layout paper limits (for scale computation)
    layout_info = _get_layout_info(doc)
    log(f"  [pdf] {len(layout_info)} elevation layouts")

    # ── 2. Equipment detection ──
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

    if items is None:
        log("  [pdf] Running equipment detection…")
        items = process_dxf(dxf_path)

    plan_x_range = None
    if dxf_legends:
        plan_x_range = (dxf_legends[0][0] - 30000,
                        dxf_legends[-1][0] + 30000)

    all_positions, sym_name_map = _collect_all_markers(
        msp, doc, entries, entry_layers, items, all_bboxes, grid_positions,
        plan_x_range, log=log,
    )

    # Sheet X boundaries (to filter markers per page)
    legend_xs = [lx for lx, _ly in dxf_legends]
    sheet_x_bounds: list[tuple[float, float]] = []
    for i, lx in enumerate(legend_xs):
        if i > 0:
            x_left = (legend_xs[i - 1] + lx) / 2
        else:
            gap = (legend_xs[1] - lx) if len(legend_xs) > 1 else 60000
            x_left = lx - gap * 0.55
        if i < len(legend_xs) - 1:
            x_right = (lx + legend_xs[i + 1]) / 2
        else:
            gap = (lx - legend_xs[-2]) if len(legend_xs) > 1 else 60000
            x_right = lx + gap * 0.55
        sheet_x_bounds.append((x_left, x_right))

    # ── 3. PDF: find plan pages + legend positions ──
    log("  [pdf] Scanning PDF pages…")
    plumber = pdfplumber.open(pdf_path)

    plan_pages: list[dict] = []
    for i, p in enumerate(plumber.pages):
        if p.width < 2000 or p.height < 3000:
            continue
        words = p.extract_words(x_tolerance=3, y_tolerance=3)
        legend_pos = None
        for w in words:
            if "Условн" in w["text"]:
                legend_pos = (w["x0"], w["top"])
                break
        plan_pages.append({
            "idx": i,
            "legend_pdf": legend_pos,
            "w": p.width, "h": p.height,
        })
    plumber.close()
    log(f"  [pdf] {len(plan_pages)} plan pages")

    # ── 4. Match legends → pages ──
    matched: list[tuple[int, dict]] = []
    li = 0
    for pp in plan_pages:
        if pp["legend_pdf"] is None:
            continue
        if li >= len(dxf_legends):
            break
        matched.append((li, pp))
        li += 1
    log(f"  [pdf] {len(matched)} pages matched")

    # ── 5. Render each page with markers ──
    pdf_fitz = fitz.open(pdf_path)
    zoom = dpi / 72.0
    base = Path(dxf_path).with_suffix("")
    saved: list[str] = []

    for li, pp in matched:
        dxf_lx, dxf_ly = dxf_legends[li]
        pdf_lx, pdf_ly = pp["legend_pdf"]
        page_w_pt, page_h_pt = pp["w"], pp["h"]

        # Scale from layout paper limits
        if li < len(layout_info):
            lo = layout_info[li]
            Sx = page_w_pt / lo["paper_range_x"]
            Sy = -page_h_pt / lo["paper_range_y"]
        else:
            Sx = page_w_pt / 84100.0
            Sy = -page_h_pt / 118900.0

        # Calibrate offset from legend match
        Tx = pdf_lx - dxf_lx * Sx
        Ty = pdf_ly - dxf_ly * Sy

        # Viewport rectangle on the PDF page (in PDF points)
        vp_clip = None
        if li < len(layout_info) and layout_info[li].get("vp"):
            lo = layout_info[li]
            vp = lo["vp"]
            vp_pdf_left = (vp["ps_cx"] - vp["ps_w"] / 2 - lo["limmin_x"]) / lo["paper_range_x"] * page_w_pt
            vp_pdf_right = (vp["ps_cx"] + vp["ps_w"] / 2 - lo["limmin_x"]) / lo["paper_range_x"] * page_w_pt
            vp_pdf_top = page_h_pt - (vp["ps_cy"] + vp["ps_h"] / 2 - lo["limmin_y"]) / lo["paper_range_y"] * page_h_pt
            vp_pdf_bottom = page_h_pt - (vp["ps_cy"] - vp["ps_h"] / 2 - lo["limmin_y"]) / lo["paper_range_y"] * page_h_pt
            pad_pt = min(page_w_pt, page_h_pt) * 0.03
            vp_clip = (vp_pdf_left - pad_pt, vp_pdf_top - pad_pt,
                        vp_pdf_right + pad_pt, vp_pdf_bottom + pad_pt)
            log(f"  [pdf] Sheet {li}: viewport clip ({vp_pdf_left:.0f},{vp_pdf_top:.0f})"
                f"-({vp_pdf_right:.0f},{vp_pdf_bottom:.0f}) pts")
        else:
            log(f"  [pdf] Sheet {li}: no viewport clip")

        pw_px = int(page_w_pt * zoom)
        ph_px = int(page_h_pt * zoom)

        # Render PDF page
        page = pdf_fitz[pp["idx"]]
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

        # Transparent overlay for markers
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)

        marker_r = max(25, min(50, pw_px // 150))
        outline_w = max(4, marker_r // 6)
        sheet_xl, sheet_xr = sheet_x_bounds[li] if li < len(sheet_x_bounds) else (-1e18, 1e18)
        ci = 0
        legend_entries: list[tuple[str, str, int]] = []

        for sym in sorted(all_positions.keys()):
            pts = all_positions[sym]
            name = sym_name_map.get(sym, sym)
            clr = MARKER_COLORS[ci % len(MARKER_COLORS)]
            rgb = _hex_to_rgb(clr)
            ci += 1

            drawn = 0
            for mx, my in pts:
                if not (sheet_xl <= mx <= sheet_xr):
                    continue
                px_pdf = mx * Sx + Tx
                py_pdf = my * Sy + Ty

                if vp_clip:
                    if not (vp_clip[0] <= px_pdf <= vp_clip[2]
                            and vp_clip[1] <= py_pdf <= vp_clip[3]):
                        continue

                ix = px_pdf * zoom
                iy = py_pdf * zoom

                if 0 <= ix <= pw_px and 0 <= iy <= ph_px:
                    # White halo for contrast
                    hr = marker_r + outline_w
                    draw.ellipse(
                        [ix - hr, iy - hr, ix + hr, iy + hr],
                        fill=None,
                        outline=(255, 255, 255, 210),
                        width=outline_w + 4,
                    )
                    # Colored circle
                    draw.ellipse(
                        [ix - marker_r, iy - marker_r,
                         ix + marker_r, iy + marker_r],
                        fill=None,
                        outline=(*rgb, 240),
                        width=outline_w,
                    )
                    drawn += 1

            if drawn:
                lbl = f"{sym}: {name[:55]}" if name != sym else sym
                legend_entries.append((clr, lbl, drawn))

        total_page = sum(c for _, _, c in legend_entries)

        if legend_entries:
            _draw_legend_panel(draw, legend_entries, pw_px, ph_px)

        img_rgba = img.convert("RGBA")
        result = Image.alpha_composite(img_rgba, overlay).convert("RGB")

        elev = layout_info[li]["elev"] if li < len(layout_info) else None
        elev_str = (f"+{elev:.3f}".replace(".", "_")
                    if elev is not None else f"sheet{li + 1}")
        out_path = str(base) + f"__pdf_{elev_str}.png"
        result.save(out_path, dpi=(dpi, dpi))

        sz = Path(out_path).stat().st_size / 1024 / 1024
        log(f"  [pdf] Page {pp['idx']} ({elev_str}): "
            f"{total_page} markers, {sz:.1f} MB")
        saved.append(out_path)

    pdf_fitz.close()

    elapsed = time.time() - t0
    log(f"  [pdf] Done: {len(saved)} overlay PNGs in {elapsed:.1f}s")
    return saved


# ── CLI ──────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Overlay equipment markers on PDF page images",
    )
    parser.add_argument("pdf", help="Path to the original PDF file")
    parser.add_argument("dxf", help="Path to the corresponding DXF file")
    parser.add_argument("--dpi", type=int, default=200)
    args = parser.parse_args()

    pdf_p = Path(args.pdf)
    dxf_p = Path(args.dxf)
    if not pdf_p.exists():
        sys.exit(f"PDF not found: {pdf_p}")
    if not dxf_p.exists():
        sys.exit(f"DXF not found: {dxf_p}")

    results = visualize_on_pdf(str(pdf_p), str(dxf_p), dpi=args.dpi)
    print(f"\nDone: {len(results)} PNGs")
    for p in results:
        print(f"  {p}")


if __name__ == "__main__":
    main()
