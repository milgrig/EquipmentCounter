"""B030 diag part 4: try multi-scale match for [16][17][18] outside legend zone."""
from __future__ import annotations
import sys
from pathlib import Path
import cv2
import numpy as np

sys.stdout.reconfigure(encoding="utf-8")

PDF_PATH = Path(
    r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf"
)
from pdf_legend_parser import parse_legend
from pdf_count_visual import (
    _extract_symbol_images, _preprocess_template, _preprocess_color_template,
    _render_page, _match_template_multi,
    SYMBOL_DPI, RENDER_DPI, _build_exclusion_zones, _pt_excluded,
)
import pdfplumber

legend = parse_legend(str(PDF_PATH))
sym_data = _extract_symbol_images(str(PDF_PATH), legend, dpi=SYMBOL_DPI)

dpi_ratio = SYMBOL_DPI / RENDER_DPI
page_bgr = _render_page(str(PDF_PATH), legend.page_index, dpi=RENDER_DPI)
page_gray = cv2.cvtColor(page_bgr, cv2.COLOR_BGR2GRAY)
px_per_pt = RENDER_DPI / 72.0

with pdfplumber.open(str(PDF_PATH)) as pdf:
    pp = pdf.pages[legend.page_index]
    excl = _build_exclusion_zones(pp, pp.lines or [], legend.legend_bbox)

scales_to_test = [0.5, 0.6, 0.7, 0.75, 0.8, 0.9, 1.0, 1.1, 1.25, 1.5, 1.75, 2.0]
target = {16, 17, 18}

for idx, item, img in sym_data:
    if idx not in target:
        continue
    gray_template = _preprocess_template(img, keep_color=getattr(item, "color", "") or "")
    th, tw = gray_template.shape[:2]
    print(f"\n[{idx}] template {th}x{tw} desc='{item.description[:50]}'")

    for sc in scales_to_test:
        raw = _match_template_multi(page_gray, gray_template, 0.70, [sc], [0], dpi_ratio)
        # filter to drawing area
        in_legend = 0
        in_drawing = 0
        for d in raw:
            x_pt = d[0] / px_per_pt
            y_pt = d[1] / px_per_pt
            if _pt_excluded(x_pt, y_pt, excl):
                in_legend += 1
            else:
                in_drawing += 1
        # Top conf in drawing area
        drawing_confs = []
        for d in raw:
            x_pt = d[0] / px_per_pt
            y_pt = d[1] / px_per_pt
            if not _pt_excluded(x_pt, y_pt, excl):
                drawing_confs.append(d[6])
        drawing_confs.sort(reverse=True)
        top5 = [f"{c:.3f}" for c in drawing_confs[:5]]
        print(f"   scale={sc:.2f}: raw={len(raw)} in_legend={in_legend} in_drawing={in_drawing} top5_drawing_conf={top5}")
