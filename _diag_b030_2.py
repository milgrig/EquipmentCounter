"""B030 diagnostic part 2: examine template build path + raw match counts."""
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
    SYMBOL_DPI, RENDER_DPI, MIN_SYMBOL_SIZE_PX,
)

legend = parse_legend(str(PDF_PATH))
sym_data = _extract_symbol_images(str(PDF_PATH), legend, dpi=SYMBOL_DPI)

dpi_ratio = SYMBOL_DPI / RENDER_DPI

# Render target page
page_bgr = _render_page(str(PDF_PATH), legend.page_index, dpi=RENDER_DPI)
page_gray = cv2.cvtColor(page_bgr, cv2.COLOR_BGR2GRAY)
print(f"page_gray shape={page_gray.shape}")
print(f"SYMBOL_DPI={SYMBOL_DPI} RENDER_DPI={RENDER_DPI} dpi_ratio={dpi_ratio}")

target = {16, 17, 18}
for idx, item, img in sym_data:
    if idx not in target:
        continue
    cropped, det_color = _preprocess_color_template(img)
    iso_color = getattr(item, "color", "") or det_color
    gray_template = _preprocess_template(img, keep_color=iso_color)
    th, tw = gray_template.shape[:2]

    # page-scale template
    from pdf_count_visual import _scale_image
    page_scale_tpl = _scale_image(gray_template, 1.0 / dpi_ratio)
    psh, psw = page_scale_tpl.shape[:2]
    _, tpl_bin = cv2.threshold(page_scale_tpl, 200, 255, cv2.THRESH_BINARY_INV)
    content_px_at_page = int(np.count_nonzero(tpl_bin))

    print(f"\n[{idx}] template at SYMBOL_DPI={th}x{tw}, at RENDER_DPI={psh}x{psw}, content_px_at_page={content_px_at_page}")
    cv2.imwrite(f"_b030_crops/item_{idx:02d}_pagescale.png", page_scale_tpl)

    # Run raw template match at threshold 0.50 to see raw landscape
    scales = [1.0]
    rotations = [0]
    raw = _match_template_multi(page_gray, gray_template, 0.50, scales, rotations, dpi_ratio)
    confs = sorted([d[6] for d in raw], reverse=True)
    print(f"   raw_detections @ thr=0.50: total={len(raw)} top10_conf={[f'{c:.3f}' for c in confs[:10]]}")

    # Show distribution buckets
    buckets = [0,0,0,0,0,0]
    for c in confs:
        if c >= 0.95: buckets[0]+=1
        elif c >= 0.90: buckets[1]+=1
        elif c >= 0.85: buckets[2]+=1
        elif c >= 0.80: buckets[3]+=1
        elif c >= 0.70: buckets[4]+=1
        else: buckets[5]+=1
    print(f"   buckets: >=0.95:{buckets[0]} 0.90-0.95:{buckets[1]} 0.85-0.90:{buckets[2]} 0.80-0.85:{buckets[3]} 0.70-0.80:{buckets[4]} <0.70:{buckets[5]}")
