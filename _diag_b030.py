"""B030 diagnostic: what happens to legend items [16][17][18] (symbol-less)?

Goals:
  * Confirm _extract_symbol_images returns a non-None crop for each.
  * Show crop dimensions, content_px, whether template survives the size /
    aspect / non-equipment filters in match_symbols build loop.
  * Render the crops to .png so we can eyeball them.
  * For surviving templates, run a quick template-match against the page
    and report raw_detection counts + best confidences.
"""
from __future__ import annotations

import sys
from pathlib import Path

import cv2
import numpy as np

sys.stdout.reconfigure(encoding="utf-8")

PDF_PATH = Path(
    r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf"
)
OUT_DIR = Path("_b030_crops")
OUT_DIR.mkdir(exist_ok=True)

if not PDF_PATH.exists():
    print(f"[ERR] PDF not found: {PDF_PATH}")
    sys.exit(1)

from pdf_legend_parser import parse_legend
from pdf_count_visual import (
    _extract_symbol_images,
    _preprocess_template,
    _preprocess_color_template,
    _render_page,
    _match_template_multi,
    MIN_SYMBOL_SIZE_PX,
    NON_EQUIPMENT_KEYWORDS,
    SYMBOL_DPI,
    RENDER_DPI,
)

legend = parse_legend(str(PDF_PATH))
print(f"Legend page_index={legend.page_index} bbox={legend.legend_bbox}")
print(f"Items: {len(legend.items)}")

target = {16, 17, 18}
sym_data = _extract_symbol_images(str(PDF_PATH), legend, dpi=SYMBOL_DPI)

print(f"\n_extract_symbol_images returned {len(sym_data)} entries")
print(f"[idx] sym_text desc[:40]  img_present  shape   has_content")
for idx, item, img in sym_data:
    if idx not in target:
        continue
    sym = item.symbol or "(empty)"
    desc = (item.description or "")[:40]
    if img is None:
        print(f"  [{idx:2}] {sym!r:10} {desc!r:42}  None")
        continue
    h, w = img.shape[:2]
    gray_template = _preprocess_template(img, keep_color=getattr(item, "color", "") or "")
    th, tw = gray_template.shape[:2]
    _, binary = cv2.threshold(gray_template, 200, 255, cv2.THRESH_BINARY_INV)
    content_px = int(np.count_nonzero(binary))
    aspect = max(th, tw) / max(min(th, tw), 1)
    desc_lower = item.description.lower()
    nonequip_hit = [kw for kw in NON_EQUIPMENT_KEYWORDS if kw in desc_lower]
    too_small = th < MIN_SYMBOL_SIZE_PX or tw < MIN_SYMBOL_SIZE_PX
    too_long = aspect > 8.0
    skipped_by = []
    if too_small:
        skipped_by.append(f"min_size({th}x{tw}<{MIN_SYMBOL_SIZE_PX})")
    if too_long:
        skipped_by.append(f"aspect({aspect:.1f}>8)")
    if nonequip_hit:
        skipped_by.append(f"nonequip({nonequip_hit})")
    print(f"  [{idx:2}] {sym!r:10} {desc!r:42}  shape={h}x{w}  content_px={content_px}  aspect={aspect:.2f}  skipped_by={skipped_by}")

    # save a PNG so we can eyeball
    cv2.imwrite(str(OUT_DIR / f"item_{idx:02d}_raw.png"), img)
    cv2.imwrite(str(OUT_DIR / f"item_{idx:02d}_gray.png"), gray_template)
    cv2.imwrite(str(OUT_DIR / f"item_{idx:02d}_bin.png"), binary)

# Now check NON_EQUIPMENT_KEYWORDS list
print(f"\nNON_EQUIPMENT_KEYWORDS = {NON_EQUIPMENT_KEYWORDS}")
