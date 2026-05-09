"""Render PDF page with cable polylines overlay.

Visualises the output of pdf_cables_by_method.measure_cables on a PDF
page. Saves a PNG with:
  * the original page rendered at 200 dpi as background
  * detected red polylines highlighted in solid red
  * detected blue polylines highlighted in solid blue
  * circle markers (gofra/truba) drawn as orange squares
  * tray rects (lotok) drawn as green outlined boxes
  * the legend bbox drawn in cyan
  * the title-block bbox drawn in magenta
"""
from __future__ import annotations

import os
import sys

import fitz
import numpy as np
from PIL import Image, ImageDraw

from pdf_cables_by_method import measure_cables


def render_overlay(pdf_path: str, out_path: str,
                   scale_mm_per_pt: float | None = None,
                   page_index: int = 0,
                   dpi: int = 200) -> None:
    rep = measure_cables(pdf_path, page_index=page_index,
                         scale_mm_per_pt=scale_mm_per_pt)
    doc = fitz.open(pdf_path)
    page = doc[page_index]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    draw = ImageDraw.Draw(img)

    def to_px(p):
        return (p[0] * zoom, p[1] * zoom)

    # Polylines
    for pl in rep.polylines:
        col = (255, 0, 0) if pl.color == "red" else (0, 0, 255)
        # Each polyline.points is a flat bag of segment endpoints in
        # appearance order, so points come in pairs (a,b,a,b,...)
        pts = pl.points
        for i in range(0, len(pts) - 1, 2):
            a = to_px(pts[i]); b = to_px(pts[i + 1])
            draw.line([a, b], fill=col, width=4)

    doc.close()
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    img.save(out_path)
    print(f"Saved {out_path}  size={img.size}")
    print(f"Polylines drawn: {len(rep.polylines)}, "
          f"markers: {rep.circle_markers}, trays: {rep.tray_rects}")
    print(f"Scale: {rep.scale_mm_per_pt} mm/pt  ({rep.scale_source})")


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        print("Usage: python _render_cable_overlay.py <pdf> <out.png> "
              "[--scale-mm-per-pt 35.29] [--page N]")
        sys.exit(1)
    pdf, out = args[0], args[1]
    scale = None
    page = 0
    i = 2
    while i < len(args):
        if args[i] == "--scale-mm-per-pt" and i + 1 < len(args):
            scale = float(args[i + 1]); i += 2
        elif args[i] == "--page" and i + 1 < len(args):
            page = int(args[i + 1]); i += 2
        else:
            i += 1
    render_overlay(pdf, out, scale_mm_per_pt=scale, page_index=page)


if __name__ == "__main__":
    main()
