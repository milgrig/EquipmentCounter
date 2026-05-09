"""
Recon script: explore vector rectangles in 005-Планы освещения-отм. 0.000.pdf
Goal: determine if light fixtures are vector primitives we can read directly.
"""
from __future__ import annotations
import io
import sys
from collections import Counter, defaultdict
from pathlib import Path

import pdfplumber
import fitz  # PyMuPDF

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")


def fmt_color(c):
    if c is None:
        return "None"
    if isinstance(c, (tuple, list)):
        return "(" + ",".join(f"{v:.2f}" for v in c) + ")"
    return str(c)


def round_size(w: float, h: float, step: float = 1.0) -> tuple[float, float]:
    return (round(w / step) * step, round(h / step) * step)


print(f"PDF: {PDF.name}")
print(f"Size: {PDF.stat().st_size / 1024 / 1024:.2f} MB")
print()

# ---- pdfplumber: read rects -------------------------------------------------
with pdfplumber.open(str(PDF)) as pdf:
    print(f"Pages: {len(pdf.pages)}")
    for pi, page in enumerate(pdf.pages):
        rects = page.rects or []
        lines = page.lines or []
        curves = page.curves or []
        chars = page.chars or []
        print(f"  page {pi}: rects={len(rects)}  lines={len(lines)}  curves={len(curves)}  chars={len(chars)}")
    print()

    # focus on page with most rects (likely the plan)
    page_idx = max(range(len(pdf.pages)), key=lambda i: len(pdf.pages[i].rects or []))
    page = pdf.pages[page_idx]
    rects = page.rects or []
    print(f"=== Analyzing page {page_idx} ({len(rects)} rects) ===")
    print(f"Page size: {page.width:.1f} x {page.height:.1f} pt")
    print()

    # group by (width, height) rounded to 1 pt
    sizes = Counter()
    for r in rects:
        w = r["width"]
        h = r["height"]
        sizes[round_size(w, h, 1.0)] += 1

    print(f"Unique (W,H) groups: {len(sizes)}")
    print("Top 30 size groups:")
    print(f"  {'W':>7s} x {'H':>7s}  count")
    for (w, h), cnt in sizes.most_common(30):
        print(f"  {w:>7.1f} x {h:>7.1f}  {cnt}")
    print()

    # group by colors for medium-sized rects (potential fixtures)
    print("Color distribution for rects with W*H between 5..10000 sq.pt:")
    color_groups = Counter()
    for r in rects:
        a = r["width"] * r["height"]
        if 5 <= a <= 10000:
            sc = fmt_color(r.get("stroke_color"))
            fc = fmt_color(r.get("non_stroke_color"))
            color_groups[(sc, fc)] += 1
    for (sc, fc), cnt in color_groups.most_common(20):
        print(f"  stroke={sc:30s} fill={fc:30s} {cnt}")
    print()

    # show a sample rect
    if rects:
        r = rects[0]
        print("Sample rect keys:")
        for k, v in r.items():
            print(f"  {k}: {v}")

# ---- PyMuPDF: get_drawings on the same page --------------------------------
print()
print("=== PyMuPDF get_drawings ===")
doc = fitz.open(str(PDF))
mp = doc[page_idx]
drawings = mp.get_drawings()
print(f"Drawings: {len(drawings)}")
type_counter = Counter()
fill_counter = Counter()
stroke_counter = Counter()
rect_only = 0
for d in drawings:
    type_counter[d.get("type", "?")] += 1
    if d.get("fill") is not None:
        fill_counter[fmt_color(d["fill"])] += 1
    if d.get("color") is not None:
        stroke_counter[fmt_color(d["color"])] += 1
    items = d.get("items", [])
    # rectangle == single 're' item
    if len(items) == 1 and items[0][0] == "re":
        rect_only += 1
print(f"Pure rectangle drawings: {rect_only}")
print(f"Drawing types: {dict(type_counter)}")
print(f"Top fills:    {fill_counter.most_common(10)}")
print(f"Top strokes:  {stroke_counter.most_common(10)}")
print()

# show first 5 rect drawings
print("First 5 rectangle drawings:")
shown = 0
for d in drawings:
    items = d.get("items", [])
    if len(items) == 1 and items[0][0] == "re":
        rect = items[0][1]  # fitz.Rect
        print(f"  rect={rect}, fill={fmt_color(d.get('fill'))}, color={fmt_color(d.get('color'))}, "
              f"width={d.get('width')}, type={d.get('type')}")
        shown += 1
        if shown >= 5:
            break

doc.close()
