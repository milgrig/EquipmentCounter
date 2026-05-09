"""Inspect what drawing primitives sit inside each legend row."""
from __future__ import annotations
import io, sys
from collections import Counter
from pathlib import Path
import fitz

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LEG_X0 = legend.legend_bbox[0]

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
drawings = mp.get_drawings()
print(f"Total drawings on page: {len(drawings)}")
print(f"Legend bbox: {legend.legend_bbox}")
print()

def fmt(c):
    if c is None: return "—"
    return "(" + ",".join(f"{v:.2f}" for v in c) + ")"

for item in legend.items:
    bx0, by0, bx1, by1 = item.bbox
    sx0 = LEG_X0
    pad_y = (by1 - by0) * 0.3
    sy0, sy1 = by0 - pad_y, by1 + pad_y
    if not item.symbol: continue
    if item.symbol not in {"1","2","3","4","5АЭ","6АЭ","7АЭ","1А","2А"}: continue

    types = Counter()
    rect_count = 0
    line_count = 0
    pict_xrange = []
    pict_yrange = []
    fills = Counter()
    strokes = Counter()
    for d in drawings:
        items = d.get("items", [])
        if not items: continue
        # get drawing bbox
        x0s, y0s, x1s, y1s = [], [], [], []
        for it in items:
            if it[0] == "re":
                r = it[1]
                x0s.append(r.x0); y0s.append(r.y0); x1s.append(r.x1); y1s.append(r.y1)
            elif it[0] in ("l", "m"):
                p = it[1]
                x0s.append(p.x); y0s.append(p.y); x1s.append(p.x); y1s.append(p.y)
                if len(it) > 2:
                    p2 = it[2]
                    x0s.append(p2.x); y0s.append(p2.y); x1s.append(p2.x); y1s.append(p2.y)
        if not x0s: continue
        cx = (min(x0s) + max(x1s)) / 2
        cy = (min(y0s) + max(y1s)) / 2
        if not (sx0 <= cx <= bx0 and sy0 <= cy <= sy1):
            continue
        # this drawing belongs to pictogram column for this row
        kind = "rect" if (len(items) == 1 and items[0][0] == "re") else f"compound({len(items)})"
        types[kind] += 1
        if d.get("fill"):
            fills[fmt(d["fill"])] += 1
        if d.get("color"):
            strokes[fmt(d["color"])] += 1

    print(f"-- '{item.symbol}': {item.description[:50]}")
    print(f"   drawings in pictogram column: {sum(types.values())}")
    print(f"   types: {dict(types)}")
    print(f"   fills: {fills.most_common(5)}")
    print(f"   strokes: {strokes.most_common(5)}")
    print()

doc.close()
