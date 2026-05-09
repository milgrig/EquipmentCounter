"""
Cable diagnostic for one PDF.

For each linear primitive on the page (excluding legend bbox):
  - color class (red/blue/black/green/grey)
  - length in pt
  - dash pattern (None = solid, else tuple)
  - line width

We summarize:
  - how many lines per (color, dash) bucket
  - total length per (color, dash) bucket
  - typical lengths (median, max)
"""
from __future__ import annotations
import io, sys
from collections import defaultdict
from pathlib import Path
import math
import fitz

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF = ROOT / r"Data\ДБТ разделы для ИИ\30. КПП\03_PDF\007-План освещения.pdf"

sys.path.insert(0, str(ROOT))
from pdf_legend_parser import parse_legend

def color_name(c):
    if c is None: return "none"
    if isinstance(c, (tuple,list)) and len(c) >= 3:
        r,g,b = c[0],c[1],c[2]
        if max(r,g,b) < 0.2: return "black"
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if b > 0.6 and r < 0.4 and g < 0.5: return "blue"
        if g > 0.5 and r < 0.4 and b < 0.4: return "green"
        if r > 0.4 and g > 0.4 and b > 0.4 and abs(r-g) < 0.15 and abs(g-b) < 0.15:
            return "grey"
        return f"other({r:.1f},{g:.1f},{b:.1f})"
    return "weird"

def line_len(p, q):
    return math.hypot(q.x-p.x, q.y-p.y)

print(f"PDF: {PDF.name}")
leg = parse_legend(str(PDF))
LX0,LY0,LX1,LY1 = leg.legend_bbox
print(f"  legend bbox: ({LX0:.0f},{LY0:.0f},{LX1:.0f},{LY1:.0f})  page={leg.page_index}")

doc = fitz.open(str(PDF)); mp = doc[leg.page_index]
print(f"  page size: {mp.rect.width:.0f} x {mp.rect.height:.0f}")

# bucket by (color, dash_signature)
buckets = defaultdict(lambda: {"n": 0, "total_len": 0.0, "lens": []})
total_drawings = 0
total_l_items = 0

for d in mp.get_drawings():
    total_drawings += 1
    col = color_name(d.get("color"))
    dashes = d.get("dashes")
    width = d.get("width") or 0
    # Normalize dash to a hashable signature
    if dashes is None or dashes == "[] 0":
        dash_sig = "solid"
    else:
        dash_sig = str(dashes).replace(" ", "")

    # Skip drawings inside legend bbox
    skip_legend = False

    for it in d.get("items", []):
        if it[0] != "l": continue
        total_l_items += 1
        p1, p2 = it[1], it[2]
        # midpoint
        mx, my = (p1.x+p2.x)/2, (p1.y+p2.y)/2
        if LX0 <= mx <= LX1 and LY0 <= my <= LY1:
            continue
        L = line_len(p1, p2)
        if L < 0.5: continue
        key = (col, dash_sig)
        b = buckets[key]
        b["n"] += 1
        b["total_len"] += L
        if len(b["lens"]) < 20:
            b["lens"].append(round(L, 2))

doc.close()

print(f"\n  total drawings on page: {total_drawings}")
print(f"  total 'l' items:        {total_l_items}")
print(f"  buckets after legend filter: {len(buckets)}")
print()
print(f"{'color':<14} {'dash':<25} {'count':>6} {'sum_len':>10} {'mean':>7} {'median':>7} {'max':>7}")
print("-"*90)
items = sorted(buckets.items(), key=lambda kv: -kv[1]["total_len"])
for (col, dash), b in items:
    n = b["n"]; tot = b["total_len"]
    if n == 0: continue
    lens = sorted(b["lens"])
    med = lens[len(lens)//2] if lens else 0
    mx = max(lens) if lens else 0
    print(f"{col:<14} {dash:<25} {n:>6} {tot:>10.1f} {tot/n:>7.2f} {med:>7.2f} {mx:>7.2f}")
