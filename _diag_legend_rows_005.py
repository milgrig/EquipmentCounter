"""Print legend item bboxes + count of red/blue strokes inside extended row bands."""
from __future__ import annotations
import io, sys
from collections import Counter, defaultdict
from pathlib import Path
import fitz

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
print(f"Legend bbox: {legend.legend_bbox}")
LEG_X0, LEG_Y0, LEG_X1, LEG_Y1 = legend.legend_bbox

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

def color_class(c):
    if c is None: return None
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b = c[0],c[1],c[2]
        if r>0.6 and g<0.4 and b<0.4: return "red"
        if r<0.4 and g<0.4 and b>0.6: return "blue"
        if r<0.15 and g<0.15 and b<0.15: return "black"
    return None

def dbox(d):
    xs,ys=[],[]
    for it in d.get("items",[]):
        if it[0]=="re":
            r=it[1]; xs+=[r.x0,r.x1]; ys+=[r.y0,r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None

strokes=[]
for d in mp.get_drawings():
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col is None: continue
    bb = dbox(d)
    if bb is None: continue
    if max(bb[2]-bb[0],bb[3]-bb[1])>50: continue
    strokes.append({"bbox":bb,"cx":(bb[0]+bb[2])/2,"cy":(bb[1]+bb[3])/2,"color":col})

print(f"Total coloured strokes: {len(strokes)}")
print()
print("Legend rows (sorted by Y):")
items_sorted = sorted(legend.items, key=lambda it: it.bbox[1])
for it in items_sorted:
    if not it.symbol or not it.symbol.startswith(("1","2","3","4","5","6","7")):
        continue
    bx0,by0,bx1,by1 = it.bbox
    # check strokes left of bx0, in row band ±5pt
    left = [s for s in strokes if LEG_X0 <= s["cx"] <= bx0 and by0-5 <= s["cy"] <= by1+5]
    print(f"  {it.symbol:>6s}: bbox=({bx0:.0f},{by0:.0f},{bx1:.0f},{by1:.0f}) h={by1-by0:.1f}  "
          f"left_strokes={len(left)} colors={Counter(s['color'] for s in left).most_common()}")
    # also strokes in the WHOLE row width including text (debug)
    full = [s for s in strokes if LEG_X0 <= s["cx"] <= LEG_X1 and by0-2 <= s["cy"] <= by1+2]
    print(f"           full_row_strokes={len(full)} colors={Counter(s['color'] for s in full).most_common()}")
doc.close()
