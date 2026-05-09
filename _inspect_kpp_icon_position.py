"""
For each KPP legend row, scan a WIDE strip around the bbox to find where the icon is.
Search: y in [bbox.y - 5, bbox.y + 5], x in [legend_bbox.x0, bbox.x0 + 20].
"""
from __future__ import annotations
import io, sys
from collections import defaultdict
from pathlib import Path
import fitz

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF = ROOT / r"Data\ДБТ разделы для ИИ\30. КПП\03_PDF\007-План освещения.pdf"
sys.path.insert(0, str(ROOT))
from pdf_legend_parser import parse_legend

def is_red(c):
    if c is None: return False
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b=c[0],c[1],c[2]; return r>0.6 and g<0.4 and b<0.4
    return False

def dbox(d):
    xs,ys=[],[]
    for it in d.get("items",[]):
        if it[0]=="re":
            r=it[1]; xs+=[r.x0,r.x1]; ys+=[r.y0,r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None

leg = parse_legend(str(PDF))
LX0,LY0,LX1,LY1 = leg.legend_bbox
print(f"Legend bbox: ({LX0:.0f},{LY0:.0f},{LX1:.0f},{LY1:.0f})  page={leg.page_index}")

doc=fitz.open(str(PDF)); mp=doc[leg.page_index]
prims=[]
for d in mp.get_drawings():
    bb=dbox(d)
    if not bb: continue
    cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if max(w,h)>60.0: continue
    c=d.get("fill") or d.get("color")
    is_r = is_red(c)
    prims.append({"bbox":bb,"cx":cx,"cy":cy,"red":is_r})
print(f"Total small prims (any color): {len(prims)},  red: {sum(1 for p in prims if p['red'])}")

# For each row, find any prims whose y is within row.y_center +/- 8pt and x < bbox.x0 + 30
print()
print(f"{'sym':<8} {'bbox':<25} | searching x in [LX0={LX0:.0f}, bbox.x0+30] y±8")
for it in leg.items:
    rx0,ry0,rx1,ry1 = it.bbox
    cy_row = (ry0+ry1)/2
    near = [p for p in prims
            if abs(p["cy"]-cy_row) <= 8.0
            and LX0 <= p["cx"] <= rx0 + 30.0]
    red_n = sum(1 for p in near if p["red"])
    if not near:
        print(f"  {it.symbol:<6} ({rx0:.0f},{ry0:.0f},{rx1:.0f},{ry1:.0f})  no prims found")
        continue
    xs=[p["cx"] for p in near]
    print(f"  {it.symbol:<6} ({rx0:.0f},{ry0:.0f})  near={len(near)} red={red_n}  x={min(xs):.0f}..{max(xs):.0f}")
doc.close()
