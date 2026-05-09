"""For each legend row, find ALL small prims in left strip (LX0..bbox.x0+30) regardless of y,
   then print their y-distance to the row center."""
from __future__ import annotations
import io, sys
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
doc=fitz.open(str(PDF)); mp=doc[leg.page_index]
prims=[]
for d in mp.get_drawings():
    bb=dbox(d)
    if not bb: continue
    cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if max(w,h)>60.0: continue
    if not (LY0-5<=cy<=LY1+5): continue
    c=d.get("fill") or d.get("color")
    prims.append({"bbox":bb,"cx":cx,"cy":cy,"red":is_red(c)})
print(f"Prims in legend y-range: {len(prims)},  red: {sum(1 for p in prims if p['red'])}")

print("\nFor each row, prims with x in [LX0, bbox.x0+30]:")
for it in leg.items:
    rx0,ry0,rx1,ry1=it.bbox
    cy_row=(ry0+ry1)/2
    in_strip = [p for p in prims if LX0 <= p["cx"] <= rx0 + 30.0]
    if not in_strip:
        print(f"  {it.symbol:<6} cy={cy_row:.0f}  no prims in x-strip"); continue
    # Bin by y
    in_strip.sort(key=lambda p: abs(p["cy"]-cy_row))
    near5 = [p for p in in_strip if abs(p["cy"]-cy_row) <= 5]
    near15 = [p for p in in_strip if abs(p["cy"]-cy_row) <= 15]
    near30 = [p for p in in_strip if abs(p["cy"]-cy_row) <= 30]
    closest = in_strip[0]
    print(f"  {it.symbol:<6} cy={cy_row:>5.0f}  in_strip={len(in_strip):>4}  "
          f"|d|<=5:{len(near5):>3} <=15:{len(near15):>3} <=30:{len(near30):>3}  "
          f"closest: cx={closest['cx']:.0f} cy={closest['cy']:.0f} dy={closest['cy']-cy_row:+.1f}")
doc.close()
