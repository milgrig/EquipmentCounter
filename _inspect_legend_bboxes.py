"""
Inspect what LegendItem.bbox actually contains for 5АЭ/6АЭ/7АЭ across the 7 lighting plans.
We need to know:
  * Is bbox the full row (symbol+icon+description) or just the icon?
  * Are bboxes consistent across files?
  * How many red primitives fall inside each bbox?
"""
from __future__ import annotations
import io, sys
from collections import defaultdict
from pathlib import Path
import fitz

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF_DIR = ROOT / r"Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF"
PDFS = [
    "005-Планы освещения-отм. 0.000.pdf",
    "006-Планы освещения-отм. +4.200.pdf",
    "007-Планы освещения-отм. +7.800 +9.000.pdf",
    "011-Планы освещения-отм. +28.200.pdf",
]
TARGET = {"5АЭ", "6АЭ", "7АЭ"}

sys.path.insert(0, str(ROOT))
from pdf_legend_parser import parse_legend

def is_red(c):
    if c is None: return False
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b=c[0],c[1],c[2]
        return r>0.6 and g<0.4 and b<0.4
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

for pname in PDFS:
    pdf=PDF_DIR/pname
    print(f"\n========== {pname[:3]} ==========")
    leg=parse_legend(str(pdf))
    print(f"  legend.legend_bbox = {tuple(round(v,1) for v in leg.legend_bbox)}")
    print(f"  page_index         = {leg.page_index}")
    doc=fitz.open(str(pdf)); mp=doc[leg.page_index]
    # collect all red primitives on the page
    prims=[]
    for d in mp.get_drawings():
        bb=dbox(d)
        if not bb: continue
        c=d.get("fill") or d.get("color")
        if not is_red(c): continue
        cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
        w,h=bb[2]-bb[0],bb[3]-bb[1]
        if max(w,h)>60.0: continue
        prims.append({"bbox":bb,"cx":cx,"cy":cy})

    items_targeted = [it for it in leg.items if it.symbol in TARGET]
    print(f"  legend rows of interest: {[it.symbol for it in items_targeted]}")
    for it in items_targeted:
        x0,y0,x1,y1=it.bbox
        in_box=[p for p in prims if x0<=p["cx"]<=x1 and y0<=p["cy"]<=y1]
        print(f"  {it.symbol:<5} bbox=({x0:.1f},{y0:.1f},{x1:.1f},{y1:.1f}) "
              f"W={x1-x0:.1f} H={y1-y0:.1f}  red prims inside: {len(in_box)}")
        # group by sub-region
        if in_box:
            xs=[p["cx"] for p in in_box]; ys=[p["cy"] for p in in_box]
            print(f"         red x-range: {min(xs):.1f}..{max(xs):.1f}  "
                  f"y-range: {min(ys):.1f}..{max(ys):.1f}")
    doc.close()
