"""Inspect KPP plan-007 legend: what symbols, where bboxes, how many red/blue prims inside."""
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
def is_blue(c):
    if c is None: return False
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b=c[0],c[1],c[2]; return b>0.6 and g<0.4 and r<0.4
    return False
def is_black(c):
    if c is None: return False
    if isinstance(c,(tuple,list)) and len(c)>=3:
        r,g,b=c[0],c[1],c[2]; return max(r,g,b)<0.3
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

print(f"PDF: {PDF.name}")
leg = parse_legend(str(PDF))
print(f"  page_index = {leg.page_index}")
print(f"  legend_bbox= {tuple(round(v,1) for v in leg.legend_bbox)}")
print(f"  rows: {len(leg.items)}")
print()
print(f"{'sym':<8} {'cat':<22} {'bbox':<40} {'descr (50ch)':<60}")
for it in leg.items:
    bbstr=f"({it.bbox[0]:.0f},{it.bbox[1]:.0f},{it.bbox[2]:.0f},{it.bbox[3]:.0f})"
    print(f"  {it.symbol:<6} {it.category:<22} {bbstr:<40} {it.description[:58]}")

# now scan the page for primitives in icon ROI of each row
doc=fitz.open(str(PDF)); mp=doc[leg.page_index]
all_prims=[]
for d in mp.get_drawings():
    bb=dbox(d)
    if not bb: continue
    cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if max(w,h)>60.0: continue
    c=d.get("fill") or d.get("color")
    cls = "red" if is_red(c) else ("blue" if is_blue(c) else ("black" if is_black(c) else "other"))
    all_prims.append({"bbox":bb,"cx":cx,"cy":cy,"col":cls})

print(f"\nTotal primitives on page: {len(all_prims)}")
col_counts=defaultdict(int)
for p in all_prims: col_counts[p["col"]]+=1
print(f"  by color: {dict(col_counts)}")

# for each row, count prims in left strip
print(f"\nPer-row icon ROI (leftmost 30pt of bbox):")
for it in leg.items:
    rx0,ry0,rx1,ry1=it.bbox
    ix0,ix1=rx0,rx0+30.0
    iy0,iy1=ry0-2.0,ry1+2.0
    inside=[p for p in all_prims if ix0<=p["cx"]<=ix1 and iy0<=p["cy"]<=iy1]
    by_col=defaultdict(int)
    for p in inside: by_col[p["col"]]+=1
    print(f"  {it.symbol:<6} -> {len(inside):>3} prims  by_color={dict(by_col)}")
doc.close()
