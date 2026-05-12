"""Look for 6АЭ candidates: tall narrow red clusters."""
from __future__ import annotations
import io, sys
from collections import defaultdict
from pathlib import Path
import fitz
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox
doc = fitz.open(str(PDF)); mp = doc[legend.page_index]
def cc(c):
    if not c: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r,g,b=c[0],c[1],c[2]
        if r>0.6 and g<0.4 and b<0.4: return "red"
    return None
def dbox(d):
    xs,ys=[],[]
    for it in d.get("items", []):
        if it[0]=="re":
            r=it[1]; xs+=[r.x0,r.x1]; ys+=[r.y0,r.y1]
        elif it[0] in ("l","m","c"):
            for p in it[1:]:
                if hasattr(p,"x"): xs.append(p.x); ys.append(p.y)
    return (min(xs),min(ys),max(xs),max(ys)) if xs else None
prims=[]
for d in mp.get_drawings():
    bb=dbox(d)
    if not bb: continue
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if max(w,h)>60: continue
    col=cc(d.get("fill")) or cc(d.get("color"))
    if col!="red": continue
    cx,cy=(bb[0]+bb[2])/2,(bb[1]+bb[3])/2
    if LX0-2<=cx<=LX1+2 and LY0-2<=cy<=LY1+2: continue
    prims.append({"bbox":bb,"cx":cx,"cy":cy,"w":w,"h":h})
def cluster(prims, link):
    n=len(prims); par=list(range(n))
    def f(a):
        while par[a]!=a: par[a]=par[par[a]]; a=par[a]
        return a
    def u(a,b):
        ra,rb=f(a),f(b)
        if ra!=rb: par[ra]=rb
    bins=defaultdict(list)
    for i,p in enumerate(prims):
        bins[(int(p["cx"]//link),int(p["cy"]//link))].append(i)
    for (bx,by),idxs in bins.items():
        for dx in (-1,0,1):
            for dy in (-1,0,1):
                ne=bins.get((bx+dx,by+dy),[])
                for i in idxs:
                    pi=prims[i]
                    for j in ne:
                        if j<=i: continue
                        pj=prims[j]
                        if abs(pi["cx"]-pj["cx"])<=link and abs(pi["cy"]-pj["cy"])<=link:
                            u(i,j)
    g=defaultdict(list)
    for i in range(n): g[f(i)].append(prims[i])
    return list(g.values())

cls = cluster(prims, 5.0)
# show all clusters with H>20 OR W between 6-20 and H 20-40
print("Clusters with H >= 18:")
for c in sorted(cls, key=lambda c: -len(c)):
    xs0=[p["bbox"][0] for p in c]; ys0=[p["bbox"][1] for p in c]
    xs1=[p["bbox"][2] for p in c]; ys1=[p["bbox"][3] for p in c]
    bb=(min(xs0),min(ys0),max(xs1),max(ys1))
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if h < 18: continue
    print(f"  n={len(c):>3d}  W={w:>6.2f}  H={h:>6.2f}  cx={(bb[0]+bb[2])/2:>7.1f}  cy={(bb[1]+bb[3])/2:>7.1f}")
print()
print("Clusters with 7<=H<18 and 18<=n<=120 (potential 7AE / small 6AE):")
for c in sorted(cls, key=lambda c: -len(c)):
    xs0=[p["bbox"][0] for p in c]; ys0=[p["bbox"][1] for p in c]
    xs1=[p["bbox"][2] for p in c]; ys1=[p["bbox"][3] for p in c]
    bb=(min(xs0),min(ys0),max(xs1),max(ys1))
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    if not (7<=h<18 and 18<=len(c)<=120): continue
    print(f"  n={len(c):>3d}  W={w:>6.2f}  H={h:>6.2f}  cx={(bb[0]+bb[2])/2:>7.1f}  cy={(bb[1]+bb[3])/2:>7.1f}")
doc.close()
