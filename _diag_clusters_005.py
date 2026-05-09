"""Diagnose red cluster size/n distribution to tune detector thresholds."""
from __future__ import annotations
import io, sys, json
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
doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

def color_class(c):
    if c is None: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if r < 0.4 and g < 0.4 and b > 0.6: return "blue"
    return None

def dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"): xs.append(p.x); ys.append(p.y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None

prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None: continue
    w, h = bb[2]-bb[0], bb[3]-bb[1]
    if max(w,h) > 60: continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col != "red": continue
    cx, cy = (bb[0]+bb[2])/2, (bb[1]+bb[3])/2
    if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
    prims.append({"bbox": bb, "cx": cx, "cy": cy, "w": w, "h": h})

print(f"Red plan primitives: {len(prims)}")

def cluster(prims, link):
    n = len(prims); parent = list(range(n))
    def find(a):
        while parent[a]!=a:
            parent[a]=parent[parent[a]]; a=parent[a]
        return a
    def union(a,b):
        ra,rb=find(a),find(b)
        if ra!=rb: parent[ra]=rb
    bins=defaultdict(list)
    for i,p in enumerate(prims):
        bins[(int(p["cx"]//link), int(p["cy"]//link))].append(i)
    for (bx,by), idxs in bins.items():
        for dx in (-1,0,1):
            for dy in (-1,0,1):
                ne = bins.get((bx+dx,by+dy),[])
                for i in idxs:
                    pi=prims[i]
                    for j in ne:
                        if j<=i: continue
                        pj=prims[j]
                        if abs(pi["cx"]-pj["cx"])<=link and abs(pi["cy"]-pj["cy"])<=link:
                            union(i,j)
    g=defaultdict(list)
    for i in range(n): g[find(i)].append(prims[i])
    return list(g.values())

for link in (2.0, 3.5, 5.0, 7.0):
    cls = cluster(prims, link)
    sizes = sorted([len(c) for c in cls], reverse=True)
    # bucket by n_parts
    buckets = {"1-5":0, "6-15":0, "16-30":0, "31-50":0, "51-80":0, "80+":0}
    for n in sizes:
        if n <= 5: buckets["1-5"] += 1
        elif n <= 15: buckets["6-15"] += 1
        elif n <= 30: buckets["16-30"] += 1
        elif n <= 50: buckets["31-50"] += 1
        elif n <= 80: buckets["51-80"] += 1
        else: buckets["80+"] += 1
    print(f"link={link:>4.1f}pt  total_clusters={len(cls):>4d}  top_n={sizes[:8]}  buckets={buckets}")

# Show top 10 clusters at link=5 with bbox+n
print("\nTop clusters at link=5pt:")
cls = cluster(prims, 5.0)
cls_sorted = sorted(cls, key=lambda c: -len(c))
for i, c in enumerate(cls_sorted[:15]):
    xs0=[p["bbox"][0] for p in c]; ys0=[p["bbox"][1] for p in c]
    xs1=[p["bbox"][2] for p in c]; ys1=[p["bbox"][3] for p in c]
    bb=(min(xs0),min(ys0),max(xs1),max(ys1))
    w,h=bb[2]-bb[0],bb[3]-bb[1]
    print(f"  #{i:>2d}  n={len(c):>3d}  W={w:>6.2f}  H={h:>6.2f}  cx={bb[0]+w/2:>7.1f}  cy={bb[1]+h/2:>7.1f}")

doc.close()
