"""
Geometric pictogram detector v4 — full-row stroke collection.

Insight from diagnostics:
  parse_legend() splits paired rows ('1' working / '1А' emergency) into
  two LegendItems with identical X-range but consecutive Y-bands.
  Each row's pictogram strokes lie in the SAME Y-band as its text bbox
  (not in a separate left column), occupying the FULL row width.

So: for each row, take all coloured strokes in (LEG_X0..LEG_X1) ×
(by0..by1) and cluster by colour.
"""
from __future__ import annotations
import io, sys, re
from collections import Counter, defaultdict
from pathlib import Path

import pdfplumber
import fitz
import numpy as np
import cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_geom_v4_005_out")
OUT.mkdir(exist_ok=True)

DPI = 300
SIZE_TOL = 0.40
SIM_THRESH = 0.30
DXF_GT = {"1":33, "2":15, "3":7, "4":4, "5АЭ":4, "6АЭ":6, "7АЭ":7}

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

sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LEG_X0, LEG_Y0, LEG_X1, LEG_Y1 = legend.legend_bbox
print(f"Legend: items={len(legend.items)}, bbox={legend.legend_bbox}")

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

# All coloured strokes
strokes = []
for d in mp.get_drawings():
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col is None: continue
    bb = dbox(d)
    if bb is None: continue
    w = bb[2]-bb[0]; h = bb[3]-bb[1]
    if max(w,h) > 50 or max(w,h) < 0.3: continue
    strokes.append({"bbox":bb,"cx":(bb[0]+bb[2])/2,"cy":(bb[1]+bb[3])/2,
                    "w":w,"h":h,"color":col})
print(f"Coloured strokes total: {len(strokes)}")

# Render
zoom=DPI/72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom,zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n==4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)

def render_crop(bbox, pad=1.0):
    x0,y0,x1,y1=bbox
    px0=max(0,int((x0-pad)*zoom)); py0=max(0,int((y0-pad)*zoom))
    px1=min(img.shape[1],int((x1+pad)*zoom)); py1=min(img.shape[0],int((y1+pad)*zoom))
    if px1<=px0 or py1<=py0: return None
    return img[py0:py1, px0:px1].copy()

def normalize(crop, target=(80,40)):
    if crop is None or crop.size==0: return None
    return cv2.resize(crop, target, interpolation=cv2.INTER_AREA)

def similarity(a,b):
    if a is None or b is None or a.shape!=b.shape: return 0.0
    ag=cv2.cvtColor(a,cv2.COLOR_RGB2GRAY).astype(np.float32)
    bg=cv2.cvtColor(b,cv2.COLOR_RGB2GRAY).astype(np.float32)
    ag=(ag-ag.mean())/(ag.std()+1e-6); bg=(bg-bg.mean())/(bg.std()+1e-6)
    return float(np.clip((ag*bg).mean(),0.0,1.0))

# Cluster by color + proximity (UF, binned)
def cluster(items, eps=1.0):
    n=len(items)
    if n==0: return []
    parent=list(range(n))
    def find(i):
        while parent[i]!=i: parent[i]=parent[parent[i]]; i=parent[i]
        return i
    def union(i,j):
        ri,rj=find(i),find(j)
        if ri!=rj: parent[ri]=rj
    bins=defaultdict(list)
    for i,s in enumerate(items):
        bins[(int(s["cx"]/5), int(s["cy"]/5))].append(i)
    for (bx,by), ids in bins.items():
        for dbx in (-1,0,1):
            for dby in (-1,0,1):
                neigh=bins.get((bx+dbx,by+dby))
                if not neigh: continue
                for i in ids:
                    a=items[i]["bbox"]
                    for j in neigh:
                        if j<=i: continue
                        if items[i]["color"]!=items[j]["color"]: continue
                        b=items[j]["bbox"]
                        if (a[2]+eps>=b[0] and b[2]+eps>=a[0] and
                            a[3]+eps>=b[1] and b[3]+eps>=a[1]):
                            union(i,j)
    groups=defaultdict(list)
    for i in range(n): groups[find(i)].append(items[i])
    return list(groups.values())

def make_pict(group):
    xs0=[g["bbox"][0] for g in group]; ys0=[g["bbox"][1] for g in group]
    xs1=[g["bbox"][2] for g in group]; ys1=[g["bbox"][3] for g in group]
    bb=(min(xs0),min(ys0),max(xs1),max(ys1))
    return {"bbox":bb,"cx":(bb[0]+bb[2])/2,"cy":(bb[1]+bb[3])/2,
            "w":bb[2]-bb[0],"h":bb[3]-bb[1],
            "color":group[0]["color"],"n_parts":len(group)}

# ---- Build descriptors using FULL-ROW strokes ----
descriptors = []
print()
print("=== Legend descriptors (full-row stroke collection) ===")
print(f"{'sym':>6s}{'col':>6s}  {'W':>6s} x {'H':>6s}  parts  desc")
for item in legend.items:
    if not item.symbol: continue
    bx0,by0,bx1,by1 = item.bbox
    # full row band: legend X-range, row Y-range (no Y-pad — keep tight to avoid neighbour bleed)
    band = [s for s in strokes
            if LEG_X0 <= s["cx"] <= LEG_X1 and by0-1 <= s["cy"] <= by1+1]
    if not band: continue
    groups = cluster(band, eps=1.0)
    by_col = defaultdict(list)
    for g in groups:
        p = make_pict(g)
        by_col[p["color"]].append(p)
    for col, lst in by_col.items():
        lst.sort(key=lambda p: p["w"]*p["h"], reverse=True)
        best = lst[0]
        crop = render_crop(best["bbox"])
        norm = normalize(crop)
        safe = re.sub(r"[^\w]", "_", f"{item.symbol}_{col}")
        if crop is not None:
            cv2.imwrite(str(OUT/f"legend_{safe}.png"), cv2.cvtColor(crop, cv2.COLOR_RGB2BGR))
        descriptors.append({
            "symbol": item.symbol, "desc": (item.description or "")[:50],
            "w":best["w"],"h":best["h"],"color":col,"n_parts":best["n_parts"],
            "norm_ref":norm,"bbox":best["bbox"],
        })
        print(f"{item.symbol:>6s}{col:>6s}  {best['w']:>6.2f} x {best['h']:>6.2f}  {best['n_parts']:>5d}  {(item.description or '')[:50]}")

# ---- Plan strokes (outside legend bbox) ----
plan = [s for s in strokes
        if not (LEG_X0-2 <= s["cx"] <= LEG_X1+2 and LEG_Y0-2 <= s["cy"] <= LEG_Y1+2)]
plan_clusters = cluster(plan, eps=0.5)
plan_picts = [make_pict(g) for g in plan_clusters if 2 <= len(g) <= 30]
print()
print(f"Plan pictograms: {len(plan_picts)}")
print(f"By color: {Counter(p['color'] for p in plan_picts).most_common()}")

# ---- Match ----
def best_match(p):
    best, best_score = None, 0.0
    crop = render_crop(p["bbox"]); norm = normalize(crop)
    for d in descriptors:
        if d["color"] != p["color"]: continue
        def size_ok(dw, dh):
            return (abs(p["w"]-dw)/max(dw,1e-6) <= SIZE_TOL and
                    abs(p["h"]-dh)/max(dh,1e-6) <= SIZE_TOL)
        rotated = False
        if size_ok(d["w"], d["h"]): pass
        elif size_ok(d["h"], d["w"]): rotated = True
        else: continue
        if rotated and norm is not None and d["norm_ref"] is not None:
            n = cv2.rotate(norm, cv2.ROTATE_90_CLOCKWISE)
            n = cv2.resize(n, (d["norm_ref"].shape[1], d["norm_ref"].shape[0]))
            sim = similarity(d["norm_ref"], n)
        else:
            sim = similarity(d["norm_ref"], norm)
        parts_match = 1.0 - min(abs(p["n_parts"]-d["n_parts"])/max(d["n_parts"],1),1.0)
        score = 0.5*sim + 0.3*parts_match + 0.2*(1.0 - min(abs(p["w"]*p["h"] - d["w"]*d["h"])/max(d["w"]*d["h"],1e-6),1.0))
        if score > best_score:
            best_score = score; best = d
    return best, best_score

assignments = []
for p in plan_picts:
    d, s = best_match(p)
    if d and s >= SIM_THRESH:
        assignments.append((p, d, s))

# Aggregate by (symbol, color) — because '1'+blue ≠ '1'+red
counts_full = Counter(); sims_full = defaultdict(list)
for p, d, s in assignments:
    key = (d["symbol"], d["color"])
    counts_full[key] += 1; sims_full[key].append(s)

print()
print("=== Per-descriptor counts ===")
print(f"{'sym':>6s}{'col':>6s}  {'GEO':>4s}  avg_score  desc")
for d in descriptors:
    key = (d["symbol"], d["color"])
    geo = counts_full.get(key, 0)
    avg = sum(sims_full[key])/len(sims_full[key]) if sims_full[key] else 0.0
    print(f"{d['symbol']:>6s}{d['color']:>6s}  {geo:>4d}  {avg:>9.3f}  {d['desc']}")

# DXF mapping: '1' (blue=working) + '1А' (red=emergency) → DXF '1' total
# Sum every pair (sym, sym+'А' or sym+'АЭ') and any descriptor for that DXF symbol
print()
print("=== Aggregate vs DXF ===")
print(f"{'DXF':>5s}  {'GEO':>4s}  {'DXF_GT':>6s}")
for sym, gt in DXF_GT.items():
    # collect all descriptor keys that map to this DXF symbol
    geo_total = 0
    if sym.endswith("АЭ"):
        # just that symbol
        for (s, c), n in counts_full.items():
            if s == sym: geo_total += n
    else:
        # sym (blue) + symА (red) + symАЭ (red) all count toward DXF[sym]
        targets = {sym, f"{sym}А", f"{sym}АЭ"}
        for (s, c), n in counts_full.items():
            if s in targets: geo_total += n
    print(f"{sym:>5s}  {geo_total:>4d}  {gt:>6d}")

# Text control
LABEL_RE = re.compile(r"^[1-9](?:[АA][ЭE]?)?$")
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    words = page.extract_words() or []
text_counts = Counter()
for w in words:
    t = (w.get("text") or "").strip()
    if not LABEL_RE.match(t): continue
    cx=(w["x0"]+w["x1"])/2; cy=(w["top"]+w["bottom"])/2
    if LEG_X0-2 <= cx <= LEG_X1+2 and LEG_Y0-2 <= cy <= LEG_Y1+2: continue
    text_counts[t] += 1

print()
print("=== Final triple comparison ===")
print(f"{'sym':>5s}  {'TEXT(А+АЭ)':>11s}  {'GEO':>4s}  {'DXF':>4s}")
for sym, gt in DXF_GT.items():
    if sym.endswith("АЭ"):
        tc = text_counts.get(sym, 0)
        geo = sum(n for (s,c),n in counts_full.items() if s == sym)
    else:
        tc = text_counts.get(f"{sym}А", 0) + text_counts.get(f"{sym}АЭ", 0)
        targets = {sym, f"{sym}А", f"{sym}АЭ"}
        geo = sum(n for (s,c),n in counts_full.items() if s in targets)
    print(f"{sym:>5s}  {tc:>11d}  {geo:>4d}  {gt:>4d}")

# Overlay
overlay = img.copy()
sym_color = {"1":(255,0,255),"2":(0,255,255),"3":(255,128,0),"4":(128,0,255),
             "5АЭ":(0,255,0),"6АЭ":(0,200,200),"7АЭ":(200,0,200)}
for p, d, s in assignments:
    x0,y0,x1,y1 = p["bbox"]
    color = sym_color.get(d["symbol"] or "", (200,200,200))
    cv2.rectangle(overlay, (int(x0*zoom)-2,int(y0*zoom)-2), (int(x1*zoom)+2,int(y1*zoom)+2), color, 2)
cv2.imwrite(str(OUT/"overlay.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
print(f"\nOverlay: {OUT/'overlay.png'}")

doc.close()
