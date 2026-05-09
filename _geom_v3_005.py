"""
Geometric pictogram detector v3.

Key fixes from v2:
  * Cluster strokes WITHIN one legend row (constrained by row Y-range)
    so that adjacent rows don't merge into one super-cluster.
  * Cluster strokes globally on the plan with small eps (0.5 pt) AND
    require same color.
  * Use 'fill' or 'stroke' color (whichever is set).
  * Allow rotated matching and a wider size tolerance (because the
    legend pictogram may be drawn at a different scale than on the plan).

Design:
  1. Read all "color drawings" (color or fill is non-grey) on the page.
  2. For each legend row:
        - constrain to its X-range (left of text) and Y-range (row band)
        - cluster strokes by color and proximity → 1..N pictogram instances
        - dedupe by color (one canonical rep per color)
  3. On the plan: cluster all strokes outside legend_bbox.
  4. Match each plan cluster to the best legend descriptor (color + size).
  5. Cross-check vs DXF and text counts.
"""
from __future__ import annotations
import io, sys, re
from collections import Counter, defaultdict
from pathlib import Path
from statistics import median

import pdfplumber
import fitz
import numpy as np
import cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_geom_v3_005_out")
OUT.mkdir(exist_ok=True)

DPI = 300
SIZE_TOL = 0.50   # generous to allow legend ≠ plan scale
SIM_THRESH = 0.30
DXF_GT = {"1":33, "2":15, "3":7, "4":4, "5АЭ":4, "6АЭ":6, "7АЭ":7}

# ---- color helpers ----
def color_class(c):
    if c is None: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if r < 0.4 and g < 0.4 and b > 0.6: return "blue"
        if r > 0.85 and g > 0.85 and b > 0.85: return None  # white = noise
        if r < 0.15 and g < 0.15 and b < 0.15: return "black"
        # mid-grey, beige etc → ignore (background hatching)
        return None
    return None

def drawing_bbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]
            xs.extend([r.x0, r.x1]); ys.extend([r.y0, r.y1])
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"):
                    xs.append(p.x); ys.append(p.y)
    if not xs: return None
    return (min(xs), min(ys), max(xs), max(ys))

# ---- load ----
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
print(f"Legend: items={len(legend.items)}, bbox={legend.legend_bbox}")
LEG_X0, LEG_Y0, LEG_X1, LEG_Y1 = legend.legend_bbox

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

# ---- collect coloured drawings (red/blue/black only) ----
strokes = []
for d in mp.get_drawings():
    # prefer fill colour, else stroke colour
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col is None: continue
    bbox = drawing_bbox(d)
    if bbox is None: continue
    w = bbox[2]-bbox[0]; h = bbox[3]-bbox[1]
    if w <= 0 or h <= 0: continue
    if max(w, h) > 50: continue       # huge = page frame
    if max(w, h) < 0.3: continue      # tiny = noise
    strokes.append({
        "bbox": bbox, "cx":(bbox[0]+bbox[2])/2, "cy":(bbox[1]+bbox[3])/2,
        "w": w, "h": h, "color": col,
    })
print(f"Coloured strokes: {len(strokes)}")
print(f"By color: {Counter(s['color'] for s in strokes).most_common()}")
print()

# ---- render page ----
zoom = DPI/72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)

def render_crop(bbox, pad=1.0):
    x0,y0,x1,y1 = bbox
    px0=max(0,int((x0-pad)*zoom)); py0=max(0,int((y0-pad)*zoom))
    px1=min(img.shape[1],int((x1+pad)*zoom)); py1=min(img.shape[0],int((y1+pad)*zoom))
    if px1<=px0 or py1<=py0: return None
    return img[py0:py1, px0:px1].copy()

def normalize(crop, target=(80, 40)):
    if crop is None or crop.size==0: return None
    return cv2.resize(crop, target, interpolation=cv2.INTER_AREA)

def similarity(a, b):
    if a is None or b is None or a.shape != b.shape: return 0.0
    ag = cv2.cvtColor(a, cv2.COLOR_RGB2GRAY).astype(np.float32)
    bg = cv2.cvtColor(b, cv2.COLOR_RGB2GRAY).astype(np.float32)
    ag = (ag-ag.mean())/(ag.std()+1e-6)
    bg = (bg-bg.mean())/(bg.std()+1e-6)
    return float(np.clip((ag*bg).mean(), 0.0, 1.0))

# ---- cluster strokes by color + proximity (Union-Find) ----
def cluster_strokes(strokes, eps=1.0):
    n = len(strokes)
    if n == 0: return []
    parent = list(range(n))
    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]; i = parent[i]
        return i
    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj: parent[ri] = rj
    # spatial bin to avoid O(n^2)
    bins = defaultdict(list)
    for i, s in enumerate(strokes):
        bx = int(s["cx"]/5); by = int(s["cy"]/5)
        bins[(bx,by)].append(i)
    for (bx,by), ids in bins.items():
        # check 3x3 neighbourhood
        for dbx in (-1,0,1):
            for dby in (-1,0,1):
                neigh = bins.get((bx+dbx, by+dby))
                if not neigh: continue
                for i in ids:
                    a = strokes[i]["bbox"]
                    for j in neigh:
                        if j <= i: continue
                        if strokes[i]["color"] != strokes[j]["color"]: continue
                        b = strokes[j]["bbox"]
                        if (a[2]+eps>=b[0] and b[2]+eps>=a[0] and
                            a[3]+eps>=b[1] and b[3]+eps>=a[1]):
                            union(i,j)
    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(strokes[i])
    return list(groups.values())

def make_pictogram(group):
    xs0=[g["bbox"][0] for g in group]; ys0=[g["bbox"][1] for g in group]
    xs1=[g["bbox"][2] for g in group]; ys1=[g["bbox"][3] for g in group]
    bbox = (min(xs0), min(ys0), max(xs1), max(ys1))
    return {
        "bbox": bbox,
        "cx":(bbox[0]+bbox[2])/2, "cy":(bbox[1]+bbox[3])/2,
        "w": bbox[2]-bbox[0], "h": bbox[3]-bbox[1],
        "color": group[0]["color"], "n_parts": len(group),
    }

# ---- Build descriptors per legend row, constrained Y-band ----
descriptors = []
print("=== Legend descriptors (per-row clustering) ===")
print(f"{'sym':>6s}{'col':>6s}  {'W':>6s} x {'H':>6s}  parts  desc")
for i, item in enumerate(legend.items):
    bx0, by0, bx1, by1 = item.bbox
    # left-of-row pictogram band
    sx0 = LEG_X0
    sx1 = bx0
    # next row Y for upper boundary
    next_top = LEG_Y1
    for j, oth in enumerate(legend.items):
        if j == i: continue
        if oth.bbox[1] > by1 and oth.bbox[1] < next_top:
            next_top = oth.bbox[1]
    prev_bot = LEG_Y0
    for j, oth in enumerate(legend.items):
        if j == i: continue
        if oth.bbox[3] < by0 and oth.bbox[3] > prev_bot:
            prev_bot = oth.bbox[3]
    # row band: between prev and next neighbour
    row_y0 = (prev_bot + by0) / 2
    row_y1 = (by1 + next_top) / 2

    in_row = [s for s in strokes
              if sx0 <= s["cx"] <= sx1 and row_y0 <= s["cy"] <= row_y1]
    if not in_row: continue

    groups = cluster_strokes(in_row, eps=1.0)
    by_color = defaultdict(list)
    for g in groups:
        p = make_pictogram(g)
        by_color[p["color"]].append(p)
    for col, lst in by_color.items():
        lst.sort(key=lambda p: p["w"]*p["h"], reverse=True)
        best = lst[0]
        crop = render_crop(best["bbox"])
        norm = normalize(crop)
        safe = re.sub(r"[^\w]", "_", f"{item.symbol or 'X'}_{col}")
        if crop is not None:
            cv2.imwrite(str(OUT/f"legend_{safe}.png"), cv2.cvtColor(crop, cv2.COLOR_RGB2BGR))
        descriptors.append({
            "symbol": item.symbol or f"x{i}",
            "desc": (item.description or "")[:50],
            "w": best["w"], "h": best["h"],
            "color": col, "n_parts": best["n_parts"],
            "norm_ref": norm, "bbox": best["bbox"],
        })
        print(f"{item.symbol or '—':>6s}{col:>6s}  {best['w']:>6.2f} x {best['h']:>6.2f}  {best['n_parts']:>5d}  {(item.description or '')[:50]}")
print()

# ---- Plan strokes (outside legend) ----
plan_strokes = [s for s in strokes
                if not (LEG_X0-2 <= s["cx"] <= LEG_X1+2 and LEG_Y0-2 <= s["cy"] <= LEG_Y1+2)]
plan_clusters = cluster_strokes(plan_strokes, eps=0.5)
plan_picts = [make_pictogram(g) for g in plan_clusters]
print(f"Plan pictograms: {len(plan_picts)}")
print(f"By color: {Counter(p['color'] for p in plan_picts).most_common()}")
print()

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
        # parts similarity
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
print(f"Plan pictograms classified: {len(assignments)} / {len(plan_picts)}")
print()

# ---- Aggregate ----
counts = Counter(); sims = defaultdict(list)
for p, d, s in assignments:
    counts[d["symbol"]] += 1; sims[d["symbol"]].append(s)

# Group by symbol+color (since '1'+blue and '1'+red are different physical fixtures)
counts_by_full = Counter(); sims_by_full = defaultdict(list)
for p, d, s in assignments:
    key = f"{d['symbol']}/{d['color']}"
    counts_by_full[key] += 1; sims_by_full[key].append(s)

print("=== Per-descriptor counts vs DXF ===")
print(f"{'sym':>5s}{'col':>5s}  {'GEO':>4s}  avg_score  desc")
for d in descriptors:
    key = f"{d['symbol']}/{d['color']}"
    geo = counts_by_full.get(key, 0)
    avg = sum(sims_by_full[key])/len(sims_by_full[key]) if sims_by_full[key] else 0.0
    print(f"{d['symbol']:>5s}{d['color']:>5s}  {geo:>4d}  {avg:>9.3f}  {d['desc']}")
print()

# ---- Text control ----
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

print("=== Comparison (text-control + GEO + DXF) ===")
print(f"{'DXF#':>5s}  {'text(А+АЭ)':>11s}  {'GEO':>4s}  {'DXF':>4s}")
geo_per_dxf = Counter()
# aggregate geo per DXF symbol: 1 = (1+1А+1АЭ summed by color: blue → 1А (working), red → 1АЭ (emergency))
for sym, gt in DXF_GT.items():
    if sym.endswith("АЭ"):
        text_total = text_counts.get(sym, 0)
        # In legend, '5АЭ'/'6АЭ'/'7АЭ' have a 'red' descriptor.
        geo = counts.get(sym, 0)
    else:
        text_total = text_counts.get(f"{sym}А", 0) + text_counts.get(f"{sym}АЭ", 0)
        # symbol '1' descriptor produces both blue + red picts
        geo = counts.get(sym, 0)
    print(f"{sym:>5s}  {text_total:>11d}  {geo:>4d}  {gt:>4d}")
print()

# ---- Overlay ----
overlay = img.copy()
sym_color_map = {
    "1":(255,0,255),"2":(0,255,255),"3":(255,128,0),"4":(128,0,255),
    "5АЭ":(0,255,0),"6АЭ":(0,200,200),"7АЭ":(200,0,200),
}
for p, d, s in assignments:
    x0,y0,x1,y1 = p["bbox"]
    color = sym_color_map.get(d["symbol"] or "", (200,200,200))
    cv2.rectangle(overlay,
                  (int(x0*zoom)-2, int(y0*zoom)-2),
                  (int(x1*zoom)+2, int(y1*zoom)+2), color, 2)
cv2.imwrite(str(OUT/"overlay.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
print(f"Overlay: {OUT/'overlay.png'}")
print(f"Legend crops: {OUT}")

doc.close()
