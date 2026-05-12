"""
Geometric pictogram detector v2 — based on STROKE drawings (not fills).

Key insight from inspection:
  Lighting fixtures in the legend are stroke-only compound drawings,
  one per fixture. Each row has 4-6 of them (red + blue variants).
  We need to:
    1. cluster drawings into pictogram instances by spatial proximity,
    2. build a descriptor (W, H, stroke_color, item_count) per instance,
    3. compare plan instances against legend instances.
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
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_geom_v2_005_out")
OUT.mkdir(exist_ok=True)

DPI = 300
SIZE_TOL = 0.30
SIM_THRESH = 0.40
DXF_GT = {"1":33, "2":15, "3":7, "4":4, "5АЭ":4, "6АЭ":6, "7АЭ":7}

# ---------- color helpers ----------
def color_class(c):
    if c is None: return "none"
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if r < 0.4 and g < 0.4 and b > 0.6: return "blue"
        if r > 0.85 and g > 0.85 and b > 0.85: return "white"
        if r < 0.15 and g < 0.15 and b < 0.15: return "black"
        return f"o:{r:.1f},{g:.1f},{b:.1f}"
    return "?"

def drawing_bbox(d):
    """Compute bbox of a drawing from its items."""
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]
            xs.extend([r.x0, r.x1]); ys.extend([r.y0, r.y1])
        elif it[0] in ("l", "m"):
            for p in it[1:]:
                if hasattr(p, "x"):
                    xs.append(p.x); ys.append(p.y)
        elif it[0] == "c":
            for p in it[1:]:
                if hasattr(p, "x"):
                    xs.append(p.x); ys.append(p.y)
    if not xs: return None
    return (min(xs), min(ys), max(xs), max(ys))

# ---------- 1. Read drawings ----------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
print(f"Legend: items={len(legend.items)}, bbox={legend.legend_bbox}")
LEG_X0, LEG_Y0, LEG_X1, LEG_Y1 = legend.legend_bbox

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

# Collect "stroke drawings" (compound with color, no fill, small enough)
stroke_drawings = []
for d in mp.get_drawings():
    if d.get("color") is None: continue
    if d.get("fill") is not None: continue
    bbox = drawing_bbox(d)
    if bbox is None: continue
    w = bbox[2] - bbox[0]; h = bbox[3] - bbox[1]
    if w <= 0 or h <= 0: continue
    # too big = page frame; too small = noise dot
    if w > 50 or h > 50: continue
    if w < 0.5 and h < 0.5: continue
    stroke_drawings.append({
        "bbox": bbox,
        "cx": (bbox[0]+bbox[2])/2,
        "cy": (bbox[1]+bbox[3])/2,
        "w": w, "h": h,
        "color": color_class(d["color"]),
    })
print(f"Stroke drawings (filter w,h <50pt): {len(stroke_drawings)}")
print(f"By color: {Counter(d['color'] for d in stroke_drawings).most_common()}")
print()

# ---------- 2. Cluster drawings into PICTOGRAM INSTANCES ----------
# A pictogram is a tight cluster of strokes (e.g. 4-6 strokes within ~3pt).
def cluster(drawings, eps=2.0):
    """Simple grid-based clustering: merge if bbox overlap or within eps."""
    n = len(drawings)
    parent = list(range(n))
    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj: parent[ri] = rj

    for i in range(n):
        for j in range(i+1, n):
            a = drawings[i]["bbox"]; b = drawings[j]["bbox"]
            if drawings[i]["color"] != drawings[j]["color"]: continue
            # overlap or within eps
            if (a[2]+eps >= b[0] and b[2]+eps >= a[0] and
                a[3]+eps >= b[1] and b[3]+eps >= a[1]):
                union(i, j)
    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(drawings[i])
    return list(groups.values())

clusters = cluster(stroke_drawings, eps=1.5)
print(f"Clusters: {len(clusters)}")

# Build pictogram instances: bbox = union, color from majority, n_parts = len(group)
pictograms = []
for grp in clusters:
    xs0 = [g["bbox"][0] for g in grp]; ys0 = [g["bbox"][1] for g in grp]
    xs1 = [g["bbox"][2] for g in grp]; ys1 = [g["bbox"][3] for g in grp]
    bbox = (min(xs0), min(ys0), max(xs1), max(ys1))
    pictograms.append({
        "bbox": bbox,
        "cx": (bbox[0]+bbox[2])/2,
        "cy": (bbox[1]+bbox[3])/2,
        "w": bbox[2]-bbox[0],
        "h": bbox[3]-bbox[1],
        "color": grp[0]["color"],
        "n_parts": len(grp),
    })
print(f"Pictogram instances: {len(pictograms)}")
print(f"By color: {Counter(p['color'] for p in pictograms).most_common()}")
print()

# ---------- 3. Render page ----------
zoom = DPI/72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)

def render_crop(bbox, pad=1.0):
    x0,y0,x1,y1 = bbox
    px0 = max(0, int((x0-pad)*zoom)); py0 = max(0, int((y0-pad)*zoom))
    px1 = min(img.shape[1], int((x1+pad)*zoom)); py1 = min(img.shape[0], int((y1+pad)*zoom))
    if px1 <= px0 or py1 <= py0: return None
    return img[py0:py1, px0:px1].copy()

def normalize(crop, target=(80, 40)):
    if crop is None or crop.size == 0: return None
    return cv2.resize(crop, target, interpolation=cv2.INTER_AREA)

def similarity(a, b):
    if a is None or b is None or a.shape != b.shape: return 0.0
    ag = cv2.cvtColor(a, cv2.COLOR_RGB2GRAY).astype(np.float32)
    bg = cv2.cvtColor(b, cv2.COLOR_RGB2GRAY).astype(np.float32)
    ag = (ag - ag.mean()) / (ag.std() + 1e-6)
    bg = (bg - bg.mean()) / (bg.std() + 1e-6)
    return float(np.clip((ag*bg).mean(), 0.0, 1.0))

# ---------- 4. Build legend descriptors ----------
descriptors = []
print("=== Legend descriptors (per row + color) ===")
print(f"{'sym':>6s}{'fill':>6s}  {'W':>6s} x {'H':>6s}  parts  desc")
for item in legend.items:
    bx0, by0, bx1, by1 = item.bbox
    pad_y = max(2.0, (by1-by0)*0.5)
    sx0 = LEG_X0; sy0, sy1 = by0-pad_y, by1+pad_y
    inside = [p for p in pictograms if (sx0 <= p["cx"] <= bx0+1) and (sy0 <= p["cy"] <= sy1)]
    if not inside: continue

    by_color = defaultdict(list)
    for p in inside: by_color[p["color"]].append(p)
    for col, lst in by_color.items():
        lst.sort(key=lambda p: p["w"]*p["h"], reverse=True)
        best = lst[0]
        crop = render_crop(best["bbox"])
        norm = normalize(crop)
        safe = re.sub(r"[^\w]", "_", f"{item.symbol or 'X'}_{col}")
        if crop is not None:
            cv2.imwrite(str(OUT / f"legend_{safe}.png"), cv2.cvtColor(crop, cv2.COLOR_RGB2BGR))
        descriptors.append({
            "symbol": item.symbol,
            "desc": (item.description or "")[:50],
            "w": best["w"], "h": best["h"],
            "color": col,
            "n_parts": best["n_parts"],
            "norm_ref": norm,
            "bbox": best["bbox"],
        })
        print(f"{item.symbol or '—':>6s}{col:>6s}  {best['w']:>6.2f} x {best['h']:>6.2f}  {best['n_parts']:>5d}  {(item.description or '')[:50]}")
        # mark all in cluster as legend-used
print()

# ---------- 5. Plan pictograms (outside legend bbox) ----------
plan_picts = []
for p in pictograms:
    cx, cy = p["cx"], p["cy"]
    if LEG_X0-2 <= cx <= LEG_X1+2 and LEG_Y0-2 <= cy <= LEG_Y1+2:
        continue
    plan_picts.append(p)
print(f"Plan pictograms (outside legend): {len(plan_picts)}")
print(f"By color: {Counter(p['color'] for p in plan_picts).most_common()}")
print()

# ---------- 6. Classify ----------
def best_match(p):
    best, best_score = None, 0.0
    crop = render_crop(p["bbox"]); norm = normalize(crop)
    for d in descriptors:
        if d["color"] != p["color"]: continue
        # size with optional rotation
        def size_ok(dw, dh):
            return (abs(p["w"] - dw)/max(dw,1e-6) <= SIZE_TOL and
                    abs(p["h"] - dh)/max(dh,1e-6) <= SIZE_TOL)
        rotated = False
        if size_ok(d["w"], d["h"]): pass
        elif size_ok(d["h"], d["w"]): rotated = True
        else: continue
        # similarity
        if rotated and norm is not None and d["norm_ref"] is not None:
            n = cv2.rotate(norm, cv2.ROTATE_90_CLOCKWISE)
            n = cv2.resize(n, (d["norm_ref"].shape[1], d["norm_ref"].shape[0]))
            sim = similarity(d["norm_ref"], n)
        else:
            sim = similarity(d["norm_ref"], norm)
        # n_parts bonus
        parts_match = 1.0 - min(abs(p["n_parts"]-d["n_parts"])/max(d["n_parts"],1),1.0)
        score = 0.6*sim + 0.4*parts_match
        if score > best_score:
            best_score = score; best = d
    return best, best_score

assignments = []
for p in plan_picts:
    d, s = best_match(p)
    if d and s >= SIM_THRESH:
        assignments.append((p, d, s))

# ---------- 7. Aggregate ----------
counts = Counter(); sims = defaultdict(list)
for p, d, s in assignments:
    counts[d["symbol"]] += 1; sims[d["symbol"]].append(s)

print("=== Geometric counts vs DXF ===")
print(f"{'symbol':>8s} {'GEO':>5s} {'DXF':>5s} {'avg_score':>10s}  desc")
for d in descriptors:
    s = d["symbol"]; geo = counts.get(s,0); gt = DXF_GT.get(s,"—")
    avg = sum(sims[s])/len(sims[s]) if sims[s] else 0.0
    print(f"{s:>8s} {geo:>5d} {str(gt):>5s} {avg:>10.3f}  {d['desc']}")
print()

# ---------- 8. Text control ----------
LABEL_RE = re.compile(r"^[1-9](?:[АA][ЭE]?)?$")
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    words = page.extract_words() or []
text_counts = Counter()
for w in words:
    t = (w.get("text") or "").strip()
    if not LABEL_RE.match(t): continue
    cx = (w["x0"]+w["x1"])/2; cy = (w["top"]+w["bottom"])/2
    if LEG_X0-2 <= cx <= LEG_X1+2 and LEG_Y0-2 <= cy <= LEG_Y1+2: continue
    text_counts[t] += 1

print("=== Text-control vs GEO vs DXF ===")
print(f"{'sym':>5s}  text(А+АЭ)  geo  DXF")
for sym, gt in DXF_GT.items():
    if sym.endswith("АЭ"):
        tc = text_counts.get(sym, 0)
    else:
        tc = text_counts.get(f"{sym}А", 0) + text_counts.get(f"{sym}АЭ", 0)
    print(f"{sym:>5s}  {tc:>9d}  {counts.get(sym,0):>3d}  {gt:>3d}")

# overlay
overlay = img.copy()
sym_color = {"1":(255,0,255),"2":(0,255,255),"3":(255,128,0),"4":(128,0,255),
             "5АЭ":(0,255,0),"6АЭ":(0,200,200),"7АЭ":(200,0,200)}
for p, d, s in assignments:
    x0,y0,x1,y1 = p["bbox"]
    color = sym_color.get(d["symbol"] or "", (255,255,255))
    cv2.rectangle(overlay, (int(x0*zoom)-2,int(y0*zoom)-2), (int(x1*zoom)+2,int(y1*zoom)+2), color, 2)
cv2.imwrite(str(OUT / "overlay.png"), cv2.cvtColor(overlay, cv2.COLOR_RGB2BGR))
print(f"\nOverlay: {OUT/'overlay.png'}")

doc.close()
