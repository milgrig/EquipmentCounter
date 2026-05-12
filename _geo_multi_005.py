"""
Multi-prototype variant of geometry-features classifier.

Same as _geo_features_005.py but each class can have several sub-prototypes
(automatically discovered by simple agglomerative merging of anchor feature
vectors). This handles the situation where one logical class has multiple
visual variants (e.g. 5АЭ as a square vs. wide-rectangular icon).
"""
from __future__ import annotations
import io, sys, json, math
from collections import defaultdict, Counter
from pathlib import Path
from statistics import median

import fitz
import numpy as np
import cv2
import pdfplumber

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_geo_multi_005_out")
OUT.mkdir(exist_ok=True)

DXF_GT = {"5АЭ": 4, "6АЭ": 6, "7АЭ": 7}
TARGETS = ["5АЭ", "6АЭ", "7АЭ"]
LINK_DIST = 5.0
MAX_PRIM = 60.0
ANCHOR_R = 15.0


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
                if hasattr(p, "x"):
                    xs.append(p.x); ys.append(p.y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None


sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend
legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox
doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None: continue
    w, h = bb[2]-bb[0], bb[3]-bb[1]
    if max(w, h) > MAX_PRIM: continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col != "red": continue
    cx, cy = (bb[0]+bb[2])/2, (bb[1]+bb[3])/2
    if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
    prims.append({"bbox": bb, "cx": cx, "cy": cy, "w": w, "h": h})
print(f"Red primitives: {len(prims)}")

n = len(prims); par = list(range(n))
def find(a):
    while par[a]!=a: par[a]=par[par[a]]; a=par[a]
    return a
def union(a,b):
    ra,rb=find(a),find(b)
    if ra!=rb: par[ra]=rb
bins = defaultdict(list)
for i,p in enumerate(prims):
    bins[(int(p["cx"]//LINK_DIST), int(p["cy"]//LINK_DIST))].append(i)
for (bx,by), idxs in bins.items():
    for dx in (-1,0,1):
        for dy in (-1,0,1):
            for i in idxs:
                pi=prims[i]
                for j in bins.get((bx+dx,by+dy), []):
                    if j<=i: continue
                    pj=prims[j]
                    if abs(pi["cx"]-pj["cx"])<=LINK_DIST and abs(pi["cy"]-pj["cy"])<=LINK_DIST:
                        union(i,j)
groups=defaultdict(list)
for i in range(n): groups[find(i)].append(i)
clusters = list(groups.values())
print(f"Clusters: {len(clusters)}")


def features(idx_list):
    xs0=[prims[i]["bbox"][0] for i in idx_list]
    ys0=[prims[i]["bbox"][1] for i in idx_list]
    xs1=[prims[i]["bbox"][2] for i in idx_list]
    ys1=[prims[i]["bbox"][3] for i in idx_list]
    bx0,by0,bx1,by1 = min(xs0),min(ys0),max(xs1),max(ys1)
    W, H = bx1-bx0, by1-by0
    n_p = len(idx_list)
    ws=[prims[i]["w"] for i in idx_list]; hs=[prims[i]["h"] for i in idx_list]
    cxs=[prims[i]["cx"] for i in idx_list]; cys=[prims[i]["cy"] for i in idx_list]
    horiz = sum(1 for i in idx_list if prims[i]["w"]>=prims[i]["h"])
    vert = n_p - horiz
    sx = float(np.std(cxs)) if n_p>1 else 0.0
    sy = float(np.std(cys)) if n_p>1 else 0.0
    return [W, H, H/max(W,0.01), math.log(max(n_p,1)),
            float(np.mean(ws)), float(np.mean(hs)),
            sx, sy, horiz/max(n_p,1), vert/max(n_p,1)], (bx0,by0,bx1,by1)


cluster_feats = []; cluster_bboxes = []
for cl in clusters:
    f, bb = features(cl); cluster_feats.append(f); cluster_bboxes.append(bb)
F_all = np.array(cluster_feats, dtype=np.float32)
mu = F_all.mean(axis=0); sd = F_all.std(axis=0) + 1e-6
F_z = (F_all - mu) / sd

cluster_centers = np.array([[(b[0]+b[2])/2, (b[1]+b[3])/2] for b in cluster_bboxes], dtype=np.float32)
cluster_n = np.array([len(c) for c in clusters])

# ---------------------------------------------------------------------------
# Anchor by text labels
# ---------------------------------------------------------------------------
label_pts = defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t = (w.get("text") or "").strip()
        if t not in TARGETS: continue
        cx = (w["x0"]+w["x1"])/2; cy = (w["top"]+w["bottom"])/2
        if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
        label_pts[t].append((cx, cy))


# Multi-prototype: agglomerative merging
weights_proto = np.array([1.0, 1.0, 1.8, 2.0, 0.5, 0.5, 1.2, 1.2, 2.5, 2.5], dtype=np.float32)
MERGE_THR = 1.8

prototypes = {}        # lab -> list of sub-centroids (z-vectors)
proto_anchors = {}     # lab -> indices into clusters

# Per-class anchoring with closest-label tiebreak: a cluster is anchored to
# class L only if L's nearest text-label is closer than any other class's
# nearest text-label.
ANCHOR_R_LARGE = 30.0
lab_pts_arr = {l: np.array(label_pts.get(l, []), dtype=np.float32)
               for l in TARGETS if label_pts.get(l)}

lab_to_anchors = defaultdict(set)
for lab in TARGETS:
    radius = ANCHOR_R_LARGE if lab == "5АЭ" else ANCHOR_R
    pts = lab_pts_arr.get(lab)
    if pts is None or len(pts) == 0: continue
    for (cx, cy) in pts:
        d = np.hypot(cluster_centers[:,0]-cx, cluster_centers[:,1]-cy)
        mask = (d <= radius) & (cluster_n >= 20)
        if not mask.any(): continue
        cand = np.where(mask)[0]
        nearest_ci = int(cand[d[cand].argmin()])
        # check that this cluster center is closer to a label of `lab`
        # than to any label of another class.
        ccx, ccy = cluster_centers[nearest_ci]
        own = float(d[nearest_ci])
        is_own = True
        for other_lab, other_pts in lab_pts_arr.items():
            if other_lab == lab: continue
            od = np.hypot(other_pts[:,0]-ccx, other_pts[:,1]-ccy)
            if od.min() < own:
                is_own = False; break
        if is_own:
            lab_to_anchors[lab].add(nearest_ci)

print()
for lab in TARGETS:
    used = lab_to_anchors.get(lab, set())
    if not used: continue
    proto_anchors[lab] = sorted(used)
    anchor_z = np.stack([F_z[i] for i in proto_anchors[lab]], axis=0)
    sub = [anchor_z[i:i+1] for i in range(len(anchor_z))]
    while len(sub) > 1:
        cents = [s.mean(axis=0) for s in sub]
        best = (None, None, 1e9)
        for ii in range(len(cents)):
            for jj in range(ii+1, len(cents)):
                dd = float(np.linalg.norm((cents[ii]-cents[jj])*weights_proto))
                if dd < best[2]:
                    best = (ii, jj, dd)
        if best[2] >= MERGE_THR: break
        ii, jj, _ = best
        sub[ii] = np.vstack([sub[ii], sub[jj]])
        del sub[jj]
    centroids = [s.mean(axis=0) for s in sub]
    prototypes[lab] = [c.tolist() for c in centroids]
    raw = F_all[proto_anchors[lab]]
    print(f"  {lab}: anchors={len(used)}  sub-protos={len(sub)}  "
          f"med W={float(np.median(raw[:,0])):.1f} H={float(np.median(raw[:,1])):.1f} "
          f"aspect={float(np.median(raw[:,2])):.2f} n={int(np.median(np.exp(raw[:,3])))}")
    for k, c in enumerate(centroids):
        rc = np.array(c) * sd + mu
        print(f"     sub#{k}: W={rc[0]:.1f} H={rc[1]:.1f} aspect={rc[2]:.2f} "
              f"n={int(np.exp(rc[3]))} horiz={rc[8]:.2f} sy={rc[7]:.1f}")


# ---------------------------------------------------------------------------
# Classify all candidate clusters: nearest sub-prototype across all classes
# ---------------------------------------------------------------------------
flat_protos = []   # list of (label, z_vector)
for lab, subs in prototypes.items():
    for c in subs:
        flat_protos.append((lab, np.array(c, dtype=np.float32)))

print(f"\nFlat sub-prototypes: {len(flat_protos)}  ({Counter(l for l,_ in flat_protos)})")

THR = 3.0
assignments = []
for i, cn in enumerate(cluster_n):
    if cn < 20:
        assignments.append(None); continue
    fz = F_z[i]
    dists = []
    for lab, p in flat_protos:
        diff = (p - fz) * weights_proto
        dists.append((float(np.linalg.norm(diff)), lab))
    dists.sort()
    if dists[0][0] > THR:
        assignments.append(None)
    else:
        assignments.append((dists[0][1], dists[0][0]))

# 5АЭ fallback: any large unclassified cluster (n>=100) that does NOT lie
# close to a 6АЭ/7АЭ text label is assumed to be 5АЭ. This is a residual
# rule for the "biggest icon on the plan" class which often loses its
# text-anchored prototypes due to long leader lines.
LARGE_N = 100
DIST_TO_OTHER = 25.0
other_pts_all = []
for ol in ("6АЭ", "7АЭ"):
    other_pts_all.extend(label_pts.get(ol, []))
other_arr = np.array(other_pts_all, dtype=np.float32) if other_pts_all else None

if "5АЭ" not in {fp[0] for fp in flat_protos}:
    for i, cn in enumerate(cluster_n):
        if assignments[i] is not None: continue
        if cn < LARGE_N: continue
        ccx, ccy = cluster_centers[i]
        if other_arr is not None:
            od = np.hypot(other_arr[:,0]-ccx, other_arr[:,1]-ccy)
            if od.min() < DIST_TO_OTHER: continue
        assignments[i] = ("5АЭ", 0.0)

cnt = Counter(a[0] for a in assignments if a is not None)
print()
print("=== Multi-prototype classification vs DXF ===")
for lab in TARGETS:
    det = cnt.get(lab, 0); exp = DXF_GT.get(lab, "?")
    diff = det - exp if isinstance(exp, int) else None
    mark = "OK" if diff == 0 else f"{diff:+d}"
    print(f"  {lab:>5s}: detected={det:>3d}  expected={exp}  {mark}")

# Save
classified = []
for i, a in enumerate(assignments):
    if a is None: continue
    classified.append({"label": a[0], "dist_z": round(a[1],3),
                       "bbox": list(cluster_bboxes[i]), "n": int(cluster_n[i])})
(OUT / "classified.json").write_text(json.dumps(classified, ensure_ascii=False, indent=2), encoding="utf-8")
(OUT / "prototypes.json").write_text(json.dumps(prototypes, ensure_ascii=False, indent=2), encoding="utf-8")

# Render
DPI_OVER = 200; zoom = DPI_OVER/72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4: img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
canvas = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
cv2.rectangle(canvas, (int(LX0*zoom), int(LY0*zoom)), (int(LX1*zoom), int(LY1*zoom)), (180,180,180), 1)
PALETTE = {"5АЭ": (0,200,0), "6АЭ": (255,0,255), "7АЭ": (0,165,255)}
ASCII = {"5АЭ":"5AE","6АЭ":"6AE","7АЭ":"7AE"}
for c in classified:
    bb = c["bbox"]; lab = c["label"]
    x0,y0,x1,y1 = int(bb[0]*zoom), int(bb[1]*zoom), int(bb[2]*zoom), int(bb[3]*zoom)
    col = PALETTE.get(lab, (0,200,0))
    cv2.rectangle(canvas, (x0-3,y0-3), (x1+3,y1+3), col, 2)
    cv2.putText(canvas, f"{ASCII.get(lab,lab)}({c['dist_z']:.1f})", (x0-3, y0-6),
                cv2.FONT_HERSHEY_SIMPLEX, 0.4, col, 1, cv2.LINE_AA)
cv2.imwrite(str(OUT / "overlay.png"), canvas)
print(f"\nSaved: {OUT}")
doc.close()
