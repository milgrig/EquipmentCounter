"""
Geometry-only feature vectors for cluster discrimination.

Key insight: bitmap rasterisation of stroke primitives is unreliable because
their bbox-thickness is near zero. Instead, describe each cluster by purely
geometric statistics computed from its primitive bboxes:

  f0  W (pt)
  f1  H (pt)
  f2  H/W (aspect)
  f3  log(n_parts)
  f4  mean primitive width
  f5  mean primitive height
  f6  std of primitive cx (spread along X)
  f7  std of primitive cy (spread along Y)
  f8  ratio of "horizontal" primitives  (w >= h)  / n
  f9  ratio of "vertical"   primitives  (h >  w)  / n

For each reliable class (5АЭ, 6АЭ, 7АЭ) we anchor by text labels, identify
the matching cluster (largest red cluster within R=15pt of label center),
collect feature vectors, and store the median as a class prototype.

Then we classify ALL red clusters by 1-NN on z-scored features. A cluster is
assigned to its nearest prototype only if distance is below per-class
threshold; otherwise it stays unassigned.

Output: _geo_005_out/{prototypes.json, classified.json, overlay.png}
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
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_geo_005_out")
OUT.mkdir(exist_ok=True)

DXF_GT = {"5АЭ": 4, "6АЭ": 6, "7АЭ": 7}
TARGETS = ["5АЭ", "6АЭ", "7АЭ"]
LINK_DIST = 5.0
MAX_PRIM = 60.0
ANCHOR_R = 15.0   # find matching cluster within this radius of label center


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

# Index red plan primitives.
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

# Cluster all primitives once via Union-Find (link=5pt)
n = len(prims); par = list(range(n))
def find(a):
    while par[a] != a:
        par[a] = par[par[a]]; a = par[a]
    return a
def union(a, b):
    ra, rb = find(a), find(b)
    if ra != rb: par[ra] = rb
bins = defaultdict(list)
for i, p in enumerate(prims):
    bins[(int(p["cx"]//LINK_DIST), int(p["cy"]//LINK_DIST))].append(i)
for (bx, by), idxs in bins.items():
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for i in idxs:
                pi = prims[i]
                for j in bins.get((bx+dx, by+dy), []):
                    if j <= i: continue
                    pj = prims[j]
                    if abs(pi["cx"]-pj["cx"])<=LINK_DIST and abs(pi["cy"]-pj["cy"])<=LINK_DIST:
                        union(i, j)
groups = defaultdict(list)
for i in range(n):
    groups[find(i)].append(i)
clusters = list(groups.values())
print(f"Clusters: {len(clusters)}")


def features(idx_list):
    """Return 10-dim feature vector for a cluster."""
    if not idx_list:
        return None
    xs0=[prims[i]["bbox"][0] for i in idx_list]
    ys0=[prims[i]["bbox"][1] for i in idx_list]
    xs1=[prims[i]["bbox"][2] for i in idx_list]
    ys1=[prims[i]["bbox"][3] for i in idx_list]
    bx0,by0,bx1,by1 = min(xs0),min(ys0),max(xs1),max(ys1)
    W, H = bx1-bx0, by1-by0
    n_p = len(idx_list)
    ws = [prims[i]["w"] for i in idx_list]
    hs = [prims[i]["h"] for i in idx_list]
    cxs = [prims[i]["cx"] for i in idx_list]
    cys = [prims[i]["cy"] for i in idx_list]
    horiz = sum(1 for i in idx_list if prims[i]["w"] >= prims[i]["h"])
    vert  = n_p - horiz
    sx = float(np.std(cxs)) if n_p > 1 else 0.0
    sy = float(np.std(cys)) if n_p > 1 else 0.0
    return [
        W, H,
        H / max(W, 0.01),
        math.log(max(n_p, 1)),
        float(np.mean(ws)),
        float(np.mean(hs)),
        sx, sy,
        horiz / max(n_p, 1),
        vert / max(n_p, 1),
    ], (bx0, by0, bx1, by1)


cluster_feats = []
cluster_bboxes = []
for cl in clusters:
    f, bb = features(cl)
    cluster_feats.append(f)
    cluster_bboxes.append(bb)
F_all = np.array(cluster_feats, dtype=np.float32)
print(f"Feature matrix: {F_all.shape}")

# z-score globally
mu = F_all.mean(axis=0)
sd = F_all.std(axis=0) + 1e-6
F_z = (F_all - mu) / sd


# ---------------------------------------------------------------------------
# Anchor by text labels and gather per-class prototypes
# ---------------------------------------------------------------------------
label_pts = defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t = (w.get("text") or "").strip()
        if t not in TARGETS: continue
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        if LX0-2 <= cx <= LX1+2 and LY0-2 <= cy <= LY1+2: continue
        label_pts[t].append((cx, cy))

# For each label center, find nearest cluster (by cluster center) within ANCHOR_R.
# Filter: keep only clusters that are "real" pictograms (n_parts >= 20).
cluster_centers = np.array([
    [(b[0]+b[2])/2, (b[1]+b[3])/2] for b in cluster_bboxes
], dtype=np.float32)
cluster_n = np.array([len(c) for c in clusters])

prototypes = {}
proto_examples = {}
for lab in TARGETS:
    pts = label_pts.get(lab, [])
    if not pts:
        continue
    feats_for_lab = []
    used_indices = set()
    for (cx, cy) in pts:
        # candidate clusters
        d = np.hypot(cluster_centers[:, 0]-cx, cluster_centers[:, 1]-cy)
        mask = (d <= ANCHOR_R) & (cluster_n >= 20)
        if not mask.any():
            continue
        # nearest among substantial clusters
        cand_idx = np.where(mask)[0]
        nearest = cand_idx[d[cand_idx].argmin()]
        if nearest in used_indices:
            continue
        used_indices.add(int(nearest))
        feats_for_lab.append(F_z[nearest])
    if not feats_for_lab:
        continue
    arr = np.stack(feats_for_lab, axis=0)
    proto = np.median(arr, axis=0)
    # report in original units too
    raw = F_all[[i for i in used_indices]]
    prototypes[lab] = {
        "n_anchors": len(used_indices),
        "proto_z": proto.tolist(),
        "median_W": float(np.median(raw[:, 0])),
        "median_H": float(np.median(raw[:, 1])),
        "median_aspect": float(np.median(raw[:, 2])),
        "median_n": int(np.median(np.exp(raw[:, 3]))),
        "median_horiz_ratio": float(np.median(raw[:, 8])),
        "median_vert_ratio": float(np.median(raw[:, 9])),
        "median_sx": float(np.median(raw[:, 6])),
        "median_sy": float(np.median(raw[:, 7])),
    }
    proto_examples[lab] = sorted(used_indices)
    print(f"  {lab}: anchors={len(used_indices)}  W={prototypes[lab]['median_W']:.1f}  H={prototypes[lab]['median_H']:.1f}  "
          f"aspect={prototypes[lab]['median_aspect']:.2f}  n={prototypes[lab]['median_n']}  "
          f"horiz={prototypes[lab]['median_horiz_ratio']:.2f}  sx={prototypes[lab]['median_sx']:.1f}  sy={prototypes[lab]['median_sy']:.1f}")

(OUT / "prototypes.json").write_text(json.dumps(prototypes, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Classify ALL clusters (n>=20) as nearest prototype
# ---------------------------------------------------------------------------
# Build prototype matrix
proto_labels = list(prototypes.keys())
P = np.array([prototypes[l]["proto_z"] for l in proto_labels], dtype=np.float32)
print(f"\nClassify {int((cluster_n>=20).sum())} candidate clusters against {len(proto_labels)} prototypes")

# Optional per-class distance threshold (in z-units, learned from anchor spread)
THR = {l: 3.5 for l in proto_labels}

assignments = []
for i, cn in enumerate(cluster_n):
    if cn < 20:
        assignments.append(None); continue
    fz = F_z[i]
    # Euclidean distance in z-space; horiz/vert ratios and n_parts are the
    # most discriminative features; weight them up.
    weights = np.array([1.0, 1.0, 1.8, 2.0, 0.5, 0.5, 1.2, 1.2, 2.5, 2.5], dtype=np.float32)
    diff = (P - fz) * weights
    dists = np.linalg.norm(diff, axis=1)
    best = int(dists.argmin())
    if dists[best] > THR[proto_labels[best]]:
        assignments.append(None); continue
    assignments.append((proto_labels[best], float(dists[best])))

# Counts
cnt = Counter(a[0] for a in assignments if a is not None)
print()
print("=== Classification result vs DXF ===")
for lab in TARGETS:
    det = cnt.get(lab, 0)
    exp = DXF_GT.get(lab, "?")
    diff = det - exp if isinstance(exp, int) else None
    mark = "OK" if diff == 0 else f"{diff:+d}"
    print(f"  {lab:>5s}: detected={det:>3d}  expected={exp}  {mark}")

# Save classified bboxes
classified = []
for i, a in enumerate(assignments):
    if a is None: continue
    bb = cluster_bboxes[i]
    classified.append({
        "label": a[0], "dist_z": round(a[1], 3),
        "bbox": list(bb),
        "n": int(cluster_n[i]),
    })
(OUT / "classified.json").write_text(json.dumps(classified, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------------------------------------------------------------------------
# Render overlay
# ---------------------------------------------------------------------------
DPI_OVER = 200
zoom = DPI_OVER / 72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4:
    img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
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

# Stats panel
panel_w, panel_h = 420, 160
px, py = 30, 30
cv2.rectangle(canvas, (px,py), (px+panel_w,py+panel_h), (255,255,255), -1)
cv2.rectangle(canvas, (px,py), (px+panel_w,py+panel_h), (0,0,0), 2)
cv2.putText(canvas, "Geometry-features classifier", (px+10, py+24),
            cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0,0,0), 1, cv2.LINE_AA)
y = py+50
for lab in TARGETS:
    det = cnt.get(lab, 0); exp = DXF_GT.get(lab, "?")
    diff = det-exp if isinstance(exp,int) else None
    s = "OK" if diff==0 else (f"{diff:+d}" if diff is not None else "-")
    col = PALETTE.get(lab, (0,0,0))
    cv2.putText(canvas, f"{ASCII.get(lab,lab):<5s} det={det:>3d}  dxf={exp:>3}  {s}",
                (px+10, y), cv2.FONT_HERSHEY_SIMPLEX, 0.5, col, 1, cv2.LINE_AA)
    y += 28

out_png = OUT / "overlay.png"
cv2.imwrite(str(out_png), canvas)
print(f"\nSaved overlay: {out_png}")
doc.close()
