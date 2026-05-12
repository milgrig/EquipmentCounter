"""
Text-free pictogram detector using ground-truth descriptors.

Loads _gt_005_out/gt_descriptors.json and scans the WHOLE plan page (excluding
the legend bbox) for clusters of coloured drawing primitives that match each
descriptor's size/colour/density signature. NO text is used during detection
- only geometry and colour. Text labels are loaded only at the END for
control/comparison.

Outputs (NEW directory _detect_005_out/):
  - detect_results.json   : per-label detected count + cluster centres
  - overlay.png           : full-page render with detected clusters boxed
  - overlay_<label>.png   : per-label crops for the first few hits
"""
from __future__ import annotations
import io, sys, json
from collections import Counter, defaultdict
from pathlib import Path
from statistics import median

import fitz
import numpy as np
import cv2
import pdfplumber

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
GT_JSON = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_gt_005_out\gt_descriptors.json")
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_detect_005_out")
OUT.mkdir(exist_ok=True)

DPI_OVER = 200          # for full-page overlay
MAX_PRIM = 60.0         # ignore drawings whose bbox is larger than this
DXF_GT = {"1": 33, "2": 15, "3": 7, "4": 4, "5АЭ": 4, "6АЭ": 6, "7АЭ": 7}

# Descriptor classes we trust as text-free templates (rich pictograms).
RELIABLE = ["5АЭ", "6АЭ", "7АЭ"]


# ---------------------------------------------------------------------------
def color_class(c):
    if c is None:
        return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4:
            return "red"
        if r < 0.4 and g < 0.4 and b > 0.6:
            return "blue"
        if r < 0.15 and g < 0.15 and b < 0.15:
            return "black"
    return None


def color_bgr(name):
    return {"red": (0, 0, 255), "blue": (255, 0, 0), "black": (40, 40, 40)}.get(name, (0, 200, 0))


def dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"):
                    xs.append(p.x); ys.append(p.y)
    if not xs:
        return None
    return (min(xs), min(ys), max(xs), max(ys))


# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox
print(f"Legend bbox: ({LX0:.1f}, {LY0:.1f}) - ({LX1:.1f}, {LY1:.1f})")

with GT_JSON.open(encoding="utf-8") as f:
    gt = json.load(f)

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
page_w = mp.rect.width
page_h = mp.rect.height
print(f"Page size: {page_w:.0f} x {page_h:.0f} pt")

# Index plan-area drawings (excluding legend bbox).
plan_draw_by_color = defaultdict(list)
total_indexed = 0
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None:
        continue
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    if max(w, h) > MAX_PRIM:
        continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col not in ("red", "blue"):
        continue
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    # exclude legend
    if LX0 - 2 <= cx <= LX1 + 2 and LY0 - 2 <= cy <= LY1 + 2:
        continue
    plan_draw_by_color[col].append({
        "bbox": bb, "cx": cx, "cy": cy,
        "w": w, "h": h,
    })
    total_indexed += 1
print(f"Indexed plan drawings: red={len(plan_draw_by_color['red'])}, "
      f"blue={len(plan_draw_by_color['blue'])} (total {total_indexed})")


# ---------------------------------------------------------------------------
# Cluster primitives by spatial proximity (Union-Find via grid bins)
# ---------------------------------------------------------------------------
def cluster(prims, link_dist):
    """Return list of clusters; each cluster = list of primitives."""
    n = len(prims)
    parent = list(range(n))

    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    # bin by grid cell of size link_dist
    bins = defaultdict(list)
    for i, p in enumerate(prims):
        bx = int(p["cx"] // link_dist)
        by = int(p["cy"] // link_dist)
        bins[(bx, by)].append(i)

    for (bx, by), idxs in bins.items():
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                neigh = bins.get((bx + dx, by + dy), [])
                for i in idxs:
                    pi = prims[i]
                    for j in neigh:
                        if j <= i:
                            continue
                        pj = prims[j]
                        if abs(pi["cx"] - pj["cx"]) <= link_dist and abs(pi["cy"] - pj["cy"]) <= link_dist:
                            union(i, j)

    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(prims[i])
    return list(groups.values())


def cluster_bbox(cl):
    xs0 = [p["bbox"][0] for p in cl]
    ys0 = [p["bbox"][1] for p in cl]
    xs1 = [p["bbox"][2] for p in cl]
    ys1 = [p["bbox"][3] for p in cl]
    return (min(xs0), min(ys0), max(xs1), max(ys1))


# Cluster red primitives once (link_dist=2pt joins primitives that share the
# same icon body without merging two icons that sit further apart).
LINK_DIST = 5.0
red_clusters = cluster(plan_draw_by_color["red"], LINK_DIST)
print(f"Red clusters formed (link={LINK_DIST}pt): {len(red_clusters)}")


# ---------------------------------------------------------------------------
# Non-max suppression: drop overlapping hits across labels (keep richer one)
# ---------------------------------------------------------------------------
def iou(a, b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0, iy0 = max(ax0, bx0), max(ay0, by0)
    ix1, iy1 = min(ax1, bx1), min(ay1, by1)
    iw, ih = max(0.0, ix1 - ix0), max(0.0, iy1 - iy0)
    inter = iw * ih
    if inter <= 0:
        return 0.0
    ua = (ax1 - ax0) * (ay1 - ay0) + (bx1 - bx0) * (by1 - by0) - inter
    return inter / ua if ua > 0 else 0.0


# ---------------------------------------------------------------------------
# Match clusters against each reliable descriptor
# ---------------------------------------------------------------------------
def match_against(label, gt_entry, clusters):
    # Override with empirical bounds derived from cluster diagnostic
    # (the GT radius=10pt undercounted n_parts and W/H).
    overrides = {
        "5АЭ": {"w": (8.0, 28.0),  "h": (8.0, 16.0), "n": (110, 200)},
        "6АЭ": {"w": (8.0, 18.0),  "h": (20.0, 35.0), "n": (25, 110)},
        "7АЭ": {"w": (7.5, 12.5),  "h": (8.0, 16.0),  "n": (60, 110)},
    }
    if label in overrides:
        o = overrides[label]
        w_lo, w_hi = o["w"]; h_lo, h_hi = o["h"]; n_lo = o["n"][0]; n_hi = o["n"][1]
    else:
        mw = gt_entry["median_w_pt"]; mh = gt_entry["median_h_pt"]; mn = gt_entry["median_n_parts"]
        w_lo, w_hi = 0.65 * mw, 1.45 * mw
        h_lo, h_hi = 0.65 * mh, 1.45 * mh
        n_lo, n_hi = max(10, int(0.55 * mn)), 10**9

    hits = []
    for cl in clusters:
        bb = cluster_bbox(cl)
        w = bb[2] - bb[0]
        h = bb[3] - bb[1]
        n = len(cl)
        if not (w_lo <= w <= w_hi and h_lo <= h <= h_hi):
            continue
        if not (n_lo <= n <= n_hi):
            continue
        cx = (bb[0] + bb[2]) / 2
        cy = (bb[1] + bb[3]) / 2
        hits.append({
            "bbox": bb, "cx": cx, "cy": cy,
            "w": round(w, 2), "h": round(h, 2), "n": n,
        })
    return hits


results = {}
for lab in RELIABLE:
    if lab not in gt or gt[lab].get("n_instances", 0) == 0:
        continue
    hits = match_against(lab, gt[lab], red_clusters)
    results[lab] = hits

# Cross-label NMS: prefer label whose hit has more parts at a given location.
def cross_nms(results, iou_thr=0.4):
    flat = []
    for lab, hits in results.items():
        for h in hits:
            flat.append((lab, h))
    keep = [True] * len(flat)
    for i in range(len(flat)):
        if not keep[i]:
            continue
        for j in range(i + 1, len(flat)):
            if not keep[j]:
                continue
            if iou(flat[i][1]["bbox"], flat[j][1]["bbox"]) >= iou_thr:
                # keep the one with more parts; tie -> earlier
                if flat[j][1]["n"] > flat[i][1]["n"]:
                    keep[i] = False
                    break
                else:
                    keep[j] = False
    out = defaultdict(list)
    for k, (lab, h) in enumerate(flat):
        if keep[k]:
            out[lab].append(h)
    return dict(out)

results = cross_nms(results)
for lab in RELIABLE:
    expected = DXF_GT.get(lab, "?")
    print(f"  {lab}: detected {len(results.get(lab, [])):>3d}  (DXF expects {expected})")


# ---------------------------------------------------------------------------
# Render overlay
# ---------------------------------------------------------------------------
zoom = DPI_OVER / 72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4:
    img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
canvas = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

palette = {"5АЭ": (0, 200, 0), "6АЭ": (255, 0, 255), "7АЭ": (0, 165, 255)}
for lab, hits in results.items():
    col = palette.get(lab, (0, 200, 0))
    for h in hits:
        bb = h["bbox"]
        x0 = int(bb[0] * zoom); y0 = int(bb[1] * zoom)
        x1 = int(bb[2] * zoom); y1 = int(bb[3] * zoom)
        cv2.rectangle(canvas, (x0 - 3, y0 - 3), (x1 + 3, y1 + 3), col, 2)
        cv2.putText(canvas, lab.encode("ascii", "replace").decode(), (x0, y0 - 4),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.4, col, 1, cv2.LINE_AA)

# legend frame
cv2.rectangle(canvas,
              (int(LX0 * zoom), int(LY0 * zoom)),
              (int(LX1 * zoom), int(LY1 * zoom)),
              (180, 180, 180), 1)

over_path = OUT / "overlay.png"
cv2.imwrite(str(over_path), canvas)
print(f"Saved overlay: {over_path}  ({canvas.shape[1]}x{canvas.shape[0]})")


# ---------------------------------------------------------------------------
# Save JSON results + comparison
# ---------------------------------------------------------------------------
out_obj = {
    "link_dist_pt": LINK_DIST,
    "labels": {
        lab: {
            "detected": len(hits),
            "expected_dxf": DXF_GT.get(lab, None),
            "hits": hits,
        }
        for lab, hits in results.items()
    },
}
(OUT / "detect_results.json").write_text(json.dumps(out_obj, ensure_ascii=False, indent=2), encoding="utf-8")

print()
print("=== Comparison vs DXF ===")
for lab in RELIABLE:
    det = len(results.get(lab, []))
    exp = DXF_GT.get(lab, "?")
    mark = "OK" if det == exp else f"diff {det - exp:+d}" if isinstance(exp, int) else "-"
    print(f"  {lab:>5s}: detected={det:>3d}  expected={exp}  {mark}")

doc.close()
print()
print(f"Output: {OUT}")
