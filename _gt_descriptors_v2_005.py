"""
Ground-truth pictogram descriptors v2 — improved.

Improvements over v1:
  * RADIUS_PT raised to 20pt (was 10) so the whole pictogram is captured.
  * Connectivity-based growth: starting from a small seed radius, add
    nearby coloured primitives that touch the cluster (link_dist=4pt) and
    stop expanding when no new neighbours are found.
  * Add base-row labels "1", "2", "3", "4" (no А/АЭ suffix) which appear
    on the plan as fixture series labels (mapped to the corresponding
    legend rows of the same family).
  * Outlier filtering by IQR for W/H so a stray label-glued primitive
    cannot blow up max bounds.

Outputs (NEW directory _gt_v2_005_out/):
  - gt_descriptors_v2.json
  - gt_<label>_inst_<i>.png
"""
from __future__ import annotations
import io, sys, re, json
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
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_gt_v2_005_out")
OUT.mkdir(exist_ok=True)

DPI = 600
SEED_RADIUS = 6.0       # initial neighbourhood around label centre
LINK_DIST = 4.0         # connectivity link distance for growth
GROW_LIMIT_PT = 35.0    # do not let a cluster exceed this on either dim
MAX_PRIM = 60.0         # ignore primitives bigger than this (background)

TARGET_LABELS = [
    "1", "2", "3", "4",        # base labels (working fixtures)
    "1А", "1АЭ", "2А", "2АЭ",  # 1/2 with suffixes
    "3А", "3АЭ", "4А", "4АЭ",
    "5АЭ", "6АЭ", "7АЭ",
]


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
    return {"red": (0, 0, 255), "blue": (255, 0, 0), "black": (40, 40, 40)}.get(name, (180, 180, 180))


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


# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox
print(f"Legend bbox: ({LX0:.1f}, {LY0:.1f}) - ({LX1:.1f}, {LY1:.1f})")

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
zoom = DPI / 72.0

# Index small coloured drawings (excluding legend area).
all_draw = []
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
    if LX0 - 2 <= cx <= LX1 + 2 and LY0 - 2 <= cy <= LY1 + 2:
        continue
    all_draw.append({
        "bbox": bb, "cx": cx, "cy": cy,
        "w": w, "h": h, "color": col,
    })
print(f"Indexed plan-area coloured drawings: {len(all_draw)}")

# Spatial grid for fast neighbour queries (cell = LINK_DIST).
grid = defaultdict(list)
for i, p in enumerate(all_draw):
    grid[(int(p["cx"] // LINK_DIST), int(p["cy"] // LINK_DIST))].append(i)


def near(cx, cy, radius):
    cells_x = range(int((cx - radius) // LINK_DIST), int((cx + radius) // LINK_DIST) + 1)
    cells_y = range(int((cy - radius) // LINK_DIST), int((cy + radius) // LINK_DIST) + 1)
    out = []
    for bx in cells_x:
        for by in cells_y:
            for idx in grid.get((bx, by), []):
                p = all_draw[idx]
                if abs(p["cx"] - cx) <= radius and abs(p["cy"] - cy) <= radius:
                    out.append(idx)
    return out


def grow(seed_idx_list, exclude_box=None):
    """Region growing: start from seed indices, add neighbours within
    LINK_DIST until cluster stabilises, but bounded by GROW_LIMIT_PT."""
    if not seed_idx_list:
        return []
    cluster_idx = set(seed_idx_list)
    frontier = list(seed_idx_list)
    # initial bbox from seeds
    bb = _bbox_for(cluster_idx)
    while frontier:
        new_frontier = []
        for i in frontier:
            p = all_draw[i]
            for j in near(p["cx"], p["cy"], LINK_DIST):
                if j in cluster_idx:
                    continue
                pj = all_draw[j]
                # skip text glyphs by exclude_box
                if exclude_box:
                    ex0, ey0, ex1, ey1 = exclude_box
                    if ex0 - 0.5 <= pj["cx"] <= ex1 + 0.5 and ey0 - 0.5 <= pj["cy"] <= ey1 + 0.5:
                        continue
                # tentative new bbox; reject if exceeds GROW_LIMIT
                nbx0 = min(bb[0], pj["bbox"][0]); nby0 = min(bb[1], pj["bbox"][1])
                nbx1 = max(bb[2], pj["bbox"][2]); nby1 = max(bb[3], pj["bbox"][3])
                if (nbx1 - nbx0) > GROW_LIMIT_PT or (nby1 - nby0) > GROW_LIMIT_PT:
                    continue
                cluster_idx.add(j)
                new_frontier.append(j)
                bb = (nbx0, nby0, nbx1, nby1)
        frontier = new_frontier
    return list(cluster_idx)


def _bbox_for(idx_iter):
    xs0, ys0, xs1, ys1 = [], [], [], []
    for i in idx_iter:
        b = all_draw[i]["bbox"]
        xs0.append(b[0]); ys0.append(b[1]); xs1.append(b[2]); ys1.append(b[3])
    return (min(xs0), min(ys0), max(xs1), max(ys1))


# ---------------------------------------------------------------------------
# Collect text labels on the plan (outside the legend)
# ---------------------------------------------------------------------------
label_pts = defaultdict(list)
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    for w in page.extract_words() or []:
        t = (w.get("text") or "").strip()
        if t not in TARGET_LABELS:
            continue
        cx = (w["x0"] + w["x1"]) / 2
        cy = (w["top"] + w["bottom"]) / 2
        if LX0 - 2 <= cx <= LX1 + 2 and LY0 - 2 <= cy <= LY1 + 2:
            continue
        label_pts[t].append((cx, cy, w["x0"], w["top"], w["x1"], w["bottom"]))

print()
print("Plan label counts:")
for k in TARGET_LABELS:
    print(f"  {k:>5s}: {len(label_pts.get(k, []))}")


# ---------------------------------------------------------------------------
# Per label: seed -> grow -> filter by dominant colour
# ---------------------------------------------------------------------------
def iqr_filter(values, k=1.5):
    if len(values) < 4:
        return values
    s = sorted(values)
    q1 = s[len(s) // 4]
    q3 = s[(3 * len(s)) // 4]
    iqr = q3 - q1
    lo, hi = q1 - k * iqr, q3 + k * iqr
    return [v for v in values if lo <= v <= hi]


descriptors = {}

for label in TARGET_LABELS:
    pts = label_pts.get(label, [])
    if not pts:
        descriptors[label] = {"n_instances": 0, "label_text_count": 0}
        continue

    instances = []
    for idx, (cx, cy, tx0, ty0, tx1, ty1) in enumerate(pts):
        seeds = near(cx, cy, SEED_RADIUS)
        # remove text-glyph primitives
        seeds = [i for i in seeds
                 if not (tx0 - 0.5 <= all_draw[i]["cx"] <= tx1 + 0.5
                         and ty0 - 0.5 <= all_draw[i]["cy"] <= ty1 + 0.5)]
        if not seeds:
            continue
        cluster_idx = grow(seeds, exclude_box=(tx0, ty0, tx1, ty1))
        # dominant colour
        cols = Counter(all_draw[i]["color"] for i in cluster_idx)
        dom_col = cols.most_common(1)[0][0]
        cluster_idx = [i for i in cluster_idx if all_draw[i]["color"] == dom_col]
        if not cluster_idx:
            continue
        bb = _bbox_for(cluster_idx)
        w_pt = bb[2] - bb[0]
        h_pt = bb[3] - bb[1]
        instances.append({
            "label": label, "idx": idx,
            "text_center": (cx, cy),
            "bbox_pt": bb, "w_pt": w_pt, "h_pt": h_pt,
            "color": dom_col, "n_parts": len(cluster_idx),
            "color_distribution": dict(cols),
        })

    if not instances:
        descriptors[label] = {"n_instances": 0, "label_text_count": len(pts)}
        continue

    # Outlier filtering
    ws = iqr_filter([i["w_pt"] for i in instances])
    hs = iqr_filter([i["h_pt"] for i in instances])
    ns = iqr_filter([i["n_parts"] for i in instances])

    descriptors[label] = {
        "label_text_count": len(pts),
        "n_instances": len(instances),
        "median_w_pt": round(median(ws), 2),
        "median_h_pt": round(median(hs), 2),
        "min_w_pt": round(min(ws), 2),
        "max_w_pt": round(max(ws), 2),
        "min_h_pt": round(min(hs), 2),
        "max_h_pt": round(max(hs), 2),
        "median_n_parts": int(median(ns)),
        "min_n_parts": int(min(ns)),
        "max_n_parts": int(max(ns)),
        "dominant_color": Counter(i["color"] for i in instances).most_common(1)[0][0],
        "color_distribution": dict(Counter(i["color"] for i in instances)),
    }

    # Save overlay crops (up to 3)
    for k_idx, inst in enumerate(instances[:3]):
        bx0, by0, bx1, by1 = inst["bbox_pt"]
        pad = 5.0
        rx0, ry0, rx1, ry1 = bx0 - pad, by0 - pad, bx1 + pad, by1 + pad
        clip = fitz.Rect(rx0, ry0, rx1, ry1)
        try:
            pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        except Exception:
            continue
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        over = img.copy()
        for s in all_draw:
            sb = s["bbox"]
            if sb[2] < rx0 or sb[0] > rx1 or sb[3] < ry0 or sb[1] > ry1:
                continue
            px0 = int((sb[0] - rx0) * zoom); py0 = int((sb[1] - ry0) * zoom)
            px1 = int((sb[2] - rx0) * zoom); py1 = int((sb[3] - ry0) * zoom)
            cv2.rectangle(over, (px0, py0), (px1, py1), color_bgr(s["color"]), 1)
        h_, w_ = img.shape[:2]
        canvas = np.full((h_, w_ * 2 + 20, 3), 255, dtype=np.uint8)
        canvas[:h_, :w_] = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        canvas[:h_, w_ + 20: w_ * 2 + 20] = cv2.cvtColor(over, cv2.COLOR_RGB2BGR)
        safe = re.sub(r"[^\w]", "_", label)
        cv2.imwrite(str(OUT / f"gt_{safe}_inst_{k_idx}.png"), canvas)


# ---------------------------------------------------------------------------
out_json = OUT / "gt_descriptors_v2.json"
out_json.write_text(json.dumps(descriptors, ensure_ascii=False, indent=2), encoding="utf-8")

print()
print("=== Descriptor summary v2 ===")
for label in TARGET_LABELS:
    d = descriptors[label]
    if d.get("n_instances", 0) == 0:
        print(f"  {label:>5s}: NO INSTANCES (text labels: {d.get('label_text_count', 0)})")
        continue
    print(f"  {label:>5s}: n={d['n_instances']:>3d}  "
          f"W={d['median_w_pt']:>6.2f}pt [{d['min_w_pt']:>5.1f}-{d['max_w_pt']:>5.1f}]  "
          f"H={d['median_h_pt']:>6.2f}pt [{d['min_h_pt']:>5.1f}-{d['max_h_pt']:>5.1f}]  "
          f"col={d['dominant_color']:>5s}  parts={d['median_n_parts']:>4d} [{d['min_n_parts']}-{d['max_n_parts']}]")

doc.close()
print()
print(f"Saved: {out_json}")
