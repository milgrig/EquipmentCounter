"""
Recon: link text labels (1, 1А, 2, ..., P1..P5) on the plan to nearby rectangles.

Goal:
  1. Build empirical (label -> rect descriptor) mapping from the plan itself.
  2. Compare descriptors with rectangles inside the legend.
  3. Compare visual crops (plan rect vs legend rect) -- pixel-level similarity.
"""
from __future__ import annotations
import io
import sys
import re
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
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_recon_005_out")
OUT.mkdir(exist_ok=True)

# Labels we care about: "1", "1А", "2", "2А", "3", "3А", "4", "4АЭ", "5АЭ", "6АЭ", "7АЭ"
LABEL_RE = re.compile(r"^([1-9](?:[АA]?[ЭE]?)?|[Pp][1-9])$")
SEARCH_RADIUS = 30.0  # pt -- max distance from text center to nearest rect

# ---- helpers ---------------------------------------------------------------
def fmt_color(c):
    if c is None: return "—"
    if isinstance(c, (tuple, list)):
        return "(" + ",".join(f"{v:.2f}" for v in c) + ")"
    return str(c)

def round_size(w, h, step=1.0):
    return (round(w/step)*step, round(h/step)*step)

# ---- 1. Get legend bbox via parse_legend -----------------------------------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
print(f"Legend page: {legend.page_index}, items: {len(legend.items)}")
legend_bbox = getattr(legend, "bbox", None)
print(f"Legend bbox: {legend_bbox}")
print()

# ---- 2. Extract text labels from plan via pdfplumber -----------------------
labels: list[dict] = []  # {text, cx, cy}
with pdfplumber.open(str(PDF)) as pdf:
    page = pdf.pages[legend.page_index]
    page_h = page.height
    words = page.extract_words(extra_attrs=["size"]) or []
    for w in words:
        t = (w.get("text") or "").strip()
        if not t: continue
        if not LABEL_RE.match(t): continue
        cx = (w["x0"] + w["x1"]) / 2.0
        cy = (w["top"] + w["bottom"]) / 2.0
        # filter out labels inside legend bbox (those are legend headers)
        if legend_bbox:
            x0, top, x1, bot = legend_bbox
            if x0 <= cx <= x1 and top <= cy <= bot:
                continue
        labels.append({"text": t, "cx": cx, "cy": cy})

print(f"Plan labels (outside legend): {len(labels)}")
counts_by_text = Counter(l["text"] for l in labels)
print("Label counts on plan (top 20):")
for t, c in counts_by_text.most_common(20):
    print(f"  '{t}': {c}")
print()

# ---- 3. Get all rectangles (bbox + descriptor) -----------------------------
doc = fitz.open(str(PDF))
mp = doc[legend.page_index]

rects = []
for d in mp.get_drawings():
    items = d.get("items", [])
    if len(items) == 1 and items[0][0] == "re":
        r = items[0][1]
        rects.append({
            "bbox": (r.x0, r.y0, r.x1, r.y1),
            "cx": (r.x0 + r.x1) / 2.0,
            "cy": (r.y0 + r.y1) / 2.0,
            "w": r.x1 - r.x0,
            "h": r.y1 - r.y0,
            "fill": d.get("fill"),
            "color": d.get("color"),
        })
# also include stroke-only rectangles found via 'l' segments?  Skip — they are
# rare for fixtures (legend uses filled rects). We'll add later if needed.
print(f"Total pure rects on page: {len(rects)}")
print()

# ---- 4. For each label, find nearest rectangle -----------------------------
def nearest_rect(cx, cy, max_dist=SEARCH_RADIUS):
    best = None
    best_d = max_dist + 1
    for r in rects:
        # distance from label to rect bbox center (or to closest point on rect)
        dx = max(r["bbox"][0] - cx, 0, cx - r["bbox"][2])
        dy = max(r["bbox"][1] - cy, 0, cy - r["bbox"][3])
        d = (dx*dx + dy*dy) ** 0.5
        if d < best_d:
            best_d = d
            best = r
    return best, best_d

# Group labels by text, collect rect descriptors
group_descriptors: dict[str, list] = defaultdict(list)
unmatched: dict[str, int] = Counter()

for l in labels:
    r, dist = nearest_rect(l["cx"], l["cy"])
    if r is None:
        unmatched[l["text"]] += 1
        continue
    group_descriptors[l["text"]].append({
        "w": r["w"], "h": r["h"],
        "fill": fmt_color(r["fill"]),
        "color": fmt_color(r["color"]),
        "dist": dist,
        "bbox": r["bbox"],
        "label_xy": (l["cx"], l["cy"]),
    })

print("=== Empirical descriptor per label ===")
print(f"{'label':>6s} {'cnt':>4s}  {'med_W':>6s} x {'med_H':>6s}  {'fill':>20s}  {'med_dist':>9s}")
for t in sorted(group_descriptors.keys()):
    g = group_descriptors[t]
    ws = [x["w"] for x in g]
    hs = [x["h"] for x in g]
    fills = Counter(x["fill"] for x in g)
    fill_top = fills.most_common(1)[0][0]
    dists = [x["dist"] for x in g]
    print(f"{t:>6s} {len(g):>4d}  {median(ws):>6.2f} x {median(hs):>6.2f}  {fill_top:>20s}  {median(dists):>9.2f}")
print()

if unmatched:
    print("Unmatched labels (no rect within radius):")
    for t, c in unmatched.most_common():
        print(f"  '{t}': {c}")
    print()

# ---- 5. Compare with DXF ground truth --------------------------------------
DXF_GT = {"1":33, "2":15, "3":7, "4":4, "5АЭ":4, "6АЭ":6, "7АЭ":7}
print("=== Comparison with DXF ground truth ===")
print(f"{'label':>6s} {'DXF':>5s} {'plan_text':>10s} {'matched_rect':>13s}")
for k, v in DXF_GT.items():
    plan_t = counts_by_text.get(k, 0)
    matched = len(group_descriptors.get(k, []))
    print(f"{k:>6s} {v:>5d} {plan_t:>10d} {matched:>13d}")
print()

# ---- 6. Compare rectangles inside legend with plan rectangles --------------
if legend_bbox:
    x0, top, x1, bot = legend_bbox
    legend_rects = [r for r in rects if x0 <= r["cx"] <= x1 and top <= r["cy"] <= bot]
    plan_rects = [r for r in rects if not (x0 <= r["cx"] <= x1 and top <= r["cy"] <= bot)]
    print(f"Legend rects: {len(legend_rects)}, Plan rects: {len(plan_rects)}")

    # group legend rects by row (cy)
    legend_rects.sort(key=lambda r: r["cy"])
    print()
    print("Legend rects (sorted by Y):")
    for lr in legend_rects:
        print(f"  bbox=({lr['bbox'][0]:.1f},{lr['bbox'][1]:.1f},{lr['bbox'][2]:.1f},{lr['bbox'][3]:.1f})"
              f" {lr['w']:.2f}x{lr['h']:.2f}  fill={fmt_color(lr['fill'])}  stroke={fmt_color(lr['color'])}")
    print()

# ---- 7. Visual comparison: render crops at high DPI, compute SSIM -----------
DPI = 300
zoom = DPI / 72.0
mat = fitz.Matrix(zoom, zoom)
pix = mp.get_pixmap(matrix=mat, alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4:
    img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
print(f"Page rendered: {img.shape}")

def crop_rect(bbox, pad=2.0):
    x0, y0, x1, y1 = bbox
    px0 = max(0, int((x0-pad)*zoom))
    py0 = max(0, int((y0-pad)*zoom))
    px1 = min(img.shape[1], int((x1+pad)*zoom))
    py1 = min(img.shape[0], int((y1+pad)*zoom))
    return img[py0:py1, px0:px1].copy()

def normalize_crop(crop, target=(40, 80)):
    if crop.size == 0: return None
    return cv2.resize(crop, target, interpolation=cv2.INTER_AREA)

def ssim_simple(a, b):
    """Quick SSIM-like score: 1 - mean(|a-b|)/255 in grayscale."""
    if a is None or b is None or a.shape != b.shape:
        return 0.0
    ag = cv2.cvtColor(a, cv2.COLOR_RGB2GRAY).astype(np.float32)
    bg = cv2.cvtColor(b, cv2.COLOR_RGB2GRAY).astype(np.float32)
    return float(1.0 - np.mean(np.abs(ag - bg)) / 255.0)

# For each label group, take a few sample plan rects and the most likely
# legend rect, compute similarity.
if legend_bbox:
    print("=== Visual similarity (plan rect vs legend rect) ===")

    # Match each label's median descriptor to closest legend rect by (W, H, fill).
    def find_legend_match(target_w, target_h, target_fill):
        best, best_score = None, 1e9
        for lr in legend_rects:
            dw = abs(lr["w"] - target_w)
            dh = abs(lr["h"] - target_h)
            color_match = 0 if fmt_color(lr["fill"]) == target_fill else 5
            score = dw + dh + color_match
            if score < best_score:
                best_score = score
                best = lr
        return best, best_score

    for label in sorted(group_descriptors.keys()):
        g = group_descriptors[label]
        if not g: continue
        med_w = median(x["w"] for x in g)
        med_h = median(x["h"] for x in g)
        fills = Counter(x["fill"] for x in g)
        fill_top = fills.most_common(1)[0][0]

        leg, score = find_legend_match(med_w, med_h, fill_top)
        if leg is None:
            print(f"  [{label}] no legend candidate")
            continue

        leg_crop = normalize_crop(crop_rect(leg["bbox"]))
        # pick first 3 plan rects for this label
        sims = []
        for entry in g[:5]:
            plan_crop = normalize_crop(crop_rect(entry["bbox"]))
            sims.append(ssim_simple(leg_crop, plan_crop))
        avg = sum(sims)/len(sims) if sims else 0.0

        # save side-by-side for the first one
        if g:
            pc = crop_rect(g[0]["bbox"])
            lc = crop_rect(leg["bbox"])
            ph, pw = pc.shape[:2]
            lh, lw = lc.shape[:2]
            H = max(ph, lh, 30); W = pw + lw + 10
            canvas = np.full((H, W, 3), 240, dtype=np.uint8)
            canvas[:ph, :pw] = pc
            canvas[:lh, pw+10:pw+10+lw] = lc
            cv2.imwrite(str(OUT / f"compare_{label}.png"), cv2.cvtColor(canvas, cv2.COLOR_RGB2BGR))

        print(f"  [{label}] plan({med_w:.1f}x{med_h:.1f},{fill_top}) vs legend({leg['w']:.1f}x{leg['h']:.1f},{fmt_color(leg['fill'])})"
              f"  score={score:.2f}  avg_sim={avg:.3f}  samples={len(sims)}")

doc.close()
print()
print(f"Side-by-side images saved to: {OUT}")
