"""
Build text-anchored ground-truth pictogram descriptors from the plan.

Idea: every text label on the plan (1А, 1АЭ, 2А, ..., 7АЭ) sits next to its
pictogram. Use the label coordinate as a supervisor signal: collect all
coloured drawings within a small radius around the label, cluster them into
one pictogram, then aggregate per-label across all instances to obtain a
clean reference (W, H, dominant colour, n_parts).

Outputs (NEW directory _gt_005_out/):
  - gt_descriptors.json    : one entry per label
  - gt_<label>_inst_<i>.png: per-instance crops with bbox overlay
  - gt_<label>_summary.txt : per-label aggregated stats
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
OUT = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_gt_005_out")
OUT.mkdir(exist_ok=True)

DPI = 600
RADIUS_PT = 10.0     # collect drawings within this radius of label centre
MAX_PRIM = 60.0      # ignore drawings whose bbox is larger than this (background)

# Labels we want to learn from the plan. Mapping: text on plan -> canonical key.
TARGET_LABELS = ["1А", "1АЭ", "2А", "2АЭ", "3А", "3АЭ", "4А", "4АЭ", "5АЭ", "6АЭ", "7АЭ"]


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
    if not xs:
        return None
    return (min(xs), min(ys), max(xs), max(ys))


# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LEG_X0, LEG_Y0, LEG_X1, LEG_Y1 = legend.legend_bbox
print(f"Legend bbox: ({LEG_X0:.1f}, {LEG_Y0:.1f}) - ({LEG_X1:.1f}, {LEG_Y1:.1f})")

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
zoom = DPI / 72.0

# Index all candidate drawings (small + coloured) once.
all_draw = []
for d in mp.get_drawings():
    bb = dbox(d)
    if bb is None:
        continue
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    if max(w, h) > MAX_PRIM:
        continue
    col = color_class(d.get("fill")) or color_class(d.get("color"))
    if col is None:
        continue
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    all_draw.append({
        "bbox": bb, "cx": cx, "cy": cy,
        "w": w, "h": h, "color": col,
    })
print(f"Indexed {len(all_draw)} small coloured drawings")


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
        # exclude legend interior
        if LEG_X0 - 2 <= cx <= LEG_X1 + 2 and LEG_Y0 - 2 <= cy <= LEG_Y1 + 2:
            continue
        label_pts[t].append((cx, cy, w["x0"], w["top"], w["x1"], w["bottom"]))

print()
print("Plan label counts:")
for k in TARGET_LABELS:
    print(f"  {k:>5s}: {len(label_pts.get(k, []))}")


# ---------------------------------------------------------------------------
# Per label: gather pictogram primitives near every instance
# ---------------------------------------------------------------------------
def neighbours(cx, cy, exclude_box=None):
    """Drawings within RADIUS_PT of (cx,cy), not overlapping the text bbox."""
    out = []
    for s in all_draw:
        if abs(s["cx"] - cx) > RADIUS_PT or abs(s["cy"] - cy) > RADIUS_PT:
            continue
        if exclude_box:
            ex0, ey0, ex1, ey1 = exclude_box
            sx, sy = s["cx"], s["cy"]
            if ex0 - 0.5 <= sx <= ex1 + 0.5 and ey0 - 0.5 <= sy <= ey1 + 0.5:
                # skip drawings sitting on top of the text glyphs
                continue
        out.append(s)
    return out


def cluster_bbox(strokes):
    if not strokes:
        return None
    xs0 = [s["bbox"][0] for s in strokes]
    ys0 = [s["bbox"][1] for s in strokes]
    xs1 = [s["bbox"][2] for s in strokes]
    ys1 = [s["bbox"][3] for s in strokes]
    return (min(xs0), min(ys0), max(xs1), max(ys1))


descriptors = {}

for label in TARGET_LABELS:
    pts = label_pts.get(label, [])
    if not pts:
        descriptors[label] = {"n_instances": 0}
        continue

    instances = []
    for idx, (cx, cy, tx0, ty0, tx1, ty1) in enumerate(pts):
        ngh = neighbours(cx, cy, exclude_box=(tx0, ty0, tx1, ty1))
        if not ngh:
            continue
        # dominant colour
        col_count = Counter(s["color"] for s in ngh)
        dom_col = col_count.most_common(1)[0][0]
        ngh_dom = [s for s in ngh if s["color"] == dom_col]
        bb = cluster_bbox(ngh_dom)
        if bb is None:
            continue
        w_pt = bb[2] - bb[0]
        h_pt = bb[3] - bb[1]
        instances.append({
            "label": label,
            "idx": idx,
            "text_center": (cx, cy),
            "bbox_pt": bb,
            "w_pt": w_pt,
            "h_pt": h_pt,
            "color": dom_col,
            "n_parts": len(ngh_dom),
            "all_color_counts": dict(col_count),
        })

    if not instances:
        descriptors[label] = {"n_instances": 0, "label_text_count": len(pts)}
        continue

    descriptors[label] = {
        "label_text_count": len(pts),
        "n_instances": len(instances),
        "median_w_pt": round(median(i["w_pt"] for i in instances), 2),
        "median_h_pt": round(median(i["h_pt"] for i in instances), 2),
        "min_w_pt": round(min(i["w_pt"] for i in instances), 2),
        "max_w_pt": round(max(i["w_pt"] for i in instances), 2),
        "min_h_pt": round(min(i["h_pt"] for i in instances), 2),
        "max_h_pt": round(max(i["h_pt"] for i in instances), 2),
        "dominant_color": Counter(i["color"] for i in instances).most_common(1)[0][0],
        "color_distribution": dict(Counter(i["color"] for i in instances)),
        "median_n_parts": int(median(i["n_parts"] for i in instances)),
    }

    # Save up to 3 instance crops with overlay
    for k, inst in enumerate(instances[:3]):
        bx0, by0, bx1, by1 = inst["bbox_pt"]
        pad = 4.0
        rx0, ry0, rx1, ry1 = bx0 - pad, by0 - pad, bx1 + pad, by1 + pad
        clip = fitz.Rect(rx0, ry0, rx1, ry1)
        try:
            pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        except Exception as e:
            print(f"  pixmap err for {label}#{k}: {e}")
            continue
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
        over = img.copy()
        for s in all_draw:
            sb = s["bbox"]
            if sb[2] < rx0 or sb[0] > rx1 or sb[3] < ry0 or sb[1] > ry1:
                continue
            px0 = int((sb[0] - rx0) * zoom)
            py0 = int((sb[1] - ry0) * zoom)
            px1 = int((sb[2] - rx0) * zoom)
            py1 = int((sb[3] - ry0) * zoom)
            cv2.rectangle(over, (px0, py0), (px1, py1), color_bgr(s["color"]), 1)
        h, w = img.shape[:2]
        canvas = np.full((h, w * 2 + 20, 3), 255, dtype=np.uint8)
        canvas[:h, :w] = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        canvas[:h, w + 20: w * 2 + 20] = cv2.cvtColor(over, cv2.COLOR_RGB2BGR)
        safe = re.sub(r"[^\w]", "_", label)
        cv2.imwrite(str(OUT / f"gt_{safe}_inst_{k}.png"), canvas)


# ---------------------------------------------------------------------------
# Save descriptors JSON + summary text
# ---------------------------------------------------------------------------
out_json = OUT / "gt_descriptors.json"
with out_json.open("w", encoding="utf-8") as f:
    json.dump(descriptors, f, ensure_ascii=False, indent=2)

print()
print("=== Descriptor summary ===")
for label in TARGET_LABELS:
    d = descriptors[label]
    if d.get("n_instances", 0) == 0:
        print(f"  {label:>5s}: NO INSTANCES (text labels: {d.get('label_text_count', 0)})")
        continue
    print(f"  {label:>5s}: n={d['n_instances']:>3d}  "
          f"W={d['median_w_pt']:>5.2f}pt  H={d['median_h_pt']:>5.2f}pt  "
          f"colour={d['dominant_color']:>5s}  parts~{d['median_n_parts']}  "
          f"color_dist={d['color_distribution']}")

doc.close()
print()
print(f"Saved descriptors: {out_json}")
print(f"Output: {OUT}")
