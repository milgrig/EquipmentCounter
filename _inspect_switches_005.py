"""Inspect legend rows for switches: list every row with its symbol & black/red primitive counts."""
import io, sys
from collections import Counter
from pathlib import Path
import fitz

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
doc = fitz.open(str(PDF)); mp = doc[legend.page_index]


def cc(c):
    if not c: return None
    if isinstance(c, (tuple, list)) and len(c) >= 3:
        r, g, b = c[0], c[1], c[2]
        if r > 0.6 and g < 0.4 and b < 0.4: return "red"
        if r < 0.4 and g < 0.4 and b > 0.6: return "blue"
        if r < 0.2 and g < 0.2 and b < 0.2: return "black"
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


# Index ALL drawings with their bbox + color + a basic shape tag
prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if not bb: continue
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    if max(w, h) > 60: continue
    col = cc(d.get("fill")) or cc(d.get("color"))
    if col is None: continue
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    # detect circle-like by checking item types
    types = Counter(it[0] for it in d.get("items", []))
    is_curve = "c" in types  # bezier ~ circle
    prims.append({"bbox": bb, "cx": cx, "cy": cy, "color": col,
                  "w": w, "h": h, "is_curve": is_curve, "types": dict(types)})

print(f"Total small coloured/black primitives: {len(prims)}")
print(f"Legend rows total: {len(legend.items)}")
print()
print(f"{'idx':>3s} {'symbol':<8s} {'desc(60)':<60s} bbox(x0,y0,x1,y1)")
for i, it in enumerate(legend.items):
    sym = (it.symbol or "").strip()
    desc = (it.description or "").strip()[:60]
    print(f"{i:>3d} {sym:<8s} {desc:<60s} {tuple(round(v,1) for v in it.bbox)}")
print()

# For each row, count primitives by colour & curve
print("Per-row primitive counts (color: total / curves):")
for i, it in enumerate(legend.items):
    rx0, ry0, rx1, ry1 = it.bbox
    # left half = pictogram column
    px1 = rx0 + (rx1 - rx0) * 0.45
    in_row = [p for p in prims if rx0 - 1 <= p["cx"] <= px1 + 1 and ry0 - 1 <= p["cy"] <= ry1 + 1]
    by_col = Counter(p["color"] for p in in_row)
    by_curve = Counter(p["color"] for p in in_row if p["is_curve"])
    sym = (it.symbol or "").strip()
    desc_short = (it.description or "")[:35]
    print(f"  row {i:>2d} [{sym:<6s}|{desc_short:<35s}] "
          f"counts={dict(by_col)} curves={dict(by_curve)}")
doc.close()
