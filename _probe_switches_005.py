"""Probe row 16/17/18 with NO colour filter to find switch primitives."""
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
LX0 = legend.legend_bbox[0]


def dbox(d):
    xs, ys = [], []
    for it in d.get("items", []):
        if it[0] == "re":
            r = it[1]; xs += [r.x0, r.x1]; ys += [r.y0, r.y1]
        elif it[0] in ("l", "m", "c"):
            for p in it[1:]:
                if hasattr(p, "x"): xs.append(p.x); ys.append(p.y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None


# Print every drawing whose centre is in the switch region.
target_rows = [16, 17, 18]
print("All drawings inside row 16/17/18 (NO colour filter):\n")
for ridx in target_rows:
    it = legend.items[ridx]
    rx0, ry0, rx1, ry1 = it.bbox
    desc = (it.description or "")[:50]
    sy0, sy1 = ry0 - 10, ry1 + 4
    sx0, sx1 = LX0, rx0 + 5
    print(f"\n--- row {ridx} '{desc}' Y[{sy0:.1f}..{sy1:.1f}] X[{sx0:.1f}..{sx1:.1f}]")
    found = []
    for d in mp.get_drawings():
        bb = dbox(d)
        if not bb: continue
        cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
        if not (sx0 - 1 <= cx <= sx1 + 1 and sy0 - 1 <= cy <= sy1 + 1):
            continue
        types = Counter(it[0] for it in d.get("items", []))
        found.append({
            "bbox": bb, "cx": cx, "cy": cy,
            "fill": d.get("fill"),
            "color": d.get("color"),
            "stroke_w": d.get("width"),
            "type": d.get("type"),
            "items": dict(types),
            "n_items": len(d.get("items", [])),
        })
    print(f"  found {len(found)} drawings")
    for k, p in enumerate(sorted(found, key=lambda x: (x["cy"], x["cx"]))[:25]):
        print(f"  [{k:>2d}] bbox=({p['bbox'][0]:>7.1f},{p['bbox'][1]:>6.1f},{p['bbox'][2]:>7.1f},{p['bbox'][3]:>6.1f}) "
              f"items={p['items']} type={p['type']} fill={p['fill']} color={p['color']} sw={p['stroke_w']}")
doc.close()
