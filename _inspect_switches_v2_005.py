"""Find switch pictograms in legend with relaxed filters."""
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
        if r < 0.25 and g < 0.25 and b < 0.25: return "black"
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


prims = []
for d in mp.get_drawings():
    bb = dbox(d)
    if not bb: continue
    w, h = bb[2] - bb[0], bb[3] - bb[1]
    # NO size cap here
    col = cc(d.get("fill")) or cc(d.get("color"))
    if col != "black": continue
    cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
    types = Counter(it[0] for it in d.get("items", []))
    is_curve = "c" in types
    prims.append({"bbox": bb, "cx": cx, "cy": cy, "color": col,
                  "w": w, "h": h, "is_curve": is_curve, "types": dict(types)})

print(f"Black primitives total: {len(prims)}")

# Examine rows 16, 17, 18 with FULL row bbox and PICTOGRAM column
target_rows = [16, 17, 18]
for ridx in target_rows:
    it = legend.items[ridx]
    rx0, ry0, rx1, ry1 = it.bbox
    desc = (it.description or "")[:50]
    print(f"\n--- row {ridx}: '{desc}' bbox=({rx0:.1f},{ry0:.1f},{rx1:.1f},{ry1:.1f})")
    # extend search a bit above & below
    sy0 = ry0 - 8; sy1 = ry1 + 4
    # pictogram column ≈ 50pt to the LEFT of description start (rx0 is description x0)
    # Looking at image, pictogram column is the 60-100pt before the description text.
    # Use absolute column [legend_bbox.x0, rx0]
    LX0 = legend.legend_bbox[0]
    sx0 = LX0; sx1 = rx0 + 5
    in_col = [p for p in prims
              if sx0 - 1 <= p["cx"] <= sx1 + 1 and sy0 - 1 <= p["cy"] <= sy1 + 1]
    print(f"  search col [{sx0:.1f}..{sx1:.1f}], y [{sy0:.1f}..{sy1:.1f}]: {len(in_col)} prims")
    by_curve = sum(1 for p in in_col if p["is_curve"])
    by_type = Counter()
    for p in in_col:
        for k, v in p["types"].items():
            by_type[k] += v
    print(f"  curves={by_curve}  item_types={dict(by_type)}")
    # show first 15 prims with size
    for p in sorted(in_col, key=lambda p: (p["cy"], p["cx"]))[:20]:
        print(f"    bbox=({p['bbox'][0]:>7.1f},{p['bbox'][1]:>7.1f},{p['bbox'][2]:>7.1f},{p['bbox'][3]:>7.1f}) "
              f"w={p['w']:>5.2f} h={p['h']:>5.2f} curve={p['is_curve']} types={p['types']}")
doc.close()
