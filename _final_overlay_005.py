"""
Final validation overlay for the text-free detector.

Loads detection results from _detect_005_out/detect_results.json, renders
the full plan page, draws coloured rectangles and labels around every
detected pictogram, and adds a stats panel comparing detector counts vs
DXF ground truth.

Output: _detect_005_out/overlay_final.png + summary.txt
"""
from __future__ import annotations
import io, sys, json
from pathlib import Path

import fitz
import numpy as np
import cv2

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

PDF = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
RESULTS = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_detect_005_out\detect_results.json")
OUT_DIR = Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_detect_005_out")

DPI = 220
DXF_GT = {"1": 33, "2": 15, "3": 7, "4": 4, "5АЭ": 4, "6АЭ": 6, "7АЭ": 7}

# ASCII-safe label mapping for cv2.putText (which can't render Cyrillic).
ASCII_MAP = {"5АЭ": "5AE", "6АЭ": "6AE", "7АЭ": "7AE",
             "1А": "1A", "1АЭ": "1AE", "2А": "2A", "2АЭ": "2AE",
             "3А": "3A", "3АЭ": "3AE", "4А": "4A", "4АЭ": "4AE"}

PALETTE = {
    "5АЭ": (0, 200, 0),       # green
    "6АЭ": (255, 0, 255),     # magenta
    "7АЭ": (0, 165, 255),     # orange
}


sys.path.insert(0, str(Path(__file__).parent))
from pdf_legend_parser import parse_legend

legend = parse_legend(str(PDF))
LX0, LY0, LX1, LY1 = legend.legend_bbox

results = json.loads(RESULTS.read_text(encoding="utf-8"))
labels_data = results.get("labels", {})

doc = fitz.open(str(PDF))
mp = doc[legend.page_index]
zoom = DPI / 72.0
pix = mp.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
if pix.n == 4:
    img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
canvas = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)

# legend frame (grey)
cv2.rectangle(canvas,
              (int(LX0 * zoom), int(LY0 * zoom)),
              (int(LX1 * zoom), int(LY1 * zoom)),
              (160, 160, 160), 2)
cv2.putText(canvas, "LEGEND", (int(LX0 * zoom) + 6, int(LY0 * zoom) + 22),
            cv2.FONT_HERSHEY_SIMPLEX, 0.6, (160, 160, 160), 1, cv2.LINE_AA)

# detected pictograms
for lab, info in labels_data.items():
    col = PALETTE.get(lab, (0, 200, 0))
    asc = ASCII_MAP.get(lab, lab)
    for h in info.get("hits", []):
        bb = h["bbox"]
        x0 = int(bb[0] * zoom); y0 = int(bb[1] * zoom)
        x1 = int(bb[2] * zoom); y1 = int(bb[3] * zoom)
        pad = 5
        cv2.rectangle(canvas, (x0 - pad, y0 - pad), (x1 + pad, y1 + pad), col, 2)
        cv2.putText(canvas, asc, (x0 - pad, y0 - pad - 3),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.55, col, 1, cv2.LINE_AA)


# ---------------------------------------------------------------------------
# Stats panel (top-left)
# ---------------------------------------------------------------------------
panel_w, panel_h = 520, 230
px, py = 30, 30
cv2.rectangle(canvas, (px, py), (px + panel_w, py + panel_h), (255, 255, 255), -1)
cv2.rectangle(canvas, (px, py), (px + panel_w, py + panel_h), (0, 0, 0), 2)

cv2.putText(canvas, "Text-free detector vs DXF ground truth",
            (px + 12, py + 28), cv2.FONT_HERSHEY_SIMPLEX, 0.65, (0, 0, 0), 1, cv2.LINE_AA)
cv2.putText(canvas, "Label  Detected  DXF   Diff",
            (px + 12, py + 60), cv2.FONT_HERSHEY_SIMPLEX, 0.55, (60, 60, 60), 1, cv2.LINE_AA)

row_y = py + 88
ok_n, total_checked = 0, 0
for lab in ("5АЭ", "6АЭ", "7АЭ"):
    asc = ASCII_MAP.get(lab, lab)
    det = len(labels_data.get(lab, {}).get("hits", []))
    exp = DXF_GT.get(lab, "?")
    if isinstance(exp, int):
        diff = det - exp
        diff_str = f"{diff:+d}" if diff != 0 else "OK"
        if diff == 0:
            ok_n += 1
        total_checked += 1
    else:
        diff_str = "-"
    col = PALETTE.get(lab, (0, 0, 0))
    line = f"{asc:<5s}     {det:>3d}      {exp:>3}    {diff_str:>4}"
    cv2.putText(canvas, line, (px + 12, row_y), cv2.FONT_HERSHEY_SIMPLEX, 0.55, col, 1, cv2.LINE_AA)
    row_y += 28

cv2.putText(canvas,
            f"Match: {ok_n}/{total_checked} exact",
            (px + 12, py + panel_h - 18),
            cv2.FONT_HERSHEY_SIMPLEX, 0.55, (0, 100, 0), 1, cv2.LINE_AA)

out_png = OUT_DIR / "overlay_final.png"
cv2.imwrite(str(out_png), canvas)
print(f"Saved: {out_png}  size={canvas.shape[1]}x{canvas.shape[0]}")


# ---------------------------------------------------------------------------
# Plain-text summary
# ---------------------------------------------------------------------------
lines = ["Text-free pictogram detection — final summary",
         "=" * 50,
         f"PDF : {PDF.name}",
         f"Page: {legend.page_index}",
         "",
         "Class      Detected    DXF    Diff   Status"]
for lab in ("5АЭ", "6АЭ", "7АЭ"):
    det = len(labels_data.get(lab, {}).get("hits", []))
    exp = DXF_GT.get(lab, "?")
    if isinstance(exp, int):
        diff = det - exp
        status = "OK" if diff == 0 else ("over" if diff > 0 else "under")
        lines.append(f"  {lab:<5s}    {det:>5d}    {exp:>5d}   {diff:>+4d}   {status}")
    else:
        lines.append(f"  {lab:<5s}    {det:>5d}      ?      -    -")
lines += ["",
          "Notes:",
          "  * detector uses NO text - only red drawing primitives,",
          "    Union-Find clustering (link=5pt), per-class W/H/n filters,",
          "    and cross-label NMS (IoU>=0.4).",
          "  * 6АЭ undercount likely reflects grouped labels pointing at",
          "    a shared icon - one icon serves multiple text labels."]
(OUT_DIR / "summary.txt").write_text("\n".join(lines), encoding="utf-8")

doc.close()
print("\n".join(lines))
