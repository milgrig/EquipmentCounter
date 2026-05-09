"""Render small crops of the 'horizontal' clusters to see what they are."""
import sys, io
from pathlib import Path
import fitz, numpy as np, cv2
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
PDF=Path(r"C:\Cursor\TayfaProject\EquipmentCounter\Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF\005-Планы освещения-отм. 0.000.pdf")
OUT=Path(r"C:\Cursor\TayfaProject\EquipmentCounter\_shape_005_out")
doc=fitz.open(str(PDF)); mp=doc[0]
zoom=600/72.0
# coordinates of the horizontal "mystery" clusters from diag
candidates = [
    ("horiz1", 665, 496, 25, 11),
    ("horiz2", 704, 1200, 21, 10),
    ("vert1",  1717, 2114, 11, 13),  # known 5АЭ vertical
    ("6AE_canonical", 1677, 617, 14, 29),
    ("7AE_canonical", 991, 1118, 10, 10),
]
for name,cx,cy,w,h in candidates:
    pad=8
    x0=cx-w/2-pad; y0=cy-h/2-pad; x1=cx+w/2+pad; y1=cy+h/2+pad
    clip=fitz.Rect(x0,y0,x1,y1)
    pix=mp.get_pixmap(matrix=fitz.Matrix(zoom,zoom),clip=clip,alpha=False)
    img=np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width,pix.n)
    if pix.n==4: img=cv2.cvtColor(img,cv2.COLOR_RGBA2RGB)
    cv2.imwrite(str(OUT/f"view_{name}.png"),cv2.cvtColor(img,cv2.COLOR_RGB2BGR))
    print(f"  saved view_{name}.png  cx={cx} cy={cy} w={w} h={h}")
doc.close()
