"""Render the KPP legend region as a PNG so we can see what's there."""
import sys, io
from pathlib import Path
import fitz, numpy as np, cv2
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
ROOT=Path(r"C:\Cursor\TayfaProject\EquipmentCounter")
PDF=ROOT/r"Data\ДБТ разделы для ИИ\30. КПП\03_PDF\007-План освещения.pdf"
OUT=ROOT/"_kpp_inspect_out"; OUT.mkdir(exist_ok=True)
sys.path.insert(0,str(ROOT))
from pdf_legend_parser import parse_legend
leg=parse_legend(str(PDF))
LX0,LY0,LX1,LY1=leg.legend_bbox
doc=fitz.open(str(PDF)); mp=doc[leg.page_index]
zoom=400/72.0
# Render legend region with 50pt padding all around
clip=fitz.Rect(LX0-50,LY0-20,LX1+20,LY1+20)
pix=mp.get_pixmap(matrix=fitz.Matrix(zoom,zoom),clip=clip,alpha=False)
img=np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width,pix.n)
if pix.n==4: img=cv2.cvtColor(img,cv2.COLOR_RGBA2RGB)
out_path=OUT/"007_legend.png"
cv2.imwrite(str(out_path),cv2.cvtColor(img,cv2.COLOR_RGB2BGR))
print(f"Saved {out_path} ({pix.width}x{pix.height})")
doc.close()
