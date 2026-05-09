"""Render KPP legend region with overlaid LegendItem bboxes to verify alignment."""
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
print(f"Page rect: {mp.rect}")
print(f"Page rotation: {mp.rotation}")
print(f"Legend bbox: ({LX0:.1f},{LY0:.1f},{LX1:.1f},{LY1:.1f})")
zoom=400/72.0
clip=fitz.Rect(LX0-50,LY0-20,LX1+20,LY1+20)
pix=mp.get_pixmap(matrix=fitz.Matrix(zoom,zoom),clip=clip,alpha=False)
img=np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width,pix.n).copy()
if pix.n==4: img=cv2.cvtColor(img,cv2.COLOR_RGBA2RGB)
img_bgr=cv2.cvtColor(img,cv2.COLOR_RGB2BGR)
ox,oy=clip.x0,clip.y0
def to_px(x,y): return (int((x-ox)*zoom), int((y-oy)*zoom))
# legend bbox in green
(x1,y1)=to_px(LX0,LY0); (x2,y2)=to_px(LX1,LY1)
cv2.rectangle(img_bgr,(x1,y1),(x2,y2),(0,200,0),3)
# each row in red
for it in leg.items:
    x0,y0,x1,y1=it.bbox
    (px0,py0)=to_px(x0,y0); (px1,py1)=to_px(x1,y1)
    cv2.rectangle(img_bgr,(px0,py0),(px1,py1),(0,0,255),2)
    cv2.putText(img_bgr, f"[{it.symbol}]", (px0,py0-5),
                cv2.FONT_HERSHEY_SIMPLEX, 0.6,(0,0,255),2)
out=OUT/"007_legend_with_bboxes.png"
cv2.imwrite(str(out), img_bgr)
print(f"Saved {out}")
doc.close()
