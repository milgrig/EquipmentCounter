"""Render СО-СО_03 and СО-СО_04 (spec sheets) to PNG."""
import fitz, pathlib, io

root = pathlib.Path(__file__).parent
data_root = None
for p in root.rglob("02_PDF"):
    if "3-" in str(p) and "Файлы" not in str(p):
        data_root = p
        break

buf = io.StringIO()
def w(s): buf.write(s+"\n")

if data_root:
    for f in sorted(data_root.iterdir()):
        if f.name.startswith("\u0421\u041e-\u0421\u041e_0") and f.suffix.lower()==".pdf":
            doc = fitz.open(str(f))
            w(f"=== {f.name} pages={doc.page_count}")
            for i in range(doc.page_count):
                pix = doc[i].get_pixmap(dpi=130)
                out = root / f"_render_so{f.stem[-2:]}_p{i:02d}.png"
                pix.save(str(out))
                w(f"  p{i} -> {out.name} {pix.width}x{pix.height}")
            doc.close()

(root/"_render_so_out.txt").write_text(buf.getvalue(), encoding="utf-8")
print("Done")
