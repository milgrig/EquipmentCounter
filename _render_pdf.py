"""Render select PDFs from GPK 3-я захватка to PNG using fitz."""
import fitz
import os
import glob
import io

# Use glob with literal Cyrillic in the .py source (UTF-8 encoded)
import pathlib
root = pathlib.Path(__file__).parent
# Walk the data folder by listing instead of typing Cyrillic
data_root = None
for p in root.rglob("02_PDF"):
    if "3-" in str(p) and "Файлы" not in str(p):
        data_root = p
        break

buf = io.StringIO()
def w(s):
    buf.write(s + "\n")

if data_root is None:
    w("DATA ROOT NOT FOUND")
else:
    w(f"DATA ROOT: {data_root}")
    files = sorted(data_root.iterdir())
    w(f"FILES: {len(files)}")
    for f in files:
        w("  " + f.name)
    # Render 005, 012, 019 = same elevation, 3 sheets
    prefixes = ("005", "012", "019")
    for f in files:
        if f.name.startswith(prefixes) and f.suffix.lower() == ".pdf":
            doc = fitz.open(str(f))
            w(f"\n== {f.name} pages={doc.page_count}")
            for i in range(doc.page_count):
                page = doc[i]
                pix = page.get_pixmap(dpi=110)
                out = root / f"_render_{f.name[:3]}_p{i:02d}.png"
                pix.save(str(out))
                w(f"   p{i} -> {out.name} {pix.width}x{pix.height}")
            doc.close()

(root / "_render_pdf_out.txt").write_text(buf.getvalue(), encoding="utf-8")
print("Done")
