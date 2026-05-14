"""Diagnostic: does _detect_duplicate_floor_plans actually fire on GPK 3-я захватка?"""
import pathlib, io, sys
from vor_compose import _detect_duplicate_floor_plans, _floor_mark

root = pathlib.Path(__file__).parent
data_root = None
for p in root.rglob("02_PDF"):
    if "3-" in str(p) and "Файлы" not in str(p):
        data_root = p
        break

buf = io.StringIO()
def w(s): buf.write(s+"\n")

if not data_root:
    w("NOT FOUND");
else:
    files = sorted(f.name for f in data_root.iterdir() if f.suffix.lower()==".pdf")
    w(f"TOTAL PDF FILES: {len(files)}")
    w("")
    w("Per-file floor_mark + classification:")
    privyazka = "\u043f\u0440\u0438\u0432\u044f\u0437\u043a"
    privyazka_typo = "\u043f\u0440\u0438\u044f\u0432\u044f\u0437\u043a"
    osvesh = "\u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438"
    for nm in files:
        mark = _floor_mark(nm)
        lower = nm.lower()
        cls = ""
        if privyazka in lower or privyazka_typo in lower:
            cls = "PRIVYAZKA"
        elif osvesh in lower:
            cls = "OSVESHCHENIYA"
        w(f"  mark={mark!r:>10}  cls={cls:13}  {nm}")
    w("")
    skip = _detect_duplicate_floor_plans(files)
    w(f"SKIP SET ({len(skip)} files):")
    for nm in sorted(skip):
        w(f"  - {nm}")

(root/"_check_dedup_out.txt").write_text(buf.getvalue(), encoding="utf-8")
print("Done")
