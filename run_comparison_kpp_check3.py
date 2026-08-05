"""Run VOR comparison for КПП to check for regressions."""
import sys, os
os.environ["PYTHONUNBUFFERED"] = "1"
os.environ["VOR_VISUAL"] = "0"
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)

from vor_comparison_xlsx import generate_comparison_xlsx

folder = 'Data/ДБТ разделы для ИИ/30. КПП/03_PDF'
output = folder + '/VOR_COMPARISON_KPP_CHECK3.xlsx'

print(f"KPP Check 3: luminaire model guard, all fixes", flush=True)
generate_comparison_xlsx(folder, output)
print(f"\nDone! {output}", flush=True)
