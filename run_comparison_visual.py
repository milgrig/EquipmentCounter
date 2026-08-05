"""Run VOR comparison with visual counting for ГПК 3-я захватка."""
import sys, os
os.environ["PYTHONUNBUFFERED"] = "1"
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)

from vor_comparison_xlsx import generate_comparison_xlsx

folder = 'Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF'
output = folder + '/VOR_COMPARISON_VISUAL6.xlsx'

print(f"Visual6: no Ex-filter, numeric guard, model_key 100, indicator smooth, visual merge", flush=True)
generate_comparison_xlsx(folder, output)
print(f"\nDone! {output}", flush=True)
