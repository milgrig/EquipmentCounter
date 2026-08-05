"""Run VOR comparison WITHOUT visual counting for ГПК 3-я захватка."""
import sys, os
os.environ["PYTHONUNBUFFERED"] = "1"
os.environ["VOR_VISUAL"] = "0"  # Disable visual counting
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)

from vor_comparison_xlsx import generate_comparison_xlsx

folder = 'Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF'
output = folder + '/VOR_COMPARISON_NOVISUAL9.xlsx'

print(f"NoVisual9: + luminaire model guard in comparison", flush=True)
generate_comparison_xlsx(folder, output)
print(f"\nDone! {output}", flush=True)
