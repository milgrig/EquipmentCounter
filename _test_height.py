"""Тест расчёта высоты кабелей на 3-й захватке ГПК."""
import os, sys, glob

sys.stdout.reconfigure(encoding='utf-8')

from pdf_cable_height import (
    collect_page_annotations,
    summarize_cable_heights,
    format_summary,
    parse_elevations, parse_groups,
)

# Юнит-тесты
print('=== Unit tests ===')
samples = [
    'Гр.1-Гр.8, Гр.15-Гр.31, Гр.34, Гр.35 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200',
    'Гр.32 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200',
    'Гр.1А-Гр.6А на отм. 0.000, +9.000, +13.800',
]
for s in samples:
    elevs = parse_elevations(s)
    grps = parse_groups(s)
    h = max(float(e) for e in [e.replace(",", ".") for e in elevs]) - min(float(e) for e in [e.replace(",", ".") for e in elevs])
    print('  groups=%-3d elevs=%-3d height=%.1fm  total=%.1fm' % (
        len(grps), len(elevs), h, h * len(grps)))
    print('    text:', s[:80])

# Реальные PDF
print()
print('=== Per-PDF annotation extraction ===')
folder = r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF'
pdfs = sorted(glob.glob(os.path.join(folder, '*.pdf')))
for p in pdfs:
    try:
        anns_by_page = collect_page_annotations(p)
    except Exception as e:
        print('  [err]', os.path.basename(p), e)
        continue
    total = sum(len(v) for v in anns_by_page.values())
    if total:
        print('  %s  → %d annotations' % (os.path.basename(p)[:60], total))
        for page_anns in anns_by_page.values():
            for a in page_anns[:3]:
                print('    p%d  groups=%d elevs=%d  height=%.1fm  total=%.1fm  col=%s' % (
                    a.page_num, a.n_groups, a.n_elevations,
                    a.riser_height_m, a.total_riser_length_m, a.color))
                print('       %s' % a.text[:100])

# Сводка по всей папке
print()
print('=== TOTAL FOLDER SUMMARY ===')
summary = summarize_cable_heights(pdfs, dedupe_global=True)
print(format_summary(summary))

print()
print('=== Without dedup (count duplicates from each floor sheet) ===')
summary_dup = summarize_cable_heights(pdfs, dedupe_global=False)
print(format_summary(summary_dup))
