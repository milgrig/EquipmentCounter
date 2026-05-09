"""Проверка извлечения аннотаций (v2 через fitz)."""
import sys
import os
from pdf_cable_annotations_v2 import (
    collect_page_annotations,
    vertical_riser_length_m,
    parse_elevations, parse_groups,
    page_summary,
)

sys.stdout.reconfigure(encoding='utf-8')

print('=== Unit tests ===')
samples = [
    'Гр.1-Гр.8, Гр.15-Гр.31, Гр.34, Гр.35 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200',
    'Гр.32 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200',
    'на отм. +4.200',
]
for s in samples:
    e = parse_elevations(s)
    g = parse_groups(s)
    r = vertical_riser_length_m(e, n_groups=len(g) or 1)
    print('  %s' % s[:70])
    print('    elevs=%s groups=%d risers=%.1fm' % (e, len(g), r))

folder = r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF'
print()
print('=== Real PDFs (all 02_PDF/) ===')
all_files = sorted(os.listdir(folder))
for fn in all_files:
    if not fn.lower().endswith('.pdf'):
        continue
    fp = os.path.join(folder, fn)
    try:
        anns = collect_page_annotations(fp)
    except Exception as e:
        print('  [error] %s: %s' % (fn, e))
        continue
    total = sum(len(v) for v in anns.values())
    if total == 0:
        continue
    summary_total = {'avg': 0, 'max': 0, 'grp': 0}
    for pn, lst in anns.items():
        s = page_summary(lst)
        summary_total['avg'] = max(summary_total['avg'], s['avg_multiplier'])
        summary_total['max'] = max(summary_total['max'], s['max_multiplier'])
        summary_total['grp'] += s['total_groups']
    print('  %-55s anns=%d max_mult=%d total_groups=%d' % (
        fn, total, summary_total['max'], summary_total['grp']))
    # покажем 2 примера
    shown = 0
    for pn, lst in anns.items():
        for a in lst:
            if shown >= 2:
                break
            print('     [p%d col=%s elev=%d grp=%d] %s' % (
                pn, a.color, len(a.elevations), len(a.groups), a.text[:120]))
            shown += 1
