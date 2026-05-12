"""Проверка извлечения аннотаций из реальных PDF 02_PDF/3-я захватка."""
import sys
import os
from pdf_cable_annotations import (
    collect_page_annotations,
    vertical_riser_length_m,
    parse_elevations, parse_groups,
)

sys.stdout.reconfigure(encoding='utf-8')

# Парсинг текста — юнит-тест
print('=== Unit tests ===')
samples = [
    'Гр.1-Гр.8, Гр.15-Гр.31, Гр.34, Гр.35 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200',
    'Гр.32 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200',
    'на отм. +4.200',
    'Гр.1А-Гр.6А, Гр.13А-Гр.32А, Гр.34А, Гр.35А, Гр.36А на отм. 0.000, +9.000, +13.800',
]
for s in samples:
    elevs = parse_elevations(s)
    grps = parse_groups(s)
    risers = vertical_riser_length_m(elevs, n_groups=len(grps) or 1)
    print(f'  text: {s[:70]}')
    print(f'    elevs={elevs}  groups={len(grps)}  risers={risers:.2f}m')

# Реальные PDF
folder = r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\02_PDF'
print()
print('=== Real PDFs ===')
candidates = [
    '019-Лотки отм. 0.000.pdf',
    '020-Лотки отм. +4.200.pdf',
    '021-Лотки на отм. +9.000.pdf',
    '005-Планы освещения-отм. 0.000.pdf',
    '012-План привязки на отм. 0.000.pdf',
]
for fn in candidates:
    fp = os.path.join(folder, fn)
    if not os.path.exists(fp):
        print(f'  [missing] {fn}')
        continue
    try:
        anns = collect_page_annotations(fp)
    except Exception as e:
        print(f'  [error] {fn}: {e}')
        continue
    total = sum(len(v) for v in anns.values())
    print(f'  {fn}: pages_with_anns={len(anns)} total_anns={total}')
    for page_no, lst in list(anns.items())[:1]:
        for a in lst[:5]:
            print(f'    p{page_no} ({a.x:.0f},{a.y:.0f}) col={a.color} '
                  f'elev={len(a.elevations)} grp={len(a.groups)} | '
                  f'{a.text[:90]}')
