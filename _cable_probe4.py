"""Inspect biggest TABLECONTENT block cells."""
import sys
sys.stdout.reconfigure(encoding='utf-8')

from _cable_probe2 import iter_entities, _clean
import re

p = ('Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/'
     '01_DWG/003, 004 - Схемы ЩО, ЩАО.dxf')

# Найдём самый большой TC с кабелями
best = None
best_n = 0
for ent, pairs in iter_entities(p):
    if ent == 'TABLECONTENT' and len(pairs) > best_n:
        if any('ВБШвнг' in v or 'ППГнг' in v for _, v in pairs):
            best = pairs
            best_n = len(pairs)

print(f'Biggest TC: {best_n} pairs')

# Собираем cells
cells = []
prev_content = False
for code, val in best:
    if code != '302':
        continue
    v = val.strip()
    if v == 'CONTENT':
        prev_content = True
        continue
    if v == 'GRIDFORMAT' or not v:
        prev_content = False
        continue
    if prev_content:
        cells.append(_clean(v))
    prev_content = False

print(f'Total cells: {len(cells)}')
cable_cells = [(i, c) for i, c in enumerate(cells)
               if re.search(r'ВБШвнг|ППГнг', c)]
len_cells = [(i, c) for i, c in enumerate(cells)
             if re.search(r'L\s*[=\-]\s*\d', c)]
print(f'Cells with cable brand: {len(cable_cells)}')
print(f'Cells with L=: {len(len_cells)}')

# Соседние клетки для первого и последнего
if cable_cells:
    for pick_idx in (0, len(cable_cells)//2, len(cable_cells)-1):
        i0 = cable_cells[pick_idx][0]
        print(f'\nContext around cable cell [{i0}]:')
        for k in range(max(0, i0-2), min(len(cells), i0+8)):
            mark = '<<' if k == i0 else '  '
            print(f'  {mark}[{k:6}] {cells[k][:100]}')

# Распределение длин и кабелей по расстоянию
print('\nDistribution of cable-brand cells in sequence:')
brand_positions = [i for i, _ in cable_cells]
len_positions = [i for i, _ in len_cells]
print(f'  brand first: {brand_positions[:10]}')
print(f'  brand last: {brand_positions[-10:]}')
print(f'  len first: {len_positions[:10]}')
print(f'  len last: {len_positions[-10:]}')
