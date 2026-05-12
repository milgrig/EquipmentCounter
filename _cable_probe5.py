"""Inspect all TC tables with cables — counts of brand/length cells."""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
from _cable_probe2 import iter_entities, _clean

p = ('Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/'
     '01_DWG/003, 004 - Схемы ЩО, ЩАО.dxf')

tables = []
for ent, pairs in iter_entities(p):
    if ent != 'TABLECONTENT':
        continue
    if not any('ВБШвнг' in v or 'ППГнг' in v for _, v in pairs):
        continue
    cells = []
    prev_content = False
    for code, val in pairs:
        if code != '302':
            continue
        v = val.strip()
        if v == 'CONTENT':
            prev_content = True; continue
        if v == 'GRIDFORMAT' or not v:
            prev_content = False; continue
        if prev_content:
            cells.append(_clean(v))
        prev_content = False
    tables.append(cells)

for idx, cells in enumerate(tables):
    brands = [i for i, c in enumerate(cells) if re.search(r'ВБШвнг|ППГнг', c)]
    lens = [i for i, c in enumerate(cells) if re.search(r'L\s*[=\-]\s*\d', c)]
    secs = [i for i, c in enumerate(cells)
            if re.fullmatch(r'\d+[.,]\d+', c.strip())]
    print(f'--- Table #{idx}: cells={len(cells)} brands={len(brands)} '
          f'lens={len(lens)} secs={len(secs)}')
    # Print bounding indices
    if brands:
        print(f'  brand indices: {brands[0]}..{brands[-1]}')
    if lens:
        print(f'  len indices: {lens[0]}..{lens[-1]}')
    if secs:
        print(f'  sec indices: {secs[0]}..{secs[-1]}')
    # Show snippets
    print('  first 40 cells:')
    for k in range(min(40, len(cells))):
        print(f'    [{k:4}] {cells[k][:60]}')
    print('  ...')
    # Show cells around first brand
    if brands:
        i0 = brands[0]
        print(f'  around first brand [{i0}]:')
        for k in range(max(0, i0-3), min(len(cells), i0+10)):
            print(f'    [{k:4}] {cells[k][:60]}')
    print()
