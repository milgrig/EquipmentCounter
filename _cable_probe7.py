"""Scan СО.dxf for cables."""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
from _cable_probe2 import iter_entities, _clean
from collections import defaultdict

CABLE_FULL_RE = re.compile(
    r'(ВБШвнг\(А\)-FRLS|ВБШвнг\(А\)-LS|ППГнг\(А\)-FRHF|ППГнг\(А\)-HF)'
    r'\s+(\d+)[хx×](\d+(?:[.,]\d+)?)'
    r'(.*?)'
    r'L\s*[=\-]\s*(\d+)',
    re.IGNORECASE | re.DOTALL,
)


def scan(p):
    total_tc = 0
    total_mt = 0
    cables_tc = 0
    cables_mt = 0
    items = []
    for ent, pairs in iter_entities(p):
        if ent == 'TABLECONTENT':
            total_tc += 1
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
            for c in cells:
                m = CABLE_FULL_RE.search(c)
                if m:
                    cables_tc += 1
                    items.append({
                        "src": "TC",
                        "brand": m.group(1),
                        "section": f"{m.group(2)}х{m.group(3).replace('.', ',')}",
                        "length": int(m.group(5)),
                        "text": c[:90],
                    })
        elif ent == 'MTEXT':
            total_mt += 1
            text_parts = []
            for code, val in pairs:
                if code in ('1', '3'):
                    text_parts.append(val)
            text = _clean(''.join(text_parts))
            m = CABLE_FULL_RE.search(text)
            if m:
                cables_mt += 1
                items.append({
                    "src": "MT",
                    "brand": m.group(1),
                    "section": f"{m.group(2)}х{m.group(3).replace('.', ',')}",
                    "length": int(m.group(5)),
                    "text": text[:90],
                })
    return total_tc, total_mt, cables_tc, cables_mt, items


import os
folder = 'Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG'
grand = []
for f in sorted(os.listdir(folder)):
    if not f.lower().endswith('.dxf'):
        continue
    p = os.path.join(folder, f)
    t_tc, t_mt, c_tc, c_mt, items = scan(p)
    if c_tc or c_mt:
        L_tc = sum(i['length'] for i in items if i['src'] == 'TC')
        L_mt = sum(i['length'] for i in items if i['src'] == 'MT')
        print(f'  {f[:55]:55s}  TC: {t_tc:3d} entities '
              f'({c_tc:3d} cables, {L_tc:5d} m)  MT: {c_mt:3d} cables ({L_mt} m)')
    grand.extend(items)

print(f'\n--- GRAND: {len(grand)} cable rows ---')
print(f'    TC only: {sum(1 for i in grand if i["src"]=="TC")}  '
      f'L={sum(i["length"] for i in grand if i["src"]=="TC")}')
print(f'    MT only: {sum(1 for i in grand if i["src"]=="MT")}  '
      f'L={sum(i["length"] for i in grand if i["src"]=="MT")}')

by_bs = defaultdict(lambda: [0, 0])
for c in grand:
    k = f"{c['brand']} {c['section']}"
    by_bs[k][0] += 1
    by_bs[k][1] += c['length']
print(f'\n  По маркам (raw, both sources):')
for k, (n, L) in sorted(by_bs.items(), key=lambda x: -x[1][1]):
    print(f'    {k:40s}  {n:4d} каб  {L:6d} м')
