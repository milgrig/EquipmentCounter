import sys
from docx import Document
from vor_elevation_grouper import regroup_aggregate_by_elevation, extract_elevation

sys.stdout.reconfigure(encoding='utf-8')

tests = [
    '005-Планы освещения-отм. 0.000.pdf',
    '006-Планы освещения-отм. +4.200.pdf',
    '007-Планы освещения-отм. +7.800 +9.000.pdf',
    '012-План привязки на отм. 0.000.pdf',
    '019-Лотки отм. 0.000.pdf',
    '003 - Схемы ЩО, ЩАО-ЩО-3.pdf',
]
for t in tests:
    print('  %-55s -> %r' % (t, extract_elevation(t)))

print()
cur = Document(r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\VOR_02_PDF.docx')
tbl = cur.tables[0]
agg = []
for ri, row in enumerate(tbl.rows[1:], start=1):
    cells = [c.text.strip() for c in row.cells]
    if not cells[1]:
        continue
    try:
        total = float(cells[3].replace(',', '.'))
    except ValueError:
        total = cells[3]
    agg.append({
        'row': int(cells[0]) if cells[0].isdigit() else ri,
        'name': cells[1], 'unit': cells[2], 'total': total,
        'formula': cells[4], 'drawing_refs': cells[5],
    })

print('BEFORE regroup: %d rows' % len(agg))
print('  [0] %s | total=%s' % (agg[0]['name'][:40], agg[0]['total']))
print('      formula = %s' % agg[0]['formula'])
print('      refs    = %s' % agg[0]['drawing_refs'][:80])

regrouped = regroup_aggregate_by_elevation(agg, labeled=False)
print()
print('AFTER regroup (labeled=False):')
for it in regrouped[:6]:
    pe = it.get('per_elevation', {})
    print('  %-55s | total=%-5s | formula=%s' % (
        it['name'][:55], it['total'], it['formula']))
    if pe:
        pe_str = ', '.join('%s:%g' % (k, v) for k, v in pe.items())
        print('      per_elevation: %s' % pe_str)

regrouped_lbl = regroup_aggregate_by_elevation(agg, labeled=True)
print()
print('AFTER regroup (labeled=True):')
for it in regrouped_lbl[:3]:
    print('  %-55s | %s' % (it['name'][:55], it['formula']))
