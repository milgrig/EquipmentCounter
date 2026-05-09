"""Derive coefficients from etalon for postprocessor."""
import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

doc = Document('Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/ВОР ЭО, Захватка 3_ГПК.docx')
t = doc.tables[0]

# Collect tray metrics and hardware from etalon
tray_m = 0
cables_all = 0  # "каб." count
total_lum = 0  # all luminaires

section = ""
for row in t.rows:
    cells = [c.text.strip() for c in row.cells]
    while len(cells) < 7:
        cells.append('')
    num, name, unit, qty_raw = cells[0], cells[1], cells[2], cells[3]
    if not num and not unit and not qty_raw and name:
        section = name
        continue
    # parse qty
    q = 0
    m = re.match(r'(\d+(?:[.,]\d+)?)', qty_raw)
    if m:
        q = float(m.group(1).replace(',', '.'))

    low = name.lower()
    if section.startswith('Монтаж кабельных лотков') and 'лоток' in low and 'штампованный' in low and unit == 'м':
        tray_m += q
    if 'измерение сопротивления изоляции' in low and unit == 'каб.':
        cables_all = int(q)
    if section.startswith('Светотехническое') and unit == 'шт' and 'монтаж' in low:
        total_lum += int(q)

print(f"Total tray (m): {tray_m}")
print(f"Total cables: {cables_all}")
print(f"Total luminaires work rows sum: {total_lum}")

# Now collect all hardware amounts from etalon's tray section
print("\n=== Hardware in lotki section ===")
section = ""
for row in t.rows:
    cells = [c.text.strip() for c in row.cells]
    while len(cells) < 7:
        cells.append('')
    num, name, unit, qty_raw = cells[0], cells[1], cells[2], cells[3]
    if not num and not unit and not qty_raw and name:
        section = name
        continue
    if not section.startswith('Монтаж кабельных лотков'):
        continue
    m = re.match(r'(\d+(?:[.,]\d+)?)', qty_raw)
    q = float(m.group(1).replace(',', '.')) if m else 0
    if q > 0 and unit == 'шт':
        coef = q / tray_m if tray_m else 0
        print(f"  {q:>6.0f} шт | coef/m={coef:.4f} | {name[:60]}")

# Also: how much cable per meter of tray
print(f"\nCables in etalon: 116 каб.")
print(f"Tray total: {tray_m}")
print(f"Approx cables/tray ratio: {116/tray_m if tray_m else 0:.4f}")
