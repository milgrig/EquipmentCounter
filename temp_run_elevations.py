import sys, os, glob
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, '.')

from equipment_counter import process_dxf

dxf_dir = r'Data\ДБТ разделы для ИИ\03_ГПК_\3-я захватка\_converted_dxf\01_DWG'
dxf_files = sorted(glob.glob(os.path.join(dxf_dir, '*.dxf')))

# Filter to only files 005-011 (План освещения)
dxf_files = [f for f in dxf_files if os.path.basename(f)[:3] in ('005','006','007','008','009','010','011')]

print(f"Found {len(dxf_files)} DXF files\n")

totals = {}

for fpath in dxf_files:
    fname = os.path.basename(fpath)
    print(f"\n{'='*60}")
    print(f"FILE: {fname}")
    
    try:
        items = process_dxf(fpath)
        print(f"  Equipment types: {len(items)}")
        file_total = 0
        for item in items:
            symbol = getattr(item, 'symbol', '?')
            name = getattr(item, 'name', '?')
            count = getattr(item, 'count', 0)
            count_ae = getattr(item, 'count_ae', 0)
            total = getattr(item, 'total', count)
            print(f"  {symbol}: {name} = {total} (count={count}, ae={count_ae})")
            file_total += total
            
            key = f"{symbol}: {name}"
            totals[key] = totals.get(key, 0) + total
        print(f"  FILE TOTAL: {file_total}")
    except Exception as e:
        print(f"  ERROR: {e}")

print(f"\n{'='*60}")
print("GRAND TOTALS (all elevations combined):")
print(f"{'='*60}")
grand = 0
for key in sorted(totals.keys()):
    print(f"  {key}: {totals[key]}")
    grand += totals[key]
print(f"\n  GRAND TOTAL: {grand}")
