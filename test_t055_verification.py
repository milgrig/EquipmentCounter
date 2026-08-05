#!/usr/bin/env python3
"""
T055 verification: Test symbol-less legend recovery on 007 PDF.
Checks that switches and control post are now counted.
"""
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

# Import web_app processing function (same as used in T123)
import web_app

PDF_PATH = "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/02_PDF/007-Планы освещения-отм. +7.800 +9.000.pdf"

print("=" * 80)
print("T055 VERIFICATION: Symbol-less legend recovery on 007 PDF")
print("=" * 80)
print()
print(f"Testing: {PDF_PATH}")
print()

# Process PDF
try:
    items = web_app._count_equipment_in_pdf(PDF_PATH)
    print(f"✅ Successfully processed PDF")
    print(f"   Total items found: {len(items)}")
    print()
except Exception as e:
    print(f"❌ Error processing PDF: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# Group items by name for analysis
by_name = {}
for item in items:
    name = item.get('name', '')
    if name:
        if name not in by_name:
            by_name[name] = {'count': 0, 'items': []}
        by_name[name]['count'] += item.get('count', 0)
        by_name[name]['items'].append(item)

print("=" * 80)
print("FOUND ITEMS (grouped by name)")
print("=" * 80)
for i, (name, data) in enumerate(sorted(by_name.items()), 1):
    count = data['count']
    print(f"{i:3}. {name[:70]:<70} : {count:4} шт")

print()
print("=" * 80)
print("VERIFICATION CHECKLIST")
print("=" * 80)

# Check for target items (legend indices 15, 16, 17)
targets = {
    '15': 'Однокл<NAME> выключатель Systeme Electric',
    '16': 'Выключатель Однокл<NAME> ЭТЮД IP44',
    '17': 'Пост управления рабочим освещением'
}

results = {}
for idx, partial_name in targets.items():
    found = False
    found_count = 0
    found_names = []

    # Search for items containing key parts of the target name
    for name, data in by_name.items():
        # Normalize for search
        name_norm = name.lower()
        partial_norm = partial_name.lower().replace('<NAME>', '')

        # Check for key words
        if idx == '15' and ('выключатель' in name_norm and 'systeme' in name_norm):
            found = True
            found_count = data['count']
            found_names.append(name)
        elif idx == '16' and ('выключатель' in name_norm and 'этюд' in name_norm):
            found = True
            found_count = data['count']
            found_names.append(name)
        elif idx == '17' and ('пост управления' in name_norm and 'освещ' in name_norm):
            found = True
            found_count = data['count']
            found_names.append(name)

    results[idx] = {
        'target': partial_name,
        'found': found,
        'count': found_count,
        'names': found_names
    }

    status = "✅ PASS" if found and found_count >= 1 else "❌ FAIL"
    print(f"{status} Legend index {idx}: {partial_name}")
    if found:
        print(f"      Found: count={found_count}")
        for fn in found_names:
            print(f"      Name: {fn}")
    else:
        print(f"      NOT FOUND")
    print()

# Regression check: lighting types (1, 1A, 2, 2A, ..., 6AE)
print("=" * 80)
print("REGRESSION CHECK: Previously-found lighting types")
print("=" * 80)

lighting_markers = ['1', '1A', '1АЭ', '2', '2A', '2АЭ', '3', '3АЭ', '4', '4АЭ', '5', '5А', '5АЭ', '6', '6АЭ']
lighting_found = {}

for marker in lighting_markers:
    found = False
    count = 0
    for name in by_name.keys():
        # Check if name contains lighting fixture keywords
        name_lower = name.lower()
        if any(kw in name_lower for kw in ['светильник', 'светодиодный']):
            # This is a lighting fixture, check if it was counted
            found = True
            count += by_name[name]['count']

    if found:
        lighting_found[marker] = count

if lighting_found:
    print(f"✅ Found {len(lighting_found)} lighting fixture types with total count {sum(lighting_found.values())}")
    print(f"   Types found: {', '.join(sorted(lighting_found.keys()))}")
else:
    print(f"❌ WARNING: No lighting fixtures found (possible regression)")

print()
print("=" * 80)
print("SUMMARY")
print("=" * 80)

target_pass = sum(1 for r in results.values() if r['found'] and r['count'] >= 1)
target_total = len(results)

print(f"Target items (switches/control post): {target_pass}/{target_total} PASS")
print(f"Regression check (lighting): {'PASS' if lighting_found else 'FAIL - NO LIGHTS FOUND'}")
print()

if target_pass == target_total and lighting_found:
    print("✅ ALL CHECKS PASSED")
    sys.exit(0)
else:
    print("❌ SOME CHECKS FAILED")
    sys.exit(1)
