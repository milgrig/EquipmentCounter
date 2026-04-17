#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Improved VOR parser - extracts exact equipment counts
"""

import re
import sys
from pathlib import Path

# Fix Windows console encoding
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

def parse_vor_improved(txt_path: str):
    """Parse VOR document with improved logic"""
    with open(txt_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    equipment = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()

        # Look for equipment entries - check if next line has "шт" (pieces)
        if i + 2 < len(lines):
            next_line = lines[i + 1].strip()
            third_line = lines[i + 2].strip()

            # Pattern: Equipment name, then "шт", then count
            if next_line == 'шт' and third_line.isdigit():
                count = int(third_line)
                # Determine category
                category = 'Unknown'
                if 'светильник' in line.lower() or 'светодиодн' in line.lower() or 'led' in line.lower():
                    category = 'Luminaire'
                elif 'монтаж светильник' in line.lower():
                    category = 'Luminaire Installation'
                elif 'щит' in line.lower() or 'що' in line.lower() or 'щао' in line.lower():
                    category = 'Panel'
                elif 'выключатель' in line.lower():
                    category = 'Switch'
                elif 'розетка' in line.lower():
                    category = 'Socket'
                elif 'кабел' in line.lower() or 'провод' in line.lower():
                    category = 'Cable'
                elif 'подключение' in line.lower() and 'жил' in line.lower():
                    category = 'Cable Connection'

                # Extract height category if present
                height_cat = 'N/A'
                if 'до 5 метр' in line.lower():
                    height_cat = 'до 5м'
                elif 'от 5 до 13 метр' in line.lower() or '5 до 13 метр' in line.lower():
                    height_cat = '5-13м'
                elif 'от 13 до 20 метр' in line.lower() or '13 до 20 метр' in line.lower():
                    height_cat = '13-20м'
                elif 'от 20 до 35 метр' in line.lower() or '20 до 35 метр' in line.lower():
                    height_cat = '20-35м'

                equipment.append({
                    'name': line,
                    'category': category,
                    'count': count,
                    'height': height_cat
                })

                i += 3  # Skip past the count
                continue

        i += 1

    return equipment

def main():
    vor_3 = Path('dataset/3-я захватка/ВОР ЭО, Захватка 3_ГПК.txt')

    if not vor_3.exists():
        print(f"Error: File not found: {vor_3}")
        return 1

    equipment = parse_vor_improved(str(vor_3))

    print("\n" + "=" * 80)
    print("VOR DOCUMENT PARSING RESULTS")
    print("=" * 80 + "\n")

    # Group by category
    by_category = {}
    by_height = {}

    for eq in equipment:
        cat = eq['category']
        height = eq['height']

        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(eq)

        if height != 'N/A':
            if height not in by_height:
                by_height[height] = 0
            by_height[height] += eq['count']

    # Print summary by category
    print("EQUIPMENT BY CATEGORY:")
    print("-" * 80)

    for cat in sorted(by_category.keys()):
        items = by_category[cat]
        total_count = sum(item['count'] for item in items)
        print(f"\n{cat}: {total_count} items")
        for item in items:
            # Truncate long names
            name = item['name'][:60] + '...' if len(item['name']) > 60 else item['name']
            print(f"  - {name}: {item['count']} шт [Height: {item['height']}]")

    # Print height category summary
    if by_height:
        print("\n" + "=" * 80)
        print("EQUIPMENT BY HEIGHT CATEGORY:")
        print("-" * 80)
        for height in ['до 5м', '5-13м', '13-20м', '20-35м']:
            if height in by_height:
                print(f"{height}: {by_height[height]} items")

    # Print totals
    total_equipment = sum(eq['count'] for eq in equipment)
    luminaire_count = sum(eq['count'] for eq in equipment if eq['category'] == 'Luminaire')
    panel_count = sum(eq['count'] for eq in equipment if eq['category'] == 'Panel')

    print("\n" + "=" * 80)
    print("TOTALS:")
    print("-" * 80)
    print(f"Total equipment items parsed: {total_equipment}")
    print(f"Total unique entries: {len(equipment)}")
    print(f"Luminaires: {luminaire_count}")
    print(f"Panels: {panel_count}")
    print()

    return 0

if __name__ == '__main__':
    exit(main())
