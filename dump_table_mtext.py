"""
Dump ALL MTEXT content from *T blocks (ACAD_TABLE data) in schema DXF files.
Lists block name, x, y coordinates, and plain_text content.
"""
import sys
import os
import ezdxf

sys.stdout.reconfigure(encoding='utf-8')

DXF_DIR = r'C:\Cursor\TayfaProject\EquipmentCounter\Data\test\_converted_dxf'


def dump_table_blocks(filepath):
    """Read a DXF file and dump all MTEXT entities from *T blocks."""
    print(f"\n{'='*120}")
    print(f"FILE: {os.path.basename(filepath)}")
    print(f"{'='*120}")

    doc = ezdxf.readfile(filepath)

    # First, list ALL block names to understand structure
    all_blocks = [b.name for b in doc.blocks]
    t_blocks = [name for name in all_blocks if name.startswith('*T')]
    other_blocks = [name for name in all_blocks if not name.startswith('*') and not name.startswith('_')]

    print(f"\nTotal blocks: {len(all_blocks)}")
    print(f"*T blocks (tables): {len(t_blocks)}")
    print(f"*T block names: {t_blocks}")
    print(f"\nOther named blocks (non-system): {other_blocks[:50]}")

    # Dump MTEXT from each *T block
    for block_name in sorted(t_blocks):
        block = doc.blocks.get(block_name)
        if block is None:
            continue

        print(f"\n{'─'*120}")
        print(f"BLOCK: {block_name}")
        print(f"{'─'*120}")

        # Collect all entities in this block
        entities = list(block)
        entity_types = {}
        for e in entities:
            entity_types[e.dxftype()] = entity_types.get(e.dxftype(), 0) + 1
        print(f"  Entity types: {entity_types}")

        # Collect MTEXT entries with coordinates
        mtext_entries = []
        for e in entities:
            if e.dxftype() == 'MTEXT':
                try:
                    x = e.dxf.insert.x if hasattr(e.dxf, 'insert') else 0
                    y = e.dxf.insert.y if hasattr(e.dxf, 'insert') else 0
                    raw = e.text
                    plain = ezdxf.tools.text.plain_mtext(raw, split=False) if raw else ""
                    mtext_entries.append((x, y, plain, raw))
                except Exception as ex:
                    print(f"  ERROR reading MTEXT: {ex}")

        # Sort by Y descending (top to bottom), then X ascending (left to right)
        mtext_entries.sort(key=lambda m: (-m[1], m[0]))

        print(f"  MTEXT count: {len(mtext_entries)}")

        # Try to detect table structure by grouping by Y coordinate
        if mtext_entries:
            # Group by approximate Y coordinate (within tolerance of 1 unit)
            rows = []
            current_row = [mtext_entries[0]]
            current_y = mtext_entries[0][1]

            for entry in mtext_entries[1:]:
                if abs(entry[1] - current_y) < 1.0:
                    current_row.append(entry)
                else:
                    rows.append(current_row)
                    current_row = [entry]
                    current_y = entry[1]
            rows.append(current_row)

            print(f"  Detected {len(rows)} rows\n")

            for ri, row in enumerate(rows):
                # Sort cells in row by X coordinate
                row.sort(key=lambda m: m[0])
                print(f"  ROW {ri:3d} (y={row[0][1]:10.2f}):")
                for ci, (x, y, plain, raw) in enumerate(row):
                    display = plain.replace('\n', ' | ')
                    if display.strip():
                        print(f"    COL {ci:2d} (x={x:10.2f}): [{display}]")
                    else:
                        print(f"    COL {ci:2d} (x={x:10.2f}): [<empty>]")

    # Also check model space for any ACAD_TABLE entities directly
    msp = doc.modelspace()
    print(f"\n{'─'*120}")
    print(f"MODEL SPACE - ACAD_TABLE and INSERT entities referencing *T blocks")
    print(f"{'─'*120}")

    for e in msp:
        if e.dxftype() == 'ACAD_TABLE':
            print(f"  ACAD_TABLE at ({e.dxf.insert.x:.2f}, {e.dxf.insert.y:.2f})")
        elif e.dxftype() == 'INSERT' and e.dxf.name.startswith('*T'):
            print(f"  INSERT of {e.dxf.name} at ({e.dxf.insert.x:.2f}, {e.dxf.insert.y:.2f})")


def main():
    # Find schema DXF files (003, 004)
    for fname in os.listdir(DXF_DIR):
        if fname.endswith('.dxf') and ('003' in fname or '004' in fname):
            filepath = os.path.join(DXF_DIR, fname)
            dump_table_blocks(filepath)


if __name__ == '__main__':
    main()
