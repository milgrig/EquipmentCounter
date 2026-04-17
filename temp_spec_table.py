import sys, re
sys.stdout.reconfigure(encoding='utf-8')

import ezdxf
from ezdxf.entities.acad_table import read_acad_table_content as _read_acad_table

dxf_path = r"C:\Cursor\TayfaProject\EquipmentCounter\Data\test2\_converted_dxf\1-ä-24-29-ØÄî.dxf"

def _clean_mtext(raw):
    if not raw:
        return ""
    t = re.sub(r'\[A-Za-z][^;]*;', '', raw)
    t = re.sub(r'[{}]', '', t)
    t = t.replace('\P', '\n').replace('\p', '\n')
    t = re.sub(r'\[A-Za-z]', '', t)
    return t.strip()

doc = ezdxf.readfile(dxf_path)

# Find tables in *Model_Space block
ms_block = doc.blocks.get("*Model_Space")
tables_in_block = [e for e in ms_block if e.dxftype() == "ACAD_TABLE"]

print(f"Tables in *Model_Space block: {len(tables_in_block)}")

for ti, ent in enumerate(tables_in_block):
    try:
        rows = _read_acad_table(ent)
    except Exception as e:
        print(f"\nTable {ti}: ERROR: {e}")
        continue
    
    ncols = max(len(row) for row in rows) if rows else 0
    print(f"\n{'='*120}")
    print(f"TABLE {ti}: {len(rows)} rows x {ncols} cols")
    try:
        pos = ent.dxf.insert
        print(f"Position: ({pos.x:.1f}, {pos.y:.1f})")
    except:
        pass
    print(f"{'='*120}")
    
    for ri, row in enumerate(rows):
        cells = []
        for ci, cell in enumerate(row):
            val = _clean_mtext(cell) if cell else ""
            if val:
                # Truncate long values but show enough to understand
                display = val.replace('\n', ' / ')[:120]
                cells.append(f"[{ci}]={display}")
            else:
                cells.append(f"[{ci}]=")
        print(f"Row {ri:3d}: {' | '.join(cells)}")

# Also check: what's in the raw OBJECTS for the elevation headers?
# Let's search for what kinds of entities reference elevation
print(f"\n{'='*120}")
print("SEARCHING FOR ENTITIES WITH ELEVATION TEXT IN ENTIRE ENTITYDB")
print(f"{'='*120}")
elev_re = re.compile(r'на\s+отм', re.IGNORECASE)
count = 0
for handle, entity in doc.entitydb.items():
    # Check various text fields
    try:
        for attr in ['text', 'plain_text']:
            if hasattr(entity, attr):
                val = getattr(entity, attr)
                if callable(val):
                    val = val()
                if isinstance(val, str) and elev_re.search(val):
                    count += 1
                    if count <= 20:
                        print(f"  Handle={handle} Type={entity.dxftype()}: {val[:150]}")
    except:
        pass
    # Check dxf attributes
    try:
        for attr in ['text', 'default_value', 'tag']:
            if hasattr(entity.dxf, attr):
                val = getattr(entity.dxf, attr)
                if isinstance(val, str) and elev_re.search(val):
                    count += 1
                    if count <= 20:
                        print(f"  Handle={handle} Type={entity.dxftype()} dxf.{attr}: {val[:150]}")
    except:
        pass

print(f"Total entities with elevation text: {count}")
