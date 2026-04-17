import sys, re
sys.stdout.reconfigure(encoding='utf-8')

import ezdxf
from ezdxf.entities.acad_table import read_acad_table_content as _read_acad_table

dxf_path = r"C:\Cursor\TayfaProject\EquipmentCounter\Data\test2\_converted_dxf\1-ä-24-29-ØÄî.dxf"

# Use the same _clean_mtext as in equipment_counter
def _clean_mtext(raw: str) -> str:
    if not raw:
        return ""
    t = re.sub(r'\[A-Za-z][^;]*;', '', raw)
    t = re.sub(r'[{}]', '', t)
    t = t.replace('\P', '\n').replace('\p', '\n')
    t = re.sub(r'\[A-Za-z]', '', t)
    t = re.sub(r'%%[cCuUdDpP]', '', t)
    return t.strip()

cable_re = re.compile(r'ППГнг\(А\)-(?:FR)?HF\s+\d+[хx×]\d+[,.]?\d*', re.IGNORECASE)
length_re = re.compile(r'L\s*=\s*(\d+)\s*м')
elev_re = re.compile(r'(?:Освещение|[Аа]варийн|[Ээ]вакуацион|[Сс]иловое|[Рр]озетк)\S*\s+.*?на\s+отм[.\s]+([+\-]?\d+[.,]\d+)', re.IGNORECASE)
elev_broad_re = re.compile(r'на\s+отм[.\s]+([+\-]?\d+[.,]\d+)', re.IGNORECASE)

doc = ezdxf.readfile(dxf_path)
msp = doc.modelspace()

print("="*100)
print("DETAILED ACAD_TABLE ANALYSIS")
print("="*100)

tables = list(msp.query("ACAD_TABLE"))
print(f"\nFound {len(tables)} ACAD_TABLE entities\n")

for ti, ent in enumerate(tables):
    try:
        rows = _read_acad_table(ent)
    except Exception as e:
        print(f"Table {ti}: ERROR: {e}")
        continue
    
    # Check if this table has any cables
    has_cables = False
    has_elev = False
    for row in rows:
        for cell in row:
            if cell:
                clean = _clean_mtext(cell)
                if cable_re.search(clean):
                    has_cables = True
                if elev_broad_re.search(clean):
                    has_elev = True
    
    if not has_cables and not has_elev:
        print(f"Table {ti}: {len(rows)} rows - no cables or elevations, skipping")
        continue
    
    try:
        pos = ent.dxf.insert
        print(f"\nTable {ti}: {len(rows)} rows, position=({pos.x:.1f}, {pos.y:.1f}), has_cables={has_cables}, has_elev={has_elev}")
    except:
        print(f"\nTable {ti}: {len(rows)} rows, has_cables={has_cables}, has_elev={has_elev}")
    
    # Print ALL rows of this table
    for ri, row in enumerate(rows):
        cells_clean = []
        for ci, cell in enumerate(row):
            if cell and cell.strip():
                cells_clean.append(f"[{ci}]={_clean_mtext(cell)}")
        if cells_clean:
            row_text = " | ".join(cells_clean)
            # Mark special rows
            marker = ""
            if any(cable_re.search(c) for c in [_clean_mtext(cell) for cell in row if cell]):
                marker = " <<< CABLE"
            if any(elev_broad_re.search(c) for c in [_clean_mtext(cell) for cell in row if cell]):
                marker = " <<< ELEVATION"
            print(f"  Row {ri:3d}: {row_text[:250]}{marker}")

# Also check: are there MTEXT entities near the tables that contain elevation?
print("\n" + "="*100)
print("MODELSPACE MTEXT WITH ELEVATION REFERENCES")
print("="*100)

for ent in msp.query("MTEXT"):
    try:
        clean = _clean_mtext(ent.text)
        if elev_broad_re.search(clean):
            x = ent.dxf.insert.x if hasattr(ent.dxf, 'insert') else 0
            y = ent.dxf.insert.y if hasattr(ent.dxf, 'insert') else 0
            print(f"  MTEXT at ({x:.1f}, {y:.1f}): {clean[:200]}")
    except:
        continue

# Check TEXT entities too
for ent in msp.query("TEXT"):
    try:
        text = ent.dxf.get("text", "")
        if text and elev_broad_re.search(text):
            x = ent.dxf.insert.x if hasattr(ent.dxf, 'insert') else 0
            y = ent.dxf.insert.y if hasattr(ent.dxf, 'insert') else 0
            print(f"  TEXT at ({x:.1f}, {y:.1f}): {text[:200]}")
    except:
        continue
