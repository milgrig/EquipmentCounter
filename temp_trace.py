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

cable_re = re.compile(r'(ППГнг\(А\)-(?:FR)?HF\s+\d+[хx\u00d7]\d+[,.]?\d*)', re.IGNORECASE)
length_re = re.compile(r'L\s*=\s*(\d+)\s*м')

doc = ezdxf.readfile(dxf_path)
msp = doc.modelspace()

# Source 1: Modelspace MTEXT
print("=== SOURCE 1: Modelspace MTEXT ===")
count = 0
for ent in msp.query("MTEXT"):
    try:
        clean = _clean_mtext(ent.text)
        if clean and cable_re.search(clean):
            lm = length_re.search(clean)
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            count += 1
            if count <= 30:
                print(f"  ({x:.1f}, {y:.1f}): {clean[:200]}")
    except:
        continue
print(f"  Total: {count} MTEXT with cables")

# Source 2: Modelspace TEXT
print("\n=== SOURCE 2: Modelspace TEXT ===")
count = 0
for ent in msp.query("TEXT"):
    try:
        text = ent.dxf.get("text", "")
        if text and cable_re.search(text):
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            count += 1
            print(f"  ({x:.1f}, {y:.1f}): {text[:200]}")
    except:
        continue
print(f"  Total: {count} TEXT with cables")

# Source 3: ACAD_TABLE
print("\n=== SOURCE 3: ACAD_TABLE ===")
total_cable_rows = 0
for ti, ent in enumerate(msp.query("ACAD_TABLE")):
    try:
        rows = _read_acad_table(ent)
        for ri, row in enumerate(rows):
            for cell in row:
                if cell:
                    clean = _clean_mtext(cell)
                    cm = cable_re.search(clean)
                    if cm:
                        lm = length_re.search(clean)
                        total_cable_rows += 1
                        print(f"  Table {ti} Row {ri}: {clean[:200]}")
                        break
    except:
        continue
print(f"  Total: {total_cable_rows} rows with cables")

# Source 4: Blocks
print("\n=== SOURCE 4: Blocks referenced by INSERT ===")
inserted_names = set()
for ent in msp.query("INSERT"):
    try:
        inserted_names.add(ent.dxf.name)
    except:
        continue

skip = {"*Model_Space", "*Paper_Space", "*Paper_Space0"}
block_cable_count = 0
for block in doc.blocks:
    if block.name in skip or block.name not in inserted_names:
        continue
    for ent in block:
        try:
            if ent.dxftype() == "MTEXT":
                clean = _clean_mtext(ent.text)
                if clean and cable_re.search(clean):
                    block_cable_count += 1
                    if block_cable_count <= 30:
                        print(f"  Block '{block.name}': {clean[:200]}")
            elif ent.dxftype() == "TEXT":
                text = ent.dxf.get("text", "")
                if text and cable_re.search(text):
                    block_cable_count += 1
                    print(f"  Block '{block.name}': {text[:200]}")
        except:
            continue
print(f"  Total: {block_cable_count} entities with cables in referenced blocks")

# Now check ACAD_TABLE in blocks with elevation tables
print("\n=== ACAD_TABLE in block definitions (including A$C* blocks) ===")
for block in doc.blocks:
    for ent in block:
        if ent.dxftype() == "ACAD_TABLE":
            try:
                rows = _read_acad_table(ent)
                has_cable = False
                has_elev = False
                elev_re2 = re.compile(r'(?:Освещение|[Аа]варийн|[Сс]иловое|[Рр]озетк)', re.IGNORECASE)
                for row in rows:
                    for cell in row:
                        if cell:
                            clean = _clean_mtext(cell)
                            if cable_re.search(clean):
                                has_cable = True
                            if elev_re2.search(clean):
                                has_elev = True
                if has_cable or has_elev:
                    print(f"\n  Block '{block.name}': TABLE with {len(rows)} rows, cable={has_cable}, elev={has_elev}")
                    for ri, row in enumerate(rows):
                        cells = []
                        for cell in row:
                            if cell:
                                cells.append(_clean_mtext(cell)[:80])
                        if cells:
                            print(f"    Row {ri}: {' | '.join(cells)}")
            except Exception as e:
                print(f"  Block '{block.name}': TABLE ERROR: {e}")
