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

cable_re = re.compile(r'(ППГнг\(А\)-FRHF\s+\d+[хx\u00d7]\d+[,.]?\d*)', re.IGNORECASE)
length_re = re.compile(r'L\s*=\s*(\d+)\s*м')

doc = ezdxf.readfile(dxf_path)
msp = doc.modelspace()

# Check all sources for FRHF cables

# Source 1 & 2: MTEXT and TEXT in modelspace
print("=== MTEXT/TEXT in modelspace ===")
for ent in list(msp.query("MTEXT")) + list(msp.query("TEXT")):
    try:
        if ent.dxftype() == "MTEXT":
            clean = _clean_mtext(ent.text)
        else:
            clean = ent.dxf.get("text", "")
        if clean and cable_re.search(clean):
            cm = cable_re.search(clean)
            lm = length_re.search(clean)
            length = lm.group(1) if lm else "?"
            print(f"  {ent.dxftype()}: {cm.group(1)} L={length}м | Full: {clean.replace(chr(10), ' / ')[:200]}")
    except:
        continue

# Source 3: ACAD_TABLE
print("\n=== ACAD_TABLE (all tables in modelspace) ===")
for ti, ent in enumerate(msp.query("ACAD_TABLE")):
    try:
        rows = _read_acad_table(ent)
        for ri, row in enumerate(rows):
            for ci, cell in enumerate(row):
                if cell:
                    clean = _clean_mtext(cell)
                    if cable_re.search(clean):
                        # Split by newlines and process each line
                        for line in clean.split('\n'):
                            cm = cable_re.search(line)
                            lm = length_re.search(line)
                            if cm:
                                length = lm.group(1) if lm else "?"
                                print(f"  Table {ti} Row {ri} Col {ci}: {cm.group(1)} L={length}м | {line.strip()[:150]}")
    except Exception as e:
        print(f"  Table {ti}: ERROR: {e}")

# Source 4: Blocks
print("\n=== Blocks referenced by INSERT ===")
inserted_names = set()
for ent in msp.query("INSERT"):
    try:
        inserted_names.add(ent.dxf.name)
    except:
        continue

skip = {"*Model_Space", "*Paper_Space", "*Paper_Space0"}
for block in doc.blocks:
    if block.name in skip or block.name not in inserted_names:
        continue
    for ent in block:
        try:
            if ent.dxftype() == "MTEXT":
                clean = _clean_mtext(ent.text)
            elif ent.dxftype() == "TEXT":
                clean = ent.dxf.get("text", "")
            else:
                continue
            if clean and cable_re.search(clean):
                for line in clean.split('\n'):
                    cm = cable_re.search(line)
                    lm = length_re.search(line)
                    if cm:
                        length = lm.group(1) if lm else "?"
                        print(f"  Block '{block.name}': {cm.group(1)} L={length}м | {line.strip()[:150]}")
        except:
            continue

# Also check spec table (Table 2 with 75 rows) for the reference values
print("\n=== SPEC TABLE (Table 2) - Cable section ===")
ms_block = doc.blocks.get("*Model_Space")
tables = [e for e in ms_block if e.dxftype() == "ACAD_TABLE"]
for ti, ent in enumerate(tables):
    try:
        rows = _read_acad_table(ent)
        if len(rows) == 75:
            # This is the spec table, show cable rows
            for ri, row in enumerate(rows):
                if ri < 28 or ri > 42:
                    continue
                cells_clean = [_clean_mtext(c)[:60] if c else "" for c in row]
                print(f"  Row {ri}: {' | '.join(c for c in cells_clean if c)}")
    except:
        continue
