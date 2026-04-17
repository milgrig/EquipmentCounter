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

elev_re = re.compile(r'на\s+отм[.\s]+([+\-]?\d+[.,]\d+)', re.IGNORECASE)
cable_re = re.compile(r'ППГнг\(А\)-(?:FR)?HF\s+\d+[хx\u00d7]\d+', re.IGNORECASE)

doc = ezdxf.readfile(dxf_path)

# 1) Check all layouts (paper spaces)
print("="*100)
print("CHECKING ALL LAYOUTS")
print("="*100)
for layout_name in doc.layout_names():
    layout = doc.layout(layout_name)
    tables = list(layout.query("ACAD_TABLE"))
    mtext_count = len(list(layout.query("MTEXT")))
    text_count = len(list(layout.query("TEXT")))
    print(f"\nLayout '{layout_name}': {len(tables)} ACAD_TABLEs, {mtext_count} MTEXTs, {text_count} TEXTs")
    
    # Check MTEXT for elevation
    for ent in layout.query("MTEXT"):
        try:
            clean = _clean_mtext(ent.text)
            if elev_re.search(clean):
                x = ent.dxf.insert.x if hasattr(ent.dxf, 'insert') else 0
                y = ent.dxf.insert.y if hasattr(ent.dxf, 'insert') else 0
                print(f"  MTEXT with elev at ({x:.1f}, {y:.1f}): {clean[:150]}")
        except:
            continue
    
    # Check ACAD_TABLE for elevation
    for ti, ent in enumerate(tables):
        try:
            rows = _read_acad_table(ent)
            has_elev = False
            has_cable = False
            for row in rows:
                for cell in row:
                    if cell:
                        clean = _clean_mtext(cell)
                        if elev_re.search(clean):
                            has_elev = True
                        if cable_re.search(clean):
                            has_cable = True
            if has_elev or has_cable:
                print(f"  TABLE {ti}: {len(rows)} rows, elev={has_elev}, cable={has_cable}")
                # Print rows with elevation
                for ri, row in enumerate(rows):
                    for cell in row:
                        if cell:
                            clean = _clean_mtext(cell)
                            if elev_re.search(clean):
                                print(f"    Row {ri}: {clean[:200]}")
        except Exception as e:
            print(f"  TABLE {ti}: ERROR: {e}")

# 2) Check block definitions for ACAD_TABLE
print("\n" + "="*100)
print("CHECKING BLOCK DEFINITIONS FOR ACAD_TABLE")
print("="*100)
for block in doc.blocks:
    tables = [e for e in block if e.dxftype() == "ACAD_TABLE"]
    if tables:
        print(f"\nBlock '{block.name}': {len(tables)} ACAD_TABLE entities")
        for ti, ent in enumerate(tables):
            try:
                rows = _read_acad_table(ent)
                print(f"  TABLE {ti}: {len(rows)} rows")
            except Exception as e:
                print(f"  TABLE {ti}: ERROR: {e}")

# 3) Check the OBJECTS section entities directly
print("\n" + "="*100) 
print("CHECKING OBJECTS SECTION FOR TABLE-LIKE ENTITIES")
print("="*100)
objects = doc.objects
table_like = []
for handle, entity in doc.entitydb.items():
    etype = entity.dxftype()
    if etype in ('ACAD_TABLE', 'TABLE', 'TABLESTYLE', 'TABLECONTENT', 'CELLCONTENT'):
        table_like.append((handle, etype))

print(f"Found {len(table_like)} table-like entities in entitydb")
for handle, etype in table_like[:20]:
    print(f"  Handle={handle}, Type={etype}")

# 4) Check if elevation is in the drawing title blocks / stamp
print("\n" + "="*100)
print("CHECKING ATTRIBS ON INSERT ENTITIES FOR ELEVATION")
print("="*100)
msp = doc.modelspace()
for ent in msp.query("INSERT"):
    try:
        attribs = list(ent.attribs)
        for a in attribs:
            tag = a.dxf.tag if hasattr(a.dxf, 'tag') else ''
            val = a.dxf.text if hasattr(a.dxf, 'text') else ''
            if val and elev_re.search(val):
                print(f"  INSERT block='{ent.dxf.name}' ATTRIB tag='{tag}': {val[:150]}")
            # Also check for elevation-related tags
            tag_lower = tag.lower()
            if any(kw in tag_lower for kw in ['отм', 'elev', 'высот', 'этаж', 'floor', 'level']):
                print(f"  INSERT block='{ent.dxf.name}' ATTRIB tag='{tag}': {val[:150]}")
    except:
        continue
