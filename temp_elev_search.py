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
msp = doc.modelspace()

# 1) Search for elevation-related keywords in ALL MTEXT in modelspace
print("=== ALL MTEXT with elevation/floor/level keywords ===")
keywords = re.compile(r'(отм|elev|высот|этаж|floor|level|\+\d+[.,]\d{3}|план\s+на|лист\s+\d)', re.IGNORECASE)
for ent in msp.query("MTEXT"):
    try:
        clean = _clean_mtext(ent.text)
        if clean and keywords.search(clean):
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            display = clean.replace('\n', ' / ')[:200]
            print(f"  ({x:.1f}, {y:.1f}): {display}")
    except:
        continue

# 2) Search TEXT entities for same
print("\n=== ALL TEXT with elevation/floor/level keywords ===")
for ent in msp.query("TEXT"):
    try:
        text = ent.dxf.get("text", "")
        if text and keywords.search(text):
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            print(f"  ({x:.1f}, {y:.1f}): {text[:200]}")
    except:
        continue

# 3) Check ALL INSERT attribs for anything elevation/sheet related
print("\n=== ALL INSERT ATTRIBS (showing all non-empty tags) ===")
seen_blocks = set()
for ent in msp.query("INSERT"):
    try:
        block_name = ent.dxf.name
        attribs = list(ent.attribs)
        if attribs and block_name not in seen_blocks:
            seen_blocks.add(block_name)
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            print(f"\n  Block '{block_name}' at ({x:.1f}, {y:.1f}):")
            for a in attribs:
                tag = a.dxf.tag if hasattr(a.dxf, 'tag') else ''
                val = a.dxf.text if hasattr(a.dxf, 'text') else ''
                if val.strip():
                    print(f"    TAG='{tag}': '{val}'")
    except:
        continue

# 4) Check layout names and properties
print("\n=== LAYOUT NAMES AND PROPERTIES ===")
for layout_name in doc.layout_names():
    layout = doc.layout(layout_name)
    print(f"  Layout: '{layout_name}'")
    try:
        # Check for description, plot settings etc
        dxf_layout = layout.dxf_layout
        if hasattr(dxf_layout.dxf, 'name'):
            print(f"    dxf.name: {dxf_layout.dxf.name}")
    except:
        pass

# 5) Look for sheet title in the title block area
# Title blocks typically have drawing name, sheet number, etc.
print("\n=== TITLE BLOCK SEARCH: MTEXT with 'план' or 'схема' or 'лист' ===")
title_re = re.compile(r'(план|схема|лист|чертеж|sheet)', re.IGNORECASE)
for ent in msp.query("MTEXT"):
    try:
        clean = _clean_mtext(ent.text)
        if clean and title_re.search(clean):
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            display = clean.replace('\n', ' / ')[:200]
            print(f"  ({x:.1f}, {y:.1f}): {display}")
    except:
        continue

for ent in msp.query("TEXT"):
    try:
        text = ent.dxf.get("text", "")
        if text and title_re.search(text):
            x = ent.dxf.insert.x
            y = ent.dxf.insert.y
            print(f"  ({x:.1f}, {y:.1f}): {text[:200]}")
    except:
        continue
