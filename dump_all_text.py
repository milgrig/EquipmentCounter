"""
Dump ALL text entities (MTEXT and TEXT) from model space and non-system blocks
in the schema DXF file, searching for junction boxes, sleeves, shrink tubes, etc.
"""
import sys
import os
import ezdxf

sys.stdout.reconfigure(encoding='utf-8')

DXF_DIR = r'C:\Cursor\TayfaProject\EquipmentCounter\Data\test\_converted_dxf'


def dump_all_text(filepath):
    print(f"\n{'='*120}")
    print(f"FILE: {os.path.basename(filepath)}")
    print(f"{'='*120}")

    doc = ezdxf.readfile(filepath)
    msp = doc.modelspace()

    # 1. Dump ALL MTEXT/TEXT from model space
    print(f"\n--- MODEL SPACE TEXT/MTEXT (all) ---")
    msp_texts = []
    for e in msp:
        if e.dxftype() == 'MTEXT':
            raw = e.text
            plain = ezdxf.tools.text.plain_mtext(raw, split=False) if raw else ""
            x = e.dxf.insert.x if hasattr(e.dxf, 'insert') else 0
            y = e.dxf.insert.y if hasattr(e.dxf, 'insert') else 0
            if plain.strip():
                msp_texts.append(('MTEXT', x, y, plain))
        elif e.dxftype() == 'TEXT':
            text = e.dxf.text if hasattr(e.dxf, 'text') else ""
            x = e.dxf.insert.x if hasattr(e.dxf, 'insert') else 0
            y = e.dxf.insert.y if hasattr(e.dxf, 'insert') else 0
            if text.strip():
                msp_texts.append(('TEXT', x, y, text))

    msp_texts.sort(key=lambda m: (-m[2], m[1]))
    print(f"  Total text entities in model space: {len(msp_texts)}")
    for typ, x, y, text in msp_texts:
        display = text.replace('\n', ' | ')
        print(f"  {typ:5s} ({x:10.2f}, {y:10.2f}): {display}")

    # 2. Dump ATTRIB from INSERT entities in model space (block references with attributes)
    print(f"\n--- MODEL SPACE INSERT ATTRIBS ---")
    for e in msp:
        if e.dxftype() == 'INSERT' and hasattr(e, 'attribs'):
            attribs = list(e.attribs)
            if attribs:
                x = e.dxf.insert.x
                y = e.dxf.insert.y
                block_name = e.dxf.name
                for a in attribs:
                    tag = a.dxf.tag if hasattr(a.dxf, 'tag') else ''
                    val = a.dxf.text if hasattr(a.dxf, 'text') else ''
                    if val.strip():
                        print(f"  BLOCK={block_name} at ({x:.2f},{y:.2f}) TAG={tag}: {val}")


def main():
    for fname in os.listdir(DXF_DIR):
        if fname.endswith('.dxf') and ('003' in fname or '004' in fname):
            filepath = os.path.join(DXF_DIR, fname)
            dump_all_text(filepath)


if __name__ == '__main__':
    main()
