import sys, re
sys.stdout.reconfigure(encoding='utf-8')

import ezdxf
from ezdxf.entities.acad_table import read_acad_table_content as _read_acad_table

import pathlib
dxf_dir = pathlib.Path(r'C:/Cursor/TayfaProject/EquipmentCounter/Data/test2/_converted_dxf')
candidates = list(dxf_dir.glob('1-*-24-29-*'))
dxf_path = str(candidates[0]) if candidates else None
print('DXF path:', dxf_path)

doc = ezdxf.readfile(dxf_path)
msp = doc.modelspace()

elev_re = re.compile(r'(?:\u041e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u0435|[\u0410\u0430]\u0432\u0430\u0440\u0438\u0439\u043d\u043e\u0435\s+\u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u0435|[\u042d\u044d]\u0432\u0430\u043a\u0443\u0430\u0446\u0438\u043e\u043d\u043d\u043e\u0435\s+\u043e\u0441\u0432\u0435\u0449\u0435\u043d\u0438\u0435|\u0421\u0438\u043b\u043e\u0432\u043e\u0435|\u0420\u043e\u0437\u0435\u0442\u043a\u0438)\s+\u043d\u0430\s+\u043e\u0442\u043c[.\s]+([+\-]?\d+[.,]\d+)', re.IGNORECASE)
cable_re = re.compile(r'\u041f\u041f\u0413\u043d\u0433\(\u0410\)-(?:FR)?HF\s+\d+[\u0445x\u00d7]\d+', re.IGNORECASE)
length_re = re.compile(r'L\s*=\s*(\d+)\s*\u043c')

def clean(text):
    if not text:
        return ''
    t = re.sub(r'\[fFpPsSlL][^;]*;', '', text)
    t = re.sub(r'[{}]', '', t)
    t = t.replace(chr(92), '')
    return t.strip()

tables = list(msp.query("ACAD_TABLE"))
print("Found %d ACAD_TABLE entities in modelspace" % len(tables))

for ti, ent in enumerate(tables):
    try:
        rows = _read_acad_table(ent)
    except Exception as e:
        print("Table %d: ERROR reading: %s" % (ti, e))
        continue

    print("")
    print("Table %d: %d rows" % (ti, len(rows)))

    try:
        pos = ent.dxf.insert
        print("  Position: (%.2f, %.2f)" % (pos.x, pos.y))
    except:
        pass

    current_elev = None
    cable_count_under_elev = {}

    for ri, row in enumerate(rows):
        row_texts = []
        for ci, cell in enumerate(row):
            if cell and cell.strip():
                row_texts.append(clean(cell))

        full_row = " | ".join(row_texts)

        for cell_text in row_texts:
            em = elev_re.search(cell_text)
            if em:
                current_elev = em.group(1)
                if ri < 20 or cable_re.search(full_row):
                    print("  Row %4d: ELEV=%s | %s" % (ri, current_elev, full_row[:150]))
                break

        for cell_text in row_texts:
            cm = cable_re.search(cell_text)
            if cm:
                lm = length_re.search(cell_text)
                length = lm.group(1) if lm else "?"
                key = current_elev or "NONE"
                cable_count_under_elev[key] = cable_count_under_elev.get(key, 0) + 1
                count = cable_count_under_elev[key]
                if count <= 3:
                    print("  Row %4d: CABLE (elev=%s) | %s" % (ri, current_elev, full_row[:200]))
                elif count == 4:
                    print("  ... more cables under elev=%s" % current_elev)
                break

    print("")
    print("  === Cable count by elevation ===")
    total = 0
    for elev, count in sorted(cable_count_under_elev.items()):
        print("    Elevation %s: %d cables" % (elev, count))
        total += count
    print("    TOTAL: %d cables" % total)
