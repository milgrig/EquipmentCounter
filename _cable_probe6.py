"""Extract cables from TABLE #1, #2 — tables that have full info per cell."""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
from _cable_probe2 import iter_entities, _clean
from collections import defaultdict

p = ('Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/'
     '01_DWG/003, 004 - Схемы ЩО, ЩАО.dxf')

# Regex to extract (brand, section, length, laying) from one-cell text
CABLE_FULL_RE = re.compile(
    r"(ВБШвнг\(А\)-FRLS|ВБШвнг\(А\)-LS|ППГнг\(А\)-FRHF|ППГнг\(А\)-HF)"
    r"\s+(\d+)[хx×](\d+(?:[.,]\d+)?)"
    r"(.*?)"
    r"L\s*[=\-]\s*(\d+)",
    re.IGNORECASE | re.DOTALL,
)
LAYING_RE = re.compile(r"в\s*(кабель[- ]?канале|гофре|лотке|земле|трубе|траншее)",
                        re.IGNORECASE)


def classify_laying(text):
    ms = [m.group(1).lower() for m in LAYING_RE.finditer(text)]
    has_tray = any("лотк" in x for x in ms)
    has_gofra = any("гофр" in x for x in ms)
    if has_tray and has_gofra:
        return "в лотке+гофре"
    if has_tray:
        return "в лотке"
    if has_gofra:
        return "в гофре"
    if any("земле" in x or "траншее" in x or "трубе" in x for x in ms):
        return "в земле"
    if any("кабель" in x or "канал" in x for x in ms):
        return "в кабель-канале"
    return ""


all_items = []
tn = 0
for ent, pairs in iter_entities(p):
    if ent != 'TABLECONTENT':
        continue
    if not any('ВБШвнг' in v or 'ППГнг' in v for _, v in pairs):
        continue
    tn += 1
    cells = []
    prev_content = False
    for code, val in pairs:
        if code != '302':
            continue
        v = val.strip()
        if v == 'CONTENT':
            prev_content = True; continue
        if v == 'GRIDFORMAT' or not v:
            prev_content = False; continue
        if prev_content:
            cells.append(_clean(v))
        prev_content = False

    for c in cells:
        m = CABLE_FULL_RE.search(c)
        if not m:
            continue
        brand = m.group(1)
        sec = f"{m.group(2)}х{m.group(3).replace('.', ',')}"
        middle = m.group(4)
        length = int(m.group(5))
        laying = classify_laying(middle)
        all_items.append({
            "table": tn,
            "brand": brand,
            "section": sec,
            "length": length,
            "laying": laying,
            "text": c[:90],
        })

print(f"Cables from full-info TABLECONTENT cells: {len(all_items)}")
print(f"Total length: {sum(i['length'] for i in all_items)}")

by_bs = defaultdict(lambda: [0, 0])
for c in all_items:
    k = f"{c['brand']} {c['section']}"
    by_bs[k][0] += 1
    by_bs[k][1] += c['length']
print("\nПо маркам × сечениям:")
for k, (n, L) in sorted(by_bs.items(), key=lambda x: -x[1][1]):
    print(f"  {k:35s}  {n:4d} каб  {L:6d} м")

by_lay = defaultdict(lambda: [0, 0])
for c in all_items:
    by_lay[c['laying'] or '(none)'][0] += 1
    by_lay[c['laying'] or '(none)'][1] += c['length']
print("\nПо способу прокладки:")
for k, (n, L) in sorted(by_lay.items(), key=lambda x: -x[1][1]):
    print(f"  {k:25s}  {n:4d} каб  {L:6d} м")

# Ethalon
print("\n--- Etalon: 116 каб, 9864 м ---")
print("ВБШвнг(А)-LS 3х1,5: 2317")
print("ВБШвнг(А)-LS 3х2,5: 1127")
print("ВБШвнг(А)-FRLS 3х1,5: 4460")
print("ВБШвнг(А)-FRLS 3х2,5: 180")
print("ППГнг(A)-HF 3х1,5: 842")
print("ППГнг(A)-FRHF 3х1,5: 1300")
