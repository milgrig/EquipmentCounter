"""Analyze unique triples from TABLECONTENT + MTEXT."""
import sys, re
sys.stdout.reconfigure(encoding='utf-8')
from collections import Counter, defaultdict
from _cable_probe2 import iter_entities, _clean

CABLE_FULL_RE = re.compile(
    r'(ВБШвнг\(А\)-FRLS|ВБШвнг\(А\)-LS|ППГнг\(А\)-FRHF|ППГнг\(А\)-HF)'
    r'\s+(\d+)[хx×](\d+(?:[.,]\d+)?)'
    r'(.*?)'
    r'L\s*[=\-]\s*(\d+)',
    re.IGNORECASE | re.DOTALL,
)

p = ('Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/'
     '01_DWG/003, 004 - Схемы ЩО, ЩАО.dxf')
all_items = []
for ent, pairs in iter_entities(p):
    if ent == 'TABLECONTENT':
        cells = []
        pc = False
        for code, val in pairs:
            if code != '302':
                continue
            v = val.strip()
            if v == 'CONTENT':
                pc = True; continue
            if v == 'GRIDFORMAT' or not v:
                pc = False; continue
            if pc:
                cells.append(_clean(v))
            pc = False
        for c in cells:
            m = CABLE_FULL_RE.search(c)
            if m:
                all_items.append({'src': 'TC', 'brand': m.group(1),
                                  'sec': f"{m.group(2)}х{m.group(3).replace('.', ',')}",
                                  'length': int(m.group(5))})
    elif ent == 'MTEXT':
        tps = []
        for code, val in pairs:
            if code in ('1', '3'):
                tps.append(val)
        text = _clean(''.join(tps))
        m = CABLE_FULL_RE.search(text)
        if m:
            all_items.append({'src': 'MT', 'brand': m.group(1),
                              'sec': f"{m.group(2)}х{m.group(3).replace('.', ',')}",
                              'length': int(m.group(5))})

# Статистика по источникам
src_counts = Counter(i['src'] for i in all_items)
print(f"By source: {dict(src_counts)}")

# Уникальные (brand, sec, length)
keys = Counter()
for i in all_items:
    keys[(i['brand'], i['sec'], i['length'])] += 1
print(f"Total items: {len(all_items)}")
print(f"Unique triples: {len(keys)}")
dist = Counter(keys.values())
print(f"Multiplicity: {sorted(dist.items())}")

# Сколько кабелей если оставить каждую тройку по 1 шт:
unique_total = sum(l for (_, _, l) in keys.keys())
print(f"\nUNIQUE triples → count={len(keys)}, total_length={unique_total}")

# Если оставлять каждую тройку по 2 шт (если встретилась >=2):
approx = []
for (brand, sec, length), mult in keys.items():
    # count = max(1, mult//2) — если встречается 2 раза в MT+TC, это один кабель
    approx.append((brand, sec, length, min(mult, 2)))
acount = sum(k[3] for k in approx)
alen = sum(k[2]*k[3] for k in approx)
print(f"Keep min(mult, 2) → count={acount}, length={alen}")

# Если max 1:
acount1 = len(approx)
alen1 = sum(k[2] for k in approx)
print(f"Keep each triple once → count={acount1}, length={alen1}")

by_bs = defaultdict(lambda: [0, 0])
for (brand, sec, length), mult in keys.items():
    k = f"{brand} {sec}"
    by_bs[k][0] += 1
    by_bs[k][1] += length
print("\nBy brand × section (unique triples once):")
for k, (n, L) in sorted(by_bs.items(), key=lambda x: -x[1][1]):
    print(f"  {k:40s}  {n:3d} каб  {L:5d} м")

# Этaлон
print("\nETALON:")
print("  ВБШвнг(А)-LS 3х1,5:      2317 м")
print("  ВБШвнг(А)-FRLS 3х1,5:    4460 м")
print("  ВБШвнг(А)-LS 3х2,5:      1127 м")
print("  ВБШвнг(А)-FRLS 3х2,5:     180 м")
print("  ППГнг(А)-HF 3х1,5:        842 м")
print("  ППГнг(А)-FRHF 3х1,5:     1300 м")
print("  TOTAL:                  10226 м")
