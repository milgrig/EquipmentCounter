#!/usr/bin/env python3
"""
Read-only: где физически лежат FRLS/FRHF 3х1,5 в DXF?
Обходим ВСЕ секции и перечисляем entity-types.
"""
from __future__ import annotations

import io
import re
import sys
from pathlib import Path
from collections import Counter

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

DXF_PATH = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/"
    "003, 004 - Схемы ЩО, ЩАО.dxf"
)

FRLS_RE = re.compile(
    r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*1[,.]5",
    re.IGNORECASE,
)
FRHF_RE = re.compile(
    r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*1[,.]5",
    re.IGNORECASE,
)


def main() -> None:
    # Читаем DXF как сырой текст и ищем: в какой секции встречаются марки?
    data = DXF_PATH.read_text(encoding="utf-8", errors="ignore")
    print(f"DXF размер: {len(data):,} байт")

    # Найти все позиции вхождений FRLS
    frls_positions = [m.start() for m in FRLS_RE.finditer(data)]
    frhf_positions = [m.start() for m in FRHF_RE.finditer(data)]
    print(f"FRLS 3х1,5 вхождений: {len(frls_positions)}")
    print(f"FRHF 3х1,5 вхождений: {len(frhf_positions)}")

    # Найти границы секций
    section_starts = []
    for m in re.finditer(r"\n\s*2\s*\n(HEADER|CLASSES|TABLES|BLOCKS|ENTITIES|OBJECTS)\s*\n", data):
        section_starts.append((m.start(), m.group(1)))
    section_starts.sort()

    def find_section(pos: int) -> str:
        current = "?"
        for start, name in section_starts:
            if start <= pos:
                current = name
            else:
                break
        return current

    # Найти entity type каждого вхождения
    def find_entity_type(pos: int) -> str:
        # Идём назад, ищем ближайший "  0\n<TYPE>\n"
        search_start = max(0, pos - 5000)
        chunk = data[search_start:pos]
        # Находим все "  0\n<ANYTHING>\n" и берём последний
        matches = list(re.finditer(r"\n\s*0\s*\n([A-Z_][A-Z_0-9]*)\s*\n", chunk))
        if matches:
            return matches[-1].group(1)
        return "?"

    print("\n--- Распределение FRLS по секциям и entity types ---")
    counter = Counter()
    for pos in frls_positions:
        section = find_section(pos)
        entity = find_entity_type(pos)
        counter[(section, entity)] += 1
    for (section, entity), count in counter.most_common():
        print(f"  {section:10} / {entity:20}  {count:5}")

    print("\n--- Распределение FRHF по секциям и entity types ---")
    counter = Counter()
    for pos in frhf_positions:
        section = find_section(pos)
        entity = find_entity_type(pos)
        counter[(section, entity)] += 1
    for (section, entity), count in counter.most_common():
        print(f"  {section:10} / {entity:20}  {count:5}")

    # Показать фрагмент вокруг первого FRLS
    if frls_positions:
        pos = frls_positions[0]
        start = max(0, pos - 600)
        end = min(len(data), pos + 400)
        print(f"\n--- Фрагмент DXF вокруг первого FRLS (@ {pos}) ---")
        fragment = data[start:end]
        # Помечаем найденное место
        rel = pos - start
        marked = fragment[:rel] + "<<<HIT>>>" + fragment[rel:]
        print(marked[:1500])


if __name__ == "__main__":
    main()
