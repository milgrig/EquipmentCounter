#!/usr/bin/env python3
"""Показывает сырой DXF фрагмент вокруг вхождений FRLS/FRHF, чтобы увидеть
точную entity-структуру (полный фрагмент 2000 символов до и после).
"""
from __future__ import annotations
import io, re, sys
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

DXF = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/"
    "003, 004 - Схемы ЩО, ЩАО.dxf"
)
FRLS_RE = re.compile(
    r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*1[,.]5",
    re.IGNORECASE,
)


def main() -> None:
    data = DXF.read_text(encoding="utf-8", errors="ignore")
    frls_positions = [m.start() for m in FRLS_RE.finditer(data)]
    print(f"FRLS вхождений: {len(frls_positions)}")

    # Найдём одно вхождение из BLOCKS (первое; мы знаем, что 17 FRLS там)
    # Ищем границы секций
    blocks_start = data.find("\nBLOCKS\n")
    blocks_end = data.find("\nENDSEC\n", blocks_start) if blocks_start != -1 else -1
    entities_start = data.find("\nENTITIES\n")
    entities_end = data.find("\nENDSEC\n", entities_start) if entities_start != -1 else -1
    objects_start = data.find("\nOBJECTS\n")
    objects_end = data.find("\nENDSEC\n", objects_start) if objects_start != -1 else -1
    print(f"BLOCKS: {blocks_start}–{blocks_end}")
    print(f"ENTITIES: {entities_start}–{entities_end}")
    print(f"OBJECTS: {objects_start}–{objects_end}")

    # Найти одно FRLS в BLOCKS
    blocks_frls = [p for p in frls_positions if blocks_start <= p <= blocks_end]
    print(f"FRLS в BLOCKS: {len(blocks_frls)}")
    entities_frls = [p for p in frls_positions if entities_start <= p <= entities_end]
    print(f"FRLS в ENTITIES: {len(entities_frls)}")
    objects_frls = [p for p in frls_positions if objects_start <= p <= objects_end]
    print(f"FRLS в OBJECTS: {len(objects_frls)}")

    def extract(pos, radius_back=1200, radius_fwd=400):
        a = max(0, pos - radius_back)
        b = min(len(data), pos + radius_fwd)
        snippet = data[a:b]
        rel = pos - a
        return snippet[:rel] + ">>>HIT<<<" + snippet[rel:]

    if blocks_frls:
        print("\n======= FRLS в BLOCKS =======")
        print(extract(blocks_frls[0]))
    if entities_frls:
        print("\n\n======= FRLS в ENTITIES =======")
        print(extract(entities_frls[0]))
    if objects_frls:
        print("\n\n======= FRLS в OBJECTS =======")
        print(extract(objects_frls[0]))


if __name__ == "__main__":
    main()
