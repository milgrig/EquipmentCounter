#!/usr/bin/env python3
"""
Read-only: пробуем связать марку кабеля с L=NNN по геометрической близости
в BLOCKS секции, без чтения OBJECTS.

Гипотеза: в таблице кабельного журнала марка лежит в одной ячейке (MTEXT),
а длина (L=NNN) — в соседней (MTEXT) с похожим Y и x = x_brand + dx_col.
"""
from __future__ import annotations

import io
import re
import sys
from pathlib import Path
from collections import Counter, defaultdict

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import ezdxf  # type: ignore

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
L_RE = re.compile(r"L\s*=\s*(\d{1,5}(?:[,.]\d+)?)", re.IGNORECASE)


def strip_mtext_codes(text: str) -> str:
    text = re.sub(r"\\[A-Za-z]\s*[^\\;]*;?", "", text)
    text = re.sub(r"\\P", " ", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def main() -> None:
    print(f"Читаю {DXF_PATH.name} ...")
    doc = ezdxf.readfile(str(DXF_PATH))

    # Собираем ВСЕ MTEXT из BLOCKS с координатами
    all_texts: list[tuple[float, float, str]] = []
    for block in doc.blocks:
        if block.name.startswith("*"):
            continue
        for ent in block:
            if ent.dxftype() in ("MTEXT", "TEXT"):
                try:
                    x, y = ent.dxf.insert.x, ent.dxf.insert.y
                    raw = (ent.text if ent.dxftype() == "MTEXT" else ent.dxf.text) or ""
                    cleaned = strip_mtext_codes(raw) if ent.dxftype() == "MTEXT" else raw.strip()
                    all_texts.append((x, y, cleaned))
                except Exception:
                    pass

    print(f"  MTEXT/TEXT в BLOCKS: {len(all_texts)}")

    # Все точки с FRLS/FRHF 3х1,5
    targets = []
    for i, (x, y, txt) in enumerate(all_texts):
        if FRLS_RE.search(txt):
            targets.append(("FRLS", x, y, txt))
        elif FRHF_RE.search(txt):
            targets.append(("FRHF", x, y, txt))

    print(f"  FRLS/FRHF 3х1,5 точек в BLOCKS: {len(targets)}")

    # Все точки с L=NNN
    length_points = []
    for (x, y, txt) in all_texts:
        for m in L_RE.finditer(txt):
            try:
                val = float(m.group(1).replace(",", "."))
                length_points.append((x, y, val))
            except ValueError:
                pass

    print(f"  L=NNN точек: {len(length_points)}")

    # Для каждой цели ищем ближайшую L= в радиусе 2000 с ограничением Y (строки таблицы)
    matched = 0
    unmatched = 0
    lengths_sum = 0.0
    lengths_by_brand = defaultdict(list)

    used_length_points = set()
    # Сортируем чтобы приоритет было у ближайших пар
    pairs = []
    for brand, tx, ty, ttxt in targets:
        for li, (lx, ly, lval) in enumerate(length_points):
            dx = abs(lx - tx)
            dy = abs(ly - ty)
            # В таблице: та же строка (|dy|<150), соседняя колонка (dx<2000)
            if dy < 150 and dx < 2500:
                pairs.append((dx + dy, brand, tx, ty, li, lval))
    pairs.sort()

    # Жадное сопоставление: каждый target и каждый length_point используется только один раз
    target_used = set()
    length_used = set()
    for dist, brand, tx, ty, li, lval in pairs:
        key = (brand, tx, ty)
        if key in target_used or li in length_used:
            continue
        target_used.add(key)
        length_used.add(li)
        matched += 1
        lengths_sum += lval
        lengths_by_brand[brand].append(lval)

    unmatched = len(targets) - matched

    print(f"\n  ✔ Сматчено target↔L=: {matched}")
    print(f"  ✘ Без пары: {unmatched}")
    print(f"  Сумма длин: {lengths_sum:.0f} м")
    for brand, vals in lengths_by_brand.items():
        print(f"    {brand}: {len(vals)} шт, {sum(vals):.0f} м (min={min(vals)}, max={max(vals)})")

    # Посмотрим распределение dy для первых 20 matched
    print(f"\n  Эталон ВОР (раздел Кабельная продукция):")
    print(f"    ВБШвнг(А)-FRLS 3х1,5: 4460 м")
    print(f"    ППГнг(А)-FRHF 3х1,5: 1300 м")
    print(f"    ИТОГО:                5760 м")


if __name__ == "__main__":
    main()
