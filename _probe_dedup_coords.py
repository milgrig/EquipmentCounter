#!/usr/bin/env python3
"""
Низкоуровневый DXF walker: проходим по group-кодам, извлекаем все MTEXT
из каждой секции (BLOCKS / ENTITIES / OBJECTS), и для каждого извлечённого
текста регистрируем (section, x, y, cable_brand, cable_length).

Задача: понять, какая секция даёт число кабелей, близкое к эталону 116.
"""
from __future__ import annotations
import io, re, sys
from pathlib import Path
from collections import defaultdict

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

DXF = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/"
    "003, 004 - Схемы ЩО, ЩАО.dxf"
)

# Паттерны марки + L=
BRAND_L_RE = re.compile(
    r"(ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS|"
    r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF|"
    r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*LS|"
    r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*HF)"
    r"\s*(\d+\s*[xх×]\s*\d+[,.]\d+)"
    r".{0,200}?"
    r"L\s*=\s*(\d{1,5})",
    re.IGNORECASE | re.DOTALL,
)


def iter_group_codes(data: str):
    """Итератор по (code, value) парам DXF."""
    lines = data.split("\n")
    i = 0
    n = len(lines)
    while i < n - 1:
        code_line = lines[i].strip()
        if code_line.isdigit() or (code_line.startswith("-") and code_line[1:].isdigit()):
            code = int(code_line)
            value = lines[i + 1]
            yield code, value
            i += 2
        else:
            i += 1


def main() -> None:
    print(f"Читаю {DXF.name} ({DXF.stat().st_size:,} байт)...")
    data = DXF.read_text(encoding="utf-8", errors="ignore")

    # Пройдём по всему файлу и вытащим каждый MTEXT с его координатами.
    # Состояние: в каком мы сейчас entity, его code-10/20 (coords) и code-1 (text).
    entities = []  # list of (section, x, y, text)
    current_section = None
    current_ent_type = None
    current_x = None
    current_y = None
    current_text_parts = []
    current_extra_parts = []  # для code 3 (continuation)

    def flush():
        if current_ent_type == "MTEXT" and current_x is not None:
            text = "".join(current_text_parts) + "".join(current_extra_parts)
            entities.append((current_section, current_x, current_y, text))

    for code, value in iter_group_codes(data):
        if code == 0:
            # Новый entity / секция
            flush()
            current_ent_type = value.strip()
            if current_ent_type == "SECTION":
                current_ent_type = None
            current_x = None
            current_y = None
            current_text_parts = []
            current_extra_parts = []
        elif code == 2:
            v = value.strip()
            if v in ("HEADER", "CLASSES", "TABLES", "BLOCKS", "ENTITIES", "OBJECTS"):
                current_section = v
        elif code == 10 and current_ent_type == "MTEXT":
            try:
                current_x = float(value.strip())
            except ValueError:
                pass
        elif code == 20 and current_ent_type == "MTEXT":
            try:
                current_y = float(value.strip())
            except ValueError:
                pass
        elif code == 1 and current_ent_type == "MTEXT":
            current_text_parts.append(value)
        elif code == 3 and current_ent_type == "MTEXT":
            current_extra_parts.append(value)
    flush()

    print(f"  MTEXT извлечено: {len(entities)}")

    # По каждой секции: найти FRLS/FRHF, дедуплицировать по (x,y,text)
    by_section = defaultdict(list)  # section -> [(x,y,brand,section_cable,length)]
    for section, x, y, text in entities:
        for m in BRAND_L_RE.finditer(text):
            brand = m.group(1)
            cable_section = m.group(2)
            length_str = m.group(3)
            try:
                length = int(length_str)
            except ValueError:
                continue
            by_section[section].append((x, y, brand, cable_section, length, text))

    for section_name in ("BLOCKS", "ENTITIES", "OBJECTS"):
        items = by_section.get(section_name, [])
        # Дедупликация по (x,y,brand,cable_section,length) — если всё совпадает, это один и тот же кабель
        dedup_coords = set()
        for x, y, brand, cs, l, txt in items:
            dedup_coords.add((round(x or 0, 1), round(y or 0, 1), brand.replace(" ", ""), cs, l))
        # Ещё более строгая: только по (x,y)
        dedup_xy = set()
        total_by_xy = defaultdict(float)
        for x, y, brand, cs, l, txt in items:
            key = (round(x or 0, 1), round(y or 0, 1))
            if key not in dedup_xy:
                dedup_xy.add(key)
                total_by_xy[key] = l
        unique_xy_total_m = sum(total_by_xy.values())

        total_all = sum(l for _, _, _, _, l, _ in items)

        print(f"\n--- {section_name} ---")
        print(f"  всего FRLS/FRHF/LS/HF записей:       {len(items):5}")
        print(f"  уникальных по (x,y,brand,sec,l):     {len(dedup_coords):5}  сумма={sum(dc[-1] for dc in dedup_coords):,.0f} м")
        print(f"  уникальных по (x,y):                 {len(dedup_xy):5}  сумма={unique_xy_total_m:,.0f} м")
        print(f"  всего метров без дедупа:             {total_all:,.0f} м")

    # Объединённая дедупликация по (x,y)
    all_xy = set()
    all_xy_total = defaultdict(float)
    for items in by_section.values():
        for x, y, brand, cs, l, txt in items:
            key = (round(x or 0, 1), round(y or 0, 1))
            if key not in all_xy:
                all_xy.add(key)
                all_xy_total[key] = l
    total = sum(all_xy_total.values())
    print(f"\n=== Объединено по (x,y): {len(all_xy)} уникальных, {total:,.0f} м ===")


if __name__ == "__main__":
    main()
