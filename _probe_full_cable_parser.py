#!/usr/bin/env python3
"""
Read-only: полный парсер кабельного журнала из DXF.

Читает DXF как сырой текст и извлекает ВСЕ MTEXT (во всех секциях),
находит в них пары (марка кабеля) + (длина L=NNN) в одном MTEXT.
Делает дедупликацию по (brand, x, y, length).

Цель: получить полную картину кабельного журнала в
"003, 004 - Схемы ЩО, ЩАО.dxf" и сравнить с эталоном.
"""
from __future__ import annotations

import io
import re
import sys
from collections import defaultdict
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

DXF = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG/"
    "003, 004 - Схемы ЩО, ЩАО.dxf"
)

# Марка кабеля (разные варианты написания)
BRAND_PATTERNS = [
    ("ВБШвнг(А)-LS 3х1,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*LS\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ВБШвнг(А)-LS 3х2,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*LS\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
    ("ВБШвнг(А)-FRLS 3х1,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ВБШвнг(А)-FRLS 3х2,5",
     re.compile(r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-HF 3х1,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*HF\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-HF 3х2,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*HF\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-FRHF 3х1,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*1[,.]5",
                re.IGNORECASE)),
    ("ППГнг(А)-FRHF 3х2,5",
     re.compile(r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*2[,.]5",
                re.IGNORECASE)),
]

L_RE = re.compile(r"L\s*=\s*(\d{1,5}(?:[,.]\d+)?)", re.IGNORECASE)
LAYING_RE = re.compile(r"в\s+(лотке|гофре|трубе|кабель[- ]канале|земле)", re.IGNORECASE)
DELTA_U_RE = re.compile(r"[ΔΔ]U\s*=\s*(\d+[,.]\d+)\s*%", re.IGNORECASE)


def iter_dxf_entities(path: Path):
    """
    Проходит DXF как поток (code, value) пар.
    Группирует entity между маркерами "0\\n<TYPE>\\n".
    Yields (entity_type, section_name, entity_dict)
    где entity_dict = {code: [values...]}.
    """
    section = "?"
    with path.open("r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    i = 0
    n = len(lines)
    current_entity: dict | None = None
    current_type: str | None = None

    while i < n - 1:
        code_line = lines[i].strip()
        value_line = lines[i + 1].rstrip("\n").rstrip("\r")
        try:
            code = int(code_line)
        except ValueError:
            i += 1
            continue
        i += 2

        if code == 0:
            # новый entity
            if current_entity is not None and current_type is not None:
                yield current_type, section, current_entity
            current_type = value_line.strip()
            current_entity = {}
            # Отслеживаем секцию
            if current_type == "SECTION":
                # Следующий (2, <name>) задаёт имя секции
                if i < n - 1 and lines[i].strip() == "2":
                    section = lines[i + 1].strip()
                    i += 2
                current_entity = None
                current_type = None
            elif current_type == "ENDSEC":
                section = "?"
                current_entity = None
                current_type = None
        else:
            if current_entity is not None:
                current_entity.setdefault(code, []).append(value_line)

    if current_entity is not None and current_type is not None:
        yield current_type, section, current_entity


def strip_mtext_codes(text: str) -> str:
    text = re.sub(r"\\[A-Za-z]\s*[^\\;]*;?", "", text)
    text = re.sub(r"\\P", " ", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def concat_mtext_content(entity: dict) -> str:
    """
    Для MTEXT в DXF текст хранится в коде 1 (основная часть)
    и опционально в коде 3 (продолжение, если >250 символов).
    """
    parts: list[str] = []
    # Код 3 — продолжения, идут ДО кода 1.
    parts.extend(entity.get(3, []))
    parts.extend(entity.get(1, []))
    return "".join(parts)


def extract_xy(entity: dict) -> tuple[float, float]:
    try:
        x = float(entity.get(10, ["0"])[0])
        y = float(entity.get(20, ["0"])[0])
        return x, y
    except (ValueError, IndexError):
        return 0.0, 0.0


def main() -> None:
    print(f"Читаю {DXF.name} (размер {DXF.stat().st_size/1e6:.1f} МБ) ...")
    found: list[dict] = []
    mtext_total = 0

    for etype, section, entity in iter_dxf_entities(DXF):
        if etype != "MTEXT":
            continue
        mtext_total += 1
        raw = concat_mtext_content(entity)
        if not raw:
            continue
        cleaned = strip_mtext_codes(raw)

        # Найти марку
        brand = None
        for bname, pat in BRAND_PATTERNS:
            if pat.search(cleaned):
                brand = bname
                break
        if not brand:
            continue

        # Найти длину
        lm = L_RE.search(cleaned)
        if not lm:
            continue
        try:
            length = float(lm.group(1).replace(",", "."))
        except ValueError:
            continue

        # Способ прокладки (если указан)
        lays = [m.group(1).lower() for m in LAYING_RE.finditer(cleaned)]

        # ΔU
        delta_u = None
        dum = DELTA_U_RE.search(cleaned)
        if dum:
            try:
                delta_u = float(dum.group(1).replace(",", "."))
            except ValueError:
                pass

        x, y = extract_xy(entity)

        found.append({
            "brand": brand,
            "length_m": length,
            "x": x,
            "y": y,
            "section": section,
            "laying": ",".join(lays) if lays else "",
            "delta_u": delta_u,
            "text": cleaned[:150],
        })

    print(f"Всего MTEXT в DXF: {mtext_total}")
    print(f"Найдено MTEXT с маркой+L=: {len(found)}")

    # Сводка до дедупликации
    print("\n--- Без дедупликации (по секции × марке) ---")
    by = defaultdict(lambda: [0, 0.0])
    for f in found:
        key = (f["section"], f["brand"])
        by[key][0] += 1
        by[key][1] += f["length_m"]
    for (sec, br), (cnt, tot) in sorted(by.items()):
        print(f"  [{sec:10}] {br:30}  {cnt:4} шт   {tot:7.0f} м")

    # Дедупликация по (brand, x, y, length)
    unique_xy = {}
    for f in found:
        key = (f["brand"], round(f["x"], 1), round(f["y"], 1), round(f["length_m"], 1))
        if key not in unique_xy:
            unique_xy[key] = f

    print(f"\n--- Дедупликация по (brand, X, Y, length): {len(unique_xy)} уникальных ---")
    by = defaultdict(lambda: [0, 0.0])
    for f in unique_xy.values():
        by[f["brand"]][0] += 1
        by[f["brand"]][1] += f["length_m"]
    total_all = 0.0
    for br, (cnt, tot) in sorted(by.items()):
        print(f"  {br:30}  {cnt:4} шт   {tot:7.0f} м")
        total_all += tot
    print(f"  {'ИТОГО':30}  {sum(v[0] for v in by.values()):4} шт   {total_all:7.0f} м")

    # Эталон
    etalon = {
        "ВБШвнг(А)-LS 3х1,5": 2317,
        "ВБШвнг(А)-LS 3х2,5": 1127,
        "ВБШвнг(А)-FRLS 3х1,5": 4460,
        "ВБШвнг(А)-FRLS 3х2,5": 180,
        "ППГнг(А)-HF 3х1,5": 842,
        "ППГнг(А)-HF 3х2,5": 0,
        "ППГнг(А)-FRHF 3х1,5": 1300,
        "ППГнг(А)-FRHF 3х2,5": 0,
    }
    print(f"\n--- Сравнение с эталоном ---")
    print(f"  {'Марка':30} {'Найдено':>10} {'Эталон':>10} {'Δ':>8}")
    for br in sorted(etalon):
        our = by.get(br, [0, 0])[1]
        et = etalon[br]
        diff = our - et
        if et:
            pct = diff / et * 100
            print(f"  {br:30} {our:>10.0f} {et:>10} {diff:>+8.0f}  ({pct:+.0f}%)")
        else:
            print(f"  {br:30} {our:>10.0f} {et:>10} {diff:>+8.0f}")
    our_sum = sum(v[1] for v in by.values())
    et_sum = sum(etalon.values())
    print(f"  {'ИТОГО':30} {our_sum:>10.0f} {et_sum:>10} {our_sum - et_sum:>+8.0f}  ({(our_sum-et_sum)/et_sum*100:+.1f}%)")


if __name__ == "__main__":
    main()
