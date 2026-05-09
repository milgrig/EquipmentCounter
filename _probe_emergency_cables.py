#!/usr/bin/env python3
"""
Read-only диагностический скрипт.

Цель: локализовать, в каких DXF файлах 3-й захватки ГПК прячутся недостающие
огнестойкие (аварийные) кабели:
  - ВБШвнг(А)-FRLS 3х1,5      (эталон 4460 м, найдено 2134 м, недобор ~2326 м)
  - ППГнг(А)-FRHF 3х1,5       (эталон 1300 м, найдено 535 м, недобор ~765 м)

Скрипт:
  1. Проходит по всем DXF (25 файлов).
  2. Для каждого файла через equipment_counter.extract_cables_dxf (raw ezdxf path)
     и через equipment_counter._extract_cables_tablecontent (OBJECTS walker)
     считает, сколько реально нашлось этих двух марок в 3х1,5.
  3. Дополнительно делает «сырую» сканировку всего текста DXF (включая INSERT/BLOCK),
     ища пары:
         <марка> ... L=<число>
     независимо от того, где они лежат (MODEL/OBJECTS/BLOCKS/ENTITIES).
  4. Выводит таблицу: файл × (MTEXT raw / TABLECONTENT / «дикий» скан).
"""
from __future__ import annotations

import io
import os
import re
import sys
from pathlib import Path

# Force UTF-8 for stdout on Windows
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")


CAPTURE_DIR = Path(
    "Data/ДБТ разделы для ИИ/03_ГПК_/3-я захватка/_converted_dxf/01_DWG"
)

# Паттерны интересующих нас марок (упрощённо — как их пишет проектировщик).
# Учитываем типичные искажения: дефисы, пробелы, регистр, разные тире, варианты скобок.
EMERGENCY_PATTERNS = {
    "ВБШвнг(А)-FRLS 3х1,5": re.compile(
        r"ВБШ\s*внг\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*LS\s*3\s*[xх×]\s*1[,.]5",
        re.IGNORECASE,
    ),
    "ППГнг(А)-FRHF 3х1,5": re.compile(
        r"ППГ\s*нг\s*[-‐–—]?\s*\(\s*А\s*\)\s*[-‐–—]?\s*FR\s*HF\s*3\s*[xх×]\s*1[,.]5",
        re.IGNORECASE,
    ),
}

# Паттерн «L=NNN» (длина в метрах)
L_PATTERN = re.compile(r"L\s*=\s*(\d{1,5}(?:[,.]\d+)?)", re.IGNORECASE)

# Окно совместного появления марки и L=
JOINT_WINDOW = 300  # символов


def scan_raw_text(path: Path) -> dict[str, tuple[int, float]]:
    """
    Грубая сканировка: читает весь DXF как текст, ищет в окне JOINT_WINDOW пары
    (марка, L=NNN). Возвращает {марка: (count, sum_meters)}.
    """
    try:
        data = path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        print(f"  [error reading {path.name}: {e}]")
        return {}

    results: dict[str, tuple[int, float]] = {}
    for brand, pat in EMERGENCY_PATTERNS.items():
        count = 0
        total_m = 0.0
        for m in pat.finditer(data):
            start = m.end()
            window = data[start : start + JOINT_WINDOW]
            lm = L_PATTERN.search(window)
            if lm:
                try:
                    length = float(lm.group(1).replace(",", "."))
                    total_m += length
                    count += 1
                except ValueError:
                    pass
        results[brand] = (count, total_m)
    return results


def scan_via_equipment_counter(path: Path) -> dict[str, tuple[int, float]]:
    """
    Использует публичную функцию equipment_counter.extract_cables_dxf, если она есть.
    """
    try:
        sys.path.insert(0, str(Path.cwd()))
        import equipment_counter as ec
    except Exception as e:
        print(f"  [import equipment_counter failed: {e}]")
        return {}

    try:
        cables = ec.extract_cables_dxf(str(path))  # type: ignore[attr-defined]
    except AttributeError:
        # Попробуем альтернативный API (может называться иначе)
        try:
            cables = ec._extract_cables_raw_dxf(str(path))  # type: ignore[attr-defined]
        except Exception:
            return {}
    except Exception as e:
        print(f"  [extract_cables_dxf error on {path.name}: {e}]")
        return {}

    results: dict[str, tuple[int, float]] = {
        k: (0, 0.0) for k in EMERGENCY_PATTERNS.keys()
    }
    for c in cables or []:
        # c — dict или объект с полями cable_type / length_m (имена могут варьироваться)
        brand_text = getattr(c, "cable_type", None) or (
            c.get("cable_type") if isinstance(c, dict) else ""
        ) or ""
        length_m = getattr(c, "length_m", None) or (
            c.get("length_m") if isinstance(c, dict) else 0.0
        ) or 0.0
        for name, pat in EMERGENCY_PATTERNS.items():
            if pat.search(brand_text):
                cnt, tot = results[name]
                results[name] = (cnt + 1, tot + float(length_m))
                break
    return results


def main() -> None:
    if not CAPTURE_DIR.exists():
        print(f"Не найдена папка: {CAPTURE_DIR}")
        return

    dxf_files = sorted(CAPTURE_DIR.glob("*.dxf"))
    print(f"DXF файлов: {len(dxf_files)}\n")

    total_raw = {k: (0, 0.0) for k in EMERGENCY_PATTERNS.keys()}
    total_ec = {k: (0, 0.0) for k in EMERGENCY_PATTERNS.keys()}

    header = (
        f"{'Файл':<55} | "
        f"{'FRLS raw (шт/м)':<20} | {'FRHF raw (шт/м)':<20} | "
        f"{'FRLS ec (шт/м)':<20} | {'FRHF ec (шт/м)':<20}"
    )
    print(header)
    print("-" * len(header))

    for path in dxf_files:
        raw = scan_raw_text(path)
        ec_res = scan_via_equipment_counter(path)

        def fmt(d, key):
            c, m = d.get(key, (0, 0.0))
            if c == 0:
                return "-"
            return f"{c}/{m:.0f}"

        name = path.name
        if len(name) > 53:
            name = name[:52] + "…"

        print(
            f"{name:<55} | "
            f"{fmt(raw, 'ВБШвнг(А)-FRLS 3х1,5'):<20} | "
            f"{fmt(raw, 'ППГнг(А)-FRHF 3х1,5'):<20} | "
            f"{fmt(ec_res, 'ВБШвнг(А)-FRLS 3х1,5'):<20} | "
            f"{fmt(ec_res, 'ППГнг(А)-FRHF 3х1,5'):<20}"
        )

        for k in EMERGENCY_PATTERNS.keys():
            c_raw, m_raw = raw.get(k, (0, 0.0))
            c_ec, m_ec = ec_res.get(k, (0, 0.0))
            tc, tm = total_raw[k]
            total_raw[k] = (tc + c_raw, tm + m_raw)
            tc, tm = total_ec[k]
            total_ec[k] = (tc + c_ec, tm + m_ec)

    print("-" * len(header))
    print(
        f"{'ИТОГО':<55} | "
        f"{total_raw['ВБШвнг(А)-FRLS 3х1,5'][0]}/{total_raw['ВБШвнг(А)-FRLS 3х1,5'][1]:.0f} m".ljust(23)
        + " | "
        + f"{total_raw['ППГнг(А)-FRHF 3х1,5'][0]}/{total_raw['ППГнг(А)-FRHF 3х1,5'][1]:.0f} m".ljust(23)
        + " | "
        + f"{total_ec['ВБШвнг(А)-FRLS 3х1,5'][0]}/{total_ec['ВБШвнг(А)-FRLS 3х1,5'][1]:.0f} m".ljust(23)
        + " | "
        + f"{total_ec['ППГнг(А)-FRHF 3х1,5'][0]}/{total_ec['ППГнг(А)-FRHF 3х1,5'][1]:.0f} m"
    )
    print()
    print("Эталон ВОР:")
    print("  ВБШвнг(А)-FRLS 3х1,5: 4460 м")
    print("  ППГнг(А)-FRHF 3х1,5: 1300 м")


if __name__ == "__main__":
    main()
