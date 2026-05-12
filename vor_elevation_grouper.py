"""
vor_elevation_grouper.py — пере-агрегация ВОР по строительным отметкам.

Проблема: текущий агрегатор формирует формулу как сумму по именам PDF-файлов.
Это даёт длинные ряды слагаемых вида "14+6+7+6+6+6+5+13+6+7+6+6+6+5",
где 7 первых значений — листы освещения (отм. 0.000 ... +28.200),
а 7 следующих — листы плана привязки (тех же 7 отметок).

Решение: извлечь из имени файла отметку («отм. 0.000», «отм. +4.200» и т.п.)
и сложить значения по одной отметке. Тогда формула станет:
   "14+6+7+6+6+6+5"           — 7 отметок, каждая = сумма по всем листам этой отметки
или с подписями:
   "0.000:14 + +4.200:6 + +7.800:7 + ..."
"""

from __future__ import annotations

import re
from typing import Iterable

# Захватываем отметку: «отм. 0.000», «отм +4.200», «отм. +7.800 +9.000»
# (берём только первое число — это базовая отметка листа)
_ELEV_RE = re.compile(
    r"отм\.?\s*([+-]?\d+[.,]\d+)",
    re.IGNORECASE,
)


def extract_elevation(filename: str) -> str | None:
    """Вернуть отметку из имени PDF, например '+4.200' или '0.000', или None."""
    if not filename:
        return None
    m = _ELEV_RE.search(filename)
    if not m:
        return None
    raw = m.group(1).replace(",", ".")
    # нормализуем: 0.000 (без знака для нуля), +4.200 (со знаком плюс)
    try:
        v = float(raw)
    except ValueError:
        return raw
    if abs(v) < 1e-6:
        return "0.000"
    return f"{v:+.3f}"  # "+4.200", "-1.500"


def _elev_sort_key(elev: str) -> float:
    """Числовой ключ сортировки для отметки."""
    try:
        return float(elev.replace(",", ".").lstrip("+"))
    except (ValueError, AttributeError):
        return 0.0


def regroup_item_by_elevation(
    item: dict,
    *,
    files_map: dict[str, str] | None = None,
    labeled: bool = False,
) -> dict:
    """Пересобрать единичный элемент агрегата: формулу и ссылки на чертежи —
    из «по файлу» в «по отметке».

    item должен иметь поля: name, total, formula, drawing_refs, и опционально
    per_file (dict {filename: count}).

    files_map (optional) — fallback-маппинг {filename: elevation}, если
    per_file отсутствует, разберём drawing_refs + соответствующие части formula.
    """
    new_item = dict(item)
    per_file = item.get("per_file")

    if per_file and isinstance(per_file, dict):
        # Идеальный случай: есть {имя_файла: число}
        per_elev: dict[str, float] = {}
        used_files: dict[str, list[str]] = {}
        unknown_total: float = 0.0
        unknown_files: list[str] = []
        for fname, qty in per_file.items():
            elev = extract_elevation(fname)
            try:
                v = float(qty)
            except (TypeError, ValueError):
                continue
            if elev is None:
                unknown_total += v
                unknown_files.append(fname)
            else:
                per_elev[elev] = per_elev.get(elev, 0) + v
                used_files.setdefault(elev, []).append(fname)
    else:
        # Fallback: парсим formula = "14+6+7+6+..." и drawing_refs = "f1, f2, f3..."
        formula_str = str(item.get("formula", "")).strip()
        refs_str = str(item.get("drawing_refs", "")).strip()
        parts = [p.strip() for p in re.split(r"[+\s]+", formula_str) if p.strip()]
        files = [f.strip() for f in refs_str.split(",") if f.strip()]
        per_elev = {}
        used_files = {}
        unknown_total = 0.0
        unknown_files = []
        # Если количество слагаемых = количество файлов — связываем по индексу
        if len(parts) == len(files) and len(parts) > 0:
            for raw_qty, fname in zip(parts, files):
                try:
                    v = float(raw_qty.replace(",", "."))
                except ValueError:
                    continue
                elev = extract_elevation(fname)
                if elev is None:
                    unknown_total += v
                    unknown_files.append(fname)
                else:
                    per_elev[elev] = per_elev.get(elev, 0) + v
                    used_files.setdefault(elev, []).append(fname)
        else:
            # Не получилось распарсить — вернуть исходное
            return new_item

    # Сортируем отметки снизу вверх
    elev_sorted = sorted(per_elev.keys(), key=_elev_sort_key)

    # Форматируем числа без лишних .0
    def _fmt(v: float) -> str:
        if abs(v - round(v)) < 1e-6:
            return str(int(round(v)))
        return f"{v:.1f}".rstrip("0").rstrip(".")

    if labeled:
        parts_out = [f"{e}:{_fmt(per_elev[e])}" for e in elev_sorted]
    else:
        parts_out = [_fmt(per_elev[e]) for e in elev_sorted]
    if unknown_total > 0:
        parts_out.append(_fmt(unknown_total))

    new_item["formula"] = "+".join(parts_out) if parts_out else ""
    new_item["per_elevation"] = {e: per_elev[e] for e in elev_sorted}
    if unknown_total > 0:
        new_item["per_elevation"]["__other__"] = unknown_total

    # Также собрать сжатую ссылку на чертежи: "отм. 0.000, +4.200, +9.000..."
    if elev_sorted:
        elev_pretty = ", ".join(f"отм. {e}" for e in elev_sorted)
        existing_ref = str(item.get("drawing_refs", "")).strip()
        # Не заменяем, а укорачиваем: оставляем семейство файлов + отметки
        new_item["drawing_refs_by_elevation"] = elev_pretty

    return new_item


def regroup_aggregate_by_elevation(
    aggregated: list[dict],
    *,
    labeled: bool = False,
) -> list[dict]:
    """Применить пере-агрегацию ко всему списку строк ВОР."""
    return [regroup_item_by_elevation(it, labeled=labeled) for it in aggregated]
