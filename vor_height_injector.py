"""
vor_height_injector.py — добавляет в агрегат ВОР строки «Кабель по вертикали»
на основе аннотаций «Гр.X на отм. ...» из PDF.

Использует pdf_cable_height.summarize_cable_heights для расчёта метража
вертикальных стояков и инжектирует результат как отдельные строки агрегата
(одну на цвет: красный, синий) либо одну общую строку.

Назначение:
  Текущий пайплайн measure_cables считает только горизонтальные сегменты
  цветных трасс. Стояки между этажами (вертикальные участки между отметками
  0.000, +9.000, +13.800 ...) в плане отсутствуют как геометрия, но
  зафиксированы в текстовых аннотациях рядом с трассой:
        «Гр.1-Гр.8 на отм. 0.000, +9.000, +13.800, +18.600, +23.400, +28.200»
  Из такой аннотации:
        riser_height_m   = max(elev) - min(elev)         = 28.200
        total_riser_m    = riser_height_m * n_groups     = 28.2 × 8 = 225.6 м
  Этот метраж необходимо добавить к итогам ВОР отдельной строкой,
  чтобы общая длина кабеля совпала с эталонными ВОР.

Публичный API:
    inject_vertical_cable_rows(aggregated, pdf_paths, *, mode='split') -> list[dict]
        aggregated  — список строк ВОР (как формирует существующий пайплайн)
        pdf_paths   — все PDF, по которым считалась агрегация
        mode        — 'split' (две строки: красный/синий) или 'single' (одна строка)
        Возвращает НОВЫЙ список (без мутаций) со вставленными строками.

    build_vertical_cable_rows(pdf_paths, *, mode='split') -> list[dict]
        Только сборка строк, без инжекции.

    install_aggregate_height_patch(folder_to_pdfs)
        Опциональный monkey-patch, если хочется зашить инжекцию в run_web.py:
            install_aggregate_height_patch(lambda folder_id: list_of_pdfs)
"""

from __future__ import annotations

import os
import re
from typing import Callable, Iterable, Optional

from pdf_cable_height import (
    CableHeightSummary,
    summarize_cable_heights,
    COLOR_RED,
    COLOR_BLUE,
)


# ────────────────────────────────────────────────────────────────────────────
# Конфигурация имён строк
# ────────────────────────────────────────────────────────────────────────────

# Название позиции в ВОР для вертикальных стояков.
# Можно переопределить извне: vor_height_injector.NAME_TEMPLATE = "..."
NAME_RED = "Кабель силовой по вертикали (стояки), красная трасса"
NAME_BLUE = "Кабель силовой по вертикали (стояки), синяя трасса"
NAME_COMBINED = "Кабель силовой по вертикали (стояки)"
UNIT = "м"


# ────────────────────────────────────────────────────────────────────────────
# Утилиты форматирования
# ────────────────────────────────────────────────────────────────────────────

def _fmt_qty(v: float) -> str:
    """1234.5 -> '1234.5'; 1234.0 -> '1234'."""
    if abs(v - round(v)) < 1e-6:
        return str(int(round(v)))
    return f"{v:.1f}".rstrip("0").rstrip(".")


def _format_elevation_formula(s: CableHeightSummary,
                               color: Optional[str]) -> str:
    """Сформировать формулу вида 'Δ1×N1 + Δ2×N2 + ...' по парам отметок.

    Если color задан — берём только аннотации этого цвета.
    Если color is None — все аннотации.
    """
    pairs: dict[str, float] = {}
    for a in s.annotations:
        if color is not None and a.color != color:
            continue
        if a.n_elevations < 2:
            continue
        try:
            vals = sorted(float(e.replace(",", ".")) for e in a.elevations)
        except ValueError:
            continue
        n_groups = max(1, a.n_groups)
        for i in range(1, len(vals)):
            delta = vals[i] - vals[i - 1]
            key = f"{vals[i-1]:+.3f}→{vals[i]:+.3f}"
            pairs[key] = pairs.get(key, 0.0) + delta * n_groups
    if not pairs:
        return ""
    parts = [f"{_fmt_qty(v)}" for v in pairs.values()]
    return "+".join(parts)


def _drawing_refs_for_color(s: CableHeightSummary,
                              color: Optional[str]) -> str:
    """Уникальные имена PDF-источников аннотаций для цвета."""
    seen: list[str] = []
    seen_set: set[str] = set()
    for a in s.annotations:
        if color is not None and a.color != color:
            continue
        src = a.source_pdf
        if not src:
            continue
        base = os.path.basename(src)
        if base not in seen_set:
            seen_set.add(base)
            seen.append(base)
    return ", ".join(seen)


# ────────────────────────────────────────────────────────────────────────────
# Сборка строк агрегата
# ────────────────────────────────────────────────────────────────────────────

def _row_for_color(s: CableHeightSummary, color: str, name: str) -> Optional[dict]:
    qty = float(s.by_color.get(color, 0.0) or 0.0)
    if qty <= 0:
        return None
    formula = _format_elevation_formula(s, color)
    refs = _drawing_refs_for_color(s, color)
    return {
        "name": name,
        "unit": UNIT,
        "total": round(qty, 1),
        "formula": formula or _fmt_qty(qty),
        "drawing_refs": refs,
        "drawing_refs_by_elevation": refs,
        "category": "Кабельная продукция",
        "source": "vor_height_injector",
        "is_vertical_cable": True,
        "color": color,
    }


def _row_combined(s: CableHeightSummary) -> Optional[dict]:
    qty = float(s.total_vertical_m or 0.0)
    if qty <= 0:
        return None
    formula = _format_elevation_formula(s, None)
    refs = _drawing_refs_for_color(s, None)
    return {
        "name": NAME_COMBINED,
        "unit": UNIT,
        "total": round(qty, 1),
        "formula": formula or _fmt_qty(qty),
        "drawing_refs": refs,
        "drawing_refs_by_elevation": refs,
        "category": "Кабельная продукция",
        "source": "vor_height_injector",
        "is_vertical_cable": True,
        "color": None,
    }


def build_vertical_cable_rows(pdf_paths: Iterable[str],
                                *,
                                mode: str = "split") -> list[dict]:
    """Посчитать высоту кабелей и собрать строки для агрегата.

    mode:
        'split'  — две строки (красная/синяя), если оба цвета присутствуют
        'single' — одна общая строка
    """
    paths = [p for p in pdf_paths if p]
    if not paths:
        return []
    try:
        s = summarize_cable_heights(paths)
    except Exception:
        return []
    if s.total_vertical_m <= 0:
        return []

    rows: list[dict] = []
    if mode == "single":
        r = _row_combined(s)
        if r:
            rows.append(r)
    else:
        r_red = _row_for_color(s, COLOR_RED, NAME_RED)
        r_blue = _row_for_color(s, COLOR_BLUE, NAME_BLUE)
        if r_red:
            rows.append(r_red)
        if r_blue:
            rows.append(r_blue)
        # Если цвета не определились (color=None) — добавим общую строку
        # на остаток метража, который не попал в красный/синий.
        accounted = sum(float(r.get("total", 0.0)) for r in rows)
        leftover = float(s.total_vertical_m) - accounted
        if leftover > 0.5:
            r_other = {
                "name": NAME_COMBINED,
                "unit": UNIT,
                "total": round(leftover, 1),
                "formula": _fmt_qty(leftover),
                "drawing_refs": "",
                "drawing_refs_by_elevation": "",
                "category": "Кабельная продукция",
                "source": "vor_height_injector",
                "is_vertical_cable": True,
                "color": None,
            }
            rows.append(r_other)
    return rows


# ────────────────────────────────────────────────────────────────────────────
# Инжекция в готовый агрегат
# ────────────────────────────────────────────────────────────────────────────

# Регулярка кабельных позиций для точки вставки
_CABLE_NAME_RE = re.compile(
    r"кабел", re.IGNORECASE,
)


def _find_insert_position(aggregated: list[dict]) -> int:
    """Вернуть индекс ПОСЛЕ последней кабельной строки (или len(aggregated))."""
    last_cable = -1
    for i, row in enumerate(aggregated):
        name = str(row.get("name", ""))
        cat = str(row.get("category", ""))
        if _CABLE_NAME_RE.search(name) or "абельн" in cat.lower():
            last_cable = i
    return last_cable + 1 if last_cable >= 0 else len(aggregated)


def _already_injected(aggregated: list[dict]) -> bool:
    return any(row.get("source") == "vor_height_injector"
               for row in aggregated)


def inject_vertical_cable_rows(aggregated: list[dict],
                                 pdf_paths: Iterable[str],
                                 *,
                                 mode: str = "split") -> list[dict]:
    """Вернуть НОВЫЙ список с добавленными строками вертикальных стояков.

    Если строки уже были инжектированы (source == 'vor_height_injector') —
    вернуть исходный список без изменений.
    """
    if not isinstance(aggregated, list):
        return aggregated
    if _already_injected(aggregated):
        return list(aggregated)
    extra = build_vertical_cable_rows(pdf_paths, mode=mode)
    if not extra:
        return list(aggregated)
    pos = _find_insert_position(aggregated)
    out = list(aggregated[:pos]) + list(extra) + list(aggregated[pos:])
    return out


# ────────────────────────────────────────────────────────────────────────────
# Опциональный monkey-patch для run_web.py
# ────────────────────────────────────────────────────────────────────────────

# Тип резолвера: получить список PDF-путей по идентификатору папки
PdfPathsResolver = Callable[[str], list[str]]

_resolver_holder: dict[str, Optional[PdfPathsResolver]] = {"resolver": None}


def install_aggregate_height_patch(resolver: PdfPathsResolver,
                                     *,
                                     mode: str = "split") -> None:
    """Подменить get_cached_vor у vor_export_patch так, чтобы агрегат
    автоматически содержал строки вертикальных кабелей.

    resolver(folder_id) -> list[str]   — функция резолвинга PDF-путей
    """
    _resolver_holder["resolver"] = resolver
    try:
        import vor_export_patch as _vep
    except ImportError:
        return

    _orig_get_cached = getattr(_vep, "get_cached_vor", None)
    if _orig_get_cached is None:
        return

    def _patched_get_cached(folder_id: str):
        agg = _orig_get_cached(folder_id)
        if not isinstance(agg, list):
            return agg
        if _already_injected(agg):
            return agg
        try:
            paths = resolver(folder_id) or []
        except Exception:
            paths = []
        if not paths:
            return agg
        return inject_vertical_cable_rows(agg, paths, mode=mode)

    _vep.get_cached_vor = _patched_get_cached
