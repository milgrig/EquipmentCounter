"""
vor_measure_wrapper.py — обёртка над pdf_count_geometry.measure_cables.

Делает ровно две вещи поверх оригинала:
  1. Вызывает оригинальный measure_cables(pdf_path, legend, pages).
  2. Применяет коррекцию длин по найденным аннотациям «на отм. ...»
     через vor_cable_corrections.collect_corrections.

Подменяет атрибуты результата:
    total_red_length_m  → скорректированная (с учётом мультипликатора и стояков)
    total_blue_length_m → то же

В отчёте появляются дополнительные поля (через monkey-patch dataclass):
    multiplier_used
    risers_m
    n_pages_with_anns
"""

from __future__ import annotations

import os
from typing import Any

from pdf_count_geometry import measure_cables as _orig_measure_cables
from vor_cable_corrections import collect_corrections, corrected_lengths


def measure_cables_with_annotations(pdf_path, legend=None, pages=None) -> Any:
    """Обёртка: считает длины + применяет аннотационный множитель."""
    result = _orig_measure_cables(pdf_path, legend, pages) if pages is not None \
             else _orig_measure_cables(pdf_path, legend)

    # Берём raw длины
    raw_red = float(getattr(result, "total_red_length_m", 0.0) or 0.0)
    raw_blue = float(getattr(result, "total_blue_length_m", 0.0) or 0.0)

    # Получаем коррекции по аннотациям
    try:
        corr = collect_corrections(str(pdf_path))
    except Exception:
        corr = {}

    info = corrected_lengths(raw_red, raw_blue, corr)

    # Применяем (только если что-то нашли — иначе оставляем raw)
    if info["n_pages_with_anns"] > 0 and info["multiplier"] > 1.0:
        try:
            result.total_red_length_m = info["red_corrected_m"]
            result.total_blue_length_m = info["blue_corrected_m"]
        except (AttributeError, TypeError):
            pass

    # Метаданные коррекции
    try:
        result.annotation_correction = {
            "raw_red_m": raw_red,
            "raw_blue_m": raw_blue,
            "multiplier": info["multiplier"],
            "risers_m": info["riser_m"],
            "n_pages_with_anns": info["n_pages_with_anns"],
            "groups_total": info["groups_total"],
        }
    except (AttributeError, TypeError):
        pass

    return result


def install_measure_cables_patch():
    """Подменить web_app.measure_cables на обёртку.

    Должно вызываться ДО того, как web_app начнёт обрабатывать запросы.
    """
    import web_app
    web_app.measure_cables = measure_cables_with_annotations
    return measure_cables_with_annotations
