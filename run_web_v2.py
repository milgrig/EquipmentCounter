"""
run_web_v2.py — расширенный entry point поверх run_web.py.

Что добавляет:
  • Вертикальные кабельные стояки в ВОР через vor_height_injector.
    После пере-агрегации по отметкам (regroup_aggregate_by_elevation)
    в кэш ВОР дописываются строки «Кабель силовой по вертикали (стояки)»,
    рассчитанные из аннотаций «Гр.X на отм. ...» в исходных PDF.

Запуск:
    python run_web_v2.py
    # или
    uvicorn run_web_v2:app --host 0.0.0.0 --port 8050

Импортирует run_web (вместе со всеми его патчами: эталонный DOCX,
HTML-инжектор, эндпоинты экспорта) и довешивает поверх patch для
get_cached_vor.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

# 1. Стартуем все патчи из run_web (DOCX-эталон, regroup-elevation,
#    HTML-инжектор, register_vor_export). Это импорт-сайд-эффект.
import run_web  # noqa: F401

from run_web import app  # реэкспорт для uvicorn run_web_v2:app  # noqa: F401

import vor_export_patch as _vep
import web_app as wa
from vor_height_injector import (
    inject_vertical_cable_rows,
    _already_injected,
)


# ────────────────────────────────────────────────────────────────────────────
# Резолвер PDF-путей по folder_id (через web_app)
# ────────────────────────────────────────────────────────────────────────────

def _resolve_pdfs(folder_id: str) -> list[str]:
    """folder_id → абсолютные пути всех PDF в этой папке."""
    try:
        rel_folder = wa._id_to_folder(folder_id)
    except Exception:
        return []
    if not rel_folder:
        return []
    try:
        files = wa._folder_files(rel_folder)
    except Exception:
        return []
    out: list[str] = []
    for f in files:
        try:
            out.append(str(Path(f)))
        except Exception:
            continue
    return out


# ────────────────────────────────────────────────────────────────────────────
# Обёртка над уже-патченным get_cached_vor (после regroup-elevation)
# ────────────────────────────────────────────────────────────────────────────

_orig_get_cached_after_elev = _vep.get_cached_vor


def _get_cached_with_heights(app_obj, folder_id: str) -> Optional[dict]:
    cached = _orig_get_cached_after_elev(app_obj, folder_id)
    if not cached:
        return None
    if cached.get("_heights_done"):
        return cached
    aggregated = cached.get("aggregated") or []
    if _already_injected(aggregated):
        new_cached = dict(cached)
        new_cached["_heights_done"] = True
        return new_cached

    pdfs = _resolve_pdfs(folder_id)
    if not pdfs:
        new_cached = dict(cached)
        new_cached["_heights_done"] = True
        return new_cached

    try:
        new_agg = inject_vertical_cable_rows(aggregated, pdfs, mode="split")
    except Exception:
        new_agg = aggregated

    new_cached = dict(cached)
    new_cached["aggregated"] = new_agg
    new_cached["_heights_done"] = True
    return new_cached


_vep.get_cached_vor = _get_cached_with_heights


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("run_web_v2:app", host="0.0.0.0", port=8050, reload=False)
