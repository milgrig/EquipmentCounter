"""
run_web_with_height.py — расширенная точка входа: run_web.py + инжекция строк
«Кабель по вертикали» в кэшируемый агрегат ВОР.

Отличие от run_web.py:
    После того, как агрегат пере-группируется по отметкам (это уже делает
    run_web.py монкипатчем _get_cached_with_elevation), сюда добавляются
    строки вертикальных стояков, посчитанные по PDF-аннотациям
    «Гр.X на отм. ...» (модуль vor_height_injector).

Запуск:
    python run_web_with_height.py
    # или
    uvicorn run_web_with_height:app --host 0.0.0.0 --port 8050
"""

from __future__ import annotations

# 1. Загружаем все патчи из run_web (DOCX-эталон, формула по отметкам, middleware)
from run_web import app  # noqa: F401

# 2. Поверх — инжекция вертикальных кабельных стояков в get_cached_vor
import vor_export_patch as _vep
import web_app as _wa
from vor_height_injector import (
    inject_vertical_cable_rows,
    _already_injected,
)


_orig_get_cached = _vep.get_cached_vor


def _resolve_pdf_paths(rel_folder: str) -> list[str]:
    """Получить абсолютные пути к PDF файлам папки через web_app._folder_files."""
    if not rel_folder:
        return []
    try:
        files = _wa._folder_files(rel_folder)
    except Exception:
        return []
    return [str(p) for p in (files or [])]


def _get_cached_with_height(app_obj, folder_id: str):
    cached = _orig_get_cached(app_obj, folder_id)
    if not cached:
        return None
    if cached.get("_height_done"):
        return cached
    aggregated = cached.get("aggregated") or []
    rel_folder = cached.get("rel_folder", "") or ""
    pdf_paths = _resolve_pdf_paths(rel_folder)
    if not pdf_paths or _already_injected(aggregated):
        new_cached = dict(cached)
        new_cached["_height_done"] = True
        return new_cached
    new_agg = inject_vertical_cable_rows(aggregated, pdf_paths, mode="split")
    new_cached = dict(cached)
    new_cached["aggregated"] = new_agg
    new_cached["_height_done"] = True
    return new_cached


_vep.get_cached_vor = _get_cached_with_height


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("run_web_with_height:app", host="0.0.0.0", port=8050,
                reload=False)
