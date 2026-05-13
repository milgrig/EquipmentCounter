"""
run_web.py — точка входа для веб-приложения с расширением VOR-экспорта.

Использование:
    python run_web.py
    # или
    uvicorn run_web:app --host 0.0.0.0 --port 8050 --reload

Стартер импортирует web_app.app, регистрирует:
  • быстрый /api/folder/{id}/export_xlsx_v2 (из кэша, формула по отметкам)
  • новый    /api/folder/{id}/export_docx (формат ДБТ-эталона, формула по отметкам)
  • POST     /api/folder/{id}/cache_vor (приём агрегата от клиента)
  • HTML-инжектор <script src="/static/vor_patch.js"> на странице "/"
"""

from __future__ import annotations

from web_app import app  # noqa: F401 — реэкспорт для uvicorn run_web:app

# 1. Подменяем простой DOCX-рендер на эталонный (ДБТ-формат)
import vor_export_patch as _vep
from vor_docx_renderer import render_vor_docx as _render_etalon
from vor_elevation_grouper import regroup_aggregate_by_elevation


def _build_docx_bytes_etalon(aggregated, rel_folder: str) -> bytes:
    """Patched: пере-агрегируем формулу по отметкам, затем рендерим минимальный ДБТ-формат.

    Шапка содержит только строки, в которых мы уверены (по образцу ВОР_ЭО.docx):
      • "Ведомость объемов работ"
      • "Основание_<раздел>"  (раздел определяется по имени папки: ЭО/ЭМ/ЭГ)
      • "Дата составления …г."
    Шифр проекта, номер захватки, имена объекта/стройки/исполнителей не
    извлекаются автоматически и не подставляются. drawing_prefix не передаётся
    (по умолчанию пуст) — в колонке "Ссылка на чертежи" остаются только номера
    листов вида "л.5-11".
    """
    # 1. Пере-группировка формулы по строительным отметкам
    aggregated = regroup_aggregate_by_elevation(aggregated, labeled=False)

    # 2. Определяем только наименование раздела (если можем)
    folder_lower = (rel_folder or "").lower()
    if "эо" in folder_lower or "освещ" in folder_lower:
        razdel_full = "Электроосвещение"
    elif "эм" in folder_lower:
        razdel_full = "Электрооборудование"
    elif "эг" in folder_lower:
        razdel_full = "Заземление"
    else:
        razdel_full = ""

    section_basis = f"Основание_{razdel_full}" if razdel_full else ""

    return _render_etalon(
        aggregated,
        rel_folder=rel_folder,
        section_basis=section_basis,
    )


_vep._build_docx_bytes = _build_docx_bytes_etalon  # подмена


# 2. Также пере-группируем XLSX (через monkey-patch агрегата на этапе экспорта).
#    Делаем это путём wrapper'а над get_cached_vor.
_orig_get_cached = _vep.get_cached_vor


def _get_cached_with_elevation(app_obj, folder_id: str):
    cached = _orig_get_cached(app_obj, folder_id)
    if not cached:
        return None
    # Пере-группируем агрегат, если ещё не пере-группирован
    if cached.get("_elevation_done"):
        return cached
    new_agg = regroup_aggregate_by_elevation(cached["aggregated"], labeled=False)
    new_cached = dict(cached)
    new_cached["aggregated"] = new_agg
    new_cached["_elevation_done"] = True
    return new_cached


_vep.get_cached_vor = _get_cached_with_elevation


# 3. middleware-инжектор скрипта
from vor_inject_middleware import install_html_inject
install_html_inject(app)

# 4. регистрация эндпоинтов экспорта
from vor_export_patch import register_vor_export
register_vor_export(app)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("run_web:app", host="0.0.0.0", port=8050, reload=False)
