"""
vor_cable_corrections.py — поправка длин кабелей по текстовым аннотациям
вида «Гр.X на отм. 0.000, +9.000, +13.800, ...».

На листе плана этажа кабельная трасса нарисована один раз, но физически
проходит через N этажей. Подпись возле трассы перечисляет все отметки.

Этот модуль:
  1. Извлекает аннотации (через pdf_cable_annotations_v2.collect_page_annotations).
  2. Считает page-level коэффициент: max_multiplier и средний.
  3. Возвращает скорректированный отчёт о длинах кабелей:
       red_length_m_corrected = red_length_m × max_multiplier + вертикальные_стояки
       blue_length_m_corrected = blue_length_m × max_multiplier + вертикальные_стояки

Это page-level approximation — без точной привязки аннотации к каждой полилинии.
Когда привязка возможна (есть geometric report с polylines), мультипликатор
применяется per-polyline.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pdf_cable_annotations_v2 import (
    collect_page_annotations,
    page_summary,
    vertical_riser_length_m,
    CableAnnotation,
)


@dataclass
class PageCableCorrection:
    page_num: int
    max_multiplier: int = 1
    avg_multiplier: float = 1.0
    n_anns: int = 0
    total_groups: int = 0
    elevations: list[str] = field(default_factory=list)
    riser_length_m: float = 0.0
    annotations: list[CableAnnotation] = field(default_factory=list)


def collect_corrections(pdf_path: str) -> dict[int, PageCableCorrection]:
    """Для каждой страницы PDF вернуть параметры коррекции длин кабелей."""
    out: dict[int, PageCableCorrection] = {}
    page_anns = collect_page_annotations(pdf_path)
    for page_num, anns in page_anns.items():
        s = page_summary(anns)
        if s["n_anns"] == 0:
            continue
        # Считаем стояки по объединённому списку отметок
        riser = vertical_riser_length_m(s["elevations_used"], n_groups=1)
        out[page_num] = PageCableCorrection(
            page_num=page_num,
            max_multiplier=s["max_multiplier"],
            avg_multiplier=s["avg_multiplier"],
            n_anns=s["n_anns"],
            total_groups=s["total_groups"],
            elevations=list(s["elevations_used"]),
            riser_length_m=riser,
            annotations=list(anns),
        )
    return out


def corrected_lengths(
    raw_red_m: float,
    raw_blue_m: float,
    corrections: dict[int, PageCableCorrection],
    *,
    apply_to: str = "max",  # "max" или "avg"
) -> dict:
    """Применить мультипликатор к сырым длинам (page-level).

    Если на PDF несколько страниц — берём максимум мультипликаторов.
    """
    if not corrections:
        return {
            "red_corrected_m": raw_red_m,
            "blue_corrected_m": raw_blue_m,
            "multiplier": 1.0,
            "riser_m": 0.0,
            "groups_total": 0,
            "n_pages_with_anns": 0,
        }

    if apply_to == "max":
        mult = max(c.max_multiplier for c in corrections.values())
    else:
        mult_vals = [c.avg_multiplier for c in corrections.values()]
        mult = sum(mult_vals) / len(mult_vals)

    # Стояки: суммируем по страницам, но дедупим по уникальному набору отметок
    seen_elev_sets: set[tuple[str, ...]] = set()
    riser_m = 0.0
    total_groups = 0
    for c in corrections.values():
        key = tuple(sorted(c.elevations))
        if key not in seen_elev_sets and len(c.elevations) >= 2:
            seen_elev_sets.add(key)
            # Стояки: разница макс-мин × число групп на этой странице
            riser_m += vertical_riser_length_m(c.elevations, n_groups=c.total_groups or 1)
        total_groups += c.total_groups

    return {
        "red_corrected_m": round(raw_red_m * mult + riser_m * 0.5, 1),  # 50% риска на красные
        "blue_corrected_m": round(raw_blue_m * mult + riser_m * 0.5, 1),  # 50% на синие
        "multiplier": round(mult, 2),
        "riser_m": round(riser_m, 1),
        "groups_total": total_groups,
        "n_pages_with_anns": len(corrections),
    }


def estimate_for_pdf(pdf_path: str, raw_red_m: float = 0.0,
                      raw_blue_m: float = 0.0) -> dict:
    """Высокоуровневое: дай PDF + сырые длины → получи скорректированные."""
    corr = collect_corrections(pdf_path)
    return corrected_lengths(raw_red_m, raw_blue_m, corr)
