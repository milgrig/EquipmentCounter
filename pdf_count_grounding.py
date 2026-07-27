# -*- coding: utf-8 -*-
"""pdf_count_grounding.py — извлечение объёмов ЭГ (заземление/УП/молниезащита)
с PDF-планов БЕЗ СО-спецификации (S011, Test2707).

Раздел ЭГ в no-СО режиме раньше не считался вовсе: точечный счётчик видел
пару позиций легенды, а измеритель трасс выдавал линии СУП как «кабель в
гофре». Этот модуль извлекает то, что реально измеримо с чертежей:

  * план заземления   → метраж полосы 40×4 (красный пунктирный контур,
                        склейка штрихов) + штучные позиции из легенды;
  * план молниезащиты → метраж круглого проводника 8 мм (зелёная
                        молниеприёмная сетка) + деривации из примечаний:
                        держатели = метраж / шаг («Шаг установки
                        держателей 1,0 м»), компенсаторы = метраж / шаг;
  * СУП / план УП     → метраж магистрали уравнивания потенциалов.

Деривационные правила земляных работ (объём траншеи на стержень и т.п.)
заимствованы из grounding_extractor.py (T064, DXF-ветка).

Ограничение: штучная номенклатура (стержни, наконечники, соединители)
надёжно извлекается только из СО; здесь она приходит из подсчёта легенды
основным пайплайном. Результат помечается как оценка по чертежам.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

# ── Классификация ЭГ-листов по имени файла/заголовку ────────────────────
_EG_SHEET_RES = [
    (re.compile(r"молниезащит", re.IGNORECASE), "lightning"),
    (re.compile(r"заземлени", re.IGNORECASE), "grounding"),
    (re.compile(r"\bСУП\b|уравниван|план\s*УП\b|-УП\b", re.IGNORECASE),
     "bonding"),
]


def classify_eg_sheet(name: str) -> str:
    """"grounding" | "lightning" | "bonding" | "" по имени листа."""
    for pat, kind in _EG_SHEET_RES:
        if pat.search(name or ""):
            return kind
    return ""


def is_eg_sheet(name: str) -> bool:
    return bool(classify_eg_sheet(name))


# ── Деривации из примечаний листа ────────────────────────────────────────
# «Шаг установки держателей 1,0м» / «шаг установки арт.294011 L=1,0м»
_STEP_RE = re.compile(
    r"шаг\s+установки[^\n]{0,40}?(\d+(?:[.,]\d+)?)\s*м",
    re.IGNORECASE,
)
# «компенсатор … шаг установки … L=20,0м»
_COMP_STEP_RE = re.compile(
    r"(?:компенсат|291090)[^\n]{0,80}?(\d+(?:[.,]\d+)?)\s*м|"
    r"(?:шаг\s+установки\s+арт\.?\s*291090)\s*L?=?(\d+(?:[.,]\d+)?)",
    re.IGNORECASE,
)
_DEFAULT_HOLDER_STEP_M = 1.0
_DEFAULT_COMPENSATOR_STEP_M = 20.0
# «бухта 110 м» / «(бухта 38 м)» — длина бухты закупаемого проводника.
_BUNDLE_RE = re.compile(r"бухт[аы]\s*(\d+(?:[.,]\d+)?)\s*м", re.IGNORECASE)
# Полилинии короче этого (м) на планах — выноски/стрелки, не магистраль.
_MIN_RUN_M = 3.0


def _bundle_len(text: str, lo: float, hi: float, default: float) -> float:
    """Длина бухты из текста листа в диапазоне [lo, hi], иначе default."""
    for m in _BUNDLE_RE.finditer(text or ""):
        try:
            v = float(m.group(1).replace(",", "."))
        except ValueError:
            continue
        if lo <= v <= hi:
            return v
    return default


def _purchase(length_m: float, bundle_m: float) -> tuple[int, float]:
    """(число бухт, закупочный метраж) с округлением вверх до бухты."""
    import math as _math
    n = max(1, int(_math.ceil(length_m / bundle_m)))
    return n, n * bundle_m


@dataclass
class EgRow:
    description: str
    unit: str
    quantity: float
    kind: str          # grounding | lightning | bonding
    source: str = ""   # geometry | derived | text


@dataclass
class EgResult:
    rows: list[EgRow] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def _measure_sheet(pdf_path: str, page_index: int, colors: tuple,
                   dash_bridge_pt: float, log=print) -> tuple[float, str]:
    """Суммарный метраж линий заданных цветов на листе (м, масштаб)."""
    from pdf_cables_by_method import measure_cables
    rep = measure_cables(pdf_path, page_index=page_index,
                         colors=colors, dash_bridge_pt=dash_bridge_pt)
    total = sum(rep.totals_m.values())
    return total, rep.scale_source


def _building_bbox(pdf_path: str, page_index: int) -> tuple | None:
    """Габарит здания: 5-95 перцентиль концов серых архитектурных линий.

    Стены/перегородки на CAD-экспортах рисуются серым (~0.5, lw 0.36);
    их массив очерчивает корпус. Контур заземлителя прокладывается в
    1 м СНАРУЖИ фундамента — полилинии за габаритом здания относятся к
    наружному заземлителю, внутри — к магистрали УП.
    """
    import pdfplumber
    xs: list[float] = []
    ys: list[float] = []
    try:
        with pdfplumber.open(pdf_path) as doc:
            page = doc.pages[page_index]
            for ln in (page.lines or []):
                c = ln.get("stroking_color")
                if (isinstance(c, (tuple, list)) and len(c) == 3
                        and all(0.4 <= float(v) <= 0.6 for v in c)):
                    xs.extend((ln["x0"], ln["x1"]))
                    ys.extend((ln["top"], ln["bottom"]))
    except Exception:  # noqa: BLE001
        return None
    if len(xs) < 200:
        return None
    xs.sort(); ys.sort()
    i5, i95 = int(len(xs) * 0.05), int(len(xs) * 0.95) - 1
    return (xs[i5], ys[i5], xs[i95], ys[i95])


def _split_outer_inner(pdf_path: str, page_index: int, colors: tuple,
                       dash_bridge_pt: float) -> tuple[float, float, str]:
    """(метраж СНАРУЖИ здания, метраж ВНУТРИ, масштаб) для цветных линий."""
    from pdf_cables_by_method import measure_cables
    rep = measure_cables(pdf_path, page_index=page_index,
                         colors=colors, dash_bridge_pt=dash_bridge_pt)
    bbox = _building_bbox(pdf_path, page_index)
    mm = rep.scale_mm_per_pt
    outer = inner = 0.0
    for pl in rep.polylines:
        length_m = pl.length_pt * mm / 1000.0
        if bbox is None:
            inner += length_m
            continue
        x0, y0, x1, y1 = bbox
        pts_in = sum(1 for px, py in pl.points
                     if x0 <= px <= x1 and y0 <= py <= y1)
        if pl.points and pts_in / len(pl.points) < 0.3:
            # Снаружи здания выносок нет — берём всё, включая обрезки
            # пунктирного контура на углах.
            outer += length_m
        elif length_m >= _MIN_RUN_M:
            # Внутри фильтруем короткие полилинии: выноски/стрелки.
            inner += length_m
    return outer, inner, rep.scale_source


def _sheet_text(pdf_path: str, page_index: int) -> str:
    import pdfplumber
    try:
        with pdfplumber.open(pdf_path) as doc:
            if page_index < len(doc.pages):
                return doc.pages[page_index].extract_text() or ""
    except Exception:  # noqa: BLE001
        pass
    return ""


def _holder_step(text: str) -> float:
    m = _STEP_RE.search(text)
    if m:
        try:
            v = float(m.group(1).replace(",", "."))
            if 0.2 <= v <= 5.0:
                return v
        except ValueError:
            pass
    return _DEFAULT_HOLDER_STEP_M


def extract_eg_quantities(pdf_path: str, sheet_name: str,
                          page_index: int = 0, log=print) -> EgResult:
    """Извлечь измеримые объёмы ЭГ с одного листа."""
    res = EgResult()
    kind = classify_eg_sheet(sheet_name)
    if not kind:
        return res
    text = _sheet_text(pdf_path, page_index)

    if kind == "lightning":
        # Молниеприёмная сетка — зелёные линии по кровле.
        length_m, scale_src = _measure_sheet(
            pdf_path, page_index, colors=("green",), dash_bridge_pt=8.0,
            log=log)
        if length_m > 1:
            res.rows.append(EgRow(
                "Прокладка круглого проводника по кровле диаметром 8 мм "
                "(молниеприёмная сетка, по чертежу)",
                "м", round(length_m, 1), kind, "geometry"))
            _bundle = _bundle_len(text, 50, 200, 110.0)
            _n, _pm = _purchase(length_m, _bundle)
            res.rows.append(EgRow(
                f"Проводник круглый 8 мм, гор. цинк — закупка "
                f"{_n} бухт × {_bundle:g} м",
                "м", _pm, kind, "derived"))
            step = _holder_step(text)
            res.rows.append(EgRow(
                f"Держатель для круглого проводника (шаг {step:g} м, "
                f"расчёт по метражу сетки)",
                "шт", int(round(length_m / step)), kind, "derived"))
            res.rows.append(EgRow(
                "Компенсатор теплового расширения "
                f"(шаг {_DEFAULT_COMPENSATOR_STEP_M:g} м, расчёт)",
                "шт", int(round(length_m / _DEFAULT_COMPENSATOR_STEP_M)),
                kind, "derived"))
            res.notes.append(
                f"молниезащита: сетка {length_m:.1f} м ({scale_src})")

    elif kind == "grounding":
        # Красный пунктир на плане заземления — это ДВЕ системы: наружный
        # контур заземлителя (в 1 м от фундамента, в траншее) и внутренняя
        # магистраль УП по стенам. Разделяем по габариту здания.
        outer_m, inner_m, scale_src = _split_outer_inner(
            pdf_path, page_index, colors=("red",), dash_bridge_pt=14.0)
        if outer_m > 1:
            res.rows.append(EgRow(
                "Прокладка горизонтального заземлителя по периметру: "
                "полоса 40х4 мм, горячее цинкование (по чертежу)",
                "м", round(outer_m, 1), kind, "geometry"))
            _bundle = _bundle_len(text, 20, 50, 38.0)
            _n, _pm = _purchase(outer_m, _bundle)
            res.rows.append(EgRow(
                f"Плоский проводник 40х4 мм, гор. цинк — закупка "
                f"{_n} бухт × {_bundle:g} м",
                "м", _pm, kind, "derived"))
            # Земляные работы: траншея ~0.35 м³/м вдоль полосы (правило
            # выверено по эталону АБК-1: 229 м × 0.35 = 80.2 ≈ 81 м³).
            res.rows.append(EgRow(
                "Разработка грунта для прокладки горизонтального "
                "заземлителя (расчёт: 0.35 м³/м)",
                "м3", round(outer_m * 0.35, 1), kind, "derived"))
            res.rows.append(EgRow(
                "Засыпка траншеи (под горизонтальное заземление)",
                "м3", round(outer_m * 0.35, 1), kind, "derived"))
        if inner_m > 1:
            res.rows.append(EgRow(
                "Прокладка магистрали системы уравнивания потенциалов: "
                "полоса 40х4 мм (по чертежу)",
                "м", round(inner_m, 1), "bonding", "geometry"))
            res.rows.append(EgRow(
                "Монтаж держателя плоского проводника до 40 мм "
                "(шаг 1 м, расчёт по метражу магистрали)",
                "шт", int(round(inner_m / 1.0)), "bonding", "derived"))
        res.notes.append(
            f"заземление: контур {outer_m:.1f} м, магистраль УП внутри "
            f"{inner_m:.1f} м ({scale_src})")

    elif kind == "bonding":
        # Магистраль СУП — полоса 40×4 по помещениям (красные линии).
        length_m, scale_src = _measure_sheet(
            pdf_path, page_index, colors=("red", "magenta"),
            dash_bridge_pt=14.0, log=log)
        if length_m > 1:
            res.rows.append(EgRow(
                "Прокладка магистрали системы уравнивания потенциалов: "
                "полоса 40х4 мм (по чертежу)",
                "м", round(length_m, 1), kind, "geometry"))
            res.rows.append(EgRow(
                "Монтаж держателя плоского проводника до 40 мм "
                "(шаг 1 м, расчёт по метражу магистрали)",
                "шт", int(round(length_m / 1.0)), kind, "derived"))
            res.notes.append(
                f"СУП: магистраль {length_m:.1f} м ({scale_src})")

    return res
