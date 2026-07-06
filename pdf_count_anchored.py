# -*- coding: utf-8 -*-
"""pdf_count_anchored.py — подсчёт глифов оборудования по текстовым якорям (S011).

Проблема: на планах освещения текстовый маркер ставится на ГРУППУ
одинаковых светильников, а не на каждый экземпляр (ГПК-3 лист 006:
символ «5» — 5 подписей при 11 светильниках в DXF ground truth), поэтому
текстовый счётчик системно недосчитывает. Шаблон из легенды тоже не
спасает: в ячейке «Обозначение» нарисован маркер-цифра, а не глиф
светильника с плана — NCC-матчинг даёт нули на любых масштабах.

Решение — учиться виду глифа у найденных текстовых маркеров:
  1. рендерим страницу, строим цветовые маски (красная/синяя графика);
  2. связные компоненты маски = кандидаты-глифы;
  3. для каждого маркированного символа легенды берём компоненты около
     его текстовых якорей как эталоны (размер + бинарная маска);
  4. находим на странице все компоненты, похожие на эталон
     (размер ± допуск, IoU масок);
  5. каждый найденный компонент относим к символу БЛИЖАЙШЕГО якоря —
     подпись стоит рядом со своей группой, поэтому визуально одинаковые
     варианты («3» ARCTIC TH и «5» ARCTIC) разводятся геометрически.

Используется как дополнение к pdf_count_text: итог = max(text, anchored).

Известное ограничение: линейные светильники (ARCTIC 1200 и т.п.), корпус
которых лежит НА ОДНОЙ ПРЯМОЙ с кабельной трассой, разрушаются вычитанием
длинных линий (корпус+кабель образуют непрерывный прямой штрих длиннее
LINE_KERNEL_PX) — для них счётчик возвращает 0 и итог остаётся текстовым.
Лечится только геометрическим разбором (DXF-ветка) либо детектором
end-cap-пар; зафиксировано на ГПК-3 006 (символы «3», «5»).
"""
from __future__ import annotations

import math
import os
from dataclasses import dataclass, field

import cv2
import numpy as np
import pdfplumber

from pdf_count_visual import (
    RENDER_DPI,
    _build_exclusion_zones,
    _pt_excluded,
    _render_page,
)

# Радиус (pt), в котором ищем глиф-эталон вокруг текстового якоря.
ANCHOR_RADIUS_PT = 35.0
# Допуск на размер компонента относительно медианы эталонов.
SIZE_TOL = 0.35
# Минимальное IoU дилатированных бинарных масок компонент/эталон.
# Порог калиброван на ГПК-3 006: истинные экземпляры дают 0.35-0.87,
# посторонние объекты того же размера — <0.3.
MASK_IOU_MIN = 0.32
# Площадь компонента (px при RENDER_DPI), вне диапазона — шум или заливка.
COMP_MIN_AREA_PX = 40
COMP_MAX_AREA_PX = 60000
# Насколько дальше якоря может лежать назначаемый компонент (pt). Группы
# светильников тянутся вдоль помещений — радиус щедрый, но конечный.
ASSIGN_MAX_DIST_PT = 400.0
_MASK_SIZE = 32


@dataclass
class AnchoredResult:
    counts: dict[str, int] = field(default_factory=dict)   # symbol -> count
    exemplars: dict[str, int] = field(default_factory=dict)  # symbol -> n эталонов
    components: int = 0          # всего компонент-кандидатов на листе
    notes: list[str] = field(default_factory=list)


# Длина ядра (px при RENDER_DPI) для вычитания прямых кабельных линий.
# Трассы тянутся сотнями px; самый крупный глиф светильника (ARCTIC 1200 мм
# при 1:100 и 200 DPI) ≈ 95 px, поэтому 121 px режет кабели, но не глифы.
LINE_KERNEL_PX = 121


def _remove_long_lines(mask: np.ndarray) -> np.ndarray:
    """Вычесть длинные прямые ОСЕВЫЕ штрихи (кабельные трассы) из маски.

    Без этого вся синяя/красная графика листа сливается через провода в
    один гигантский компонент, и глифы светильников не отделяются.

    Только сепарабельные rect-ядра (121×1 и 1×121): дёшево даже на
    60-Мпикс рендерах. Диагональные ядра (np.eye(121)) из первой версии
    стоили O(N·k²) — минуты CPU на страницу × 8 воркеров и подвесили
    прод-сервер; при этом на контрольном листе ГПК-3 006 они не меняли
    ни одного счётчика. Наклонные трассы не вычитаются: глифы на них
    остаются слитыми и просто не учитываются якорным счётчиком
    (деградация к текстовому счёту, не ошибка).
    """
    out = mask.copy()
    k = LINE_KERNEL_PX
    grown = cv2.dilate(mask, np.ones((3, 3), np.uint8))
    for ker in (cv2.getStructuringElement(cv2.MORPH_RECT, (k, 1)),
                cv2.getStructuringElement(cv2.MORPH_RECT, (1, k))):
        lines = cv2.morphologyEx(grown, cv2.MORPH_OPEN, ker)
        out = cv2.subtract(out, cv2.dilate(lines, np.ones((5, 5), np.uint8)))
    return out


def _color_masks(bgr: np.ndarray) -> dict[str, np.ndarray]:
    """Маски красной/синей графики (терпимые к антиалиасингу пороги),
    с вычтенными кабельными линиями."""
    b = bgr[:, :, 0].astype(np.int16)
    g = bgr[:, :, 1].astype(np.int16)
    r = bgr[:, :, 2].astype(np.int16)
    red = ((r > 140) & (r - g > 60) & (r - b > 60)).astype(np.uint8) * 255
    blue = ((b > 140) & (b - r > 60) & (b - g > 40)).astype(np.uint8) * 255
    kernel = np.ones((3, 3), np.uint8)
    out = {}
    for name, m in (("red", red), ("blue", blue)):
        m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, kernel)
        m = _remove_long_lines(m)
        m = cv2.morphologyEx(m, cv2.MORPH_CLOSE, kernel)
        out[name] = m
    return out


# Радиус (px) склейки фрагментов одного глифа: вычитание кабельной линии,
# проходящей сквозь ряд светильников, режет каждый глиф на части.
CLUSTER_GAP_PX = 9


def _glyph_clusters(mask: np.ndarray) -> list[dict]:
    """Кластеры-кандидаты глифов: фрагменты маски, склеенные дилатацией.

    Возвращает [{bbox, cx, cy, area, mask32}] — bbox в px исходной маски,
    mask32 — бинарная форма кластера (по НЕдилатированной маске),
    приведённая к _MASK_SIZE×_MASK_SIZE.
    """
    k = 2 * CLUSTER_GAP_PX + 1
    grown = cv2.dilate(mask, np.ones((k, k), np.uint8))
    n, labels_img, stats, cents = cv2.connectedComponentsWithStats(grown, 8)
    out = []
    for i in range(1, n):
        gx, gy, gw, gh, _ = stats[i]
        # Снимаем поля дилатации, чтобы размер был размером глифа.
        x = gx + CLUSTER_GAP_PX
        y = gy + CLUSTER_GAP_PX
        w = max(1, gw - 2 * CLUSTER_GAP_PX)
        h = max(1, gh - 2 * CLUSTER_GAP_PX)
        crop = mask[max(0, y):y + h, max(0, x):x + w]
        area = int(np.count_nonzero(crop))
        if not (COMP_MIN_AREA_PX <= area <= COMP_MAX_AREA_PX):
            continue
        m32 = cv2.resize((crop > 0).astype(np.uint8),
                         (_MASK_SIZE, _MASK_SIZE),
                         interpolation=cv2.INTER_AREA)
        # Контуры глифов тонкие: без дилатации нормированные маски двух
        # экземпляров одного глифа почти не пересекаются (IoU ~0.05).
        mask32 = cv2.dilate(m32, np.ones((5, 5), np.uint8)) > 0
        # Канонизация ориентации: глиф может стоять горизонтально или
        # вертикально — приводим к альбомной (w >= h), чтобы эталоны
        # одного символа в разных поворотах агрегировались корректно.
        cw, ch = float(w), float(h)
        if ch > cw:
            cw, ch = ch, cw
            mask32 = np.rot90(mask32)
        out.append({
            "bbox": (x, y, w, h), "cw": cw, "ch": ch,
            "cx": float(cents[i][0]), "cy": float(cents[i][1]),
            "area": area, "mask32": mask32,
        })
    return out


def _mask_iou(a: np.ndarray, b: np.ndarray) -> float:
    inter = np.count_nonzero(a & b)
    union = np.count_nonzero(a | b)
    return inter / union if union else 0.0


def _size_match(w: float, h: float, ew: float, eh: float) -> bool:
    """Размер совпадает с эталоном (с учётом поворота на 90°)."""
    def ok(a, b):
        return abs(a - b) <= SIZE_TOL * max(b, 1.0)
    return (ok(w, ew) and ok(h, eh)) or (ok(w, eh) and ok(h, ew))


def count_by_marker_anchors(
    pdf_path: str,
    legend_result,
    text_result,
    page_index: int | None = None,
    render_dpi: int = RENDER_DPI,
) -> AnchoredResult:
    """Посчитать глифы оборудования, взяв найденные текстовые маркеры
    как якоря-эталоны.

    ``text_result`` — CountResult из pdf_count_text (нужны positions).
    Обрабатываются только маркированные символы с цветом red/blue из
    легенды (категория «светильник» и прочие точечные глифы).
    """
    res = AnchoredResult()
    # Аварийный рубильник: VOR_ANCHORED=0 полностью выключает якорный
    # подсчёт (например, на слабом сервере).
    if os.environ.get("VOR_ANCHORED", "1") != "1":
        res.notes.append("disabled by VOR_ANCHORED=0")
        return res
    page_idx = legend_result.page_index if page_index is None else page_index

    # Якоря по символам: symbol -> [(x_pt, y_pt), ...]
    anchors: dict[str, list[tuple[float, float]]] = {}
    for p in getattr(text_result, "positions", []):
        anchors.setdefault(p.symbol, []).append((p.x, p.y))

    # Символы, для которых есть смысл считать: цвет известен, якоря есть.
    sym_color: dict[str, str] = {}
    for it in legend_result.items:
        s = (it.symbol or "").strip()
        c = (getattr(it, "color", "") or "").lower()
        if s and c in ("red", "blue") and s in anchors:
            sym_color[s] = c
    if not sym_color:
        return res

    bgr = _render_page(pdf_path, page_idx, dpi=render_dpi)
    # Предохранитель по размеру: сверхбольшие рендеры (нестандартные листы)
    # съедают RAM в параллельных воркерах — пропускаем, итог остаётся
    # текстовым (деградация, не зависание).
    n_mpx = (bgr.shape[0] * bgr.shape[1]) / 1e6
    if n_mpx > 120:
        res.notes.append(f"skipped: render {n_mpx:.0f} Mpx > 120 Mpx")
        return res
    px_per_pt = render_dpi / 72.0
    masks = _color_masks(bgr)
    del bgr

    # Зоны исключения (легенда, штамп, оси) в pt.
    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return res
        pp = pdf.pages[page_idx]
        lb = legend_result.legend_bbox if page_idx == legend_result.page_index \
            else None
        zones = _build_exclusion_zones(pp, pp.lines or [], lb)

    for color in ("red", "blue"):
        syms = [s for s, c in sym_color.items() if c == color]
        if not syms:
            continue
        clusters = [c for c in _glyph_clusters(masks[color])
                    if not _pt_excluded(c["cx"] / px_per_pt,
                                        c["cy"] / px_per_pt, zones)]
        res.components += len(clusters)

        # Эталоны per symbol: крупнейший кластер возле каждого якоря.
        exemplars: dict[str, list[dict]] = {}
        r_px = ANCHOR_RADIUS_PT * px_per_pt
        for s in syms:
            for ax_pt, ay_pt in anchors[s]:
                ax, ay = ax_pt * px_per_pt, ay_pt * px_per_pt
                near = [c for c in clusters
                        if abs(c["cx"] - ax) <= r_px + c["bbox"][2] / 2
                        and abs(c["cy"] - ay) <= r_px + c["bbox"][3] / 2]
                if not near:
                    continue
                best = max(near, key=lambda c: c["area"])
                exemplars.setdefault(s, []).append(best)
        for s, ex in exemplars.items():
            res.exemplars[s] = len(ex)
        active = [s for s in syms if exemplars.get(s)]
        if not active:
            continue

        # Медианный (канонический) размер и маска эталонов per symbol.
        # Эталоны-выбросы (склейка двух глифов, обрезок) отсеиваются по
        # отклонению от медианы размеров.
        ex_feat: dict[str, tuple[float, float, np.ndarray]] = {}
        for s in active:
            ex = exemplars[s]
            ws = sorted(c["cw"] for c in ex)
            hs = sorted(c["ch"] for c in ex)
            mw, mh = ws[len(ws) // 2], hs[len(hs) // 2]
            trimmed = [c for c in ex
                       if abs(c["cw"] - mw) <= 0.4 * mw
                       and abs(c["ch"] - mh) <= 0.4 * max(mh, 1.0)]
            if trimmed:
                ex = trimmed
            ws = sorted(c["cw"] for c in ex)
            hs = sorted(c["ch"] for c in ex)
            ew, eh = float(ws[len(ws) // 2]), float(hs[len(hs) // 2])
            acc = np.zeros((_MASK_SIZE, _MASK_SIZE), dtype=np.float32)
            for c in ex:
                acc += c["mask32"].astype(np.float32)
            # Мягкое голосование (≥25% эталонов): жёсткое большинство на
            # тонких контурах оставляло пустую маску.
            ex_mask = acc >= max(1.0, len(ex) / 4.0)
            ex_feat[s] = (ew, eh, ex_mask)
            res.notes.append(
                f"{s}: эталонов {len(exemplars[s])} (после отсева {len(ex)}), "
                f"размер ~{ew:.0f}x{eh:.0f}px")

        # Классификация кластеров: похожие на эталон → символ ближайшего якоря.
        assign_r_px = ASSIGN_MAX_DIST_PT * px_per_pt
        for c in clusters:
            cand: list[str] = []
            for s in active:
                ew, eh, emask = ex_feat[s]
                if not _size_match(c["cw"], c["ch"], ew, eh):
                    continue
                # Маски уже канонизированы по ориентации; rot90 покрывает
                # зеркальные постановки глифа.
                iou = max(_mask_iou(c["mask32"], emask),
                          _mask_iou(np.rot90(np.rot90(c["mask32"])), emask))
                if iou >= MASK_IOU_MIN:
                    cand.append(s)
            if not cand:
                continue
            # Ближайший якорь среди символов-кандидатов.
            best_s, best_d = None, assign_r_px
            for s in cand:
                for ax_pt, ay_pt in anchors[s]:
                    d = math.hypot(c["cx"] - ax_pt * px_per_pt,
                                   c["cy"] - ay_pt * px_per_pt)
                    if d < best_d:
                        best_d, best_s = d, s
            if best_s is not None:
                res.counts[best_s] = res.counts.get(best_s, 0) + 1

    return res
