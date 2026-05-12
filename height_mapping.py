#!/usr/bin/env python3
"""
height_mapping.py
=================

Модуль для правильного расчёта «монтажной высоты» светильников и кабелей
по DXF/DWG-чертежам.

Предыстория (см. chat_history.json):
------------------------------------
Эталон ВОР (ВОР ЭО, Захватка 3_ГПК.docx) классифицирует светильники по
категориям «до 5 / от 5 до 13 / от 13 до 20 / от 20 до 35 метров»
НЕ по отметке пола этажа, где они установлены, а по **монтажной высоте
от уровня пола до потолка помещения**, т.е. фактически — по
**отметке потолка этажа**, к которому светильник прикреплён.

В ГПК-3 отметки этажей:
    0.000, +4.200, +7.800, +9.000, +13.800, +18.600, +23.400, +28.200

Для светильника на полу 0.000 потолок на +4.200 → 4.2 м → «до 5 метров» ✓
Для светильника на полу +4.200 потолок на +7.800 или +9.000 → 7.8-9.0 м
    → «от 5 до 13 метров» ✓
Для светильника на полу +28.200 (последний этаж) потолок определяется
    типовой высотой этажа ~4.8 м → 33.0 → «от 20 до 35 метров» ✓

Старая функция vor_generator.elevation_to_height брала отметку пола
как есть, что давало 217/85/146/138 вместо эталонных 117/182/81/206.

Использование:
--------------
    from height_mapping import build_floor_to_ceiling_map, floor_to_height_category

    # Получить список отметок из имён файлов (или передать свой)
    floors = [0.0, 4.2, 7.8, 9.0, 13.8, 18.6, 23.4, 28.2]
    mapping = build_floor_to_ceiling_map(floors)

    category = floor_to_height_category(4.2, mapping)
    # → "от 5 до 13 метров"

    # Через строковую «отметку» из имени файла
    from height_mapping import elevation_str_to_category
    category = elevation_str_to_category("+13-800", floors)
    # → "от 13 до 20 метров"

Этот модуль — чистый, без побочных эффектов, без зависимостей от
equipment_counter / vor_generator.  Его можно использовать как в
генерации ВОР, так и в постпроцессоре.
"""
from __future__ import annotations

import re
from typing import Iterable

# ---------------------------------------------------------------------------
# Константы
# ---------------------------------------------------------------------------

# Типовая высота этажа (используется для верхнего этажа, у которого нет
# вышестоящего пола).  4.8 м взято как усреднённое значение для
# индустриальных/промышленных зданий типа ГПК.  Для офисных 2.8-3.0 м.
DEFAULT_FLOOR_HEIGHT_M = 4.8

# Границы категорий монтажной высоты (в метрах)
HEIGHT_BOUNDARIES = [
    (5.0, "до 5 метров"),
    (13.0, "от 5 до 13 метров"),
    (20.0, "от 13 до 20 метров"),
    (float("inf"), "от 20 до 35 метров"),
]

# Регулярное выражение для извлечения отметки из имени файла.
# Примеры:
#   "005 - План освещения на отм- 0-000.dwg"          → 0.000
#   "006 - План освещения на отм- +4-200.dwg"         → 4.200
#   "007 - План освещения на отм- +7-800, +9-000.dwg" → 7.800 и 9.000
_ELEV_TOKEN_RE = re.compile(r"[+-]?\d{1,3}[-.,]\d{1,3}")

# Корректировка значений: «4-200» (формат DWG имени файла) это «4.200 м»;
# «4200» в миллиметрах → 4.2 м.
_MM_THRESHOLD = 100.0


# ---------------------------------------------------------------------------
# Парсинг отметок
# ---------------------------------------------------------------------------


def parse_elevation_token(token: str) -> float | None:
    """
    Парсит одиночный токен отметки ('+4-200', '4.200', '+7800').

    Возвращает значение в метрах или None, если не удалось распарсить.

    >>> parse_elevation_token("+4-200")
    4.2
    >>> parse_elevation_token("0-000")
    0.0
    >>> parse_elevation_token("+13.800")
    13.8
    >>> parse_elevation_token("+7800")  # миллиметры
    7.8
    """
    s = token.strip()
    if s.startswith("+"):
        s = s[1:]
    # Унифицируем разделитель: '-' и ',' → '.'
    # Но только один раз (между целой и дробной частями)
    s = s.replace(",", ".").replace("-", ".")
    # Если оказалось несколько точек (например, "-13.800" → "..13.800"),
    # объединяем лишние точки в минус
    if s.startswith(".."):
        s = "-" + s[2:]
    try:
        value = float(s)
    except ValueError:
        return None
    # Значения >= 100 по модулю считаем миллиметрами
    if abs(value) >= _MM_THRESHOLD:
        value = value / 1000.0
    return value


def extract_elevations_from_filename(filename: str) -> list[float]:
    """
    Извлекает все отметки из имени файла.

    >>> extract_elevations_from_filename("005 - План освещения на отм- 0-000.dwg")
    [0.0]
    >>> extract_elevations_from_filename("007 - План освещения на отм- +7-800, +9-000.dwg")
    [7.8, 9.0]
    >>> extract_elevations_from_filename("003, 004 - Схемы ЩО, ЩАО.dwg")
    []
    """
    # Отметки ищем только после слова "отм" / "отм." / "отметк" — чтобы не
    # путать с префиксом "003, 004" (это номера листов, не отметки).
    m = re.search(r"отм", filename, re.IGNORECASE)
    if not m:
        return []
    tail = filename[m.end():]
    tokens = _ELEV_TOKEN_RE.findall(tail)
    result: list[float] = []
    for t in tokens:
        v = parse_elevation_token(t)
        if v is not None and -50.0 <= v <= 200.0:  # разумный диапазон
            result.append(v)
    return result


# ---------------------------------------------------------------------------
# Карта floor → ceiling
# ---------------------------------------------------------------------------


def build_floor_to_ceiling_map(
    floor_elevations: Iterable[float],
    last_floor_height: float = DEFAULT_FLOOR_HEIGHT_M,
) -> dict[float, float]:
    """
    По списку отметок пола всех этажей здания возвращает карту
    «отметка пола → отметка потолка» (в метрах).

    Для верхнего (последнего) этажа потолок вычисляется как
    `floor + last_floor_height`, т.к. пола «над ним» нет.

    >>> m = build_floor_to_ceiling_map([0.0, 4.2, 9.0, 13.8, 18.6])
    >>> m[0.0]
    4.2
    >>> m[4.2]
    9.0
    >>> round(m[18.6], 1)
    23.4
    """
    sorted_floors = sorted(set(floor_elevations))
    mapping: dict[float, float] = {}
    for i, floor in enumerate(sorted_floors):
        if i + 1 < len(sorted_floors):
            mapping[floor] = sorted_floors[i + 1]
        else:
            mapping[floor] = round(floor + last_floor_height, 3)
    return mapping


# ---------------------------------------------------------------------------
# Классификация по категориям
# ---------------------------------------------------------------------------


def height_to_category(height_m: float) -> str:
    """
    Классифицирует монтажную высоту (в метрах) по четырём категориям ВОР.

    >>> height_to_category(3.0)
    'до 5 метров'
    >>> height_to_category(5.0)       # граница — включается в верхнюю
    'от 5 до 13 метров'
    >>> height_to_category(13.8)
    'от 13 до 20 метров'
    >>> height_to_category(30.0)
    'от 20 до 35 метров'
    """
    for boundary, label in HEIGHT_BOUNDARIES:
        if height_m < boundary:
            return label
    return HEIGHT_BOUNDARIES[-1][1]


def floor_to_height_category(
    floor_elev: float,
    floor_to_ceiling: dict[float, float],
    tolerance: float = 0.3,
) -> str:
    """
    Преобразует отметку пола этажа в категорию монтажной высоты,
    используя карту floor→ceiling.

    Допускает небольшую погрешность (tolerance) при поиске floor в карте,
    чтобы «+7.8» матчилось с ключом «7.800».

    >>> m = build_floor_to_ceiling_map([0.0, 4.2, 7.8, 9.0, 13.8, 18.6, 23.4, 28.2])
    >>> floor_to_height_category(0.0, m)
    'до 5 метров'
    >>> floor_to_height_category(4.2, m)
    'от 5 до 13 метров'
    >>> floor_to_height_category(13.8, m)
    'от 13 до 20 метров'
    >>> floor_to_height_category(28.2, m)
    'от 20 до 35 метров'
    """
    # Точное совпадение
    if floor_elev in floor_to_ceiling:
        ceiling = floor_to_ceiling[floor_elev]
        return height_to_category(ceiling)
    # Нечёткое совпадение
    for f, c in floor_to_ceiling.items():
        if abs(f - floor_elev) <= tolerance:
            return height_to_category(c)
    # Фоллбек: использовать отметку пола как есть
    return height_to_category(floor_elev)


def elevation_str_to_category(
    elev_str: str,
    all_floors: list[float],
    last_floor_height: float = DEFAULT_FLOOR_HEIGHT_M,
) -> str | None:
    """
    Удобная обёртка: принимает строку-отметку ("+4-200") и список всех
    отметок захватки, возвращает категорию монтажной высоты.

    >>> all_f = [0.0, 4.2, 7.8, 9.0, 13.8, 18.6, 23.4, 28.2]
    >>> elevation_str_to_category("+4-200", all_f)
    'от 5 до 13 метров'
    >>> elevation_str_to_category("+28-200", all_f)
    'от 20 до 35 метров'
    """
    elev = parse_elevation_token(elev_str)
    if elev is None:
        return None
    mapping = build_floor_to_ceiling_map(all_floors, last_floor_height)
    return floor_to_height_category(elev, mapping)


# ---------------------------------------------------------------------------
# Утилита: сбор всех отметок из списка имён файлов
# ---------------------------------------------------------------------------


def collect_floors_from_filenames(filenames: Iterable[str]) -> list[float]:
    """
    Собирает упорядоченный уникальный список отметок из имён файлов.

    >>> files = [
    ...     "005 - План освещения на отм- 0-000.dwg",
    ...     "006 - План освещения на отм- +4-200.dwg",
    ...     "007 - План освещения на отм- +7-800, +9-000.dwg",
    ... ]
    >>> collect_floors_from_filenames(files)
    [0.0, 4.2, 7.8, 9.0]
    """
    all_elevs: set[float] = set()
    for fn in filenames:
        for e in extract_elevations_from_filename(fn):
            all_elevs.add(round(e, 3))
    return sorted(all_elevs)


# ---------------------------------------------------------------------------
# Применение: перераспределить counts_by_height в готовом отчёте
# ---------------------------------------------------------------------------


def remap_counts_by_floor(
    per_floor_counts: dict[float, int],
    floor_to_ceiling: dict[float, float],
) -> dict[str, int]:
    """
    Имея словарь {отметка_пола: количество}, возвращает
    {категория_высоты: сумма_количества}.

    >>> mapping = build_floor_to_ceiling_map([0.0, 4.2, 9.0, 13.8, 18.6, 23.4, 28.2])
    >>> per_floor = {0.0: 79, 4.2: 64, 9.0: 52, 13.8: 43, 18.6: 43, 23.4: 49, 28.2: 38}
    >>> out = remap_counts_by_floor(per_floor, mapping)
    >>> sorted(out.items())  # doctest: +NORMALIZE_WHITESPACE
    [('до 5 метров', 79), ('от 13 до 20 метров', 95),
     ('от 20 до 35 метров', 130), ('от 5 до 13 метров', 64)]
    """
    result: dict[str, int] = {}
    for floor, count in per_floor_counts.items():
        cat = floor_to_height_category(floor, floor_to_ceiling)
        result[cat] = result.get(cat, 0) + count
    return result


# ---------------------------------------------------------------------------
# Smoke-test
# ---------------------------------------------------------------------------


def _selftest() -> None:
    """Мини-самопроверка на данных ГПК-3."""
    # Отметки из имён файлов ГПК-3 (захватка 3)
    floors = [0.0, 4.2, 7.8, 9.0, 13.8, 18.6, 23.4, 28.2]
    mapping = build_floor_to_ceiling_map(floors)

    expected = {
        0.0: "до 5 метров",          # 0 → 4.2
        4.2: "от 5 до 13 метров",    # 4.2 → 7.8
        7.8: "от 5 до 13 метров",    # 7.8 → 9.0
        9.0: "от 13 до 20 метров",   # 9.0 → 13.8
        13.8: "от 13 до 20 метров",  # 13.8 → 18.6
        18.6: "от 20 до 35 метров",  # 18.6 → 23.4
        23.4: "от 20 до 35 метров",  # 23.4 → 28.2
        28.2: "от 20 до 35 метров",  # 28.2 → 33.0 (top)
    }

    for floor, expect in expected.items():
        got = floor_to_height_category(floor, mapping)
        status = "OK" if got == expect else "FAIL"
        print(f"  [{status}] floor=+{floor:5.3f}  ceiling=+{mapping[floor]:5.3f}  "
              f"→ {got}  (expected: {expect})")


if __name__ == "__main__":
    print("height_mapping.py — самопроверка на отметках ГПК-3:\n")
    _selftest()
    print("\nDoctests:")
    import doctest
    results = doctest.testmod(verbose=False)
    print(f"  attempted={results.attempted}, failed={results.failed}")
