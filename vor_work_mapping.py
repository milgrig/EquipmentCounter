#!/usr/bin/env python3
"""vor_work_mapping.py -- Map equipment names to VOR work descriptions.

Maps equipment names extracted from PDF legends into standardized VOR
(Work Volume Statement) work description format used in construction
documentation.

The VOR reference files use a two-level structure:
  1. Work item:   "Монтаж светильника в подвесных потолках (до 5 м)"  qty=399
  2. Material row: "Светодиодный светильник 4000К 40Вт UNI/R EVO..."  qty=195

Our parser outputs equipment names like:
  "Светильник светодиодный UNI/R EVO (595x595) 40W OPL 840"

This module provides the mapping from (2)-style names to (1)-style work
descriptions, and also normalizes material names to match (2)-style.

Reference VOR naming conventions (from analysing real ВОР files):
  - Luminaires:  "Монтаж светильника {type} {model}"            шт
  - Emergency:   "Монтаж светильника аварийного {model}"         шт
  - Exit signs:  "Монтаж указателя аварийного освещения {model}" шт
  - Sockets:     "Монтаж розетки {specs}"                        шт
  - Switches:    "Монтаж выключателя {type} {specs}"             шт
  - Panels:      "Монтаж щита распределительного {designation}"  шт  / комп
  - Cables:      "Прокладка кабеля {brand} {section}"            м
  - Wire:        "Прокладка провода {brand} {section}"           м
  - Conduits:    "Монтаж трубы гофрированной {specs}"            м   / м.п.
  - Cable trays: "Монтаж лотка кабельного {specs}"               м   / м.п.
  - Jn. boxes:   "Монтаж коробки {type} {specs}"                шт
  - Grounding:   "Монтаж заземления {detail}"                    шт  / м
  - Sensors:     "Монтаж датчика {type}"                         шт
  - Breakers:    "Монтаж автоматического выключателя {specs}"    шт
  - Frames:      "Установка рамки {N}-постовой {specs}"          шт
  - Connection:  "Подключение жил кабеля {specs}"                шт
"""

from __future__ import annotations

import re
from dataclasses import dataclass


# ---------------------------------------------------------------------------
# Valid VOR units
# ---------------------------------------------------------------------------

VALID_UNITS = {"шт", "м", "м.п.", "комп", "м3", "каб.", "изм.", "маш/час",
               "измерение", "компл", "компл."}


@dataclass
class WorkMapping:
    """Result of mapping an equipment name to a VOR work description."""
    work_name: str        # VOR "Наименование вида работ" column
    unit: str             # Unit of measurement
    equipment_name: str   # Original equipment name (for "Доп.информация")
    category: str         # Equipment category key


# ---------------------------------------------------------------------------
# Detail extraction helpers
# ---------------------------------------------------------------------------

def _extract_detail(name: str) -> str:
    """Extract the 'detail' portion of an equipment name.

    Removes the leading generic noun (Светильник, Розетка, etc.)
    and returns the rest.
    """
    detail = re.sub(
        r"^(?:светильник|розетк[аи]|выключатель|коробк[аи]|датчик"
        r"|извещатель|указатель|кабельн\S*\s+вывод|кабель|провод(?:ник)?"
        r"|труб[аы]|лот[оа]к"
        r"|щит|короб|автомат(?:ический)?\s*(?:выключатель)?"
        r"|рамк[аи]|контактор|пускатель|реле|клеммник"
        r"|полос[аы]|стержен[ьи]|электрод"
        r"|молниезащит\S*|молниеприёмник|молниеприемник|токоотвод)\s*",
        "",
        name.strip(),
        flags=re.IGNORECASE,
    )
    return detail.strip() if detail.strip() else name.strip()


def _extract_luminaire_detail(name: str) -> str:
    """Extract luminaire detail, keeping model info.

    "Светильник светодиодный UNI/R EVO" -> "светодиодного UNI/R EVO"
    """
    detail = re.sub(
        r"^светильн\S*\s*",
        "",
        name.strip(),
        flags=re.IGNORECASE,
    )
    return detail.strip() if detail.strip() else name.strip()


def _extract_emergency_detail(name: str) -> str:
    """Extract emergency luminaire detail, removing 'аварийный' since
    the template already says 'аварийного'.

    "Светильник светодиодный аварийный" -> "светодиодного"
    "Светильник аварийный MARS 2223-4 LED с пиктограммой" -> "MARS 2223-4 LED с пиктограммой"
    """
    detail = re.sub(
        r"^светильн\S*\s*",
        "",
        name.strip(),
        flags=re.IGNORECASE,
    )
    # Remove "аварийный/аварийная/аварийное/эвакуационный" and surrounding spaces
    detail = re.sub(r"\s*(?:аварийн\S*|эвакуационн\S*)\s*", " ", detail, flags=re.IGNORECASE)
    return detail.strip() if detail.strip() else name.strip()


def _make_panel_detail(name: str) -> str:
    """Extract panel designation from name.

    "ЩР-12" -> "ЩР-12"
    "Щит ЩО1" -> "ЩО1"
    "ВРУ-1" -> "ВРУ-1"
    "ЩО1" -> "ЩО1"
    """
    # Try to find a panel designation like ЩР-1, ЩО1, ВРУ-1
    m = re.search(
        r"((?:ЩР|ЩО|ЩАО|ЩВН|ЩВ|ЩЭ|ВРУ|ПД|ЩУ|ГРЩ|РУ)[\s-]*\d*[\w-]*)",
        name,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()
    return name.strip()


def _make_cable_detail(name: str) -> str:
    """Extract cable brand + section from name.

    "Кабель ВВГнг 3х1.5" -> "ВВГнг 3х1.5"
    "ВВГнг(А)-LS 5x2.5" -> "ВВГнг(А)-LS 5x2.5"
    "Кабель ППГнг-(А)-HF сечением 3x1,5" -> "ППГнг-(А)-HF 3x1,5"
    "Кабель 1х40 (?-1х40)" -> "1х40"
    """
    detail = re.sub(r"^кабель\s*", "", name.strip(), flags=re.IGNORECASE)
    detail = re.sub(r"\bсечением?\b\s*", "", detail, flags=re.IGNORECASE)
    detail = re.sub(r"\bсеч\.?\s*", "", detail, flags=re.IGNORECASE)
    # Remove "(прокладка)" suffix
    detail = re.sub(r"\s*\(прокладка\)\s*", " ", detail, flags=re.IGNORECASE)
    # Remove trailing group references like "(ЩО1-Гр.7)" or "(ЩО1-ЩО1-Гр.7)"
    detail = re.sub(r"\s*\([^)]*(?:Щ|Гр\.?|гр\.?)[\w-]*\)\s*$", "", detail)
    # Remove unknown cable type markers like "(?-1х40)"
    detail = re.sub(r"\s*\(\?-[^)]*\)\s*", "", detail)
    return detail.strip() if detail.strip() else name.strip()


def _make_wire_detail(name: str) -> str:
    """Extract wire/conductor brand + section."""
    detail = re.sub(r"^провод\s*", "", name.strip(), flags=re.IGNORECASE)
    detail = re.sub(r"\bсечением?\b\s*", "", detail, flags=re.IGNORECASE)
    return detail.strip() if detail.strip() else name.strip()


def _make_conduit_detail(name: str) -> str:
    """Extract conduit pipe specification.

    "Труба ПВХ гибкая гофр. д.16мм" -> "ПВХ гибкая гофр. д.16мм"
    "Гофротруба ПВХ d=20" -> "ПВХ d=20"
    "Кабель-канал 40x25" -> "40x25"
    """
    detail = re.sub(
        r"^(?:труб[аы]|гофр[оа]?труб[аы]?|кабель-?канал|короб)\s*",
        "",
        name.strip(),
        flags=re.IGNORECASE,
    )
    return detail.strip() if detail.strip() else name.strip()


def _make_tray_detail(name: str) -> str:
    """Extract cable tray specification.

    "Лоток 300х100" -> "300х100"
    "Лоток перфорированный 200x50" -> "перфорированный 200x50"
    """
    detail = re.sub(
        r"^лот[оа]к\s*",
        "",
        name.strip(),
        flags=re.IGNORECASE,
    )
    return detail.strip() if detail.strip() else name.strip()


# ---------------------------------------------------------------------------
# Panel detection regex (separate because we need it for short names)
# ---------------------------------------------------------------------------

# Matches names that START with a panel designation code.
# Note: Python \b doesn't work well at Cyrillic/Latin boundaries so we use
# explicit look-ahead or end-of-string.
_PANEL_START_RE = re.compile(
    r"^(?:ЩР|ЩО|ЩАО|ЩВН|ЩВ|ЩЭ|ВРУ|ЩУ|ГРЩ)"
    r"(?:[-\s]\d|\d|$)",
    re.IGNORECASE,
)

_PANEL_KEYWORD_RE = re.compile(
    r"^щит\s|^пд\s|^ру\s"
    r"|щит\s+(?:распределительн|осветительн|управлен|силов)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Unit normalization
# ---------------------------------------------------------------------------

def _normalize_unit(unit: str) -> str:
    """Normalize unit string to standard VOR form.

    "шт." -> "шт", "метр" -> "м", "компл." -> "комп", etc.
    """
    u = unit.strip().rstrip(".")
    mapping = {
        "шт": "шт", "штук": "шт", "штука": "шт",
        "м": "м", "метр": "м", "метров": "м",
        "мп": "м.п.", "м п": "м.п.", "пм": "м.п.", "п.м": "м.п.",
        "м.п": "м.п.",
        "компл": "комп", "комплект": "комп",
        "комп": "комп",
        "м3": "м3", "куб": "м3",
        "каб": "каб.",
        "изм": "изм.",
    }
    return mapping.get(u.lower(), unit.strip())


# ---------------------------------------------------------------------------
# Classification function
# ---------------------------------------------------------------------------

def _classify(name: str) -> tuple[str, str, str]:
    """Classify equipment name and return (category, work_template, unit).

    The work_template contains {detail} placeholder.
    Returns ("unknown", "", "") if no match.

    Naming follows reference VOR convention (singular form):
      Монтаж светильника, Монтаж розетки, Монтаж выключателя, etc.

    S003-02 (T028) additions:
      Seven new categories are recognised before the generic fallbacks.
      See S003-01 analyst mapping in `.tayfa/common/discussions/` for
      provenance of each template.  Order matters:
        * cable_strand_connection_small/large — match "подключение жил"
          BEFORE the generic `terminal` rule catches "соединит".
        * socket_in_tray — must precede the plain `socket` rule so a
          legend like "Розетка в кабель-канале" is not mis-classified
          as a wall socket.
        * fireproof_sealant — totally new keyword ("огнестойкая пена").
        * conduit_pvc / conduit_corrugated — specialise the existing
          `conduit_pipe` catch-all with unit=м and a size-aware template.
        * cable_tray (specific "кабель-канал" form) — now produces the
          S003-01 "Монтаж кабель-канала {size} мм" template when the
          legend literally uses "кабель-канал"; the broader "лоток"
          rule continues to yield the legacy "Монтаж лотка кабельного".
    """
    lower = name.lower().strip()

    # --- 0a. Cable strand connection (S003-02 / T028) -------------------
    # Must precede generic terminal rule that catches "соединит".
    # Reference: S003-01 analyst mapping.
    if re.search(r"подключени\S*\s+жил", lower):
        # Section classification: "до 10", "до 10 мм", "10 мм2", "свыше 10",
        # "более 10", "выше 10" — all treated as size threshold around 10 mm2.
        if re.search(r"свыш\S*\s*10|более\s*10|выше\s*10", lower):
            return (
                "cable_strand_connection_large",
                "Подключение жил кабелей свыше 10 мм2",
                "шт",
            )
        return (
            "cable_strand_connection_small",
            "Подключение жил кабелей до 10 мм2",
            "шт",
        )

    # --- 0b. Fireproof sealant / two-component fire foam (S003-02) ------
    if re.search(r"огнестойк\S*\s+пен|двухкомпонент\S*\s+огнестойк", lower):
        return (
            "fireproof_sealant",
            "Двухкомпонентная огнестойкая пена {detail}",
            "компл",
        )

    # --- 0c. Socket in cable channel (S003-02) --------------------------
    # Must precede the plain `socket` rule on line ~282.
    if re.search(r"розетк\S*.*кабел[ья]?-?канал", lower):
        return (
            "socket_in_tray",
            "Монтаж розеток в кабель-канале {detail}",
            "шт",
        )

    # --- 0d. Cable-channel (S003-02) ------------------------------------
    # "Кабель-канал" must be detected BEFORE the generic `cable` rule on
    # section 7 below (which would otherwise greedily swallow the word
    # "кабель").  Matches names starting with "кабель-канал" OR "кабельканал"
    # (occasionally written without hyphen on CAD exports).
    if re.match(r"\s*кабел[ья]?-?канал", lower):
        return "cable_tray", "Монтаж кабель-канала {detail} мм", "м"

    # --- 1. Junction boxes / distribution boxes (BEFORE conduits!) ---
    if re.search(
        r"коробк.*(?:распред|ответвит|соединит|установочн|монтажн)"
        r"|коробк.*(?:ip|пластик)",
        lower,
    ):
        return "junction_box", "Монтаж коробки {detail}", "шт"

    # --- 2. Luminaires (emergency first, then general) ---
    if re.search(
        r"(?:светильник|светил\.).*(?:аварийн|эвакуац|exit|выход|указат|пиктограмм)",
        lower,
    ):
        return "luminaire_emergency", "Монтаж светильника аварийного {detail}", "шт"

    if re.search(
        r"(?:указатель|табло).*(?:exit|выход|эвакуац|аварийн)",
        lower,
    ):
        return "luminaire_exit", "Монтаж указателя аварийного освещения {detail}", "шт"

    if re.search(r"светильник|светил\.", lower):
        return "luminaire", "Монтаж светильника {detail}", "шт"

    # --- 3. Frames (before sockets/switches, ref: "Установка рамки") ---
    if re.search(r"рамк[аи]", lower):
        return "frame", "Установка рамки {detail}", "шт"

    # --- 4. Sockets ---
    if re.search(r"розетк", lower):
        return "socket", "Монтаж розетки {detail}", "шт"

    # --- 5. Switches (but NOT "автоматический выключатель" — that's a breaker) ---
    if re.search(r"выключатель|выкл\.", lower):
        if re.search(r"автомат\S*\s+выкл|дифф?\s*автомат", lower):
            return "circuit_breaker", "Монтаж автоматического выключателя {detail}", "шт"
        return "switch", "Монтаж выключателя {detail}", "шт"

    # --- 6. Panels ---
    if _PANEL_START_RE.match(name.strip()) or _PANEL_KEYWORD_RE.search(lower):
        return "panel", "Монтаж щита распределительного {detail}", "комп"

    # --- 7. Cables ---
    # "Кабельный вывод" is NOT a cable — it's a cable entry/exit point (шт)
    if re.search(r"кабельн\S*\s+вывод", lower):
        return "cable_entry", "Монтаж кабельного вывода {detail}", "шт"
    if re.search(
        r"(?:кабель|кабел[ья])"
        r"|^ввг|^ппгнг|^nym|^кг\s|^кввг|^вбшв|^авбб?шв|^аввг|^пвс\s",
        lower,
    ):
        return "cable", "Прокладка кабеля {detail}", "м"

    # --- 8. Grounding / Equipotential bonding / Lightning protection ---
    # (MUST be before Wire rule, since "проводник" also matches "провод")
    # These categories use "{fullname}" placeholder — the full name IS the detail.
    if re.search(r"стержен.*заземлен|электрод.*заземлен|забивк.*заземл", lower):
        return "ground_rod", "Забивка стержня заземления {detail}", "шт"
    if re.search(r"полос[аы]?\s*(?:стальн|оцинков|заземл|40)", lower):
        return "ground_strip", "Прокладка заземлителя горизонтального {detail}", "м"
    if re.search(r"уравнивани.*потенциал|потенциал.*уравнивани", lower):
        return "equipotential", "Прокладка проводника уравнивания потенциалов", "м"
    if re.search(r"проводник.*заземлен|заземлител.*горизонт", lower):
        return "grounding", "Прокладка горизонтального заземлителя", "м"
    if re.search(r"заземл|зазем\.|заземлител", lower):
        return "grounding", "Прокладка заземлителя {detail}", "м"
    if re.search(r"молниезащит|молниеприёмник|молниеприемник|токоотвод", lower):
        return "lightning", "Монтаж молниезащиты {detail}", "м"
    if re.search(r"проводник", lower):
        return "conductor", "Прокладка проводника {detail}", "м"

    # --- 9. Wire ---
    if re.search(r"(?:провод(?:\s|$|[а-яё]))|^пугв|^пув\s", lower):
        return "wire", "Прокладка провода {detail}", "м"

    # --- 10. Cable trays and cable channels ---
    # S003-02 (T028): cable-channel rows like "Кабель-канал 40х25" get
    # the dedicated "Монтаж кабель-канала {size} мм" template (unit m).
    # The broader "лоток / кабельрост" rule retains legacy semantics.
    if re.search(r"кабел[ья]?-?канал", lower):
        return "cable_tray", "Монтаж кабель-канала {detail} мм", "м"
    if re.search(r"лот[оа]к|кабельрост", lower):
        return "cable_tray", "Монтаж лотка кабельного {detail}", "м"

    # --- 11. Conduit pipes ---
    # S003-02 (T028): split the former catch-all into three subkinds
    # based on the material/construction keyword.  The legacy category
    # key ``conduit_pipe`` stays as the fall-through for tokens that
    # don't clearly name PVC or a corrugated pipe (e.g. plain "труба").
    if re.search(r"труб\S*.*гофр|гофр[аы]?(?:труб)?", lower):
        return "conduit_corrugated", "Монтаж трубы гофрированной {detail}", "м"
    if re.search(r"труб\S*.*пвх|^пвх\s|пнд\s|пвд\s", lower):
        return "conduit_pvc", "Монтаж трубы ПВХ {detail}", "м"
    if re.search(
        r"(?:труба)|^короб\s",
        lower,
    ):
        return "conduit_pipe", "Монтаж трубы гофрированной {detail}", "м"

    # --- 12. Sensors / detectors ---
    if re.search(r"датчик|извещатель|детектор", lower):
        return "sensor", "Монтаж датчика {detail}", "шт"

    # --- 13. Circuit breakers / fuses ---
    if re.search(
        r"автомат\S*\s+выкл|автомат(?:\s|$)|предохранитель|узо(?:\s|$)|дифф?\s*автомат",
        lower,
    ):
        return "circuit_breaker", "Монтаж автоматического выключателя {detail}", "шт"

    # --- 14. Contactors / relays ---
    if re.search(r"контактор|пускатель|реле(?:\s|$)", lower):
        return "contactor", "Монтаж контактора {detail}", "шт"

    # --- 15. DIN rail equipment (generic) ---
    if re.search(r"дин.?рейк|din.?рейк", lower):
        return "din_rail", "Монтаж оборудования на DIN-рейку {detail}", "шт"

    # --- 16. Connectors / terminals ---
    if re.search(r"клемм|наконечник|гильз|соединител", lower):
        return "terminal", "Подключение {detail}", "шт"

    return "unknown", "", ""


# ---------------------------------------------------------------------------
# Main mapping function
# ---------------------------------------------------------------------------

def map_equipment_to_work(
    equipment_name: str,
    unit: str = "шт",
) -> WorkMapping:
    """Map an equipment name to a VOR work description.

    Parameters
    ----------
    equipment_name : str
        Equipment name as extracted from PDF legend or cable schedule.
    unit : str
        Default unit from the source (may be overridden by category rules).

    Returns
    -------
    WorkMapping
        Contains work_name, unit, equipment_name (original), and category.

    Examples
    --------
    >>> m = map_equipment_to_work("Светильник светодиодный ДВО 6565-25-О")
    >>> m.work_name
    'Монтаж светильника светодиодный ДВО 6565-25-О'

    >>> m = map_equipment_to_work("Розетка РА16-003")
    >>> m.work_name
    'Монтаж розетки РА16-003'

    >>> m = map_equipment_to_work("ЩР-12")
    >>> m.work_name
    'Монтаж щита распределительного ЩР-12'

    >>> m = map_equipment_to_work("Кабель ВВГнг 3х1.5", unit="м")
    >>> m.work_name
    'Прокладка кабеля ВВГнг 3х1.5'
    """
    name = equipment_name.strip()
    if not name:
        return WorkMapping(
            work_name="",
            unit=_normalize_unit(unit),
            equipment_name=equipment_name,
            category="unknown",
        )

    category, template, default_unit = _classify(name)

    if category == "unknown":
        # Default: "Монтаж [equipment_name]"
        return WorkMapping(
            work_name=f"Монтаж {name}",
            unit=_normalize_unit(unit),
            equipment_name=equipment_name,
            category="other",
        )

    # Determine detail based on category
    if category == "panel":
        detail = _make_panel_detail(name)
    elif category == "cable":
        detail = _make_cable_detail(name)
    elif category == "wire":
        detail = _make_wire_detail(name)
    elif category == "luminaire_emergency":
        detail = _extract_emergency_detail(name)
    elif category == "luminaire":
        detail = _extract_luminaire_detail(name)
    elif category == "luminaire_exit":
        detail = _extract_detail(name)
    elif category in ("conduit_pipe", "conduit_pvc", "conduit_corrugated"):
        detail = _make_conduit_detail(name)
    elif category == "cable_tray":
        detail = _make_tray_detail(name)
    elif category == "socket_in_tray":
        # S003-02 (T028): "Розетка в кабель-канале ..." -> keep the
        # size/model descriptor after "кабель-канале".
        detail = re.sub(
            r"^розетк\S*.*кабел[ья]?-?канал[еа]?\s*",
            "",
            name.strip(),
            flags=re.IGNORECASE,
        ).strip() or _extract_detail(name)
    elif category == "fireproof_sealant":
        # S003-02 (T028): strip leading "двухкомпонентная огнестойкая
        # пена" / "огнестойкая пена" so the remainder (typically a
        # DNxxxx code) drops into {detail}.
        detail = re.sub(
            r"^(?:двухкомпонент\S*\s*)?огнестойк\S*\s*пен\S*\s*",
            "",
            name.strip(),
            flags=re.IGNORECASE,
        ).strip()
    elif category in ("cable_strand_connection_small",
                      "cable_strand_connection_large"):
        # S003-02 (T028): templates are fully literal — no detail.
        detail = ""
    elif category in ("ground_rod", "ground_strip"):
        # Strip "стержень/полоса" + "заземления/заземлен" from detail
        detail = _extract_detail(name)
        detail = re.sub(r"^заземлени\S*\s*", "", detail, flags=re.IGNORECASE).strip()
    elif category in ("grounding", "equipotential", "conductor",
                       "lightning", "cable_entry"):
        # These use {fullname} or have static templates — detail not used
        detail = ""
    else:
        detail = _extract_detail(name)

    # Build work name: {detail} = extracted detail, {fullname} = full original name
    work_name = template.replace("{fullname}", name).replace("{detail}", detail).strip()
    # Clean up double/triple spaces
    work_name = re.sub(r"\s{2,}", " ", work_name)

    return WorkMapping(
        work_name=work_name,
        unit=_normalize_unit(default_unit),
        equipment_name=equipment_name,
        category=category,
    )


# ---------------------------------------------------------------------------
# Convenience: batch mapping
# ---------------------------------------------------------------------------

def map_items(items: list[dict]) -> list[dict]:
    """Apply work mapping to a list of equipment item dicts.

    Each item dict should have at least 'name' key.
    Adds 'work_name', 'unit' (mapped), and 'equipment_name' (original).

    The original 'name' is replaced with the work_name for VOR output.
    """
    result = []
    for item in items:
        name = item.get("name", "")
        original_unit = item.get("unit", "шт")
        mapping = map_equipment_to_work(name, unit=original_unit)

        new_item = dict(item)
        new_item["work_name"] = mapping.work_name
        new_item["equipment_name"] = mapping.equipment_name
        new_item["unit"] = mapping.unit
        new_item["category"] = mapping.category
        result.append(new_item)
    return result


# ---------------------------------------------------------------------------
# CLI / debug
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")

    test_names = [
        # --- Luminaires ---
        ("Светильник светодиодный ДВО 6565-25-О", "шт"),
        ("Светильник светодиодный UNI/R EVO (595x595) 40W OPL 840", "шт"),
        ("Светильник светодиодный SLICK.PRS LED 30 with driver box", "шт"),
        ("Светильник светодиодный OWP EVO (595x595) 30W OPL 840", "шт"),
        ("Светильник светодиодный SIMPLE OPTIMA (1200) 40W OPL 840", "шт"),
        ("Светильник светодиодный OWP", "шт"),
        # --- Emergency luminaires ---
        ("Светильник светодиодный аварийный", "шт"),
        ("Светильник аварийный MARS 2223-4 LED с пиктограммой", "шт"),
        ("Светильник светодиодный эвакуационный MARS 2223-4 LED", "шт"),
        # --- Sockets ---
        ("Розетка РА16-003", "шт"),
        ("Розетка с заземлением со шторками, 16А", "шт"),
        # --- Switches ---
        ("Выключатель 1-кл", "шт"),
        ("Выключатель 1-клавишный 10А IP44", "шт"),
        # --- Frames ---
        ("Рамка 2-постовая горизонтальная", "шт"),
        # --- Panels ---
        ("ЩР-12", "шт"),
        ("ЩО1", "шт"),
        ("ЩАО1", "шт"),
        ("ВРУ-1", "шт"),
        ("Щит ЩВ", "шт"),
        # --- Cables ---
        ("ВВГнг 3х1.5", "м"),
        ("Кабель ВВГнг(А)-LS 5x2.5", "м"),
        ("Кабель ППГнг-(А)-HF сечением 3x1,5", "м"),
        ("Кабель ППГнг-(А)-HF (прокладка)", "м"),
        ("Кабель ВВГнг(А)-LS 3x1.5 (ЩО1-Гр.7)", "м"),
        ("Кабель 1х40 (?-1х40)", "м"),
        ("Кабель 1х18", "м"),
        # --- Wire ---
        ("Провод ПуГВ 1х6", "м"),
        # --- Conduits ---
        ("Труба ПВХ гибкая гофр. д.16мм", "м"),
        ("Гофротруба ПВХ d=20", "м"),
        # --- Cable trays ---
        ("Лоток 300х100", "м"),
        ("Лоток перфорированный 200x50", "м"),
        # --- Junction boxes ---
        ("Коробка распределительная IP55 100x100x50", "шт"),
        # --- Sensors ---
        ("Датчик движения", "шт"),
        # --- Exit indicators ---
        ("Указатель EXIT", "шт"),
        # --- Circuit breakers ---
        ("Автоматический выключатель 16А", "шт"),
        ("УЗО 40А 30мА", "шт"),
        # --- Terminals ---
        ("Клеммник WAGO 3x2.5", "шт"),
        # --- Contactors ---
        ("Контактор КМИ-10910 9А", "шт"),
        # --- Unknown (fallback) ---
        ("Электросчетчик Меркурий 230", "шт"),
        # --- Unit normalization ---
        ("Кабель NYM 3x1.5", "метр"),
        ("Щит ЩО2", "компл."),
    ]

    print(f"{'Equipment Name':<55s} -> {'Work Name':<65s} [{'Unit':>5s}] ({'Category'})")
    print("-" * 200)
    for name, unit in test_names:
        m = map_equipment_to_work(name, unit)
        print(f"{name:<55s} -> {m.work_name:<65s} [{m.unit:>5s}] ({m.category})")
