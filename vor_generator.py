#!/usr/bin/env python3
"""
VOR Generator — generate a Ведомость объёмов работ (.docx) from DXF drawings.

Parses all DXF files from a project folder, groups equipment by type and
installation height, extracts cables from schema drawings, and produces
a .docx document matching the standard VOR format.

Usage:
    python vor_generator.py "D:\\Project\\DWG\\_converted_dxf"
    python vor_generator.py "D:\\Project\\DWG\\_converted_dxf" -o VOR.docx
"""

import argparse
import math
import re
import sys
import time
from collections import defaultdict, OrderedDict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

from equipment_counter import (
    EquipmentItem,
    CableItem,
    SpecItem,
    process_dxf,
    extract_cables_dxf,
    parse_spec_dxf,
    classify_plan,
    extract_elevation_float,
    ELEVATION_RE,
    _HAS_DXF,
)

try:
    import ezdxf as _ezdxf
except ImportError:
    _ezdxf = None  # type: ignore[assignment]

# ── Constants ────────────────────────────────────────────────────────

HeightCategory = Literal[
    "до 5 метров",
    "от 5 до 13 метров",
    "от 13 до 20 метров",
    "от 20 до 35 метров",
]

HEIGHT_CATEGORIES: list[HeightCategory] = [
    "до 5 метров",
    "от 5 до 13 метров",
    "от 13 до 20 метров",
    "от 20 до 35 метров",
]

DXF_EXTS = {".dxf", ".DXF"}

_LUMINAIRE_PREFIXES = (
    "Светильник ", "Световые указатели ", "Световой указатель ",
)

_INDICATOR_KEYWORDS = ("указатель", "mercury", "atom", "выход", "exit")
_PICTOGRAM_KEYWORDS = ("пиктограмма", "пэу")
_SOCKET_KEYWORDS = ("розетк",)
_CABLE_OUTLET_KEYWORDS = ("кабельный вывод",)
_SWITCH_KEYWORDS = (
    "выключатель", "пост управления", "коробк", "датчик", "блок аварийн",
)
_MATERIAL_KEYWORDS = (
    "гофротруб", "стяжк", "металлоконструкц", "дюбель", "хомут",
    "лоток", "крепеж", "кронштейн", "рамка", "гильза", "трубка термо",
    "термоусад", "пена", "пистолет для", "гильза закладн",
)
_PANEL_KEYWORDS = ("щит", "що", "щао", "цсао", "вру", "панель питания",
                   "распределительное устройство")

_NON_EQUIPMENT_KEYWORDS = (
    "прокладка", "проводка", "кабельная трасса", "групповая сеть",
    "кабель ", "распаечная", "противопожарная", "подключение",
    "кабеленесущ", "точечное оборудование на плане",
    "кабельная конструкция",
)
_CABLE_SPEC_KEYWORDS = (
    "ппгнг", "ввгнг", "вбшвнг", "кгнг", "апвпу", "пвпу",
    "кабель силовой", "0,66кв", "1кв",
)

_GROUNDING_KEYWORDS = (
    "заземлен", "заземляющ", "уравнивания потенциалов", "стержень заземления",
    "контрольного соединения", "токоотвод", "проводник", "плоский проводник",
    "забивная головка", "соединитель диагональн", "соединитель плоского",
    "проводника к точке", "скоба на ленте", "шина уравнивания",
    "пугвнг", "провод пугв",
)
_LIGHTNING_KEYWORDS = (
    "молниезащит", "молниеприемни", "проводник круглый", "мостовая опора",
    "держатель на кровлю", "лента монтажная на кровельный",
    "зажим крепежный", "держатель для круглых",
    "болтовой на водосточный",
)
_PVC_CONDUIT_KEYWORDS = (
    "труба пвх", "гофр.", "с протяжкой", "держатель с защелкой",
)


# ── Helpers ──────────────────────────────────────────────────────────

def _classify_plan(filename: str) -> str:
    return classify_plan(filename)


_CONTENT_ELEV_RE = re.compile(
    r"на\s+отм[.\s_]*([+-]?\d+[.,]\d+)", re.IGNORECASE,
)
_CONTENT_ELEV_MM_RE = re.compile(
    r"на\s+отм[.\s_]*([+-]?\d{4,})\b", re.IGNORECASE,
)
_ATTRIB_ELEV_RE = re.compile(
    r"^[+-]?\d+[.,]\d+$",
)


def _parse_elev_value(raw: str) -> float | None:
    """Parse elevation string to float meters.

    Handles formats: '+7.800' (m), '+7,800' (m with comma), '+7800' (mm).
    Values >= 100 are treated as millimetres and divided by 1000.
    """
    s = raw.replace(",", ".")
    if s.startswith("+"):
        s = s[1:]
    try:
        val = float(s)
    except ValueError:
        return None
    if val >= 100.0 or val <= -100.0:
        val = val / 1000.0
    return val


def _classify_by_content(dxf_path: str) -> tuple[str, float | None]:
    """Fallback classification by scanning MTEXT/TEXT content inside the DXF.

    Returns (plan_type, elevation).  Used when filename-based classification
    yields 'другое' (e.g. because of broken encoding).
    """
    if _ezdxf is None:
        return "другое", None

    try:
        doc = _ezdxf.readfile(dxf_path)
    except Exception:
        return "другое", None

    msp = doc.modelspace()

    texts: list[str] = []
    for e in msp:
        plain: str | None = None
        if e.dxftype() == "MTEXT":
            plain = e.plain_text().strip()
        elif e.dxftype() == "TEXT":
            plain = e.dxf.text.strip() if hasattr(e.dxf, "text") else None
        if plain:
            texts.append(plain)

        # Collect ATTRIB values from INSERT entities
        if e.dxftype() == "INSERT":
            try:
                for attrib in e.attribs:
                    val = attrib.dxf.text.strip()
                    if val:
                        texts.append(val)
            except Exception:
                pass

    # Also scan block names for elevation info
    block_texts: list[str] = []
    for block in doc.blocks:
        if "отм" in block.name.lower():
            block_texts.append(block.name)

    if not texts and not block_texts:
        return "общие", None

    combined = " ".join(texts).lower()

    # Extract elevation from content, ATTRIBs, or block names
    elev: float | None = None

    # 1) Try decimal elevation in text/ATTRIB content: "на отм +7.800"
    for t in texts:
        m = _CONTENT_ELEV_RE.search(t)
        if m:
            elev = _parse_elev_value(m.group(1))
            if elev is not None:
                break

    # 2) Try millimetre elevation in text/ATTRIB: "на отм +7800"
    if elev is None:
        for t in texts:
            m = _CONTENT_ELEV_MM_RE.search(t)
            if m:
                elev = _parse_elev_value(m.group(1))
                if elev is not None:
                    break

    # 3) Try block names: "План освещение на отм_ _7800"
    if elev is None:
        for bname in block_texts:
            m = _CONTENT_ELEV_RE.search(bname)
            if m:
                elev = _parse_elev_value(m.group(1))
                if elev is not None:
                    break
            m = _CONTENT_ELEV_MM_RE.search(bname)
            if m:
                elev = _parse_elev_value(m.group(1))
                if elev is not None:
                    break

    # Classify by keywords found in content
    if "принципиальная" in combined and "схема" in combined:
        return "схема", elev
    if "общие данные" in combined:
        return "общие", elev
    if "план освещения" in combined:
        return "освещение", elev
    if "план привязки" in combined:
        return "привязка", elev
    if "план расстановки" in combined or "план расположения" in combined:
        return "расположение", elev
    # Heuristic: luminaire mentions imply lighting plan
    has_luminaire = "светильник" in combined
    has_rozetka = "розетк" in combined
    has_cable_route = "кабельная трасса" in combined
    if has_rozetka and not has_luminaire:
        return "расположение", elev
    if has_luminaire and has_cable_route:
        return "освещение", elev
    if has_luminaire:
        return "привязка", elev

    return "другое", elev


_ELEV_VALUE_RE = re.compile(r"[+-]?\d+[-.,]\d+")


def _extract_elevation(filename: str) -> float | None:
    return extract_elevation_float(filename)


def _extract_all_elevations(filename: str) -> list[float]:
    """Extract all elevation values from a filename like '+7-800, +9-000'."""
    m = ELEVATION_RE.search(filename)
    if not m:
        return []
    after = filename[m.start(1):]
    raw_values = _ELEV_VALUE_RE.findall(after)
    elevs = []
    for raw in raw_values:
        raw = raw.replace(",", ".").replace("-", ".")
        if raw.startswith("+"):
            raw = raw[1:]
        try:
            elevs.append(float(raw))
        except ValueError:
            pass
    return elevs


def elevation_to_height(elev_m: float) -> HeightCategory:
    if elev_m < 5.0:
        return "до 5 метров"
    if elev_m < 13.0:
        return "от 5 до 13 метров"
    if elev_m < 20.0:
        return "от 13 до 20 метров"
    return "от 20 до 35 метров"


def _normalize_equip_name(name: str) -> str:
    """Strip prefixes and normalize whitespace for equipment name grouping."""
    s = name.strip()
    for pfx in _LUMINAIRE_PREFIXES:
        if s.startswith(pfx):
            s = s[len(pfx):]
            break
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _classify_equipment(name: str) -> str:
    """Classify an equipment item into a VOR section."""
    nl = name.lower()
    if any(kw in nl for kw in _NON_EQUIPMENT_KEYWORDS):
        return "skip"
    if re.search(r"\d+\s*мм\s*[xх×]\s*\d+\s*мм", nl):
        return "skip"
    if any(kw in nl for kw in _MATERIAL_KEYWORDS):
        return "material"
    if any(kw in nl for kw in _PANEL_KEYWORDS):
        return "panel"
    if any(kw in nl for kw in _PICTOGRAM_KEYWORDS):
        return "pictogram"
    if any(kw in nl for kw in _INDICATOR_KEYWORDS):
        return "indicator"
    if any(kw in nl for kw in _SOCKET_KEYWORDS):
        return "socket"
    if any(kw in nl for kw in _CABLE_OUTLET_KEYWORDS):
        return "cable_outlet"
    if any(kw in nl for kw in _SWITCH_KEYWORDS):
        return "switch"
    if "[?" in name or "[Auto" in name:
        return "skip"
    return "luminaire"


def _classify_spec_item(desc: str) -> str:
    """Classify a spec item into a VOR section (extended categories)."""
    nl = desc.lower()
    if any(kw in nl for kw in _PANEL_KEYWORDS):
        return "panel"
    if any(kw in nl for kw in _PICTOGRAM_KEYWORDS):
        return "pictogram"
    if any(kw in nl for kw in _INDICATOR_KEYWORDS):
        return "indicator"
    if "светильник" in nl or ("led" in nl and "розетк" not in nl):
        return "luminaire"
    if any(kw in nl for kw in _SOCKET_KEYWORDS):
        return "socket"
    if any(kw in nl for kw in _CABLE_OUTLET_KEYWORDS):
        return "cable_outlet"
    if any(kw in nl for kw in _SWITCH_KEYWORDS):
        return "switch"
    if any(kw in nl for kw in _CABLE_SPEC_KEYWORDS):
        return "cable"
    if any(kw in nl for kw in _PVC_CONDUIT_KEYWORDS):
        return "pvc_conduit"
    if any(kw in nl for kw in _LIGHTNING_KEYWORDS):
        return "lightning"
    if any(kw in nl for kw in _GROUNDING_KEYWORDS):
        return "grounding"
    if any(kw in nl for kw in _MATERIAL_KEYWORDS):
        return "material"
    return "material"


# ── Schema parsing ───────────────────────────────────────────────────

@dataclass
class PanelInfo:
    name: str
    breaker: str = ""
    feed_cable: str = ""
    circuit_count: int = 0
    circuit_cables: list[tuple[str, int]] = field(default_factory=list)


_PANEL_NAME_RE = re.compile(
    r"(ЩР[-\s]?\d+(?:\.\d+)?|ЩО[-\s]?\d+(?:\.\d+)?|ЩАО[-\s]?\d+(?:\.\d+)?|ЦСАО[-\s]?\d+|ВРУ[-\s]?\d+(?:\.\d+)?)", re.IGNORECASE,
)
_KM_RE = re.compile(r"^KM-\d+$")
_CABLE_LEN_RE = re.compile(
    r"([\w()А-Яа-яё\-]+\s+\d+[хx×]\d[\d,]*)\s+L=(\d+)", re.IGNORECASE,
)


def extract_panels_from_schema(dxf_path: str, log=print) -> list[PanelInfo]:
    """Extract panel info (names, breakers, circuits) from a schema DXF."""
    if _ezdxf is None:
        return []
    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    panel_entries: list[tuple[str, float, float]] = []
    qf_entries: list[tuple[str, float, float]] = []
    feed_cables: list[str] = []
    circuit_cable_entries: list[tuple[str, int, float, float]] = []
    km_count = 0

    for e in msp:
        if e.dxftype() == "MTEXT":
            plain = e.plain_text().strip()
            x, y = e.dxf.insert.x, e.dxf.insert.y
            m = _PANEL_NAME_RE.search(plain)
            if m and len(plain) < 20:
                panel_entries.append((m.group(1).replace(" ", ""), x, y))
            if _KM_RE.match(plain):
                km_count += 1
            if plain.startswith("QF") and ("A/" in plain or "А/" in plain):
                qf_entries.append((plain.replace("\n", " "), x, y))
            if "ЦСАО" in plain and "Dialog" in plain and len(plain) > 20:
                panel_entries.append(("ЦСАО-3", x, y))
            if plain.startswith("н") and ("ППГ" in plain or "ВБШ" in plain):
                feed_cables.append(plain)
            m2 = _CABLE_LEN_RE.search(plain)
            if m2:
                circuit_cable_entries.append((m2.group(1).strip(), int(m2.group(2)), x, y))

    seen: set[str] = set()
    panels: list[PanelInfo] = []
    for pname, px, py in panel_entries:
        if pname in seen:
            continue
        seen.add(pname)
        info = PanelInfo(name=pname)
        best_dist = float("inf")
        for qf_text, qx, qy in qf_entries:
            d = abs(qx - px) + abs(qy - py)
            if d < best_dist:
                best_dist = d
                info.breaker = qf_text
        pname_short = pname.replace("-", "").replace(" ", "")
        for fc in feed_cables:
            fc_norm = fc.replace("-", "").replace(" ", "")
            if pname_short.lower() in fc_norm.lower():
                info.feed_cable = fc
                break
        panels.append(info)

    if km_count > 0:
        for p in panels:
            if "ЩО" in p.name and "А" not in p.name:
                p.circuit_count = km_count
                break

    panel_positions: dict[str, tuple[float, float]] = {}
    for pname, px, py in panel_entries:
        if pname not in panel_positions:
            panel_positions[pname] = (px, py)

    for ctype, clen, cx, cy in circuit_cable_entries:
        best_panel = None
        best_d = float("inf")
        for p in panels:
            pos = panel_positions.get(p.name)
            if pos is None:
                continue
            d = abs(pos[0] - cx) + abs(pos[1] - cy)
            if d < best_d:
                best_d = d
                best_panel = p
        if best_panel is not None:
            best_panel.circuit_cables.append((ctype, clen))

    # Extract circuit breaker counts from ACAD_TABLE *T blocks.
    # AutoCAD creates *T blocks and ACAD_TABLE entities in matching order
    # (sorted by handle).  Count QF-x.y entries in each *T block, then
    # use the corresponding ACAD_TABLE position to match to nearest panel.
    _QF_CIRCUIT_RE = re.compile(r"^QF[-\s]?\d+\.\d+$")
    t_blocks_sorted = sorted(
        (b for b in doc.blocks if b.name.startswith("*T")),
        key=lambda b: b.name,
    )
    acad_tables_sorted = sorted(
        (e for e in msp if e.dxftype() == "ACAD_TABLE"),
        key=lambda e: e.dxf.handle,
    )
    if len(t_blocks_sorted) == len(acad_tables_sorted):
        for tblock, atable in zip(t_blocks_sorted, acad_tables_sorted):
            qf_count = sum(
                1 for ent in tblock
                if ent.dxftype() == "MTEXT"
                and _QF_CIRCUIT_RE.match(ent.plain_text().strip())
            )
            if qf_count == 0:
                continue
            try:
                tx, ty = atable.dxf.insert[0], atable.dxf.insert[1]
            except Exception:
                continue
            best_panel = None
            best_d = float("inf")
            for p in panels:
                pos = panel_positions.get(p.name)
                if pos is None:
                    continue
                d = abs(pos[0] - tx) + abs(pos[1] - ty)
                if d < best_d:
                    best_d = d
                    best_panel = p
            if best_panel is not None and best_panel.circuit_count == 0:
                best_panel.circuit_count = qf_count

    # Sort: emergency panels (АО/AO) first, then distribution panels
    def _panel_sort_key(p: PanelInfo) -> tuple[int, str]:
        nl = p.name.upper()
        if "АО" in nl or "AO" in nl:
            return (0, p.name)
        return (1, p.name)
    panels.sort(key=_panel_sort_key)

    # Filter upstream/parent panels that don't belong to this building section.
    # If ≥2 panels share a section suffix (e.g. "3.12"), drop any that lack it.
    _SECTION_RE = re.compile(r"(\d+\.\d+)")
    suffix_counts: dict[str, int] = {}
    for p in panels:
        m = _SECTION_RE.search(p.name)
        if m:
            suffix_counts[m.group(1)] = suffix_counts.get(m.group(1), 0) + 1
    if suffix_counts:
        dominant = max(suffix_counts, key=lambda s: suffix_counts[s])
        if suffix_counts[dominant] >= 2:
            before = len(panels)
            panels = [p for p in panels if dominant in p.name]
            if len(panels) < before:
                log(f"    [panels] Filtered to section {dominant}: "
                    f"{before} → {len(panels)}")

    for p in panels:
        total_cable_m = sum(ln for _, ln in p.circuit_cables)
        log(f"    Panel: {p.name} breaker={p.breaker}  "
            f"feed={p.feed_cable}  circuits={p.circuit_count}  "
            f"cables={len(p.circuit_cables)} ({total_cable_m}m)")
    return panels


# ── Multi-sheet combined DXF parsing ─────────────────────────────────

_COUNT_TYPE_RE = re.compile(r"^(\d+)\s*[-–]\s*(.+)", re.IGNORECASE)


@dataclass
class _SheetRegion:
    """A virtual sheet within a combined DXF file."""
    legend_x: float
    x_left: float
    x_right: float
    elevation: float | None
    height_category: HeightCategory | None
    legend_items: list[str] = field(default_factory=list)
    equipment: list[EquipmentItem] = field(default_factory=list)


def _detect_sheets(dxf_path: str, log=print) -> list[_SheetRegion] | None:
    """Detect multiple legends in a DXF and split into virtual sheets.

    Returns None if file has 0-1 legends (standard single-sheet file).
    """
    if _ezdxf is None:
        return None
    doc = _ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    legend_positions: list[tuple[float, float]] = []
    for e in msp:
        if e.dxftype() == "MTEXT":
            plain = e.plain_text().strip()
            if "Условн" in plain and "обозначен" in plain.lower():
                legend_positions.append((e.dxf.insert.x, e.dxf.insert.y))

    if len(legend_positions) < 2:
        return None

    legend_positions.sort(key=lambda p: p[0])
    legend_xs = [lx for lx, _ in legend_positions]

    layout_elevations: dict[str, float] = {}
    for layout in doc.layouts:
        name = layout.name.strip()
        if name == "Model":
            continue
        m = re.match(r"[+]?(\d+[.,]\d+)", name)
        if m:
            elev = float(m.group(1).replace(",", "."))
            layout_elevations[name] = elev

    sorted_elevs = sorted(layout_elevations.values())

    sheets: list[_SheetRegion] = []
    for i, lx in enumerate(legend_xs):
        x_left = (legend_xs[i - 1] + lx) / 2 if i > 0 else lx - 40000
        x_right = (lx + legend_xs[i + 1]) / 2 if i < len(legend_xs) - 1 else lx + 40000
        elev = sorted_elevs[i] if i < len(sorted_elevs) else None
        hcat = elevation_to_height(elev) if elev is not None else None
        sheets.append(_SheetRegion(
            legend_x=lx, x_left=x_left, x_right=x_right,
            elevation=elev, height_category=hcat,
        ))

    all_mtext: list[tuple[float, float, str]] = []
    for e in msp:
        if e.dxftype() == "MTEXT":
            plain = e.plain_text().replace("\n", " ").strip()
            plain = re.sub(r"\s+", " ", plain)
            if plain:
                all_mtext.append((e.dxf.insert.x, e.dxf.insert.y, plain))

    legend_y_threshold = min(ly for _, ly in legend_positions) + 500

    _LEGEND_LUMINAIRE_KW = ["SLICK", "ARCTIC", "CD LED", "NERO", "INSEL"]
    _LEGEND_EQUIP_KW = [
        "Розетк", "Выключатель", "Кабельный вывод", "Коробк", "Щит",
        "Датчик", "Пост управления", "Блок аварийн",
    ]
    _ANNOTATION_LUMINAIRE_KW = [
        "SLICK", "ARCTIC", "LED", "CD", "OPL", "INSEL", "NERO",
    ]

    for sheet in sheets:
        for x, y, plain in all_mtext:
            if sheet.x_left <= x <= sheet.x_right and y < legend_y_threshold:
                if "Светильник" in plain or any(
                    kw in plain for kw in _LEGEND_LUMINAIRE_KW
                ) or any(kw in plain for kw in _LEGEND_EQUIP_KW):
                    sheet.legend_items.append(plain)

        equip_counts: dict[str, int] = defaultdict(int)

        for x, y, plain in all_mtext:
            if sheet.x_left <= x <= sheet.x_right and y > legend_y_threshold:
                m = _COUNT_TYPE_RE.match(plain)
                if m:
                    count = int(m.group(1))
                    etype = m.group(2).strip()
                    if any(kw in etype for kw in _ANNOTATION_LUMINAIRE_KW) or any(
                        kw.lower() in etype.lower() for kw in _LEGEND_EQUIP_KW
                    ):
                        matched_name = _match_annotation_to_legend(etype, sheet.legend_items)
                        equip_counts[matched_name] += count
                elif plain and not plain.startswith(("Обозначен", "Наименован")):
                    if len(plain) < 60 and (
                        any(kw in plain for kw in _LEGEND_LUMINAIRE_KW) or any(
                            kw.lower() in plain.lower() for kw in _LEGEND_EQUIP_KW
                        )
                    ):
                        matched_name = _match_annotation_to_legend(plain, sheet.legend_items)
                        equip_counts[matched_name] += 1

        for name, count in equip_counts.items():
            sheet.equipment.append(EquipmentItem(
                symbol="", name=name, count=count, count_ae=0,
            ))

    for s in sheets:
        eq_total = sum(it.count for it in s.equipment)
        log(f"    Sheet elev={s.elevation} ({s.height_category}): "
            f"{len(s.equipment)} types, {eq_total} items")

    return sheets


def _match_annotation_to_legend(annotation: str, legend_items: list[str]) -> str:
    """Find the best matching legend item for an abbreviated annotation."""
    ann_lower = annotation.lower().strip()
    ann_words = set(re.findall(r"[a-zA-Zа-яА-ЯёЁ]+|\d+", ann_lower))

    best_match = annotation
    best_score = 0

    for legend in legend_items:
        legend_lower = legend.lower()
        legend_words = set(re.findall(r"[a-zA-Zа-яА-ЯёЁ]+|\d+", legend_lower))
        common = ann_words & legend_words
        score = len(common)
        if score > best_score:
            best_score = score
            best_match = legend
    return best_match


# ── Data structures ──────────────────────────────────────────────────

@dataclass
class FileParseResult:
    filename: str
    plan_type: str
    elevation: float | None
    height_category: HeightCategory | None
    equipment: list[EquipmentItem] = field(default_factory=list)
    cables: list[CableItem] = field(default_factory=list)
    panels: list[PanelInfo] = field(default_factory=list)
    spec_items: list[SpecItem] = field(default_factory=list)


@dataclass
class AggregatedEquipment:
    name: str
    category: str
    counts_by_height: dict[HeightCategory, int] = field(default_factory=dict)
    total: int = 0
    unit: str = "шт"


# ── Scan & Parse ─────────────────────────────────────────────────────

def scan_and_classify(folder: Path) -> list[tuple[Path, str, float | None]]:
    """Scan folder for DXF files and classify each by plan type and elevation.

    Multi-elevation files (e.g. '+7-800, +9-000') produce one entry per
    elevation so that deduplication can match each against its привязка pair.

    Deduplicates by stem name: if the same DXF file appears in multiple
    _converted_dxf directories, only the first copy is kept.
    """
    seen_stems: set[str] = set()
    results = []
    for f in sorted(folder.rglob("*")):
        if not f.is_file() or f.suffix not in DXF_EXTS:
            continue
        if any(p.startswith(".") for p in f.relative_to(folder).parts):
            continue
        stem_lower = f.stem.lower()
        if stem_lower in seen_stems:
            continue
        seen_stems.add(stem_lower)
        plan_type = _classify_plan(f.name)
        elevs = _extract_all_elevations(f.name)
        elev = elevs[0] if elevs else None
        if plan_type == "другое" and f.suffix.lower() == ".dxf":
            plan_type, content_elev = _classify_by_content(str(f))
            if elev is None and content_elev is not None:
                elev = content_elev
        results.append((f, plan_type, elev))
    return results


def parse_all_files(
    file_list: list[tuple[Path, str, float | None]],
    log=print,
) -> list[FileParseResult]:
    """Parse all DXF files and return structured results.

    For combined DXF files with multiple legends, each virtual sheet is
    returned as a separate FileParseResult with its own elevation.
    """
    results: list[FileParseResult] = []
    parse_cache: dict[str, tuple[list[EquipmentItem], list[CableItem]]] = {}
    combined_cache: dict[str, list[_SheetRegion] | None] = {}

    for fpath, plan_type, elev in file_list:
        height_cat = elevation_to_height(elev) if elev is not None else None
        log(f"  [{plan_type}] {fpath.name}  elev={elev}  height={height_cat}")

        if plan_type in ("общие", "опросные"):
            results.append(FileParseResult(
                filename=fpath.name, plan_type=plan_type,
                elevation=elev, height_category=height_cat,
            ))
            continue

        if plan_type == "спецификация":
            try:
                spec = parse_spec_dxf(str(fpath), log=log)
                results.append(FileParseResult(
                    filename=fpath.name, plan_type=plan_type,
                    elevation=None, height_category=None,
                    spec_items=spec,
                ))
            except Exception as e:
                log(f"    Spec parse error: {e}")
            continue

        fkey = str(fpath)

        if fkey not in combined_cache:
            should_check = (
                fpath.suffix.lower() == ".dxf"
                and plan_type == "другое"
                and elev is None
            )
            if should_check:
                try:
                    log(f"    Checking for multi-sheet layout...")
                    combined_cache[fkey] = _detect_sheets(fkey, log=log)
                except Exception as e:
                    log(f"    Multi-sheet detection error: {e}")
                    combined_cache[fkey] = None
            else:
                combined_cache[fkey] = None

        sheets = combined_cache[fkey]
        if sheets is not None:
            log(f"    Combined DXF: {len(sheets)} sheets detected")
            for sheet in sheets:
                sheet_id = f"{fpath.name}__sheet_{sheet.elevation}"
                result = FileParseResult(
                    filename=sheet_id,
                    plan_type="освещение",
                    elevation=sheet.elevation,
                    height_category=sheet.height_category,
                    equipment=sheet.equipment,
                )
                results.append(result)

            cables: list[CableItem] = []
            if fpath.suffix.lower() == ".dxf":
                try:
                    cables = extract_cables_dxf(fkey)
                    if cables:
                        total_m = sum(c.total_length_m for c in cables)
                        log(f"    Cables: {len(cables)} types, {total_m}m")
                except Exception as e:
                    log(f"    Cable error: {e}")
                try:
                    panels = extract_panels_from_schema(fkey, log=log)
                except Exception as e:
                    log(f"    Panel error: {e}")
                    panels = []

            cable_result = FileParseResult(
                filename=fpath.name, plan_type="схема",
                elevation=None, height_category=None,
                cables=cables, panels=panels,
            )
            results.append(cable_result)
            continue

        result = FileParseResult(
            filename=fpath.name, plan_type=plan_type,
            elevation=elev, height_category=height_cat,
        )

        if fkey in parse_cache:
            items, cables = parse_cache[fkey]
            result.equipment = items
            result.cables = cables
            eq_total = sum(it.count + it.count_ae for it in items)
            log(f"    (cached) Equipment: {len(items)} types, {eq_total} total")
            results.append(result)
            continue

        items: list[EquipmentItem] = []
        cables: list[CableItem] = []

        try:
            items = process_dxf(str(fpath))
            result.equipment = items
            eq_total = sum(it.count + it.count_ae for it in items)
            log(f"    Equipment: {len(items)} types, {eq_total} total")
        except Exception as e:
            log(f"    Equipment error: {e}")

        if fpath.suffix.lower() == ".dxf":
            try:
                cables = extract_cables_dxf(str(fpath))
                result.cables = cables
                if cables:
                    total_m = sum(c.total_length_m for c in cables)
                    log(f"    Cables: {len(cables)} types, {total_m}m")
            except Exception as e:
                log(f"    Cable error: {e}")
            if plan_type == "схема":
                try:
                    result.panels = extract_panels_from_schema(str(fpath), log=log)
                except Exception as e:
                    log(f"    Panel error: {e}")

        parse_cache[fkey] = (items, cables)
        results.append(result)

    return results


# ── Aggregation ──────────────────────────────────────────────────────

def _merge_per_elevation(
    results: list[FileParseResult],
    log=print,
) -> list[FileParseResult]:
    """For each elevation, merge equipment from all matching plans.

    Takes the MAX count for each equipment type across all plans at the same
    elevation.  This captures items that appear in one plan but not the other
    (e.g. exit indicators only on освещение, higher switch counts on привязка).
    """
    dup_types = {"привязка", "освещение", "розетки", "расположение"}

    groups: dict[tuple[str, float | None], list[FileParseResult]] = defaultdict(list)
    out: list[FileParseResult] = []

    for r in results:
        if r.plan_type in dup_types:
            dup_key = (
                "light" if r.plan_type in ("освещение", "привязка") else "power",
                r.elevation,
            )
            groups[dup_key].append(r)
        else:
            out.append(r)

    # Merge None-elevation groups into a known-elevation group of the same
    # category when exactly one such group exists (e.g. привязка without
    # elevation matched to освещение that has elevation).
    none_keys = [k for k in groups if k[1] is None]
    for nk in none_keys:
        cat = nk[0]
        peers = [k for k in groups if k[0] == cat and k[1] is not None]
        if len(peers) == 1:
            groups[peers[0]].extend(groups.pop(nk))

    for key in sorted(groups.keys(), key=lambda k: k[1] or 0):
        group = groups[key]

        merged_equip: dict[str, EquipmentItem] = {}
        for r in group:
            for it in r.equipment:
                norm = _normalize_equip_name(it.name)
                total_new = it.count + it.count_ae
                if norm in merged_equip:
                    total_old = merged_equip[norm].count + merged_equip[norm].count_ae
                    if total_new > total_old:
                        merged_equip[norm] = EquipmentItem(
                            symbol=it.symbol, name=it.name,
                            count=it.count, count_ae=it.count_ae,
                        )
                else:
                    merged_equip[norm] = EquipmentItem(
                        symbol=it.symbol, name=it.name,
                        count=it.count, count_ae=it.count_ae,
                    )

        best = max(group, key=lambda r: sum(
            it.count + it.count_ae for it in r.equipment
        ))
        merged_result = FileParseResult(
            filename=best.filename,
            plan_type=best.plan_type,
            elevation=key[1],
            height_category=best.height_category,
            equipment=list(merged_equip.values()),
            cables=best.cables,
            panels=best.panels,
        )
        out.append(merged_result)

        merged_total = sum(it.count + it.count_ae for it in merged_result.equipment)
        sources = [f"{r.filename}({sum(it.count+it.count_ae for it in r.equipment)})"
                   for r in group]
        if len(group) > 1:
            log(f"  [merge] {key[1]}: {' + '.join(sources)} → {merged_total} items")
        else:
            log(f"  [merge] {key[1]}: {sources[0]}")

    return out


@dataclass
class SpecGroupedItem:
    """Item from equipment specification, with unit and quantity."""
    description: str
    unit: str
    quantity: int


@dataclass
class DerivedMaterials:
    """Installation materials derived from cables, panels, and luminaires."""
    # Cable connections at panels: sum of conductor counts across all cables
    cable_connections: int = 0
    # Junction boxes for power circuits (FS 100x100x50)
    junction_boxes_power: int = 0
    # Junction/distribution boxes for lighting circuits (85x85x40)
    junction_boxes_lighting: int = 0
    # Crimp sleeves (ГМЛ) = total conductor connections at junction boxes
    crimp_sleeves: int = 0
    # Heat-shrink tubes (ТТК) = connections at feeder cable entries
    heat_shrink_tubes: int = 0
    # Fire-seal foam cartridges (DN1201)
    fire_seal_foam: int = 0
    # Foam gun (DN1202) — 1 if any foam needed
    foam_gun: int = 0
    # Steel conduit sleeves for wall penetrations (Ду20)
    steel_sleeves: int = 0
    # PVC conduit by diameter: {diameter_mm: length_m}
    pvc_conduit: dict[int, int] = field(default_factory=dict)
    # PVC conduit holders by diameter: {diameter_mm: count}
    pvc_holders: dict[int, int] = field(default_factory=dict)
    # Total PVC conduit length for work item
    pvc_total_m: int = 0


# Standard conduit diameter selection by cable cross-section (mm2)
_CABLE_SECTION_RE = re.compile(
    r"(\d+)\s*[хx×]\s*(\d+[.,]?\d*)", re.IGNORECASE,
)


def _conduit_diameter(cable_type: str) -> int:
    """Determine PVC conduit diameter (mm) from cable cross-section.

    Standard rule:
      cross-section ≤ 1.5mm2 → d16
      cross-section ≤ 2.5mm2 → d20
      cross-section ≤ 6mm2   → d25
      cross-section ≤ 10mm2  → d32
    """
    m = _CABLE_SECTION_RE.search(cable_type)
    if not m:
        return 20  # default
    section = float(m.group(2).replace(",", "."))
    if section <= 1.5:
        return 16
    if section <= 2.5:
        return 20
    if section <= 6.0:
        return 25
    return 32


def _conductor_count(cable_type: str) -> int:
    """Extract conductor count from cable type string (e.g. '3x2.5' → 3, '5x4' → 5)."""
    m = _CABLE_SECTION_RE.search(cable_type)
    if not m:
        return 3  # default assumption
    return int(m.group(1))


def _derive_installation_materials(
    panels: list[PanelInfo],
    cables: list[CableItem],
    luminaire_count: int,
    lighting_groups: int,
    cable_outlet_count: int = 0,
    log=print,
) -> DerivedMaterials:
    """Derive installation material quantities from parsed cable/panel/luminaire data.

    Rules (based on standard electrical installation practices):

    1. Cable connections at panels = sum of conductor count per cable run
       (each conductor is terminated individually at the panel).

    2. Junction boxes for lighting = luminaire_count minus end-of-line
       luminaires (1 per lighting group). Rule: "near each luminaire
       except the last on the line".

    3. Junction boxes for power circuits = number of non-lighting,
       non-reserve cable runs (sockets, AC, etc.) that need junction boxes.
       Typically 2 per socket/equipment circuit that branches.

    4. Crimp sleeves (ГМЛ) = junction_box_count × conductors_per_cable (3).

    5. Heat-shrink tubes (ТТК) = used at splice points where cable
       cross-section changes. Typically at feeder-to-distribution transitions.

    6. Steel conduit sleeves = number of cable runs (wall penetrations).

    7. Fire-seal foam = 1 cartridge if any wall penetrations exist.
       Foam gun = 1 if foam is needed.

    8. PVC conduit = cable length by conduit diameter;
       holders = length / 0.8m (standard step).
    """
    dm = DerivedMaterials()

    if not cables:
        return dm

    total_cable_runs = sum(c.count for c in cables)

    # ── 1. Cable connections at panels ──
    # Each cable run has conductors terminated at both ends (panel + device)
    # The VOR item "Подключение жил кабелей" counts individual conductor
    # connections.  Each circuit cable: conductor_count × 1 (at panel end).
    # Plus feeder cables: conductor_count × 1 each.
    circuit_connections = 0
    for c in cables:
        n_cond = _conductor_count(c.cable_type)
        circuit_connections += c.count * n_cond
    # Add feeder cable connections (from panel feed_cable field)
    feeder_connections = 0
    for p in panels:
        if p.feed_cable:
            n_cond = _conductor_count(p.feed_cable)
            feeder_connections += n_cond
    dm.cable_connections = circuit_connections + feeder_connections
    log(f"    [derived] Cable connections: {dm.cable_connections} "
        f"(circuits={circuit_connections} + feeders={feeder_connections})")

    # ── 2. Junction boxes ──
    # Lighting: one per luminaire, minus end-of-line (1 per lighting group)
    if luminaire_count > 0:
        dm.junction_boxes_lighting = max(0, luminaire_count - lighting_groups)
    # Power: junction boxes at cable outlet / branching points.
    # Cable outlets are branch connection points on the equipment plan.
    # If cable_outlet_count is available, use it directly.
    # Otherwise estimate from power circuit count.
    if cable_outlet_count > 0:
        dm.junction_boxes_power = cable_outlet_count
    else:
        power_circuit_count = 0
        for p in panels:
            for ctype, clen in p.circuit_cables:
                ct_lower = ctype.lower()
                if "1,5" in ct_lower or "1.5" in ct_lower:
                    continue
                power_circuit_count += 1
        dm.junction_boxes_power = min(power_circuit_count, 2) if power_circuit_count > 0 else 0
    log(f"    [derived] Junction boxes: lighting={dm.junction_boxes_lighting}, "
        f"power={dm.junction_boxes_power}")

    # ── 3. Crimp sleeves ──
    total_boxes = dm.junction_boxes_lighting + dm.junction_boxes_power
    dm.crimp_sleeves = total_boxes * 3  # 3 conductors per 3-wire cable
    log(f"    [derived] Crimp sleeves (ГМЛ): {dm.crimp_sleeves} "
        f"({total_boxes} boxes × 3 conductors)")

    # ── 4. Heat-shrink tubes ──
    # Used at feeder cable entries to panels where cross-section changes.
    # Typically 1 per feeder cable splice × number of spliced conductors,
    # but in practice a small fixed count (reference shows 4).
    # Rule: count of feeder cables that pass through wall penetrations
    # and need splice protection.
    dm.heat_shrink_tubes = len([p for p in panels if p.feed_cable]) * 2
    if dm.heat_shrink_tubes == 0 and total_cable_runs > 0:
        dm.heat_shrink_tubes = max(2, total_cable_runs // 3)
    log(f"    [derived] Heat-shrink tubes (ТТК): {dm.heat_shrink_tubes}")

    # ── 5. Steel conduit sleeves (wall penetrations) ──
    dm.steel_sleeves = total_cable_runs
    log(f"    [derived] Steel sleeves (Ду20): {dm.steel_sleeves}")

    # ── 6. Fire-seal foam & gun ──
    if dm.steel_sleeves > 0:
        dm.fire_seal_foam = 1
        dm.foam_gun = 1
    log(f"    [derived] Fire-seal foam: {dm.fire_seal_foam}, "
        f"gun: {dm.foam_gun}")

    # ── 7. PVC conduit quantities ──
    # Conduit length ≈ 0.9 × cable length (conduit route is ~10 % shorter
    # than total cable length because cable has slack/loops at connection
    # points).
    _PVC_LENGTH_COEFF = 0.9
    conduit_lengths: dict[int, int] = {}
    for c in cables:
        diam = _conduit_diameter(c.cable_type)
        conduit_lengths[diam] = (
            conduit_lengths.get(diam, 0)
            + round(c.total_length_m * _PVC_LENGTH_COEFF)
        )
    dm.pvc_conduit = conduit_lengths
    dm.pvc_total_m = sum(conduit_lengths.values())

    # Holders: every 0.8m
    dm.pvc_holders = {
        diam: math.ceil(length / 0.8)
        for diam, length in conduit_lengths.items()
    }
    for diam in sorted(conduit_lengths):
        log(f"    [derived] PVC d.{diam}мм: {conduit_lengths[diam]}м, "
            f"holders: {dm.pvc_holders[diam]}")

    return dm


def aggregate_by_height(
    results: list[FileParseResult],
    log=print,
) -> dict:
    """Aggregate parsed results into VOR sections.

    Returns a dict with keys:
        luminaires, indicators, schema_panels, switches, sockets,
        cable_outlets, materials, cables, spec_panels,
        grounding_items, lightning_items, pvc_items, spec_cables.
    """
    deduped = _merge_per_elevation(results, log=log)

    name_counts: dict[str, dict[HeightCategory, int]] = defaultdict(
        lambda: defaultdict(int)
    )
    name_category: dict[str, str] = {}

    for r in deduped:
        if not r.equipment:
            continue
        hcat = r.height_category
        for it in r.equipment:
            norm = _normalize_equip_name(it.name)
            cat = _classify_equipment(it.name)
            if cat == "skip":
                continue
            total = it.count + it.count_ae
            if total <= 0:
                continue
            name_category[norm] = cat
            if hcat:
                name_counts[norm][hcat] += total
            else:
                name_counts[norm]["до 5 метров"] += total

    luminaires: list[AggregatedEquipment] = []
    indicators: list[AggregatedEquipment] = []
    switches: list[AggregatedEquipment] = []
    sockets: list[AggregatedEquipment] = []
    cable_outlets: list[AggregatedEquipment] = []
    materials: list[AggregatedEquipment] = []

    for norm, heights in name_counts.items():
        cat = name_category.get(norm, "luminaire")
        total = sum(heights.values())
        agg = AggregatedEquipment(
            name=norm, category=cat, counts_by_height=dict(heights), total=total,
        )
        if cat == "luminaire":
            luminaires.append(agg)
        elif cat in ("indicator", "pictogram"):
            indicators.append(agg)
        elif cat == "switch":
            switches.append(agg)
        elif cat == "socket":
            sockets.append(agg)
        elif cat == "cable_outlet":
            cable_outlets.append(agg)
        elif cat == "material":
            materials.append(agg)

    luminaires.sort(key=lambda a: -a.total)
    indicators.sort(key=lambda a: -a.total)
    switches.sort(key=lambda a: -a.total)
    sockets.sort(key=lambda a: -a.total)
    cable_outlets.sort(key=lambda a: -a.total)
    materials.sort(key=lambda a: -a.total)

    # Collect panels from schema files
    schema_panels: list[PanelInfo] = []
    seen_panels: set[str] = set()
    for r in results:
        for p in r.panels:
            if p.name not in seen_panels:
                schema_panels.append(p)
                seen_panels.add(p.name)

    all_cables: dict[str, CableItem] = {}
    for r in deduped:
        for c in r.cables:
            if c.cable_type in all_cables:
                all_cables[c.cable_type].count += c.count
                all_cables[c.cable_type].total_length_m += c.total_length_m
            else:
                all_cables[c.cable_type] = CableItem(
                    cable_type=c.cable_type,
                    count=c.count,
                    total_length_m=c.total_length_m,
                )
    cables = sorted(all_cables.values(), key=lambda c: -c.total_length_m)

    # ── Spec items (from СО.dxf ACAD_TABLE) ──────────────────────────
    spec_panels: list[SpecGroupedItem] = []
    grounding_items: list[SpecGroupedItem] = []
    lightning_items: list[SpecGroupedItem] = []
    pvc_items: list[SpecGroupedItem] = []
    spec_cables: list[SpecGroupedItem] = []

    all_spec: list[SpecItem] = []
    for r in results:
        all_spec.extend(r.spec_items)

    def _model_in_plan(model: str, existing_names: set[str]) -> bool:
        """Check if a model name from spec already exists in plan luminaires."""
        if not model:
            return False
        ml = model.lower().strip()
        if len(ml) < 3:
            return False
        for en in existing_names:
            if ml in en:
                return True
        return False

    if all_spec:
        log(f"\n  [spec] Processing {len(all_spec)} specification items")

        plan_luminaire_names = {a.name.lower() for a in luminaires}
        plan_indicator_names = {a.name.lower() for a in indicators}
        plan_switch_names = {a.name.lower() for a in switches}
        plan_socket_names = {a.name.lower() for a in sockets}

        for si in all_spec:
            cat = _classify_spec_item(si.description)
            gi = SpecGroupedItem(
                description=si.description, unit=si.unit, quantity=si.quantity,
            )
            if cat == "panel":
                spec_panels.append(gi)
            elif cat == "grounding":
                grounding_items.append(gi)
            elif cat == "lightning":
                lightning_items.append(gi)
            elif cat == "pvc_conduit":
                pvc_items.append(gi)
            elif cat == "luminaire":
                norm = _normalize_equip_name(si.description)
                if _model_in_plan(si.model, plan_luminaire_names):
                    log(f"    [spec=] luminaire (plan has): "
                        f"{si.model or norm[:40]}")
                else:
                    luminaires.append(AggregatedEquipment(
                        name=norm, category="luminaire",
                        counts_by_height={"до 5 метров": si.quantity},
                        total=si.quantity,
                    ))
                    plan_luminaire_names.add(norm.lower())
                    log(f"    [spec+] luminaire: {norm[:70]} × {si.quantity}")
            elif cat in ("indicator", "pictogram"):
                norm = _normalize_equip_name(si.description)
                if norm.lower() not in plan_indicator_names:
                    indicators.append(AggregatedEquipment(
                        name=norm, category=cat,
                        counts_by_height={"до 5 метров": si.quantity},
                        total=si.quantity,
                    ))
                    plan_indicator_names.add(norm.lower())
                    log(f"    [spec+] indicator: {norm[:70]} × {si.quantity}")
            elif cat == "switch":
                norm = _normalize_equip_name(si.description)
                if norm.lower() not in plan_switch_names:
                    switches.append(AggregatedEquipment(
                        name=norm, category="switch",
                        counts_by_height={"до 5 метров": si.quantity},
                        total=si.quantity,
                    ))
                    plan_switch_names.add(norm.lower())
                    log(f"    [spec+] switch: {norm[:70]} × {si.quantity}")
            elif cat == "socket":
                norm = _normalize_equip_name(si.description)
                if norm.lower() not in plan_socket_names:
                    sockets.append(AggregatedEquipment(
                        name=norm, category="socket",
                        counts_by_height={"до 5 метров": si.quantity},
                        total=si.quantity,
                    ))
                    plan_socket_names.add(norm.lower())
                    log(f"    [spec+] socket: {norm[:70]} × {si.quantity}")
            elif cat == "cable":
                spec_cables.append(gi)
                log(f"    [spec+] cable: {si.description[:60]} × {si.quantity} {si.unit}")
            elif cat == "material":
                materials.append(AggregatedEquipment(
                    name=si.description, category="material",
                    counts_by_height={"до 5 метров": si.quantity},
                    total=si.quantity,
                    unit=si.unit,
                ))

    # ── Derive installation materials from cables/panels/luminaires ──
    total_luminaires = sum(lum.total for lum in luminaires)
    # Count lighting groups = number of unique lighting circuit cables
    lighting_groups_count = 0
    for p in schema_panels:
        for ctype, clen in p.circuit_cables:
            ct_lower = ctype.lower()
            if "1,5" in ct_lower or "1.5" in ct_lower:
                lighting_groups_count += 1
    if lighting_groups_count == 0 and total_luminaires > 0:
        lighting_groups_count = max(1, len(schema_panels))

    total_cable_outlets = sum(co.total for co in cable_outlets)

    log(f"\n  [derived] Deriving installation materials...")
    derived = _derive_installation_materials(
        panels=schema_panels,
        cables=cables,
        luminaire_count=total_luminaires,
        lighting_groups=lighting_groups_count,
        cable_outlet_count=total_cable_outlets,
        log=log,
    )

    log(f"\n  [VOR] Luminaires:    {len(luminaires)} types")
    log(f"  [VOR] Indicators:    {len(indicators)} types")
    log(f"  [VOR] Panels:        {len(schema_panels)} from schema")
    log(f"  [VOR] Spec panels:   {len(spec_panels)}")
    log(f"  [VOR] Switches:      {len(switches)} types")
    log(f"  [VOR] Sockets:       {len(sockets)} types")
    log(f"  [VOR] Cable outlets: {len(cable_outlets)} types")
    log(f"  [VOR] Materials:     {len(materials)} types")
    log(f"  [VOR] Cables:        {len(cables)} types, "
        f"{sum(c.total_length_m for c in cables):.0f}m")
    log(f"  [VOR] Grounding:     {len(grounding_items)} items")
    log(f"  [VOR] Lightning:     {len(lightning_items)} items")
    log(f"  [VOR] PVC conduits:  {len(pvc_items)} items")
    log(f"  [VOR] Spec cables:   {len(spec_cables)} items")
    log(f"  [VOR] Derived materials: connections={derived.cable_connections}, "
        f"boxes={derived.junction_boxes_lighting}+{derived.junction_boxes_power}, "
        f"sleeves={derived.crimp_sleeves}, steel={derived.steel_sleeves}")

    return {
        "luminaires": luminaires,
        "indicators": indicators,
        "schema_panels": schema_panels,
        "switches": switches,
        "sockets": sockets,
        "cable_outlets": cable_outlets,
        "materials": materials,
        "cables": cables,
        "spec_panels": spec_panels,
        "grounding_items": grounding_items,
        "lightning_items": lightning_items,
        "pvc_items": pvc_items,
        "spec_cables": spec_cables,
        "derived": derived,
    }


# ── DOCX Generation ──────────────────────────────────────────────────

_COL_WIDTHS_CM = [1.2, 9.5, 1.5, 1.8, 4.5, 4.5, 3.0]
_HEADERS = [
    "№ п/п", "Наименование вида работ", "Ед. изм.",
    "Объем работ", "Формула расчета объемов работ и расхода материалов",
    "Ссылка на чертежи, спецификации", "Дополнительная информация",
]
_FONT_NAME = "Times New Roman"
_FONT_SIZE_BODY = 9
_FONT_SIZE_HEADER = 10


def _set_cell(cell, text, bold=False, font_size=_FONT_SIZE_BODY, align="left"):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.font.name = _FONT_NAME
    run.font.size = Pt(font_size)
    run.bold = bold
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    margins = {
        "top": "28", "bottom": "28", "start": "57", "end": "57",
    }
    for side, val in margins.items():
        tcPr.set(qn(f"w:{side}"), val)


def _add_section_header(table, text):
    row = table.add_row()
    _set_cell(row.cells[0], "", bold=True)
    _set_cell(row.cells[1], text, bold=True)
    for i in range(2, 7):
        _set_cell(row.cells[i], "")
    return row


def _add_work_row(table, num, description, unit, qty, formula="", ref="", info=""):
    row = table.add_row()
    _set_cell(row.cells[0], str(num) if num else "", align="center")
    _set_cell(row.cells[1], description)
    _set_cell(row.cells[2], unit, align="center")
    _set_cell(row.cells[3], str(qty) if qty else "", align="center")
    _set_cell(row.cells[4], formula)
    _set_cell(row.cells[5], ref)
    _set_cell(row.cells[6], info)
    return row


def _add_material_row(table, description, unit, qty, ref=""):
    row = table.add_row()
    _set_cell(row.cells[0], "")
    _set_cell(row.cells[1], description)
    _set_cell(row.cells[2], unit, align="center")
    _set_cell(row.cells[3], str(qty) if qty else "", align="center")
    _set_cell(row.cells[4], "")
    _set_cell(row.cells[5], ref)
    _set_cell(row.cells[6], "")
    return row


def generate_vor_docx(
    luminaires: list[AggregatedEquipment],
    indicators: list[AggregatedEquipment],
    panels: list[PanelInfo],
    switches: list[AggregatedEquipment],
    cables: list[CableItem],
    output_path: str,
    project_name: str = "",
    section_name: str = "Электроосвещение",
    drawing_ref: str = "",
    log=print,
    sockets: list[AggregatedEquipment] | None = None,
    cable_outlets: list[AggregatedEquipment] | None = None,
    materials: list[AggregatedEquipment] | None = None,
    spec_panels: list[SpecGroupedItem] | None = None,
    grounding_items: list[SpecGroupedItem] | None = None,
    lightning_items: list[SpecGroupedItem] | None = None,
    pvc_items: list[SpecGroupedItem] | None = None,
    spec_cables: list[SpecGroupedItem] | None = None,
    derived: DerivedMaterials | None = None,
) -> str:
    """Generate a VOR .docx file from aggregated data."""
    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = _FONT_NAME
    style.font.size = Pt(_FONT_SIZE_BODY)

    section = doc.sections[0]
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(1.0)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.0)

    # ── Title block ──
    if project_name:
        p = doc.add_paragraph()
        run = p.add_run(project_name)
        run.font.size = Pt(12)
        run.font.name = _FONT_NAME
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    run = p.add_run("Ведомость объемов работ")
    run.font.size = Pt(14)
    run.font.name = _FONT_NAME
    run.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if section_name:
        p = doc.add_paragraph()
        run = p.add_run(f"Основание_{section_name}")
        run.font.size = Pt(10)
        run.font.name = _FONT_NAME
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p = doc.add_paragraph()
    run = p.add_run(f"Дата составления {time.strftime('%d.%m.%Y')}г.")
    run.font.size = Pt(10)
    run.font.name = _FONT_NAME

    doc.add_paragraph()

    # ── Main table ──
    table = doc.add_table(rows=0, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"

    header_row = table.add_row()
    for i, hdr in enumerate(_HEADERS):
        _set_cell(header_row.cells[i], hdr, bold=True,
                  font_size=_FONT_SIZE_HEADER, align="center")

    num_row = table.add_row()
    for i in range(7):
        _set_cell(num_row.cells[i], str(i + 1), align="center")

    for i, w_cm in enumerate(_COL_WIDTHS_CM):
        for row in table.rows:
            row.cells[i].width = Cm(w_cm)

    item_num = 0

    # ── Section 1: Panels ──
    has_panels = panels or (spec_panels and len(spec_panels) > 0)
    if has_panels:
        _add_section_header(table, "Щитовое оборудование")
        schema_panel_names = set()
        if panels:
            for p in panels:
                item_num += 1
                _add_work_row(
                    table, item_num,
                    f"Монтаж щита распределительного {p.name}",
                    "шт", 1, ref=drawing_ref,
                )
                schema_panel_names.add(p.name.lower().replace("-", "").replace(" ", ""))
        if spec_panels:
            for sp in spec_panels:
                sp_norm = sp.description.lower().replace("-", "").replace(" ", "")
                if any(sn in sp_norm or sp_norm in sn for sn in schema_panel_names):
                    continue
                item_num += 1
                _add_work_row(
                    table, item_num,
                    f"Монтаж {sp.description}",
                    sp.unit, sp.quantity, ref=drawing_ref,
                )
        # Cable connections at panels (derived)
        if derived and derived.cable_connections > 0:
            item_num += 1
            _add_work_row(
                table, item_num,
                "Подключение жил кабелей до 10 мм2",
                "шт", derived.cable_connections, ref=drawing_ref,
            )

    # ── Section 2: Lighting equipment by height ──
    has_lighting = any(
        lum.counts_by_height.get(hcat, 0) > 0
        for lum in luminaires
        for hcat in HEIGHT_CATEGORIES
    ) or indicators
    if has_lighting:
        _add_section_header(table, "Светотехническое оборудование")

    for hcat in HEIGHT_CATEGORIES:
        items_in_height = [
            (lum, lum.counts_by_height.get(hcat, 0))
            for lum in luminaires
            if lum.counts_by_height.get(hcat, 0) > 0
        ]
        if not items_in_height:
            continue

        height_total = sum(cnt for _, cnt in items_in_height)
        item_num += 1
        _add_work_row(
            table, item_num,
            f"Монтаж светильников на шпильках к перекрытию "
            f"на высоте {hcat}:",
            "шт", height_total, ref=drawing_ref,
        )
        for lum, cnt in items_in_height:
            desc = f"Светодиодный светильник {lum.name}"
            _add_material_row(table, desc, "шт", cnt, ref=drawing_ref)

    # Wall-mounted indicators by height
    has_indicators = any(
        ind.counts_by_height.get(hcat, 0) > 0
        for ind in indicators
        for hcat in HEIGHT_CATEGORIES
        if ind.category == "indicator"
    )
    if has_indicators:
        for hcat in HEIGHT_CATEGORIES:
            ind_in_height = [
                (ind, ind.counts_by_height.get(hcat, 0))
                for ind in indicators
                if ind.category == "indicator"
                and ind.counts_by_height.get(hcat, 0) > 0
            ]
            if not ind_in_height:
                continue

            height_total = sum(cnt for _, cnt in ind_in_height)
            item_num += 1
            _add_work_row(
                table, item_num,
                f"Монтаж настенного указателя светодиодного "
                f"(с пиктограммой) {hcat}:",
                "шт", height_total, ref=drawing_ref,
            )
            for ind, cnt in ind_in_height:
                _add_material_row(table, ind.name, "шт", cnt, ref=drawing_ref)

            pictograms_in_height = [
                (p, p.counts_by_height.get(hcat, 0))
                for p in indicators
                if p.category == "pictogram"
                and p.counts_by_height.get(hcat, 0) > 0
            ]
            for pic, cnt in pictograms_in_height:
                _add_material_row(table, pic.name, "шт", cnt, ref=drawing_ref)

    # Standalone pictograms (if not already added with indicators)
    standalone_pics = [
        p for p in indicators
        if p.category == "pictogram" and not has_indicators
    ]
    for pic in standalone_pics:
        item_num += 1
        _add_work_row(
            table, item_num, pic.name,
            "шт", pic.total, ref=drawing_ref,
        )

    # ── Section 3: Electrical devices ──
    has_derived_boxes = (derived and (
        derived.junction_boxes_power > 0
        or derived.junction_boxes_lighting > 0
        or derived.crimp_sleeves > 0
        or derived.heat_shrink_tubes > 0
    ))
    has_devices = (switches
                   or (sockets and any(s.total > 0 for s in sockets))
                   or (cable_outlets and any(c.total > 0 for c in cable_outlets))
                   or has_derived_boxes)
    if has_devices:
        _add_section_header(table, "Монтаж электроустановочных изделий")
        for sw in (switches or []):
            item_num += 1
            _add_work_row(
                table, item_num, f"Монтаж {sw.name}",
                "шт", sw.total, ref=drawing_ref,
            )
        for sock in (sockets or []):
            if sock.total > 0:
                item_num += 1
                _add_work_row(
                    table, item_num, f"Монтаж {sock.name}",
                    "шт", sock.total, ref=drawing_ref,
                )
        for co in (cable_outlets or []):
            if co.total > 0:
                item_num += 1
                _add_work_row(
                    table, item_num, f"Монтаж {co.name}",
                    "шт", co.total, ref=drawing_ref,
                )
        # Derived junction boxes
        if derived and derived.junction_boxes_power > 0:
            item_num += 1
            _add_work_row(
                table, item_num,
                "Монтаж коробки соединительной с кабельными "
                "вводами 100x100x50 мм",
                "шт", derived.junction_boxes_power, ref=drawing_ref,
            )
        if derived and derived.junction_boxes_lighting > 0:
            item_num += 1
            _add_work_row(
                table, item_num,
                "Монтаж коробки распределительной "
                "85x85x40 мм",
                "шт", derived.junction_boxes_lighting, ref=drawing_ref,
            )
        # Derived crimping materials (sub-section)
        if derived and derived.crimp_sleeves > 0:
            item_num += 1
            _add_work_row(
                table, item_num,
                "Соединение жил кабелей методом опрессовки",
                "", "", ref=drawing_ref,
            )
            _add_material_row(
                table,
                "Луженая гильза ГМЛ",
                "шт", derived.crimp_sleeves, ref=drawing_ref,
            )
            if derived.heat_shrink_tubes > 0:
                _add_material_row(
                    table,
                    "Термоусадочная трубка ТТК (4:1)",
                    "шт", derived.heat_shrink_tubes, ref=drawing_ref,
                )

    # ── Section 3b: Materials from spec ──
    if materials:
        _add_section_header(table, "Монтажные изделия и материалы")
        for mat in materials:
            item_num += 1
            _add_work_row(
                table, item_num, mat.name,
                mat.unit, mat.total, ref=drawing_ref,
            )

    # ── Section 4: Cables ──
    if cables:
        _add_section_header(table, "Кабельная продукция")

        fire_cables = [c for c in cables if "FRHF" in c.cable_type or "FR" in c.cable_type]
        normal_cables = [c for c in cables if c not in fire_cables]

        if normal_cables:
            total_normal = sum(c.total_length_m for c in normal_cables)
            item_num += 1
            _add_work_row(
                table, item_num,
                "Прокладка кабеля в лотке/гофре/трубе:",
                "м", total_normal, ref=drawing_ref,
            )
            for c in normal_cables:
                _add_material_row(
                    table, f"Кабель {c.cable_type}",
                    "м", c.total_length_m, ref=drawing_ref,
                )

        if fire_cables:
            total_fire = sum(c.total_length_m for c in fire_cables)
            item_num += 1
            _add_work_row(
                table, item_num,
                "Прокладка огнестойкого кабеля в лотке/гофре/трубе:",
                "м", total_fire, ref=drawing_ref,
            )
            for c in fire_cables:
                _add_material_row(
                    table, f"Кабель {c.cable_type}",
                    "м", c.total_length_m, ref=drawing_ref,
                )

        total_cable_m = sum(c.total_length_m for c in cables)
        total_cable_cnt = sum(c.count for c in cables)

        item_num += 1
        _add_work_row(
            table, item_num,
            "Затяжка кабеля в трубы, блоки и на лотки",
            "м", total_cable_m,
            formula=f"{total_cable_cnt} линий × средняя длина",
            ref=drawing_ref,
        )
    if spec_cables:
        if not cables:
            _add_section_header(table, "Кабельная продукция (по спецификации)")
        existing_cable_types = {c.cable_type.lower() for c in (cables or [])}
        new_spec = [sc for sc in spec_cables
                    if not any(et in sc.description.lower() for et in existing_cable_types)]
        if new_spec:
            if cables:
                _add_section_header(table, "Кабельная продукция (дополнительно по спецификации)")
            for sc in new_spec:
                item_num += 1
                _add_work_row(
                    table, item_num, f"Кабель {sc.description}",
                    sc.unit, sc.quantity, ref=drawing_ref,
                )

    # ── Section 5: Cable penetrations (fire sealing) ──
    has_penetrations = derived and derived.steel_sleeves > 0
    if has_penetrations:
        item_num += 1
        _add_work_row(
            table, item_num,
            "Выполнение проходки кабеля через стены",
            "", "", ref=drawing_ref,
        )
        if derived.fire_seal_foam > 0:
            _add_material_row(
                table,
                "Двухкомпонентная огнестойкая пена, DN1201",
                "шт", derived.fire_seal_foam, ref=drawing_ref,
            )
        if derived.foam_gun > 0:
            _add_material_row(
                table,
                "Пистолет для двухкомпонентной пены, DN1202",
                "шт", derived.foam_gun, ref=drawing_ref,
            )
        _add_material_row(
            table,
            "Гильза закладная труба сталь ВГП Ду 20 "
            "(Дн 26,8x2,5) ГОСТ 3262-75",
            "шт", derived.steel_sleeves, ref=drawing_ref,
        )

    # ── Section 6: PVC conduits ──
    _add_section_header(table, "Монтаж ПВХ изделий и труб")
    if pvc_items:
        for pv in pvc_items:
            item_num += 1
            _add_work_row(
                table, item_num, pv.description,
                pv.unit, pv.quantity, ref=drawing_ref,
            )
    elif derived and derived.pvc_total_m > 0:
        # Use derived PVC quantities based on cable data
        item_num += 1
        _add_work_row(
            table, item_num,
            "Монтаж гофрированной трубы ПВХ гибкой с креплением "
            "клипсами каждые 0,8 м",
            "м", derived.pvc_total_m, ref=drawing_ref,
        )
        for diam in sorted(derived.pvc_conduit):
            _add_material_row(
                table,
                f"Труба ПВХ гибкая гофр. д.{diam}мм",
                "м", derived.pvc_conduit[diam], ref=drawing_ref,
            )
        for diam in sorted(derived.pvc_holders):
            _add_material_row(
                table,
                f"Держатель оцинкованный двусторонний, "
                f"д.{diam}мм",
                "шт", derived.pvc_holders[diam], ref=drawing_ref,
            )
    elif cables:
        gofra_cnt = sum(c.count for c in cables)
        item_num += 1
        _add_work_row(
            table, item_num,
            "Затяжка кабеля в гофрированные трубы",
            "м", sum(c.total_length_m for c in cables),
            formula=f"~{gofra_cnt} линий (уточнить по планам)",
            ref=drawing_ref,
        )
        _add_material_row(
            table,
            "[Гофротруба ПВХ -- марка и количество по кабеленесущим планам]",
            "м", "",
        )
    else:
        _add_material_row(
            table,
            "[Заполнить вручную]",
            "", "",
        )

    # ── Section 7: Grounding ──
    if grounding_items:
        _add_section_header(table, "Заземление")
        for gi in grounding_items:
            item_num += 1
            _add_work_row(
                table, item_num, gi.description,
                gi.unit, gi.quantity, ref=drawing_ref,
            )
    else:
        _add_section_header(table, "Заземление")
        _add_material_row(
            table,
            "[Заполнить вручную из проекта заземления]",
            "", "",
        )

    # ── Section 7b: Lightning protection ──
    if lightning_items:
        _add_section_header(table, "Молниезащита")
        for li in lightning_items:
            item_num += 1
            _add_work_row(
                table, item_num, li.description,
                li.unit, li.quantity, ref=drawing_ref,
            )
    else:
        _add_section_header(table, "Молниезащита")
        _add_material_row(
            table,
            "[Заполнить вручную из проекта молниезащиты]",
            "", "",
        )

    # ── Section 8: Commissioning (PNR) ──
    _add_section_header(table, "Пусконаладочные работы")

    cable_line_count = sum(c.count for c in cables) if cables else 0
    # Add panel feed cables (each panel with a feed_cable adds 1 line)
    cable_line_count += sum(1 for p in panels if p.feed_cable)

    if cable_line_count > 0:
        # 1) Insulation resistance measurement
        item_num += 1
        _add_work_row(
            table, item_num,
            "Измерение сопротивления изоляции",
            "каб.", cable_line_count,
            ref=drawing_ref,
        )

        # 2) Cable continuity and phasing
        item_num += 1
        _add_work_row(
            table, item_num,
            "Определение целостности жил кабеля и фазировка "
            "кабельной линии",
            "каб.", cable_line_count,
            ref=drawing_ref,
        )

        # 3) Grounding continuity check
        item_num += 1
        _add_work_row(
            table, item_num,
            "Проверка наличия цепи между заземлителями "
            "и заземленными элементами",
            "изм.", cable_line_count * 2,
            ref=drawing_ref,
        )

        # 4) Mobile testing lab
        lab_hours = math.ceil(cable_line_count / 3)
        item_num += 1
        _add_work_row(
            table, item_num,
            "Лаборатория передвижная монтажно-измерительная",
            "маш/час", lab_hours,
            ref=drawing_ref,
        )

    # 5) Circuit breaker testing per panel
    if panels:
        for p in panels:
            single_pole = p.circuit_count or len(p.circuit_cables) or 1
            # The main panel breaker (QF in p.breaker) is three-pole (1 unit)
            three_pole = 1 if p.breaker else 0
            qty_str = (f"{single_pole}/{three_pole}"
                       if three_pole else str(single_pole))
            item_num += 1
            _add_work_row(
                table, item_num,
                f"Проверка срабатывания автоматических выключателей "
                f"в щите {p.name} (однополюсных/трехполюсных)",
                "шт", qty_str,
                ref=drawing_ref,
            )

    # 6) Lighting network verification
    if panels:
        item_num += 1
        _add_work_row(
            table, item_num,
            "Проверка осветительной сети на правильность "
            "зажигания групп внутреннего освещения",
            "шт", len(panels),
            ref=drawing_ref,
        )

    # ── Save ──
    doc.save(output_path)
    log(f"  [VOR] Saved: {output_path}")
    return output_path


# ── Pipeline ─────────────────────────────────────────────────────────

def generate_vor(
    folder: Path,
    output_path: str | None = None,
    project_name: str = "",
    section_name: str = "Электроосвещение",
    drawing_ref: str = "",
    log=print,
) -> str:
    """Full pipeline: scan folder -> parse -> aggregate -> generate .docx."""
    if output_path is None:
        output_path = str(folder / "ВОР_ЭО.docx")

    log(f"\n{'=' * 50}")
    log(f"  Генерация ВОР: {folder.name}")
    log(f"{'=' * 50}")

    t0 = time.time()

    log("\n[1] Сканирование файлов")
    file_list = scan_and_classify(folder)
    by_type = defaultdict(int)
    for _, pt, _ in file_list:
        by_type[pt] += 1
    for pt, cnt in sorted(by_type.items()):
        log(f"  {pt}: {cnt}")

    log("\n[2] Парсинг чертежей")
    results = parse_all_files(file_list, log=log)

    log("\n[3] Агрегация по высотам")
    agg = aggregate_by_height(results, log=log)

    luminaires = agg["luminaires"]
    log("\n  --- Светильники по высотам ---")
    for hcat in HEIGHT_CATEGORIES:
        items = [(l, l.counts_by_height.get(hcat, 0)) for l in luminaires
                 if l.counts_by_height.get(hcat, 0) > 0]
        if items:
            total = sum(c for _, c in items)
            log(f"  {hcat}: {total}")
            for l, c in items:
                log(f"    {c:>4}  {l.name[:60]}")

    log("\n[4] Генерация .docx")
    generate_vor_docx(
        agg["luminaires"], agg["indicators"], agg["schema_panels"],
        agg["switches"], agg["cables"],
        output_path,
        project_name=project_name,
        section_name=section_name,
        drawing_ref=drawing_ref,
        log=log,
        sockets=agg["sockets"],
        cable_outlets=agg["cable_outlets"],
        materials=agg["materials"],
        spec_panels=agg["spec_panels"],
        grounding_items=agg["grounding_items"],
        lightning_items=agg["lightning_items"],
        pvc_items=agg["pvc_items"],
        spec_cables=agg["spec_cables"],
        derived=agg.get("derived"),
    )

    elapsed = time.time() - t0
    log(f"\n  Готово за {elapsed:.1f}с")
    return output_path


# ── Multi-section (combined) pipeline ────────────────────────────────

def _discover_sections(parent: Path) -> list[tuple[str, Path]]:
    """Find section subfolders (e.g. ЭО, ЭМ, ЭГ) that contain _converted_dxf."""
    sections = []
    if not parent.is_dir():
        return sections
    for sub in sorted(parent.iterdir()):
        if not sub.is_dir():
            continue
        cdxf = sub / "_converted_dxf"
        if cdxf.is_dir():
            dxf_count = sum(1 for f in cdxf.iterdir() if f.suffix.lower() == ".dxf")
            if dxf_count > 0:
                sections.append((sub.name, cdxf))
    return sections


def find_dwg_parent(folder: Path) -> Path | None:
    """Walk up from folder to find the DWG-level directory containing
    at least 2 section subfolders with _converted_dxf."""
    for candidate in [folder, folder.parent, folder.parent.parent]:
        if not candidate.is_dir():
            continue
        sections = _discover_sections(candidate)
        if len(sections) >= 2:
            return candidate
    return None


def generate_vor_combined(
    parent_folder: Path,
    output_path: str | None = None,
    project_name: str = "",
    drawing_ref: str = "",
    log=print,
) -> str:
    """Multi-section pipeline: scan ЭО+ЭМ+ЭГ -> parse -> aggregate -> .docx.

    ``parent_folder`` should be the directory that contains section
    subfolders (ЭО, ЭМ, ЭГ …), each with their own ``_converted_dxf``.
    """
    sections = _discover_sections(parent_folder)
    if not sections:
        raise ValueError(f"Не найдены разделы с _converted_dxf в {parent_folder}")

    section_names_str = " + ".join(name for name, _ in sections)
    if output_path is None:
        safe_name = "_".join(name for name, _ in sections)
        output_path = str(parent_folder / f"ВОР_{safe_name}.docx")

    log(f"\n{'=' * 60}")
    log(f"  Генерация объединённого ВОР: {section_names_str}")
    log(f"{'=' * 60}")

    t0 = time.time()

    all_results: list[FileParseResult] = []

    for sec_name, sec_folder in sections:
        log(f"\n{'─' * 50}")
        log(f"  Раздел: {sec_name}  ({sec_folder})")
        log(f"{'─' * 50}")

        log("\n  [1] Сканирование файлов")
        file_list = scan_and_classify(sec_folder)
        by_type: dict[str, int] = defaultdict(int)
        for _, pt, _ in file_list:
            by_type[pt] += 1
        for pt, cnt in sorted(by_type.items()):
            log(f"    {pt}: {cnt}")

        log("\n  [2] Парсинг чертежей")
        results = parse_all_files(file_list, log=log)
        all_results.extend(results)

    log(f"\n{'─' * 50}")
    log(f"  Агрегация (все разделы)")
    log(f"{'─' * 50}")
    agg = aggregate_by_height(all_results, log=log)

    luminaires = agg["luminaires"]
    log("\n  --- Светильники по высотам ---")
    for hcat in HEIGHT_CATEGORIES:
        items = [(l, l.counts_by_height.get(hcat, 0)) for l in luminaires
                 if l.counts_by_height.get(hcat, 0) > 0]
        if items:
            total = sum(c for _, c in items)
            log(f"  {hcat}: {total}")
            for l, c in items:
                log(f"    {c:>4}  {l.name[:60]}")

    log("\n  [3] Генерация .docx")
    generate_vor_docx(
        agg["luminaires"], agg["indicators"], agg["schema_panels"],
        agg["switches"], agg["cables"],
        output_path,
        project_name=project_name,
        section_name=section_names_str,
        drawing_ref=drawing_ref,
        log=log,
        sockets=agg["sockets"],
        cable_outlets=agg["cable_outlets"],
        materials=agg["materials"],
        spec_panels=agg["spec_panels"],
        grounding_items=agg["grounding_items"],
        lightning_items=agg["lightning_items"],
        pvc_items=agg["pvc_items"],
        spec_cables=agg["spec_cables"],
        derived=agg.get("derived"),
    )

    elapsed = time.time() - t0
    log(f"\n  Объединённый ВОР готов за {elapsed:.1f}с")
    return output_path


# ── CLI ──────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Генерация ВОР (Ведомость объёмов работ) из DXF чертежей"
    )
    parser.add_argument("folder", help="Папка с DXF файлами")
    parser.add_argument("-o", "--output", default=None, help="Путь .docx")
    parser.add_argument("--project", default="", help="Наименование стройки")
    parser.add_argument("--section", default="Электроосвещение",
                        help="Наименование раздела")
    parser.add_argument("--ref", default="", help="Ссылка на чертежи")
    args = parser.parse_args()

    folder = Path(args.folder).resolve()
    if not folder.is_dir():
        sys.exit(f"Папка не найдена: {folder}")

    generate_vor(
        folder,
        output_path=args.output,
        project_name=args.project,
        section_name=args.section,
        drawing_ref=args.ref,
    )


if __name__ == "__main__":
    main()
