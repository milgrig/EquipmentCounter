"""T064: DXF-first grounding facts and canonical VOR rows (Lever 4)."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field

try:
    import ezdxf
except ImportError:
    ezdxf = None  # type: ignore[assignment]

from equipment_counter import parse_spec_dxf

_GROUND_LAYER_RE = re.compile(
    r"(заземл|зазем|^pe$|^зк$|молниез|earth)",
    re.IGNORECASE,
)
_GROUND_INSERT_RE = re.compile(
    r"(стерж|наконеч|соедин|диагонал|sds|max|мпп-?40[хx×]4|"
    r"плоск|проводник|лента|цинк|спрей|заземл|электрод|головк)",
    re.IGNORECASE,
)
_GROUND_ROD_RE = re.compile(r"(стерж|электрод|заземлител)", re.IGNORECASE)
_GROUND_TIP_RE = re.compile(r"наконеч", re.IGNORECASE)
_GROUND_HEAD_RE = re.compile(r"(головк|sds)", re.IGNORECASE)
_GROUND_DIAG_RE = re.compile(r"диагон", re.IGNORECASE)
_GROUND_FLAT_CONN_RE = re.compile(
    r"(плоск.*соедин|соедин.*плоск|40\s*мм)",
    re.IGNORECASE,
)
_GROUND_STRIP_RE = re.compile(
    r"(полос|проводник.*40|40\s*[xх×]\s*4|мпп)",
    re.IGNORECASE,
)
_GROUND_TAPE_RE = re.compile(r"(антикорроз|лента.*50)", re.IGNORECASE)
_GROUND_SPRAY_RE = re.compile(r"(спрей|свартон|цинк)", re.IGNORECASE)
_GROUND_BUNDLE_LEN_RE = re.compile(
    r"бухт[аы]\s*(\d+(?:[.,]\d+)?)\s*м",
    re.IGNORECASE,
)
_GROUND_SPEC_RE = re.compile(
    r"(заземлен|стержень\s+заземления|наконечник\s+стержня|"
    r"забивная\s+головка|соединитель\s+диагональн|соединитель\s+плоского|"
    r"плоский\s+проводник|антикорроз|спрей\s+цинк|цинковый\s+спрей)",
    re.IGNORECASE,
)

_EARTHWORK_M3_PER_ROD = 3.2
_PERIMETER_M_PER_ROD = 22.0
_TRENCH_DEPTH_FACTOR = 0.15
_SAND_DEPTH_FACTOR = 0.03


@dataclass
class GroundingFact:
    kind: str
    qty: int = 0
    length_m: float = 0.0
    derived_qty: int = 0
    description: str = ""


@dataclass
class GroundingVorRow:
    description: str
    unit: str
    quantity: int | float


@dataclass
class _GroundingCounts:
    rods: int = 0
    tips: int = 0
    heads: int = 0
    diag_connectors: int = 0
    flat_connectors: int = 0
    strip_m: float = 0.0
    tape: int = 0
    spray: int = 0
    spec_rows: list[tuple[str, str, int | float]] = field(default_factory=list)


def _entity_layer(entity) -> str:
    return str(getattr(entity.dxf, "layer", "") or "")


def _polyline_length(entity) -> float:
    try:
        if entity.dxftype() == "LINE":
            p1, p2 = entity.dxf.start, entity.dxf.end
            return math.hypot(p2.x - p1.x, p2.y - p1.y)
        if hasattr(entity, "length"):
            return float(entity.length())
        if entity.dxftype() == "LWPOLYLINE":
            pts = list(entity.get_points("xy"))
            total = 0.0
            for i in range(1, len(pts)):
                total += math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1])
            return total
    except Exception:
        return 0.0
    return 0.0


def _classify_insert_name(name: str) -> str | None:
    if _GROUND_ROD_RE.search(name):
        return "rod"
    if _GROUND_TIP_RE.search(name):
        return "tip"
    if _GROUND_HEAD_RE.search(name):
        return "head"
    if _GROUND_DIAG_RE.search(name):
        return "diag_connector"
    if _GROUND_FLAT_CONN_RE.search(name):
        return "flat_connector"
    if _GROUND_STRIP_RE.search(name):
        return "strip"
    if _GROUND_TAPE_RE.search(name):
        return "tape"
    if _GROUND_SPRAY_RE.search(name):
        return "spray"
    return None


def _classify_spec_description(desc: str) -> str | None:
    dl = desc.lower()
    if "стержень заземления" in dl or (
        "стерж" in dl and "зазем" in dl
    ):
        return "rod"
    if "наконечник" in dl and "стерж" in dl:
        return "tip"
    if "забивная головка" in dl or ("головк" in dl and "sds" in dl):
        return "head"
    if "соединитель диагональн" in dl:
        return "diag_connector"
    if "соединитель плоского" in dl:
        return "flat_connector"
    if "плоский проводник" in dl or "40х4" in dl or "40x4" in dl:
        return "strip"
    if "антикорроз" in dl and "лента" in dl:
        return "tape"
    if "спрей" in dl and "цинк" in dl:
        return "spray"
    return None


def _parse_spec_qty(desc: str, qty: int | float, unit: str) -> tuple[int | float, float]:
    length_m = 0.0
    if unit == "м":
        length_m = float(qty)
    m = _GROUND_BUNDLE_LEN_RE.search(desc)
    if m and unit == "м":
        try:
            length_m = max(length_m, float(m.group(1).replace(",", ".")))
        except ValueError:
            pass
    return qty, length_m


def _merge_counts(counts: _GroundingCounts, other: _GroundingCounts) -> None:
    counts.rods = max(counts.rods, other.rods)
    counts.tips = max(counts.tips, other.tips)
    counts.heads = max(counts.heads, other.heads)
    counts.diag_connectors = max(counts.diag_connectors, other.diag_connectors)
    counts.flat_connectors = max(counts.flat_connectors, other.flat_connectors)
    counts.strip_m = max(counts.strip_m, other.strip_m)
    counts.tape = max(counts.tape, other.tape)
    counts.spray = max(counts.spray, other.spray)
    for row in other.spec_rows:
        if row not in counts.spec_rows:
            counts.spec_rows.append(row)


def _extract_dxf_signals(dxf_path: str) -> _GroundingCounts:
    counts = _GroundingCounts()
    if ezdxf is None:
        return counts
    try:
        doc = ezdxf.readfile(dxf_path)
    except Exception:
        return counts

    for entity in doc.modelspace():
        layer = _entity_layer(entity)
        on_ground_layer = bool(_GROUND_LAYER_RE.search(layer))
        if entity.dxftype() == "INSERT":
            name = str(entity.dxf.name or "")
            if not on_ground_layer and not _GROUND_INSERT_RE.search(name):
                continue
            kind = _classify_insert_name(name)
            if kind == "rod":
                counts.rods += 1
            elif kind == "tip":
                counts.tips += 1
            elif kind == "head":
                counts.heads += 1
            elif kind == "diag_connector":
                counts.diag_connectors += 1
            elif kind == "flat_connector":
                counts.flat_connectors += 1
            elif kind == "strip":
                counts.strip_m += _polyline_length(entity)
            elif kind == "tape":
                counts.tape += 1
            elif kind == "spray":
                counts.spray += 1
            elif on_ground_layer and "заземлител" in name.lower():
                counts.rods += 1
        elif on_ground_layer and entity.dxftype() in (
            "LINE", "LWPOLYLINE", "POLYLINE", "ARC",
        ):
            counts.strip_m += _polyline_length(entity)

    return counts


def _extract_spec_signals(dxf_path: str, log=print) -> _GroundingCounts:
    counts = _GroundingCounts()
    try:
        spec = parse_spec_dxf(dxf_path, log=log)
    except Exception:
        return counts
    for item in spec:
        if not _GROUND_SPEC_RE.search(item.description):
            continue
        kind = _classify_spec_description(item.description)
        if not kind:
            continue
        qty, length_m = _parse_spec_qty(
            item.description, item.quantity, item.unit or "шт",
        )
        counts.spec_rows.append((item.description, item.unit or "шт", qty))
        q = int(qty) if isinstance(qty, (int, float)) else 0
        if kind == "rod":
            counts.rods = max(counts.rods, q)
        elif kind == "tip":
            counts.tips = max(counts.tips, q)
        elif kind == "head":
            counts.heads = max(counts.heads, q)
        elif kind == "diag_connector":
            counts.diag_connectors = max(counts.diag_connectors, q)
        elif kind == "flat_connector":
            counts.flat_connectors = max(counts.flat_connectors, q)
        elif kind == "strip":
            counts.strip_m = max(counts.strip_m, float(length_m or qty))
        elif kind == "tape":
            counts.tape = max(counts.tape, q)
        elif kind == "spray":
            counts.spray = max(counts.spray, q)
    return counts


def extract_grounding_facts(dxf_path: str, log=print) -> list[GroundingFact]:
    """DXF-first grounding extraction: plan INSERT/layers + spec table in same file."""
    counts = _extract_dxf_signals(dxf_path)
    spec_counts = _extract_spec_signals(dxf_path, log=log)
    _merge_counts(counts, spec_counts)

    if counts.rods == 0 and counts.strip_m <= 0 and not counts.spec_rows:
        return []

    if counts.diag_connectors == 0 and counts.rods > 0:
        counts.diag_connectors = counts.rods
    if counts.flat_connectors == 0 and counts.rods > 0:
        counts.flat_connectors = counts.rods
    if counts.tips == 0 and counts.rods > 0:
        counts.tips = counts.rods
    if counts.heads == 0 and counts.rods > 0:
        counts.heads = max(1, counts.rods // 15)

    perimeter_m = counts.strip_m
    if perimeter_m <= 0 and counts.rods > 0:
        perimeter_m = counts.rods * _PERIMETER_M_PER_ROD
    trench_m3 = max(
        perimeter_m * _TRENCH_DEPTH_FACTOR,
        counts.rods * _EARTHWORK_M3_PER_ROD,
    )
    sand_m3 = max(
        perimeter_m * _SAND_DEPTH_FACTOR,
        counts.rods * _EARTHWORK_M3_PER_ROD,
    )

    facts: list[GroundingFact] = [
        GroundingFact("trench", derived_qty=round(trench_m3, 1)),
        GroundingFact("sand", derived_qty=round(sand_m3, 1)),
        GroundingFact("perimeter", length_m=perimeter_m),
    ]
    if counts.rods:
        facts.append(GroundingFact("rod", qty=counts.rods))
    if counts.tips:
        facts.append(GroundingFact("tip", qty=counts.tips))
    if counts.heads:
        facts.append(GroundingFact("head", qty=counts.heads))
    if counts.diag_connectors:
        facts.append(GroundingFact("diag_connector", qty=counts.diag_connectors))
    if counts.flat_connectors:
        facts.append(GroundingFact("flat_connector", qty=counts.flat_connectors))
    if counts.strip_m > 0:
        facts.append(GroundingFact("strip", length_m=counts.strip_m))
    if counts.tape:
        facts.append(GroundingFact("tape", qty=counts.tape))
    if counts.spray:
        facts.append(GroundingFact("spray", qty=counts.spray))

    for desc, unit, qty in counts.spec_rows:
        kind = _classify_spec_description(desc)
        if kind:
            facts.append(
                GroundingFact(kind, description=desc, qty=int(qty) if unit == "шт" else 0,
                              length_m=float(qty) if unit == "м" else 0.0),
            )
    return facts


def classify_grounding(facts: list[GroundingFact]) -> list[GroundingFact]:
    """Collapse duplicate kinds; prefer spec descriptions when present."""
    by_kind: dict[str, GroundingFact] = {}
    spec_desc: dict[str, str] = {}
    for fact in facts:
        if fact.description:
            spec_desc[fact.kind] = fact.description
        prev = by_kind.get(fact.kind)
        if prev is None:
            by_kind[fact.kind] = GroundingFact(
                kind=fact.kind,
                qty=fact.qty,
                length_m=fact.length_m,
                derived_qty=fact.derived_qty,
                description=fact.description,
            )
        else:
            prev.qty = max(prev.qty, fact.qty)
            prev.length_m = max(prev.length_m, fact.length_m)
            prev.derived_qty = max(prev.derived_qty, fact.derived_qty)
            if fact.description:
                prev.description = fact.description
    for kind, desc in spec_desc.items():
        if kind in by_kind:
            by_kind[kind].description = desc
    order = (
        "trench", "sand", "rod", "tip", "head",
        "diag_connector", "flat_connector", "strip", "tape", "spray",
        "perimeter",
    )
    return [by_kind[k] for k in order if k in by_kind]


def ground_to_vor(facts: list[GroundingFact]) -> list[GroundingVorRow]:
    """Emit canonical grounding work rows for the VOR section."""
    classified = classify_grounding(facts)
    if not classified:
        return []

    by_kind = {f.kind: f for f in classified}
    rows: list[GroundingVorRow] = []

    trench = by_kind.get("trench")
    if trench and trench.derived_qty > 0:
        qty = trench.derived_qty
        rows.append(GroundingVorRow(
            "Разработка грунта для прокладки горизонтального заземлителя.",
            "м³", qty,
        ))
        sand = by_kind.get("sand")
        sand_qty = sand.derived_qty if sand and sand.derived_qty > 0 else qty
        rows.append(GroundingVorRow(
            "Засыпка траншеи (под горизонтальное заземление).",
            "м³", sand_qty,
        ))

    def _mat_row(kind: str, fallback: str, unit: str) -> None:
        fact = by_kind.get(kind)
        if not fact:
            return
        desc = fact.description or fallback
        if unit == "м":
            q = fact.length_m or fact.qty
            if q <= 0:
                return
        else:
            q = fact.qty
            if q <= 0:
                return
        rows.append(GroundingVorRow(desc, unit, q))

    _mat_row(
        "rod",
        "Стержень заземления L=1500 мм, Ø20 мм, горячее цинкование",
        "шт",
    )
    _mat_row("tip", "Наконечник стержня заземления Ø20 мм", "шт")
    _mat_row("head", "Забивная головка SDS-MAX D20", "шт")
    _mat_row(
        "diag_connector",
        "Соединитель диагональный стержня заземления D=20 мм, термодиффузия",
        "шт",
    )
    _mat_row(
        "flat_connector",
        "Соединитель плоского проводника до 40 мм, термодиффузия",
        "шт",
    )
    _mat_row(
        "strip",
        "Плоский проводник 40х4 мм, горячее цинкование",
        "м",
    )
    _mat_row("tape", "Антикоррозийная лента шириной 50 мм, длиной 10 м", "шт")
    _mat_row("spray", "Спрей цинковый", "шт")

    return rows


def has_grounding_evidence(facts: list[GroundingFact]) -> bool:
    """True when DXF/spec produced at least one material or earthwork fact."""
    for fact in facts:
        if fact.kind in ("trench", "sand"):
            if fact.derived_qty > 0:
                return True
        elif fact.kind in ("rod", "tip", "head", "diag_connector", "flat_connector",
                           "strip", "tape", "spray"):
            if fact.qty > 0 or fact.length_m > 0:
                return True
    return False
