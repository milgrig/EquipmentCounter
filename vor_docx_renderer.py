"""
vor_docx_renderer.py — рендер ВОР в DOCX по эталонному формату ДБТ.

Эталон: «ВОР ЭО, Захватка 3_ГПК.docx»
Формат:
  • Альбомная A4, поля L=R=2см, T=0.8см, B=0.3см
  • Times New Roman 12pt в таблице, 14pt в преамбуле
  • Преамбула: стройка / объект / "Ведомость объемов работ" / основание / дата
  • Таблица 7 колонок: № п/п, Наименование вида работ, Ед.изм., РД,
    Формула расчета, Ссылка на чертежи, Дополнительная информация
  • Группировка строк по разделам (bold-строки):
    Щитовое оборудование / Светотехническое оборудование /
    Монтаж электроустановочных изделий / Кабельная продукция /
    Монтаж кабельных лотков / ПВХ изделия и трубы / Пусконаладочные работы
  • В конце: блок подписей "Составил _____ / Проверил _____"
"""

from __future__ import annotations

import io
import re
from datetime import date
from typing import Iterable

from docx import Document
from docx.shared import Cm, Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ────────────────────────────────────────────────────────────────────────────
# Конфигурация формата
# ────────────────────────────────────────────────────────────────────────────

FONT_NAME = "Times New Roman"
TABLE_FONT_PT = 12
HEADER_FONT_PT = 14

# Ширины колонок (в см) — сумма ≈ 27.24 (ширина страницы A4 landscape - поля)
COL_WIDTHS_CM = [1.23, 9.76, 2.00, 1.50, 5.00, 4.25, 3.50]
COL_HEADERS = [
    "№\nп/п",
    "Наименование вида работ",
    "Ед.\nизм.",
    "Объем работ",
    "Формула расчета объемов работ и расхода материалов",
    "Ссылка на чертежи, спецификации",
    "Дополнительная информация",
]
COL_NUMBERS = ["1", "2", "3", "4", "5", "6", "7"]


# ────────────────────────────────────────────────────────────────────────────
# Группировка элементов ВОР по разделам
# ────────────────────────────────────────────────────────────────────────────

SECTION_RULES: list[tuple[str, list[re.Pattern]]] = [
    ("Щитовое оборудование", [
        re.compile(r"монтаж\s+щит", re.IGNORECASE),
        re.compile(r"подключени[ея]\s+жил", re.IGNORECASE),
        re.compile(r"щ[оа]?[оаб]?\b", re.IGNORECASE),
    ]),
    ("Светотехническое оборудование", [
        re.compile(r"светильник", re.IGNORECASE),
        re.compile(r"светодиодн", re.IGNORECASE),
        re.compile(r"монтаж\s+светил", re.IGNORECASE),
        re.compile(r"\bвыход\b", re.IGNORECASE),
        re.compile(r"световой\s+указат", re.IGNORECASE),
        re.compile(r"эвакуац", re.IGNORECASE),
    ]),
    ("Монтаж электроустановочных изделий", [
        re.compile(r"розетк", re.IGNORECASE),
        re.compile(r"выключател[ья]", re.IGNORECASE),
        re.compile(r"\bпост\s+управл", re.IGNORECASE),
        re.compile(r"датчик", re.IGNORECASE),
        re.compile(r"кнопк", re.IGNORECASE),
        re.compile(r"коробк", re.IGNORECASE),
    ]),
    ("Кабельная продукция", [
        re.compile(r"\bкабел[ья]\b", re.IGNORECASE),
        re.compile(r"провод", re.IGNORECASE),
        re.compile(r"\bввг\w*", re.IGNORECASE),
        re.compile(r"\bппг\w*", re.IGNORECASE),
        re.compile(r"\bвбшв\w*", re.IGNORECASE),
        re.compile(r"\bnym\b", re.IGNORECASE),
        re.compile(r"\bпвс\b", re.IGNORECASE),
    ]),
    ("Монтаж кабельных лотков и соединительных деталей", [
        re.compile(r"лот[ока]к?", re.IGNORECASE),
        re.compile(r"кабельн[аы]\w*\s+полк", re.IGNORECASE),
        re.compile(r"крышк[аи]\s+лотк", re.IGNORECASE),
    ]),
    ("ПВХ изделия и трубы", [
        re.compile(r"\bтруб[аы]\s+пвх", re.IGNORECASE),
        re.compile(r"\bгофр", re.IGNORECASE),
        re.compile(r"\bпнд\b", re.IGNORECASE),
        re.compile(r"\bпвх\b", re.IGNORECASE),
    ]),
    ("Пусконаладочные работы", [
        re.compile(r"измерени[ея]\s+сопротивл", re.IGNORECASE),
        re.compile(r"проверк[аи]\s+срабат", re.IGNORECASE),
        re.compile(r"целостност[иь]\s+жил", re.IGNORECASE),
        re.compile(r"\bпнр\b", re.IGNORECASE),
        re.compile(r"пусконаладоч", re.IGNORECASE),
        re.compile(r"\bлаборатор", re.IGNORECASE),
    ]),
]
SECTION_OTHER = "Прочие работы"


def _classify_section(name: str) -> str:
    """Определить раздел ВОР по наименованию работы."""
    if not name:
        return SECTION_OTHER
    for sec_name, patterns in SECTION_RULES:
        for pat in patterns:
            if pat.search(name):
                return sec_name
    return SECTION_OTHER


def group_by_sections(items: list[dict]) -> list[tuple[str, list[dict]]]:
    """Сгруппировать строки агрегата по разделам в порядке SECTION_RULES."""
    buckets: dict[str, list[dict]] = {sec: [] for sec, _ in SECTION_RULES}
    buckets[SECTION_OTHER] = []
    for it in items:
        sec = _classify_section(it.get("name", ""))
        buckets[sec].append(it)
    out: list[tuple[str, list[dict]]] = []
    for sec, _ in SECTION_RULES:
        if buckets[sec]:
            out.append((sec, buckets[sec]))
    if buckets[SECTION_OTHER]:
        out.append((SECTION_OTHER, buckets[SECTION_OTHER]))
    return out


# ────────────────────────────────────────────────────────────────────────────
# Низкоуровневые помощники DOCX
# ────────────────────────────────────────────────────────────────────────────

def _set_cell_borders(cell):
    """Включить тонкие границы у ячейки."""
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        borders.append(b)
    tc_pr.append(borders)


def _set_cell_text(cell, text: str, *, bold=False, size_pt=TABLE_FONT_PT,
                   align=None, font=FONT_NAME):
    """Записать текст в ячейку с заданным форматом."""
    cell.text = ""
    p = cell.paragraphs[0]
    if align is not None:
        p.alignment = align
    # перенос строк
    parts = str(text or "").split("\n")
    for i, part in enumerate(parts):
        if i > 0:
            p.add_run().add_break()
        run = p.add_run(part)
        run.bold = bold
        run.font.name = font
        run.font.size = Pt(size_pt)
        # Cyrillic font hint
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), font)
        rFonts.set(qn('w:hAnsi'), font)
        rFonts.set(qn('w:cs'), font)


def _set_col_widths(table, widths_cm: list[float]):
    """Жёстко зафиксировать ширины колонок (cm)."""
    # tblLayout fixed
    tbl_pr = table._element.find(qn('w:tblPr'))
    if tbl_pr is None:
        tbl_pr = OxmlElement('w:tblPr')
        table._element.insert(0, tbl_pr)
    layout = tbl_pr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tbl_pr.append(layout)
    layout.set(qn('w:type'), 'fixed')

    for row in table.rows:
        for ci, w in enumerate(widths_cm):
            if ci < len(row.cells):
                row.cells[ci].width = Cm(w)


# ────────────────────────────────────────────────────────────────────────────
# Главный рендер
# ────────────────────────────────────────────────────────────────────────────

def render_vor_docx(
    aggregated: list[dict],
    *,
    rel_folder: str = "",
    section_basis: str = "",
    drawing_prefix: str = "",
    issue_date: str | None = None,
    **_ignored,
) -> bytes:
    """Сформировать ВОР в формате эталона ДБТ. Возвращает байты .docx."""
    doc = Document()

    # 1. Section: A4 landscape, особые поля
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width = Cm(29.7)
    sec.page_height = Cm(21.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(0.8)
    sec.bottom_margin = Cm(0.3)

    if issue_date is None:
        issue_date = date.today().strftime("%d.%m.%Y")

    def _add_par(text, *, bold=False, align=None, size=HEADER_FONT_PT):
        p = doc.add_paragraph()
        if align is not None:
            p.alignment = align
        run = p.add_run(text)
        run.bold = bold
        run.font.name = FONT_NAME
        run.font.size = Pt(size)
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:ascii'), FONT_NAME)
        rFonts.set(qn('w:hAnsi'), FONT_NAME)
        rFonts.set(qn('w:cs'), FONT_NAME)
        return p

    # 2. Преамбула (минимальная, по образцу ВОР_ЭО.docx)
    _add_par("Ведомость объемов работ", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    if section_basis:
        _add_par(section_basis)
    _add_par(f"Дата составления {issue_date}г.")
    _add_par("")

    # 3. Таблица
    grouped = group_by_sections(aggregated)
    total_data_rows = sum(len(rows) for _, rows in grouped)
    total_section_rows = len(grouped)
    # +1 шапка, +1 нумерация колонок
    total_rows = 2 + total_section_rows + total_data_rows

    tbl = doc.add_table(rows=total_rows, cols=len(COL_HEADERS))
    tbl.style = "Table Grid"

    # шапка
    for ci, hdr in enumerate(COL_HEADERS):
        c = tbl.rows[0].cells[ci]
        _set_cell_text(c, hdr, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_borders(c)
    for ci, num in enumerate(COL_NUMBERS):
        c = tbl.rows[1].cells[ci]
        _set_cell_text(c, num, align=WD_ALIGN_PARAGRAPH.CENTER)
        _set_cell_borders(c)

    # данные
    row_idx = 2
    item_no = 0  # сквозная нумерация по эталону

    for section_name, items in grouped:
        # строка-разделитель раздела (bold во 2-й колонке, остальные пустые)
        for ci in range(len(COL_HEADERS)):
            cell = tbl.rows[row_idx].cells[ci]
            if ci == 1:
                _set_cell_text(cell, section_name, bold=True)
            else:
                _set_cell_text(cell, "")
            _set_cell_borders(cell)
        row_idx += 1

        for it in items:
            item_no += 1
            name = str(it.get("name", "")).strip()
            # T078 / S019-minimal-renderer: strip any "[UNMATCHED-LEGEND]"
            # diagnostic prefix from the work-name before writing to col 2.
            # The reference docx (T065_recon.md section 0) has none of these
            # warnings; they belong in logs/audit only, not in deliverable
            # column 2.  Pattern from web_app:1540 ("[UNMATCHED-LEGEND] %s").
            name = _strip_unmatched_legend_prefix(name)
            unit = str(it.get("unit", "шт")).strip()
            total = it.get("total", 0)
            drawing_refs = str(it.get("drawing_refs", "")).strip()

            # Привести ссылку на чертежи к шифру проекта
            ref_text = _format_drawing_ref(drawing_refs, drawing_prefix)

            # T078: per T065_recon.md section 0, the human-authored
            # reference docx has columns 5 (Formula) and 7 (Additional
            # info) EMPTY in 130/130 data rows.  Emit "" for both columns
            # to match the reference layout; intentionally ignore
            # it.get("formula") and it.get("extra_info") here.  Section
            # header rows are emitted above with all non-name columns
            # already empty -- no change needed there.
            cells_text = [
                str(item_no),
                name,
                unit,
                _fmt_qty(total),
                "",  # T078: column 5 (Formula) empty for all data rows
                ref_text,
                "",  # T078: column 7 (Additional info) empty for all data rows
            ]
            for ci, text in enumerate(cells_text):
                cell = tbl.rows[row_idx].cells[ci]
                align = None
                if ci in (0, 2, 3):
                    align = WD_ALIGN_PARAGRAPH.CENTER
                _set_cell_text(cell, text, align=align)
                _set_cell_borders(cell)
            row_idx += 1

    _set_col_widths(tbl, COL_WIDTHS_CM)

    # 4. Подписи (удалены: данные исполнителей не известны автогенератору)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _fmt_qty(v) -> str:
    """Отформатировать количество (целые без .0, дробные с одной запятой)."""
    if v is None or v == "":
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(f - round(f)) < 1e-6:
        return str(int(round(f)))
    return f"{f:.1f}".replace(".", ",")


# T078: pattern for the diagnostic warning prefix produced by web_app
# legend-coverage audit (web_app.py:1540).  We strip it before writing
# col 2 so the deliverable docx never contains the audit string.  Pattern
# is anchored at start and tolerates one or more whitespace chars after
# the closing bracket.
_UNMATCHED_LEGEND_RE = re.compile(r"^\s*\[UNMATCHED-LEGEND\]\s*", re.IGNORECASE)


def _strip_unmatched_legend_prefix(name: str) -> str:
    """Remove "[UNMATCHED-LEGEND] " prefix from a work-name if present.

    Returns the cleaned name; leaves any non-prefixed name unchanged.
    Idempotent over reruns (since the regex is anchored at start).
    """
    if not name:
        return name
    return _UNMATCHED_LEGEND_RE.sub("", name, count=1)


_LIST_NUM_RE = re.compile(r"л\.?\s*[\d.,\-–\s]+", re.IGNORECASE)


def _format_drawing_ref(raw: str, prefix: str) -> str:
    """Свернуть исходный список имён файлов в формат "<prefix>, л.<N>"."""
    if not raw:
        return ""
    # Если уже похож на эталон (содержит "л." и шифр) — оставить
    if "1Д-" in raw and "л." in raw:
        return raw
    # Извлечь номера листов из строк "005-План...", "006-План..."
    nums = []
    for tok in re.findall(r"\b(\d{2,3})\b", raw):
        if tok not in nums:
            nums.append(tok)
    # Также пустить через регулярку «л.NN»
    if not nums:
        return prefix
    # Сжать в диапазоны: 5,6,7,8 -> 5-8
    nums_int = sorted({int(n) for n in nums})
    ranges: list[str] = []
    i = 0
    while i < len(nums_int):
        j = i
        while j + 1 < len(nums_int) and nums_int[j + 1] == nums_int[j] + 1:
            j += 1
        if j > i:
            ranges.append(f"{nums_int[i]}-{nums_int[j]}")
        else:
            ranges.append(str(nums_int[i]))
        i = j + 1
    sheet_part = f"л.{','.join(ranges)}"
    if prefix:
        return f"{prefix}, {sheet_part}"
    return sheet_part
