#!/usr/bin/env python3
"""Generate the commercial proposal Word document (Appendix 3 & title attributes)
directly from Gen_TKP.xlsx / Gen_TKP.xlsm using openpyxl and python-docx.

Can be used both as a library module and as a CLI script without Excel/Word COM.
"""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
import tempfile
import warnings
from collections.abc import Iterable
from copy import deepcopy
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile

import docx
import openpyxl
from docx.document import Document as DocxDocument
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor
from lxml import etree
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parent
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
NS = {"w": W_NS}

TEMPLATE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
)
DOCUMENT_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)

SPEC_HEADERS = (
    "№ п/п",
    "Наименование и техническая характеристика",
    "Артикул",
    "Ед. изм.",
    "Кол-во",
)

# The proposal is set in 12 pt throughout, the specification table included.
SPEC_FONT_SIZE = Pt(12)

# Excel error literals. These are blanked out wherever they appear; any other
# value starting with "#" (a legitimate article such as "#A-100") is kept.
EXCEL_ERRORS = frozenset(
    {
        "#REF!",
        "#N/A",
        "#VALUE!",
        "#DIV/0!",
        "#NAME?",
        "#NULL!",
        "#NUM!",
        "#SPILL!",
        "#CALC!",
        "#GETTING_DATA",
    }
)

# The proposal is a Russian document: «2,5 шт.», never «2.5 шт.».
DECIMAL_SEPARATOR = ","

EMPTY_WORKBOOK_MESSAGE = (
    "Нет позиций с количеством больше нуля. "
    "Заполните столбец G на листах оборудования, ЗИП или ПО."
)


class WorkbookError(ValueError):
    """The workbook cannot become Appendix 3 as it stands; nothing is written."""


class TemplateError(ValueError):
    """The template lacks a bookmark the generator fills; nothing is written."""


class _Unresolved:
    """A formula the generator cannot evaluate and the workbook holds no saved value for."""

    def __repr__(self) -> str:
        return "UNRESOLVED"


UNRESOLVED = _Unresolved()


@dataclass
class SpecItem:
    description: str
    article: str
    unit: str
    qty: float


@dataclass
class SpecSection:
    title: str
    items: list[SpecItem] = field(default_factory=list)


@dataclass
class TKPMetadata:
    number: str = ""
    customer: str = ""
    object_name: str = ""
    subject: str = ""


@dataclass
class TKPData:
    metadata: TKPMetadata
    sections: list[SpecSection]
    issues: list[str] = field(default_factory=list)  # cells that could not be evaluated
    demo: bool = False

    @property
    def total_items(self) -> int:
        return sum(len(s.items) for s in self.sections)


def clean_text(value: object) -> str:
    """Normalize cell values into a single-line trimmed string, ignoring errors/None."""
    if value is None or isinstance(value, bool) or value is UNRESOLVED:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)  # VBA's CStr(12345#) is "12345", not "12345.0"
    text = str(value)
    for ch in ("\r\n", "\r", "\n", "\t"):
        text = text.replace(ch, " ")
    text = text.strip()
    if text in EXCEL_ERRORS:
        return ""
    return text


def format_qty(qty: float) -> str:
    """The quantity text Word shows: plain notation, at most six decimals, a decimal comma.

    Module1.ТекстКоличества produces the same text whatever the Windows locale.
    The document is customer-facing, so scientific notation must never reach it:
    %g would render 1234567.5 as "1.23457e+06".
    """
    rounded = round(qty, 6)
    if abs(rounded - round(rounded)) < 1e-9:
        return str(round(rounded))
    return f"{rounded:.6f}".rstrip("0").rstrip(".").replace(".", DECIMAL_SEPARATOR)


def format_number(number: object) -> str:
    """Normalise the proposal number shown on the first page and in the head.

    The running header reads «Технико-коммерческое предложение № …» and repeats
    the «Номер» row of the first page, so the sign is added here instead of
    having to be typed into the sheet.
    """
    text = clean_text(number)
    if not text:
        return ""
    if text.startswith("№"):
        text = text[1:].strip()
    return f"№ {text}" if text else "№"


def _find_sheet(wb: Workbook, *names: str) -> Worksheet | None:
    for name in names:
        if name in wb.sheetnames:
            return wb[name]
    return None


def _block_from_caption(caption: str) -> int:
    low = caption.lower()
    if "нвт" in low:
        return 1
    if "эксара" in low or "инсат" in low:
        return 2
    if "3-х поставщиков" in low:
        return 3
    return 0


def _block_from_row(row: int) -> int:
    if 3 <= row <= 9:
        return 1
    if 14 <= row <= 15:
        return 2
    if 20 <= row <= 47:
        return 3
    return 0


_SHEET_REF_RE = re.compile(r"^='?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)$")
_PLAIN_NUMBER_RE = re.compile(r"[+-]?(?:[0-9]+(?:[.,][0-9]*)?|[.,][0-9]+)")
_ZIP_TERM_RE = re.compile(r"SUMIF\('([^']+)'!\$D\$3:\$D\$100,D(\d+),'([^']+)'!\$G\$3:\$G\$100\)")

CABINET_BLOCK_TITLES = {
    1: "Оборудование ПТК",
    2: "Телекоммуникационное оборудование для АСУ ТП",
    3: "Шкафное оборудование ПТК",
}

# (sheet name and alias, cabinet number or 0, fixed section title), in document order.
SPEC_SHEETS: tuple[tuple[tuple[str, str], int, str], ...] = (
    (("SW(w.o.VAT)", "ПО"), 0, "Программное обеспечение"),
    *(((str(n), f"Шкаф {n}"), n, "") for n in range(1, 18)),
    (("Spare parts", "ЗИП"), 0, "ЗИП"),
)

# Columns whose text reaches the document: description, article, unit, quantity.
# Column B only decides the grouping and the skips, so a formula there that the
# generator cannot evaluate cannot corrupt the proposal and is not an issue.
DOCUMENT_COLUMNS = frozenset({3, 4, 6, 7})


def _cached_value(ws_val: Worksheet | None, ws_fml: Worksheet, row: int, col: int) -> object:
    """The result Excel saved for a formula cell, or UNRESOLVED when there is none.

    A workbook loaded once (an in-memory Workbook) has no separate cache: its
    "value" is the formula text itself, which must never reach the document.
    """
    if ws_val is None or ws_val is ws_fml:
        return UNRESOLVED
    cell = ws_val.cell(row, col)
    if cell.value is not None:
        return cell.value
    # Excel saves an empty-text result as t="str" with an empty <v>, which openpyxl
    # reads back as None of that type; a missing cache stays type "n".
    return "" if cell.data_type in ("s", "str") else UNRESOLVED


def _zip_rollup_sheets(formula: str, row: int) -> list[str] | None:
    """Cabinet sheets of the exact ЗИП formula ROUNDUP((ΣSUMIF(...))/10,0), else None."""
    compact = re.sub(r"\s+", "", formula)
    head, tail = "=ROUNDUP((", ")/10,0)"
    if not (compact.startswith(head) and compact.endswith(tail)):
        return None
    body = compact[len(head) : -len(tail)]
    terms = _ZIP_TERM_RE.findall(body)
    rebuilt = "+".join(f"SUMIF('{a}'!$D$3:$D$100,D{r},'{b}'!$G$3:$G$100)" for a, r, b in terms)
    if not terms or rebuilt != body or any(a != b or int(r) != row for a, r, b in terms):
        return None
    return [a for a, _, _ in terms]


def _zip_rollup(ws_val, ws_fml, wb_val, wb_fml, row: int, sheets: list[str], seen) -> object:
    """The ЗИП quantity: one spare per ten items of that article across the cabinets."""
    raw_article = _eval_cell(ws_val, ws_fml, wb_val, wb_fml, row, 4, seen)
    if raw_article is UNRESOLVED:
        return UNRESOLVED
    article = clean_text(raw_article).lower()
    if not article:
        return 0
    total = 0.0
    for name in sheets:
        if name not in wb_fml.sheetnames:
            return UNRESOLVED
        c_fml = wb_fml[name]
        c_val = wb_val[name] if wb_val is not None and name in wb_val.sheetnames else None
        for r in range(3, min(c_fml.max_row, 100) + 1):
            if clean_text(_eval_cell(c_val, c_fml, wb_val, wb_fml, r, 4, seen)).lower() != article:
                continue
            g_val = _eval_cell(c_val, c_fml, wb_val, wb_fml, r, 7, seen)
            if g_val is UNRESOLVED:
                return UNRESOLVED
            if isinstance(g_val, (int, float)) and not isinstance(g_val, bool):
                total += float(g_val)
    return math.ceil(total / 10.0) if total > 0 else 0


def _eval_cell(
    ws_val: Worksheet | None,
    ws_fml: Worksheet | None,
    wb_val: Workbook | None,
    wb_fml: Workbook | None,
    row: int,
    col: int,
    _seen: frozenset[tuple[str, int, int]] = frozenset(),
) -> object:
    """Evaluate a cell the way Excel shows it, as far as the generator can.

    Literals come back as they are. Three formula shapes are evaluated here:
    anything containing #REF! (blank, as its IFERROR wrapper makes it), a single
    cross-sheet reference, and the exact ЗИП roll-up. Every other formula falls
    back to the value Excel saved, and to UNRESOLVED when there is none — the
    caller reports it; it is never read as 0. A reference cycle degrades to "".
    """
    if ws_fml is None:
        return ws_val.cell(row, col).value if ws_val is not None else None
    key = (ws_fml.title, row, col)
    if key in _seen:
        return ""
    _seen = _seen | {key}

    raw_fml = ws_fml.cell(row, col).value
    if not (isinstance(raw_fml, str) and raw_fml.startswith("=")):
        return raw_fml
    if "#REF!" in raw_fml.upper():
        return ""

    match = _SHEET_REF_RE.match(raw_fml.strip())
    if match and wb_fml is not None and match.group(1) in wb_fml.sheetnames:
        target, col_letters, row_digits = match.groups()
        t_val = wb_val[target] if wb_val is not None and target in wb_val.sheetnames else None
        resolved = _eval_cell(
            t_val,
            wb_fml[target],
            wb_val,
            wb_fml,
            int(row_digits),
            column_index_from_string(col_letters),
            _seen,
        )
        if resolved is None:
            return 0  # Excel shows a reference to an empty cell as 0
        if resolved is not UNRESOLVED and not (isinstance(resolved, str) and resolved in EXCEL_ERRORS):
            return resolved

    if wb_fml is not None and "SUMIF(" in raw_fml.upper():
        sheets = _zip_rollup_sheets(raw_fml, row)
        if sheets is not None:
            rolled = _zip_rollup(ws_val, ws_fml, wb_val, wb_fml, row, sheets, _seen)
            if rolled is not UNRESOLVED:
                return rolled

    return _cached_value(ws_val, ws_fml, row, col)


def _parse_numeric_qty(raw_qty: object) -> float:
    """Read column G exactly as Module1.ЧислоКоличества does.

    A number cell as it is; text only in plain decimal notation with "." or ","
    as the separator. float() alone would also take "inf", "1_000" and non-ASCII
    digits, which neither Excel nor the macro treat as numbers.
    """
    if raw_qty is None or isinstance(raw_qty, bool):
        return 0.0
    if isinstance(raw_qty, (int, float)):
        value = float(raw_qty)
        return value if math.isfinite(value) else 0.0
    if isinstance(raw_qty, str):
        text = raw_qty.strip(" ")  # VBA Trim$ strips spaces only
        if _PLAIN_NUMBER_RE.fullmatch(text):
            return float(text.replace(",", "."))
    return 0.0


def _describe_unresolved(ws_fml: Worksheet, row: int, col: int) -> str:
    formula = " ".join(str(ws_fml.cell(row, col).value or "").split())
    if len(formula) > 80:
        formula = formula[:77] + "..."
    return f"{ws_fml.title}!{get_column_letter(col)}{row}: {formula}"


def _sheet_last_row(ws_fml: Worksheet) -> int:
    """The last row the macro scans: the lowest filled B, C, D or G cell, at least 48."""
    for row in range(ws_fml.max_row or 0, 48, -1):
        if any(ws_fml.cell(row, col).value not in (None, "") for col in (2, 3, 4, 7)):
            return row
    return 48


def _is_service_row(caption: str, description: str) -> bool:
    """Totals and margin roll-ups, in the caption (B) or the description (C)."""
    return caption.lower().startswith("итого") or description.lower().startswith(("итого", "маржа"))


def _collect_sheet_sections(
    ws_val: Worksheet | None,
    ws_fml: Worksheet,
    wb_val: Workbook,
    wb_fml: Workbook,
    *,
    cab_no: int,
    fixed_title: str,
    issues: list[str],
) -> list[SpecSection]:
    """One top-down pass over a sheet, step for step like Module1.ОбработатьЛист.

    Sections come in sheet order. On a cabinet sheet every caption (and, without
    captions, every change of row-range block) opens a new section; the software
    and ЗИП sheets form one section each.
    """
    sections: list[SpecSection] = []
    current: SpecSection | None = None
    current_block = prev_block = 0
    has_captions = False

    def value(row: int, col: int) -> object:
        return _eval_cell(ws_val, ws_fml, wb_val, wb_fml, row, col)

    def text(row: int, col: int, raw: object) -> str:
        if raw is UNRESOLVED and col in DOCUMENT_COLUMNS:
            issues.append(_describe_unresolved(ws_fml, row, col))
        return clean_text(raw)

    for r in range(1, _sheet_last_row(ws_fml) + 1):
        caption = clean_text(value(r, 2))
        description_raw = value(r, 3)
        description = clean_text(description_raw)

        if caption.lower().startswith("спецификация "):
            current_block = _block_from_caption(caption)
            if current_block == 0:
                raise WorkbookError(
                    f"Лист «{ws_fml.title}», строка {r}: неизвестный раздел «{caption}». "
                    "Ожидается «НВТ», «Эксара»/«ИнСАТ» или «3-х поставщиков»."
                )
            has_captions = True
            if cab_no > 0:
                current = None
            continue
        if _is_service_row(caption, description):
            current_block = 0
            continue
        if caption.lower() == "№ п/п":
            continue

        block = current_block
        if block == 0 and not has_captions:
            block = _block_from_row(r)
            if cab_no > 0 and block != prev_block:
                current, prev_block = None, block
        if block == 0:
            continue

        raw_qty = value(r, 7)
        if raw_qty is UNRESOLVED:
            issues.append(_describe_unresolved(ws_fml, r, 7))
            continue
        qty = _parse_numeric_qty(raw_qty)
        if qty <= 0:
            continue

        article = text(r, 4, value(r, 4))
        unit = text(r, 6, value(r, 6)) or "шт."
        description = text(r, 3, description_raw)
        if not description or description == "0":
            description = article  # the broken #REF! price-list lookup in C
        if not description:
            continue
        if current is None:
            title = f"{CABINET_BLOCK_TITLES[block]} (шкаф {cab_no})" if cab_no > 0 else fixed_title
            current = SpecSection(title=title)
            sections.append(current)
        current.items.append(SpecItem(description=description, article=article, unit=unit, qty=qty))
    return sections


def _open_book_pair(excel_source: str | Path | Workbook) -> tuple[Workbook | None, Workbook]:
    """Open the workbook twice: cached values and formulas."""
    if isinstance(excel_source, Workbook):
        return excel_source, excel_source
    path = Path(excel_source)
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        wb_val = openpyxl.load_workbook(path, data_only=True)
        wb_fml = openpyxl.load_workbook(path, data_only=False)
    return wb_val, wb_fml


def read_tkp_data(excel_source: str | Path | Workbook) -> TKPData:
    """Read the title attributes and the Appendix 3 sections; never raises for bad cells.

    Cells that could not be evaluated are listed in TKPData.issues.
    """
    wb_val, wb_fml = _open_book_pair(excel_source)

    issues: list[str] = []
    metadata = TKPMetadata()
    info_fml = _find_sheet(wb_fml, "Наценки-Скидки")
    info_val = _find_sheet(wb_val, "Наценки-Скидки") if wb_val is not None else None
    if info_fml is not None:
        values = []
        for row in (2, 3, 4, 5):
            raw = _eval_cell(info_val, info_fml, wb_val, wb_fml, row, 3)
            if raw is UNRESOLVED:
                issues.append(_describe_unresolved(info_fml, row, 3))
            values.append(clean_text(raw))
        metadata = TKPMetadata(*values)

    sections: list[SpecSection] = []
    for names, cab_no, title in SPEC_SHEETS:
        ws_fml = _find_sheet(wb_fml, *names)
        if ws_fml is not None:
            sections.extend(
                _collect_sheet_sections(
                    wb_val[ws_fml.title] if wb_val is not None else None,
                    ws_fml,
                    wb_val,
                    wb_fml,
                    cab_no=cab_no,
                    fixed_title=title,
                    issues=issues,
                )
            )
    return TKPData(metadata=metadata, sections=sections, issues=issues)


def build_preview_rows(sections: Iterable[SpecSection]) -> list[tuple[str, str, str, str, float | int, str]]:
    """Build the preview table rows matching sheet 'Спец_ТКП' (columns A..F)."""
    rows: list[tuple[str, str, str, str, float | int, str]] = []
    section_no = 0
    for section in sections:
        if not section.items:
            continue
        section_no += 1
        rows.append((f"{section_no}.", section.title, "", "Компл.", 1, "H"))
        for item_idx, item in enumerate(section.items, start=1):
            qty_val: float | int = (
                round(item.qty) if abs(item.qty - round(item.qty)) < 1e-9 else item.qty
            )
            rows.append(
                (f"{item_idx}.", item.description, item.article, item.unit, qty_val, "I")
            )
    return rows


def update_preview_sheet(wb: Workbook, sections: Iterable[SpecSection]) -> Worksheet:
    """Populate the 'Спец_ТКП' preview worksheet in an openpyxl Workbook."""
    ws = _find_sheet(wb, "Спец_ТКП")
    if ws is None:
        ws = wb.create_sheet("Спец_ТКП")

    max_r = max(ws.max_row or 1, 1)
    for r in range(1, max_r + 1):
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.value = None
            cell.font = openpyxl.styles.Font(bold=False)

    bold_font = openpyxl.styles.Font(bold=True)
    regular_font = openpyxl.styles.Font(bold=False)
    for col_idx, header_text in enumerate(SPEC_HEADERS, start=1):
        cell = ws.cell(1, col_idx, header_text)
        cell.font = bold_font

    for row_offset, row_data in enumerate(build_preview_rows(sections), start=2):
        is_group = row_data[5] == "H"
        row_font = bold_font if is_group else regular_font
        for col_idx, val in enumerate(row_data, start=1):
            c = ws.cell(row_offset, col_idx, val)
            # A:D are text, as in the macro: in a General cell «0441504» would
            # lose its zero, «1-2» would become a date and «=…» a formula.
            if col_idx <= 4:
                c.number_format = "@"
            if col_idx <= 5:
                c.font = row_font

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 27
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].hidden = True
    return ws


def open_word_template(template_path: str | Path) -> DocxDocument:
    """Open either a .docx or .dotx Word template in python-docx."""
    path = Path(template_path)
    buf = BytesIO()
    with ZipFile(path, "r") as zin, ZipFile(buf, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == "word/_rels/settings.xml.rels":
                continue
            data = zin.read(info.filename)
            if info.filename == "[Content_Types].xml":
                text = data.decode("utf-8").replace(
                    TEMPLATE_CONTENT_TYPE, DOCUMENT_CONTENT_TYPE
                )
                data = text.encode("utf-8")
            elif info.filename == "word/settings.xml":
                root = etree.fromstring(data)
                for attached in root.findall(f"{W}attachedTemplate"):
                    root.remove(attached)
                data = etree.tostring(
                    root, encoding="UTF-8", xml_declaration=True, standalone=True
                )
            zout.writestr(info, data)
    buf.seek(0)
    return docx.Document(buf)


# Children a paragraph-mark rPr may hold but a run rPr may not (CT_ParaRPr).
_PARAGRAPH_MARK_ONLY = frozenset(f"{W}{t}" for t in ("ins", "del", "moveFrom", "moveTo", "rPrChange"))


def _bookmark_start(root: etree._Element, name: str) -> etree._Element | None:
    return next((el for el in root.iter(f"{W}bookmarkStart") if el.get(f"{W}name") == name), None)


def _next_bookmark_id(root: etree._Element) -> str:
    ids = [int(v) for el in root.iter(f"{W}bookmarkStart") if (v := el.get(f"{W}id", "")).isdigit()]
    return str(max(ids, default=0) + 1)


def _replacement_rpr(parent, start, end) -> etree._Element | None:
    """The run formatting Word gives replacement text: that of the first replaced
    character, or of the character before an empty bookmark. The paragraph-mark
    properties are the last resort, minus the revision marks a run cannot hold."""
    children = list(parent)
    first = children.index(start) + 1
    stop = children.index(end) if end is not None and end.getparent() is parent else first
    for run in children[first:stop] + children[: first - 1][::-1]:
        if run.tag == f"{W}r":
            rpr = run.find(f"{W}rPr")
            return deepcopy(rpr) if rpr is not None else None
    ppr = parent.find(f"{W}pPr")
    mark = ppr.find(f"{W}rPr") if ppr is not None else None
    if mark is None:
        return None
    rpr = deepcopy(mark)
    for child in [c for c in rpr if c.tag in _PARAGRAPH_MARK_ONLY]:
        rpr.remove(child)
    return rpr


def set_bookmark_text(root: etree._Element, name: str, text: str) -> bool:
    """Replace the text inside a bookmark; False if it is missing or not inside a paragraph."""
    start = _bookmark_start(root, name)
    parent = start.getparent() if start is not None else None
    if parent is None or parent.tag != f"{W}p":
        return False
    bm_id = start.get(f"{W}id")
    end = next((el for el in root.iter(f"{W}bookmarkEnd") if el.get(f"{W}id") == bm_id), None)
    rpr = _replacement_rpr(parent, start, end)

    if end is not None and end.getparent() is parent:
        for child in list(parent)[parent.index(start) + 1 : parent.index(end)]:
            parent.remove(child)
    else:
        if end is not None:
            end.getparent().remove(end)
        end = etree.Element(f"{W}bookmarkEnd", {f"{W}id": bm_id or "0"})
        parent.insert(parent.index(start) + 1, end)

    run = etree.Element(f"{W}r")
    if rpr is not None:
        run.append(rpr)
    t_el = etree.SubElement(run, f"{W}t")
    if text != text.strip():
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_el.text = text
    parent.insert(parent.index(end), run)
    return True


def _header_elements(doc: DocxDocument) -> list[etree._Element]:
    """The header parts the template already has. section.even_page_header
    creates a part on first access — the empty header3.xml of B4."""
    return [
        rel.target_part.element
        for rel in doc.part.rels.values()
        if rel.reltype == RT.HEADER and not rel.is_external
    ]


def _style_cell(
    cell,
    text: str,
    *,
    bold: bool = False,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.LEFT,
    bg_fill: str | None = None,
    text_color: RGBColor | None = None,
    width_pct: int | None = None,
) -> None:
    if text_color is None:
        text_color = RGBColor(0x00, 0x00, 0x00)
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    if width_pct is not None:
        for old_w in tcPr.findall(qn("w:tcW")):
            tcPr.remove(old_w)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"), str(width_pct))
        tcW.set(qn("w:type"), "pct")
        tcPr.append(tcW)

    if bg_fill is not None:
        for old_shd in tcPr.findall(qn("w:shd")):
            tcPr.remove(old_shd)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{bg_fill}"/>')
        tcPr.append(shd)

    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)

    for r in list(p.runs):
        p._p.remove(r._r)

    run = p.add_run(text)
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = SPEC_FONT_SIZE
    run.font.color.rgb = text_color


def _set_row_repeat_header(row) -> None:
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))


def populate_spec_table(doc: DocxDocument, sections: Iterable[SpecSection]) -> docx.table.Table:
    """Create and populate the Appendix 3 specification table at bookmark Spec_Table."""
    body = doc.element.find(f"{W}body")
    start = _bookmark_start(doc.element, "Spec_Table")
    anchor = start
    while anchor is not None and anchor.getparent() is not body:
        anchor = anchor.getparent()
    if body is None or anchor is None:
        raise TemplateError("В шаблоне нет закладки Spec_Table: таблице Приложения 3 некуда встать.")
    preview_rows = build_preview_rows(sections)
    total_rows = 2 + len(preview_rows)

    table = doc.add_table(rows=total_rows, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tblPr = table._tbl.tblPr
    for old_w in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old_w)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")
    tblPr.append(tblW)

    for old_borders in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old_borders)
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/>'
        "</w:tblBorders>"
    )
    tblPr.append(tblBorders)

    col_pcts = (359, 2739, 1009, 505, 388)
    grid_twips = (702, 5357, 1974, 988, 759)  # = 9780 twips, the portrait text width
    for old_grid in table._tbl.findall(qn("w:tblGrid")):
        table._tbl.remove(old_grid)
    tblGrid = OxmlElement("w:tblGrid")
    for tw in grid_twips:
        gridCol = OxmlElement("w:gridCol")
        gridCol.set(qn("w:w"), str(tw))
        tblGrid.append(gridCol)
    table._tbl.insert(list(table._tbl).index(tblPr) + 1, tblGrid)

    # Row 0: column titles, Row 1: column numbers 1..5
    _set_row_repeat_header(table.rows[0])
    _set_row_repeat_header(table.rows[1])
    for col_idx, header_text in enumerate(SPEC_HEADERS):
        _style_cell(
            table.rows[0].cells[col_idx],
            header_text,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            bg_fill="C00000",
            text_color=RGBColor(0xFF, 0xFF, 0xFF),
            width_pct=col_pcts[col_idx],
        )
        _style_cell(
            table.rows[1].cells[col_idx],
            str(col_idx + 1),
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            bg_fill="C00000",
            text_color=RGBColor(0xFF, 0xFF, 0xFF),
            width_pct=col_pcts[col_idx],
        )

    # Data rows (sections and items)
    for idx, (num_str, desc, art, unit, qty, marker) in enumerate(preview_rows, start=2):
        row = table.rows[idx]
        is_group = marker == "H"
        bg = "D9D9D9" if is_group else "FFFFFF"
        values = (num_str, desc, art, unit, format_qty(qty))
        aligns = (
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER if is_group else WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
        )
        for col_idx in range(5):
            _style_cell(
                row.cells[col_idx],
                values[col_idx],
                bold=is_group,
                align=aligns[col_idx],
                bg_fill=bg,
                width_pct=col_pcts[col_idx],
            )
        if is_group:
            for p in row.cells[0]._tc.getparent().findall(f".//{W}p"):
                pPr = p.find(f"{W}pPr")
                if pPr is not None and pPr.find(f"{W}rPr") is None:
                    rPr = etree.SubElement(pPr, f"{W}rPr")
                    etree.SubElement(rPr, f"{W}b")

    # Place the table at bookmark Spec_Table
    bm_id = start.get(f"{W}id") or _next_bookmark_id(doc.element)
    for end_el in [el for el in doc.element.iter(f"{W}bookmarkEnd") if el.get(f"{W}id") == bm_id]:
        end_el.getparent().remove(end_el)
    body.insert(body.index(anchor), table._tbl)
    body.remove(anchor)

    # Keep Spec_Table bookmark after <w:pPr> in cell(0, 0) so the bookmark remains valid
    first_p = table.rows[0].cells[0].paragraphs[0]._p
    pPr = first_p.find(f"{W}pPr")
    insert_idx = (list(first_p).index(pPr) + 1) if pPr is not None else 0
    new_start = OxmlElement("w:bookmarkStart")
    new_start.set(qn("w:id"), str(bm_id))
    new_start.set(qn("w:name"), "Spec_Table")
    new_end = OxmlElement("w:bookmarkEnd")
    new_end.set(qn("w:id"), str(bm_id))
    first_p.insert(insert_idx, new_end)
    first_p.insert(insert_idx, new_start)

    return table


def find_default_template(base_dir: Path = ROOT) -> Path:
    """Return the project's single supported default template."""
    template = base_dir / "ТКП_шаблон.dotx"
    if template.exists():
        return template
    raise FileNotFoundError(f"Word template ТКП_шаблон.dotx not found in {base_dir}")
def find_default_workbook(base_dir: Path = ROOT) -> Path:
    """The macro workbook first: a Gen_TKP.xlsx beside it is an export, not the source."""
    for name in ("Gen_TKP.xlsm", "Gen_TKP.xlsx"):
        candidate = base_dir / name
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"Excel generator workbook not found in {base_dir}")


_SHEET_ALIASES = {names[0]: names[1] for names, _, _ in SPEC_SHEETS}

DEMO_METADATA = {
    "C2": "№ 104/1 от 16.10.2026",
    "C3": "ДЕМОНСТРАЦИОННЫЕ ДАННЫЕ — заказчик не указан",
    "C4": "ДЕМО: АСУ ТП компрессорного цеха",
    "C5": "Поставка оборудования и выполнение работ по созданию АСУ ТП "
    "на базе программно-технического комплекса АРКС400",
}

# (sheet, row, description or None to keep the sheet's #REF! lookup, article, unit, qty)
DEMO_ROWS: tuple[tuple[str, int, str | None, str, str, float], ...] = (
    ("SW(w.o.VAT)", 3, None, "АРКС400-СДА", "шт.", 1),
    ("SW(w.o.VAT)", 14, "Исполнительная система для программирования контроллеров. 1000 точек",
     "MasterPLC.N1000", "шт.", 2),
    ("SW(w.o.VAT)", 15, "Лицензия на клиент визуализации. 1000 точек", "Визуализация.N1000", "шт.", 2),
    ("SW(w.o.VAT)", 25, "Операционная система Astra Linux Special Edition", "ALSE-1.7", "шт.", 2),
    ("1", 3, None, "АРКС400-ЦП-01", "шт.", 2),
    ("1", 4, "Модуль дискретного ввода 32 канала =24В", "АРКС400-ДВ-32", "шт.", 4),
    ("1", 5, "Модуль аналогового ввода 16 каналов 4..20 мА", "АРКС400-АВ-16", "шт.", 3),
    ("1", 14, "Коммутатор промышленный управляемый L2, 8x100/1000Base-T", "EX-SW-8G", "шт.", 2),
    ("1", 20, "Шкаф напольный 2000х800х600 мм, IP54", "TS-8806", "шт.", 1),
    ("1", 21, "Блок питания промышленный 24В DC, 20А", "QUINT4-PS/1AC/24DC/20", "шт.", 2),
    ("2", 3, "Модуль дискретного вывода 16 каналов =24В", "АРКС400-ДВыв-16", "шт.", 2),
    # Cabinet 2, block 2 stays empty: the block must be omitted.
    ("2", 20, "Шкаф навесной 800х600х300 мм, IP54", "AE-1076", "шт.", 1),
    ("2", 21, "Клемма проходная винтовая 2,5 мм²", "UT-2.5", "шт.", 50),
    ("Spare parts", 3, None, "АРКС400-ЦП-01", "шт.", 1),
    ("Spare parts", 19, "Коммутатор промышленный управляемый L2, 8x100/1000Base-T", "EX-SW-8G", "шт.", 1),
    ("Spare parts", 25, "Блок питания промышленный 24В DC, 20А", "QUINT4-PS/1AC/24DC/20", "шт.", 1),
)


def demo_cells() -> list[tuple[str, str, object]]:
    """The demo dataset as (sheet, cell, value): one table both engines and the tests use."""
    cells: list[tuple[str, str, object]] = [("Наценки-Скидки", ref, v) for ref, v in DEMO_METADATA.items()]
    for sheet, row, description, article, unit, qty in DEMO_ROWS:
        if description is not None:
            cells.append((sheet, f"C{row}", description))
        cells.extend([(sheet, f"D{row}", article), (sheet, f"F{row}", unit), (sheet, f"G{row}", qty)])
    return cells


def populate_demo_workbook(wb: Workbook) -> None:
    """Write the demo dataset into an in-memory workbook."""
    for sheet, ref, value in demo_cells():
        ws = _find_sheet(wb, sheet, _SHEET_ALIASES.get(sheet, sheet))
        if ws is not None:
            ws[ref] = value


def _demo_workbook(excel_source: str | Path | Workbook) -> Workbook:
    """A private copy with the demo rows; neither the file nor the caller's object changes."""
    source: BytesIO | Path
    if isinstance(excel_source, Workbook):
        source = BytesIO()
        excel_source.save(source)
        source.seek(0)
    else:
        source = Path(excel_source)
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        wb = openpyxl.load_workbook(source, data_only=False)
    populate_demo_workbook(wb)
    return wb


def fill_template(doc: DocxDocument, data: TKPData) -> None:
    """Fill the title bookmarks, the running head and Appendix 3; TemplateError if one is missing."""
    root = doc.element
    number = format_number(data.metadata.number)
    required = (("Num", number), ("Customer", data.metadata.customer), ("Object", data.metadata.object_name))
    missing = [name for name, value in required if not set_bookmark_text(root, name, value)]
    if data.metadata.subject:
        set_bookmark_text(root, "Subject", data.metadata.subject)  # optional, as in the macro
    header_found = False
    for header in _header_elements(doc):
        header_found |= set_bookmark_text(header, "HeaderNum", number)
    if not header_found:
        missing.append("HeaderNum")
    if missing:
        raise TemplateError("В шаблоне нет закладок: " + ", ".join(missing) + ". Документ не создан.")
    populate_spec_table(doc, data.sections)


def _save_atomically(doc: DocxDocument, out: Path) -> None:
    """Write beside the target, then rename: no half-written proposal on a crash,
    and a target held open by Word fails cleanly (PermissionError)."""
    out.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="~$tkp-", suffix=".docx", dir=out.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        doc.save(str(tmp))
        os.replace(tmp, out)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise


def generate_tkp(
    excel_source: str | Path | Workbook | None = None,
    template_path: str | Path | None = None,
    output_path: str | Path | None = None,
    *,
    use_demo_if_empty: bool = False,
    strict: bool = True,
) -> TKPData:
    """Assemble the proposal. Nothing is written when this raises:

    WorkbookError — no positions (unless use_demo_if_empty), an unknown section
    caption, or, with strict, a formula without a saved value; TemplateError — a
    missing bookmark.
    """
    excel_source = excel_source if excel_source is not None else find_default_workbook(ROOT)
    template_path = template_path if template_path is not None else find_default_template(ROOT)
    output_path = output_path if output_path is not None else ROOT / "Технико-коммерческое предложение.docx"

    data = read_tkp_data(excel_source)
    if data.total_items == 0:
        if not use_demo_if_empty:
            raise WorkbookError(EMPTY_WORKBOOK_MESSAGE)
        data = read_tkp_data(_demo_workbook(excel_source))
        data.demo = True
    if strict and data.issues:
        raise WorkbookError(
            "Формулы не вычислены, а сохранённых значений в книге нет "
            "(откройте книгу в Excel, пересчитайте и сохраните):\n  " + "\n  ".join(data.issues)
        )
    doc = open_word_template(template_path)
    fill_template(doc, data)
    _save_atomically(doc, Path(output_path))
    return data


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Сборка технико-коммерческого предложения (Приложение 3 и реквизиты) из Gen_TKP."
    )
    parser.add_argument("-e", "--excel", type=Path, default=None, help="Книга Gen_TKP.xlsm")
    parser.add_argument("-t", "--template", type=Path, default=None, help="Шаблон Word (.dotx / .docx)")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=ROOT / "Технико-коммерческое предложение.docx",
        help="Куда сохранить документ",
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Если в книге нет позиций, собрать документ на демонстрационных данных",
    )
    parser.add_argument(
        "--allow-unresolved",
        action="store_true",
        help="Не останавливаться на формулах без сохранённого значения, только предупредить",
    )
    args = parser.parse_args(argv)
    try:
        excel = args.excel or find_default_workbook(ROOT)
        print(f"Книга: {excel}")
        data = generate_tkp(
            excel_source=excel,
            template_path=args.template,
            output_path=args.output,
            use_demo_if_empty=args.demo,
            strict=not args.allow_unresolved,
        )
    except (WorkbookError, TemplateError, OSError, BadZipFile) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 2
    for issue in data.issues:
        print(f"ВНИМАНИЕ: формула не вычислена — {issue}", file=sys.stderr)
    if data.demo:
        print(
            "ВНИМАНИЕ: в книге нет позиций с количеством больше нуля — "
            "документ собран на демонстрационных данных.",
            file=sys.stderr,
        )
    print(f"Документ сохранен: {args.output} (разделов: {len(data.sections)}, позиций: {data.total_items})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
