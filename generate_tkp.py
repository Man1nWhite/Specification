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


class WorkbookError(ValueError):
    """The workbook cannot be used to create a complete proposal."""


class TemplateError(ValueError):
    """A required template bookmark is missing or invalid."""


class _Unresolved:
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
    issues: list[str] = field(default_factory=list)
    demo: bool = False

    @property
    def total_items(self) -> int:
        return sum(len(s.items) for s in self.sections)


def clean_text(value: object) -> str:
    """Normalize cell values into a single-line trimmed string, ignoring errors/None."""
    if value is None or value is UNRESOLVED or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value)
    for ch in ("\r\n", "\r", "\n", "\t"):
        text = text.replace(ch, " ")
    text = text.strip()
    if text in EXCEL_ERRORS:
        return ""
    return text


def format_qty(qty: float) -> str:
    """Plain decimal notation, at most six decimals, with a decimal comma."""
    scaled = qty * 1_000_000
    rounded = round(scaled) / 1_000_000 if math.isfinite(scaled) else qty
    if rounded == 0:
        return "0"
    return f"{rounded:.6f}".rstrip("0").rstrip(".").replace(".", ",")


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
_ZIP_TERM_RE = re.compile(r"SUMIF\('([^']+)'!\$D\$3:\$D\$100,D(\d+),'([^']+)'!\$G\$3:\$G\$100\)")


def _zip_rollup_sheets(formula: str, row: int) -> list[str] | None:
    compact = re.sub(r"\s+", "", formula)
    head, tail = "=ROUNDUP((", ")/10,0)"
    if not compact.startswith(head) or not compact.endswith(tail):
        return None
    body = compact[len(head):-len(tail)]
    terms = _ZIP_TERM_RE.findall(body)
    rebuilt = "+".join(f"SUMIF('{a}'!$D$3:$D$100,D{r},'{b}'!$G$3:$G$100)" for a, r, b in terms)
    if not terms or rebuilt != body or any(a != b or int(r) != row for a, r, b in terms):
        return None
    return [a for a, _, _ in terms]


def _eval_cell(
    ws_val: Worksheet | None,
    ws_fml: Worksheet | None,
    wb_val: Workbook | None,
    wb_fml: Workbook | None,
    row: int,
    col: int,
    _seen: frozenset[tuple[str, int, int]] = frozenset(),
) -> object:
    """Resolve supported formulas, otherwise use a real saved cache or UNRESOLVED."""
    if ws_fml is None:
        return ws_val.cell(row, col).value if ws_val is not None else None
    key = (ws_fml.title, row, col)
    if key in _seen:
        return ""
    seen = _seen | {key}
    cell = ws_fml.cell(row, col)
    raw = cell.value
    if cell.data_type != "f" or not isinstance(raw, str):
        return raw

    # Only the exact broken lookup in the shipped book has a known empty result.
    if re.fullmatch(r'=IFERROR\(INDEX\(#REF!,MATCH\(D[0-9]+,#REF!,0\)\),""\)', raw, re.I):
        return ""
    match = _SHEET_REF_RE.match(raw.strip())
    if match and wb_fml is not None and match.group(1) in wb_fml.sheetnames:
        target, letters, digits = match.groups()
        val = wb_val[target] if wb_val is not None and target in wb_val.sheetnames else None
        resolved = _eval_cell(val, wb_fml[target], wb_val, wb_fml, int(digits),
                              openpyxl.utils.column_index_from_string(letters), seen)
        if resolved is None:
            return 0
        if resolved is not UNRESOLVED:
            return resolved

    sheets = _zip_rollup_sheets(raw, row) if ws_fml.title in ("Spare parts", "ЗИП") and col == 7 else None
    if sheets is not None and wb_fml is not None:
        article_value = _eval_cell(ws_val, ws_fml, wb_val, wb_fml, row, 4, seen)
        if article_value is not UNRESOLVED:
            article = clean_text(article_value).lower()
            if not article:
                return 0
            # SUMIF interprets wildcards and leading comparison operators; do
            # not silently substitute literal matching for those criteria.
            special = (not isinstance(article_value, str) or article != article_value.lower()
                       or any(c in article for c in "*?~") or article.startswith(("<", ">", "=")))
            total = 0.0
            complete = not special
            for name in sheets:
                if not complete:
                    break
                if name not in wb_fml.sheetnames:
                    complete = False
                    break
                cf = wb_fml[name]
                cv = wb_val[name] if wb_val is not None and name in wb_val.sheetnames else None
                for r in range(3, min(cf.max_row, 100) + 1):
                    candidate = _eval_cell(cv, cf, wb_val, wb_fml, r, 4, seen)
                    if candidate is UNRESOLVED:
                        complete = False
                        break
                    if clean_text(candidate).lower() == article:
                        qty = _eval_cell(cv, cf, wb_val, wb_fml, r, 7, seen)
                        if qty is UNRESOLVED:
                            complete = False
                            break
                        if isinstance(qty, (int, float)) and not isinstance(qty, bool):
                            total += float(qty)
                if not complete:
                    break
            if complete:
                return math.ceil(total / 10) if total > 0 else 0

    if ws_val is None or ws_val is ws_fml:
        return UNRESOLVED
    cell = ws_val.cell(row, col)
    if cell.value is not None:
        return cell.value
    # Excel's cached empty string has a string type, unlike a missing cache.
    return "" if cell.data_type in ("s", "str") else UNRESOLVED


def _describe_unresolved(ws: Worksheet, row: int, col: int) -> str:
    formula = " ".join(str(ws.cell(row, col).value or "").split())
    return f"{ws.title}!{openpyxl.utils.get_column_letter(col)}{row}: {formula[:80]}"


_PLAIN_NUMBER_RE = re.compile(r"[+-]?(?:[0-9]+(?:[.,][0-9]*)?|[.,][0-9]+)")
CABINET_BLOCK_TITLES = {
    1: "Оборудование ПТК",
    2: "Телекоммуникационное оборудование для АСУ ТП",
    3: "Шкафное оборудование ПТК",
}


def _parse_numeric_qty(raw_qty: object) -> float:
    """Accept numeric cells and plain decimal text, independent of locale."""
    if raw_qty is None or isinstance(raw_qty, bool):
        return 0.0
    if isinstance(raw_qty, (int, float)):
        try:
            value = float(raw_qty)
        except OverflowError:
            return 0.0
    elif isinstance(raw_qty, str) and _PLAIN_NUMBER_RE.fullmatch(raw_qty.strip(" ")):
        value = float(raw_qty.strip(" ").replace(",", "."))
    else:
        return 0.0
    return value if math.isfinite(value) else 0.0


def _sheet_last_row(ws: Worksheet) -> int:
    """Mirror the macro: last populated B, C, D or G, at least row 48."""
    for row in range(ws.max_row or 0, 48, -1):
        if any(ws.cell(row, col).value not in (None, "") for col in (2, 3, 4, 7)):
            return row
    return 48


def _collect_sheet_sections(
    ws_val: Worksheet | None,
    ws_fml: Worksheet,
    wb_val: Workbook | None,
    wb_fml: Workbook,
    *,
    cab_no: int,
    fixed_title: str,
    issues: list[str],
) -> list[SpecSection]:
    """Scan once, preserving caption and section order as in the VBA macro."""
    sections: list[SpecSection] = []
    current: SpecSection | None = None
    current_block = prev_block = 0
    has_captions = False

    def value(row: int, col: int) -> object:
        return _eval_cell(ws_val, ws_fml, wb_val, wb_fml, row, col)

    for row in range(1, _sheet_last_row(ws_fml) + 1):
        caption = clean_text(value(row, 2))
        description_raw = value(row, 3)
        description = clean_text(description_raw)
        if caption.lower().startswith("спецификация "):
            current_block = _block_from_caption(caption)
            if not current_block:
                raise WorkbookError(
                    f"Лист «{ws_fml.title}», строка {row}: неизвестный раздел «{caption}». "
                    "Ожидается «НВТ», «Эксара»/«ИнСАТ» или «3-х поставщиков»."
                )
            has_captions = True
            if cab_no:
                current = None
            continue
        if caption.lower().startswith("итого") or description.lower().startswith(("итого", "маржа")):
            current_block = 0
            continue
        if caption.lower() == "№ п/п":
            continue
        block = current_block
        if block == 0 and not has_captions:
            block = _block_from_row(row)
            if cab_no and block != prev_block:
                current, prev_block = None, block
        if not block:
            continue
        raw_qty = value(row, 7)
        if raw_qty is UNRESOLVED:
            issues.append(_describe_unresolved(ws_fml, row, 7))
            continue
        qty = _parse_numeric_qty(raw_qty)
        if qty <= 0:
            continue
        values = [description_raw, value(row, 4), value(row, 6)]
        for col, cell_value in zip((3, 4, 6), values, strict=True):
            if cell_value is UNRESOLVED:
                issues.append(_describe_unresolved(ws_fml, row, col))
        description, article, unit = map(clean_text, values)
        if not description or description == "0":
            description = article
        if not description:
            continue
        if current is None:
            title = f"{CABINET_BLOCK_TITLES[block]} (шкаф {cab_no})" if cab_no else fixed_title
            current = SpecSection(title)
            sections.append(current)
        current.items.append(SpecItem(description, article, unit or "шт.", qty))
    return sections


def read_tkp_data(excel_source: str | Path | Workbook) -> TKPData:
    """Read title metadata and Appendix 3 specification sections from Excel."""
    if isinstance(excel_source, Workbook):
        wb_fml = excel_source
        wb_val = excel_source
    else:
        path = Path(excel_source)
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
            wb_val = openpyxl.load_workbook(path, data_only=True)
            wb_fml = openpyxl.load_workbook(path, data_only=False)

    issues: list[str] = []
    info_fml = _find_sheet(wb_fml, "Наценки-Скидки")
    info_val = _find_sheet(wb_val, "Наценки-Скидки") if wb_val is not None else None
    metadata = TKPMetadata()
    if info_fml is not None:
        values = []
        for r in (2, 3, 4, 5):
            value = _eval_cell(info_val, info_fml, wb_val, wb_fml, r, 3)
            if value is UNRESOLVED:
                issues.append(_describe_unresolved(info_fml, r, 3))
            values.append(clean_text(value))
        metadata = TKPMetadata(*values)

    sections: list[SpecSection] = []
    sheets = [(("SW(w.o.VAT)", "ПО"), 0, "Программное обеспечение")]
    sheets += [((str(n), f"Шкаф {n}"), n, "") for n in range(1, 18)]
    sheets.append((("Spare parts", "ЗИП"), 0, "ЗИП"))
    for names, cab_no, title in sheets:
        ws = _find_sheet(wb_fml, *names)
        if ws is not None:
            sections.extend(_collect_sheet_sections(
                wb_val[ws.title] if ws.title in wb_val.sheetnames else None,
                ws, wb_val, wb_fml, cab_no=cab_no, fixed_title=title, issues=issues,
            ))

    return TKPData(metadata=metadata, sections=sections, issues=issues)


def build_preview_rows(sections: Iterable[SpecSection]) -> list[tuple[str, str, str, str, float | int, str]]:
    """Build the preview table rows matching sheet 'Спец_ТКП' (columns A..F)."""
    rows: list[tuple[str, str, str, str, float | int, str]] = []
    sec_idx = 0
    for section in sections:
        if not section.items:
            continue
        sec_idx += 1
        rows.append((f"{sec_idx}.", section.title, "", "Компл.", 1, "H"))
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
            if col_idx <= 4:
                c.number_format = "@"
                # openpyxl otherwise infers a formula for text starting with '='.
                c.data_type = "s"
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


def set_bookmark_text(root: etree._Element, name: str, text: str) -> bool:
    """Set the text inside a Word bookmark while preserving paragraph/run formatting."""
    starts = [
        el
        for el in root.findall(f".//{W}bookmarkStart")
        if el.get(f"{W}name") == name
    ]
    if not starts:
        return False

    start = starts[0]
    bm_id = start.get(f"{W}id")
    parent = start.getparent()
    if parent is None or parent.tag != f"{W}p":
        return False

    ends = [
        el
        for el in root.findall(f".//{W}bookmarkEnd")
        if el.get(f"{W}id") == bm_id
    ]
    end = ends[0] if ends else None

    rpr_template = None
    children = list(parent)
    first = children.index(start) + 1
    stop = children.index(end) if end is not None and end.getparent() is parent else first
    for child in children[first:stop] + children[:first - 1][::-1]:
        if child.tag == f"{W}r":
            props = child.find(f"{W}rPr")
            rpr_template = deepcopy(props) if props is not None else None
            break
    else:
        ppr = parent.find(f"{W}pPr")
        props = ppr.find(f"{W}rPr") if ppr is not None else None
        if props is not None:
            rpr_template = deepcopy(props)
            for child in list(rpr_template):
                if child.tag in {f"{W}{tag}" for tag in ("ins", "del", "moveFrom", "moveTo", "rPrChange")}:
                    rpr_template.remove(child)

    if end is not None and end.getparent() is parent:
        start_idx = list(parent).index(start)
        end_idx = list(parent).index(end)
        for child in list(parent)[start_idx + 1 : end_idx]:
            parent.remove(child)
    else:
        if end is not None and end.getparent() is not None:
            end.getparent().remove(end)
        end = etree.Element(f"{W}bookmarkEnd", {f"{W}id": bm_id or "0"})
        parent.insert(list(parent).index(start) + 1, end)

    new_run = etree.Element(f"{W}r")
    if rpr_template is not None:
        new_run.append(rpr_template)
    t_el = etree.SubElement(new_run, f"{W}t")
    if text.startswith(" ") or text.endswith(" "):
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_el.text = text
    parent.insert(list(parent).index(end), new_run)
    return True


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
    grid_twips = (702, 5357, 1974, 988, 759)
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
    doc_el = doc.element
    body = doc_el.find(f".//{W}body")
    starts = [
        el
        for el in doc_el.findall(f".//{W}bookmarkStart")
        if el.get(f"{W}name") == "Spec_Table"
    ]
    if not starts or body is None:
        raise TemplateError("В шаблоне нет закладки Spec_Table.")
    anchor = starts[0]
    while anchor.getparent() is not None and anchor.getparent() is not body:
        anchor = anchor.getparent()
    if anchor.getparent() is not body:
        raise TemplateError("Закладка Spec_Table находится вне документа.")
    if starts and body is not None:
        bm_start = starts[0]
        ids = [int(v) for el in doc_el.iter(f"{W}bookmarkStart")
               if (v := el.get(f"{W}id", "")).isdigit()]
        bm_id = bm_start.get(f"{W}id") or str(max(ids, default=0) + 1)
        bm_ends = [
            el
            for el in doc_el.findall(f".//{W}bookmarkEnd")
            if el.get(f"{W}id") == bm_id
        ]
        for end_el in bm_ends:
            if end_el.getparent() is not None:
                end_el.getparent().remove(end_el)
        if bm_start.getparent() is not None:
            anchor = bm_start
            while anchor.getparent() is not None and anchor.getparent() is not body:
                anchor = anchor.getparent()
            if anchor.getparent() is body:
                body.insert(list(body).index(anchor), table._tbl)
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
    for name in ("Gen_TKP.xlsm", "Gen_TKP.xlsx"):
        candidate = base_dir / name
        if candidate.exists():
            return candidate
    raise FileNotFoundError(f"Excel generator workbook not found in {base_dir}")


DEMO_METADATA = {
    "C2": "№ 104/1 от 16.10.2026",
    "C3": "ДЕМОНСТРАЦИОННЫЕ ДАННЫЕ — заказчик не указан",
    "C4": "ДЕМО: АСУ ТП компрессорного цеха",
    "C5": "Поставка оборудования и выполнение работ по созданию АСУ ТП "
          "на базе программно-технического комплекса АРКС400",
}

# sheet, row, description (None keeps the broken lookup), article, unit, quantity
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
    ("2", 20, "Шкаф навесной 800х600х300 мм, IP54", "AE-1076", "шт.", 1),
    ("2", 21, "Клемма проходная винтовая 2,5 мм²", "UT-2.5", "шт.", 50),
    ("Spare parts", 3, None, "АРКС400-ЦП-01", "шт.", 1),
    ("Spare parts", 19, "Коммутатор промышленный управляемый L2, 8x100/1000Base-T", "EX-SW-8G", "шт.", 1),
    ("Spare parts", 25, "Блок питания промышленный 24В DC, 20А", "QUINT4-PS/1AC/24DC/20", "шт.", 1),
)


def demo_cells() -> list[tuple[str, str, object]]:
    """The same demo input cells for both engines and tests."""
    cells: list[tuple[str, str, object]] = [("Наценки-Скидки", ref, v) for ref, v in DEMO_METADATA.items()]
    for sheet, row, description, article, unit, qty in DEMO_ROWS:
        if description is not None:
            cells.append((sheet, f"C{row}", description))
        cells.extend([(sheet, f"D{row}", article), (sheet, f"F{row}", unit), (sheet, f"G{row}", qty)])
    return cells


def populate_demo_workbook(wb: Workbook) -> None:
    """Populate a private workbook copy with clearly labelled demo data."""
    aliases = {"SW(w.o.VAT)": "ПО", "Spare parts": "ЗИП"}
    aliases.update({str(n): f"Шкаф {n}" for n in range(1, 18)})
    for name, ref, value in demo_cells():
        ws = _find_sheet(wb, name, aliases.get(name, name))
        if ws is not None:
            ws[ref] = value


def generate_tkp(
    excel_source: str | Path | Workbook | None = None,
    template_path: str | Path | None = None,
    output_path: str | Path | None = None,
    *,
    use_demo_if_empty: bool = False,
    strict: bool = True,
) -> TKPData:
    """Generate a proposal, refusing unresolved formulas and empty inputs by default."""
    excel_source = excel_source if excel_source is not None else find_default_workbook(ROOT)
    template_path = template_path if template_path is not None else find_default_template(ROOT)
    output_path = Path(output_path) if output_path is not None else ROOT / "Технико-коммерческое предложение.docx"
    data = read_tkp_data(excel_source)
    # Check issues first: unresolved quantities are not proof of an empty book.
    if strict and data.issues:
        raise WorkbookError("Формулы не вычислены; откройте книгу в Excel, пересчитайте и сохраните:\n  "
                            + "\n  ".join(data.issues))
    if data.total_items == 0:
        if data.issues and use_demo_if_empty:
            raise WorkbookError("Нельзя считать книгу пустой: есть невычисленные формулы. "
                                "Пересчитайте и сохраните книгу перед запуском --demo.")
        if not use_demo_if_empty:
            raise WorkbookError("Нет позиций с количеством больше нуля. "
                                "Заполните столбец G на листах оборудования, ЗИП или ПО.")
        if isinstance(excel_source, Workbook):
            source = BytesIO()
            excel_source.save(source)
            source.seek(0)
        else:
            source = excel_source
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
            wb = openpyxl.load_workbook(source, data_only=False)
        populate_demo_workbook(wb)
        data = read_tkp_data(wb)
        data.demo = True
        if strict and data.issues:
            raise WorkbookError("Формулы не вычислены:\n  " + "\n  ".join(data.issues))

    doc = open_word_template(template_path)
    number = format_number(data.metadata.number)
    missing = [name for name, value in (("Num", number), ("Customer", data.metadata.customer),
                                        ("Object", data.metadata.object_name))
               if not set_bookmark_text(doc.element, name, value)]
    if data.metadata.subject:
        set_bookmark_text(doc.element, "Subject", data.metadata.subject)
    headers = [rel.target_part.element for rel in doc.part.rels.values()
               if rel.reltype == RT.HEADER and not rel.is_external]
    found = False
    for header in headers:
        found = set_bookmark_text(header, "HeaderNum", number) or found
    if not found:
        missing.append("HeaderNum")
    if missing:
        raise TemplateError("В шаблоне нет закладок: " + ", ".join(missing))
    populate_spec_table(doc, data.sections)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="~$tkp-", suffix=".docx", dir=output_path.parent)
    os.close(fd)
    try:
        doc.save(tmp_name)
        os.replace(tmp_name, output_path)
    finally:
        Path(tmp_name).unlink(missing_ok=True)
    return data


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Сборка ТКП из книги Gen_TKP.")
    parser.add_argument("-e", "--excel", type=Path, default=None)
    parser.add_argument("-t", "--template", type=Path, default=None)
    parser.add_argument("-o", "--output", type=Path, default=ROOT / "Технико-коммерческое предложение.docx")
    parser.add_argument("--demo", action="store_true", help="Использовать демонстрационные данные для пустой книги")
    parser.add_argument("--allow-unresolved", action="store_true",
                        help="Предупреждать вместо остановки на невычисленных формулах")
    args = parser.parse_args(argv)
    try:
        excel = args.excel or find_default_workbook(ROOT)
        print(f"Книга: {excel}")
        data = generate_tkp(excel, args.template, args.output, use_demo_if_empty=args.demo,
                            strict=not args.allow_unresolved)
    except (WorkbookError, TemplateError, OSError, BadZipFile) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 2
    for issue in data.issues:
        print(f"ВНИМАНИЕ: формула не вычислена — {issue}", file=sys.stderr)
    if data.demo:
        print("ВНИМАНИЕ: документ собран на демонстрационных данных.", file=sys.stderr)
    print(f"Документ сохранен: {args.output} (разделов: {len(data.sections)}, позиций: {data.total_items})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
