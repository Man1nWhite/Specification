"""Embed Module1.bas, prepare Word templates (.dotx and .docx), and build demo output.

Run from the repository root with Python 3.10+ after installing requirements-dev.txt.
Excel and Word are not required for this packaging step.
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import ZipFile, ZIP_DEFLATED
import os
import sys

from lxml import etree
from pyopenvba import ExcelFile

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from generate_tkp import (
    DOCUMENT_CONTENT_TYPE,
    TEMPLATE_CONTENT_TYPE,
    generate_tkp,
)

WORKBOOK = ROOT / "Gen_TKP.xlsm"
TEMPLATE_DOTX = ROOT / "ТКП_шаблон.dotx"
TEMPLATE_DOCX = ROOT / "ТКП_шаблон.docx"
TEMPLATE_KP_DOCX = ROOT / "Коммерческое предложение_шаблон.docx"
SOURCE_DOCX = ROOT / "Коммерческое предложение.docx"
RESULT_DOCX = ROOT / "Коммерческое предложение_результат.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
X_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
W = f"{{{W_NS}}}"
X = f"{{{X_NS}}}"
NS = {"w": W_NS}


def replace_zip_part(
    path: Path,
    parts: dict[str, bytes],
    *,
    remove_members: set[str] | None = None,
) -> None:
    """Replace just the requested members; preserve the other ZIP entries."""
    remove_members = remove_members or set()
    mode = path.stat().st_mode & 0o777
    with NamedTemporaryFile(dir=path.parent, suffix=path.suffix, delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        with ZipFile(path) as old, ZipFile(temp_path, "w", ZIP_DEFLATED) as new:
            for info in old.infolist():
                if info.filename in remove_members:
                    continue
                new.writestr(info, parts.get(info.filename, old.read(info.filename)))
        os.replace(temp_path, path)
        os.chmod(path, mode)
    finally:
        temp_path.unlink(missing_ok=True)


def write_xml(element: etree._Element) -> bytes:
    return etree.tostring(element, encoding="UTF-8", xml_declaration=True, standalone=True)


def word_text(element: etree._Element) -> str:
    return "".join(t.text or "" for t in element.findall(".//w:t", NS))


def bookmark(name: str, bookmark_id: int) -> tuple[etree._Element, etree._Element]:
    start = etree.Element(W + "bookmarkStart", {W + "id": str(bookmark_id), W + "name": name})
    end = etree.Element(W + "bookmarkEnd", {W + "id": str(bookmark_id)})
    return start, end


def run(text: str, example: etree._Element | None = None) -> etree._Element:
    result = etree.Element(W + "r")
    if example is not None:
        props = example.find(W + "rPr")
        if props is not None:
            result.append(deepcopy(props))
    value = etree.SubElement(result, W + "t")
    if text.startswith(" ") or text.endswith(" "):
        value.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    value.text = text
    return result


def remove_bookmark(root: etree._Element, name: str) -> None:
    starts = root.xpath(".//w:bookmarkStart[@w:name=$name]", namespaces=NS, name=name)
    if len(starts) != 1:
        raise ValueError(f"Expected one bookmark {name!r}, found {len(starts)}")
    start = starts[0]
    bm_id = start.get(W + "id")
    ends = root.xpath(".//w:bookmarkEnd[@w:id=$id]", namespaces=NS, id=bm_id)
    for end in ends:
        if end.getparent() is not None:
            end.getparent().remove(end)
    if start.getparent() is not None:
        start.getparent().remove(start)


def add_empty_bookmark(paragraph: etree._Element, name: str, number: int) -> None:
    start, end = bookmark(name, number)
    paragraph.extend((start, end))


def _apply_title_and_header_bookmarks(
    body: etree._Element, header: etree._Element
) -> None:
    number_cell = customer_cell = object_cell = subject_cell = None
    for table in body.findall(".//w:tbl", NS):
        rows = table.findall("w:tr", NS)
        if len(rows) < 4:
            continue
        cells = [row.findall("w:tc", NS) for row in rows[:4]]
        if all(len(row) >= 2 for row in cells) and [word_text(row[0]) for row in cells[:3]] == [
            "Номер",
            "Заказчик",
            "Объект",
        ]:
            number_cell, customer_cell, object_cell, subject_cell = [row[1] for row in cells]
            break
    if number_cell is None:
        raise ValueError("Could not locate the customer details table")

    number_p = number_cell.find("w:p", NS)
    assert number_p is not None
    original_run = number_p.find("w:r", NS)
    for child in list(number_p):
        if child.tag != W + "pPr":
            number_p.remove(child)
    number_start, number_end = bookmark("Num", 23)
    number_p.extend((number_start, run("__/1 от __.__.2026", original_run), number_end))

    for cell, name, number in ((customer_cell, "Customer", 24), (object_cell, "Object", 25)):
        p = cell.find("w:p", NS)
        assert p is not None
        for child in list(p):
            if child.tag != W + "pPr":
                p.remove(child)
        add_empty_bookmark(p, name, number)

    if subject_cell is not None:
        subject_p = subject_cell.find("w:p", NS)
        assert subject_p is not None
        for child in list(subject_p):
            if child.tag != W + "pPr":
                subject_p.remove(child)
        subject_start, subject_end = bookmark("Subject", 29)
        subject_p.extend(
            (
                subject_start,
                run(
                    "Поставка оборудования и выполнение работ по созданию АСУ ТП "
                    "на базе программно-технического комплекса АРКС400",
                    original_run,
                ),
                subject_end,
            )
        )

    running_header = next(
        (p for p in header.findall(".//w:p", NS) if "Коммерческое предложение" in word_text(p)),
        None,
    )
    if running_header is None:
        raise ValueError("Could not locate the running header")
    example = running_header.find("w:r", NS)
    for child in list(running_header):
        if child.tag != W + "pPr":
            running_header.remove(child)
    header_start, header_end = bookmark("HeaderNum", 1000)
    running_header.extend(
        (
            run("Коммерческое предложение ", example),
            header_start,
            run("№_/1 от ..2026", example),
            header_end,
        )
    )


def _fix_notes_paragraph(notes: etree._Element) -> None:
    for t_el in notes.findall(".//w:t", NS):
        if t_el.text and t_el.text.startswith("римечания к Приложению"):
            t_el.text = "П" + t_el.text


def patch_template() -> None:
    """Update ТКП_шаблон.dotx and write .docx template copies."""
    with ZipFile(TEMPLATE_DOTX) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        header = etree.fromstring(archive.read("word/header1.xml"))
        settings = etree.fromstring(archive.read("word/settings.xml"))

    for attached in settings.findall(W + "attachedTemplate"):
        settings.remove(attached)

    body = document.find(".//w:body", NS)
    assert body is not None
    note_paragraphs = [
        p for p in body.findall("w:p", NS) if "римечания к Приложению" in word_text(p)
    ]
    if len(note_paragraphs) != 1:
        raise ValueError("Could not locate the notes after Appendix 3")
    notes = note_paragraphs[0]
    _fix_notes_paragraph(notes)

    spec_start = document.xpath(".//w:bookmarkStart[@w:name='Spec_Table']", namespaces=NS)[0]
    old_spec_paragraph = spec_start.getparent()
    for name in ("Num", "Customer", "Object", "Spec_Table"):
        remove_bookmark(document, name)
    if document.xpath(".//w:bookmarkStart[@w:name='Subject']", namespaces=NS):
        remove_bookmark(document, "Subject")
    if old_spec_paragraph is not notes and not word_text(old_spec_paragraph):
        if old_spec_paragraph.getparent() is body:
            body.remove(old_spec_paragraph)

    _apply_title_and_header_bookmarks(body, header)

    spec_p = etree.Element(W + "p")
    add_empty_bookmark(spec_p, "Spec_Table", 26)
    body.insert(body.index(notes), spec_p)

    replace_zip_part(
        TEMPLATE_DOTX,
        {
            "word/document.xml": write_xml(document),
            "word/header1.xml": write_xml(header),
            "word/settings.xml": write_xml(settings),
        },
        remove_members={"word/_rels/settings.xml.rels"},
    )

    # Also write .docx template variants with document.main+xml ContentType
    for docx_target in (TEMPLATE_DOCX, TEMPLATE_KP_DOCX):
        with ZipFile(TEMPLATE_DOTX, "r") as zin, ZipFile(docx_target, "w", ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    data = (
                        data.decode("utf-8")
                        .replace(TEMPLATE_CONTENT_TYPE, DOCUMENT_CONTENT_TYPE)
                        .encode("utf-8")
                    )
                zout.writestr(info, data)


def patch_source_docx() -> None:
    """Ensure Коммерческое предложение.docx has clean bookmarks and no broken template link."""
    if not SOURCE_DOCX.exists():
        return
    with ZipFile(SOURCE_DOCX) as archive:
        document = etree.fromstring(archive.read("word/document.xml"))
        header = etree.fromstring(archive.read("word/header1.xml"))
        settings = etree.fromstring(archive.read("word/settings.xml"))

    for attached in settings.findall(W + "attachedTemplate"):
        settings.remove(attached)

    body = document.find(".//w:body", NS)
    assert body is not None
    note_paragraphs = [
        p for p in body.findall("w:p", NS) if "римечания к Приложению" in word_text(p)
    ]
    if note_paragraphs:
        _fix_notes_paragraph(note_paragraphs[0])

    for name in ("Num", "Customer", "Object", "Subject"):
        if document.xpath(".//w:bookmarkStart[@w:name=$name]", namespaces=NS, name=name):
            remove_bookmark(document, name)
    _apply_title_and_header_bookmarks(body, header)

    # Ensure Spec_Table bookmark inside Table 4 has a clean matching bookmarkEnd
    spec_starts = document.xpath(".//w:bookmarkStart[@w:name='Spec_Table']", namespaces=NS)
    if spec_starts:
        s = spec_starts[0]
        p = s.getparent()
        remove_bookmark(document, "Spec_Table")
        if p is not None:
            pPr = p.find(W + "pPr")
            insert_idx = (list(p).index(pPr) + 1) if pPr is not None else 0
            st, en = bookmark("Spec_Table", 26)
            p.insert(insert_idx, en)
            p.insert(insert_idx, st)

    replace_zip_part(
        SOURCE_DOCX,
        {
            "word/document.xml": write_xml(document),
            "word/header1.xml": write_xml(header),
            "word/settings.xml": write_xml(settings),
        },
        remove_members={"word/_rels/settings.xml.rels"},
    )


def inline_cell(address: str, text: str) -> etree._Element:
    cell = etree.Element(X + "c", r=address, t="inlineStr")
    etree.SubElement(etree.SubElement(cell, X + "is"), X + "t").text = text
    return cell


def patch_preview_sheet() -> None:
    """Initialize the preview and keep its helper marker column hidden."""
    part = "xl/worksheets/sheet26.xml"
    with ZipFile(WORKBOOK) as archive:
        root = etree.fromstring(archive.read(part))
    data = root.find(X + "sheetData")
    dimension = root.find(X + "dimension")
    assert data is not None and dimension is not None

    if dimension.get("ref") == "A1:G3":
        header = root.xpath(".//x:c[@r='A1']", namespaces={"x": X_NS})
        if header:
            header[0].clear()
            header[0].set("r", "A1")
            header[0].set("t", "inlineStr")
            header[0].append(etree.Element(X + "is"))
            etree.SubElement(header[0][0], X + "t").text = "№ п/п"
        columns = root.find(X + "cols")
        if columns is not None:
            for col in columns.findall(X + "col"):
                if col.get("min") == "6" and col.get("max") == "6":
                    col.set("hidden", "1")
        replace_zip_part(WORKBOOK, {part: write_xml(root)})
        return

    if len(data) != 1 or data[0].get("r") != "1":
        raise ValueError("The preview sheet is not empty; refusing to erase it")
    data.remove(data[0])
    hints = (
        "Alt+F8: СформироватьТКП (Word) или ОбновитьСпецификациюТКП (предпросмотр)",
        "Сначала заполните позиции и количество на листах 1-17, Spare parts и SW(w.o.VAT).",
        "Берутся C/D/F/G и только количество > 0; цены, итоги и работы не переносятся.",
    )
    first = etree.SubElement(data, X + "row", r="1", spans="1:7")
    for letter, value in zip(
        "ABCDE",
        ("№ п/п", "Наименование и техническая характеристика", "Артикул", "Ед. изм.", "Кол-во"),
    ):
        first.append(inline_cell(letter + "1", value))
    first.append(inline_cell("G1", hints[0]))
    for row_num in (2, 3):
        row = etree.SubElement(data, X + "row", r=str(row_num), spans="7:7")
        row.append(inline_cell(f"G{row_num}", hints[row_num - 1]))
    dimension.set("ref", "A1:G3")
    sheet_format = root.find(X + "sheetFormatPr")
    assert sheet_format is not None
    columns = etree.Element(X + "cols")
    for col_no, width in ((1, 7), (2, 65), (3, 27), (4, 13), (5, 12), (6, 3), (7, 78)):
        attributes = {
            "min": str(col_no),
            "max": str(col_no),
            "width": str(width),
            "customWidth": "1",
        }
        if col_no == 6:
            attributes["hidden"] = "1"
        etree.SubElement(columns, X + "col", **attributes)
    root.insert(root.index(sheet_format) + 1, columns)
    replace_zip_part(WORKBOOK, {part: write_xml(root)})


def embed_macro() -> None:
    source = (ROOT / "Module1.bas").read_text(encoding="utf-8").replace("\r\n", "\n")
    source = source.replace("\n", "\r\n")
    with ExcelFile(WORKBOOK) as book:
        if book.get_module("Module1").replace("\r\n", "\n").rstrip() != source.replace("\r\n", "\n").rstrip():
            book.set_module("Module1", source)
            book.save(allow_protected=True)
        if book.get_module("Module1").replace("\r\n", "\n").rstrip() != source.replace("\r\n", "\n").rstrip():
            raise ValueError("Embedded VBA source does not match Module1.bas")


if __name__ == "__main__":
    patch_template()
    patch_source_docx()
    patch_preview_sheet()
    embed_macro()
    generate_tkp(
        excel_source=WORKBOOK,
        template_path=TEMPLATE_DOCX,
        output_path=RESULT_DOCX,
        use_demo_if_empty=True,
    )
    print("Updated templates, preview sheet, embedded VBA macro, and demo output document.")
