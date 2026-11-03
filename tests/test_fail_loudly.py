"""Regression tests for incomplete workbooks and damaged templates."""

import contextlib
import io
import shutil
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from zipfile import ZipFile

import openpyxl
from lxml import etree

from generate_tkp import (
    UNRESOLVED,
    SpecItem,
    SpecSection,
    TemplateError,
    WorkbookError,
    _eval_cell,
    build_preview_rows,
    find_default_workbook,
    generate_tkp,
    main,
    read_tkp_data,
    update_preview_sheet,
)
from tools import sync_assets

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"
TEMPLATE = ROOT / "ТКП_шаблон.dotx"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


class FailLoudlyTests(unittest.TestCase):
    def test_unresolved_quantity_even_when_no_items(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "1"
        ws["C4"], ws["D4"], ws["G4"] = "Module", "A", "=2+3"
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out.docx"
            with self.assertRaisesRegex(WorkbookError, "1!G4"):
                generate_tkp(wb, output_path=out, use_demo_if_empty=True)
            self.assertFalse(out.exists())
            with self.assertRaisesRegex(WorkbookError, "невычисленные формулы"):
                generate_tkp(wb, output_path=out, use_demo_if_empty=True, strict=False)
            self.assertFalse(out.exists())

    def test_unresolved_description_never_emits_formula(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "1"
        ws["C4"], ws["D4"], ws["G4"] = '=D4&" kit"', "ART-1", 1
        data = read_tkp_data(wb)
        self.assertEqual(data.sections[0].items[0].description, "ART-1")
        self.assertEqual(data.issues, ['1!C4: =D4&" kit"'])

    def test_arbitrary_sumif_is_not_spare_parts_formula(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["G3"] = '=SUMIF(A1:A2,"x",B1:B2)'
        self.assertIs(_eval_cell(ws, ws, wb, wb, 3, 7), UNRESOLVED)

    def test_ref_inside_other_formula_is_not_assumed_blank(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["C3"] = '=IFERROR(IF(1,12,#REF!),"")'
        self.assertIs(_eval_cell(ws, ws, wb, wb, 3, 3), UNRESOLVED)

    def test_preview_preserves_formula_like_text(self):
        wb = openpyxl.Workbook()
        item = SpecItem("=not a formula", "0441504", "шт.", 1)
        ws = update_preview_sheet(wb, [SpecSection("test", [item])])
        self.assertEqual(ws["B3"].value, "=not a formula")
        self.assertEqual(ws["B3"].data_type, "s")
        self.assertEqual(ws["C3"].value, "0441504")
        self.assertEqual(ws["C3"].number_format, "@")

    def test_atomic_save_preserves_previous_file_on_replace_failure(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out.docx"
            out.write_bytes(b"previous")
            with (patch("generate_tkp.os.replace", side_effect=PermissionError("locked")),
                  self.assertRaises(PermissionError)):
                generate_tkp(WORKBOOK, output_path=out, use_demo_if_empty=True)
            self.assertEqual(out.read_bytes(), b"previous")
            self.assertEqual(list(Path(tmp).iterdir()), [out])

    def test_cached_value_and_missing_cache(self):
        formulas = openpyxl.Workbook()
        values = openpyxl.Workbook()
        wf, wv = formulas.active, values.active
        wf["G3"] = "=2+3"
        self.assertIs(_eval_cell(wv, wf, values, formulas, 3, 7), UNRESOLVED)
        wv["G3"] = 5
        self.assertEqual(_eval_cell(wv, wf, values, formulas, 3, 7), 5)
        wv["G3"] = ""
        self.assertEqual(_eval_cell(wv, wf, values, formulas, 3, 7), "")

    def test_missing_bookmark_never_replaces_existing_output(self):
        with TemporaryDirectory() as tmp:
            template = Path(tmp) / "broken.dotx"
            with ZipFile(TEMPLATE) as src, ZipFile(template, "w") as dst:
                for info in src.infolist():
                    data = src.read(info.filename)
                    if info.filename == "word/document.xml":
                        root = etree.fromstring(data)
                        for el in root.iter(f"{W}bookmarkStart"):
                            if el.get(f"{W}name") == "Customer":
                                el.getparent().remove(el)
                        data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                    dst.writestr(info, data)
            out = Path(tmp) / "out.docx"
            out.write_bytes(b"previous")
            with self.assertRaisesRegex(TemplateError, "Customer"):
                generate_tkp(WORKBOOK, template, out, use_demo_if_empty=True)
            self.assertEqual(out.read_bytes(), b"previous")

    def test_assets_have_no_author_metadata(self):
        dc = "{http://purl.org/dc/elements/1.1/}"
        cp = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
        for asset in (TEMPLATE, WORKBOOK):
            with ZipFile(asset) as archive:
                core = etree.fromstring(archive.read("docProps/core.xml"))
                if asset == TEMPLATE:
                    app = etree.fromstring(archive.read("docProps/app.xml"))
                    ep = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
                    self.assertEqual(app.findtext(f"{ep}Template"), TEMPLATE.name)
            for tag in (f"{dc}creator", f"{cp}lastModifiedBy"):
                element = core.find(tag)
                self.assertTrue(element is None or not (element.text or "").strip(), asset.name)

    def test_sync_refuses_data_inside_small_preview_dimension(self):
        ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
        with TemporaryDirectory() as tmp:
            book = Path(tmp) / "book.xlsm"
            shutil.copy(WORKBOOK, book)
            with patch.object(sync_assets, "WORKBOOK", book):
                part = sync_assets.preview_sheet_part()
                with ZipFile(book) as archive:
                    root = etree.fromstring(archive.read(part))
                data = root.find(f"{ns}sheetData")
                row = next(r for r in data if r.get("r") == "2")
                cell = etree.SubElement(row, f"{ns}c", r="A2", t="inlineStr")
                etree.SubElement(etree.SubElement(cell, f"{ns}is"), f"{ns}t").text = "private data"
                sync_assets.replace_zip_part(book, {part: sync_assets.write_xml(root)})
                before = book.read_bytes()
                with self.assertRaisesRegex(ValueError, "not empty"):
                    sync_assets.patch_preview_sheet()
                self.assertEqual(book.read_bytes(), before)

    def test_empty_cli_and_explicit_demo(self):
        with TemporaryDirectory() as tmp:
            out = Path(tmp) / "out.docx"
            err = io.StringIO()
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
                code = main(["-e", str(WORKBOOK), "-o", str(out)])
            self.assertEqual(code, 2)
            self.assertFalse(out.exists())
            self.assertIn("Нет позиций", err.getvalue())
            err = io.StringIO()
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(err):
                code = main(["-e", str(WORKBOOK), "-o", str(out), "--demo"])
            self.assertEqual(code, 0)
            self.assertIn("демонстрационных данных", err.getvalue())
            with ZipFile(TEMPLATE) as src, ZipFile(out) as dst:
                def headers(z):
                    return sorted(n for n in z.namelist() if n.startswith("word/header") and n.endswith(".xml"))

                self.assertEqual(headers(src), headers(dst))

    def test_demo_does_not_mutate_workbook(self):
        wb = openpyxl.load_workbook(WORKBOOK)
        before = wb["1"]["G4"].value
        with TemporaryDirectory() as tmp:
            data = generate_tkp(wb, output_path=Path(tmp) / "out.docx", use_demo_if_empty=True)
        self.assertTrue(data.demo)
        self.assertEqual(wb["1"]["G4"].value, before)

    def test_default_book_and_empty_section_numbering(self):
        with TemporaryDirectory() as tmp:
            base = Path(tmp)
            (base / "Gen_TKP.xlsx").touch()
            (base / "Gen_TKP.xlsm").touch()
            self.assertEqual(find_default_workbook(base), base / "Gen_TKP.xlsm")
        rows = build_preview_rows([SpecSection("empty"), SpecSection("full", [SpecItem("x", "a", "шт.", 1)])])
        self.assertEqual(rows[0][0], "1.")
