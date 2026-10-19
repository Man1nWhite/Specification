"""Static packaging checks, headless VBA preview test, and Python generator tests."""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

import docx
from lxml import etree
import openpyxl
from pyopenvba import ExcelFile
from pyopenvba.apps.excel import ExcelApplication

from generate_tkp import (
    build_preview_rows,
    find_default_template,
    generate_tkp,
    read_tkp_data,
)

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"
TEMPLATE = ROOT / "ТКП_шаблон.dotx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"w": W}
SHEET_NS = {"m": M}


def normalise_preview_rows(rows):
    """Normalise preview rows from either engine so they can be compared.

    Excel returns every number as a float, so an integral quantity comes back as
    2.0 where the Python generator builds 2. Collapsing integral floats keeps the
    comparison faithful, unlike an int() cast, which would silently turn 2.9 into
    2 and hide a real difference between the engines.
    """
    normalised = []
    for num, description, article, unit, qty, marker in rows:
        if isinstance(qty, float) and qty.is_integer():
            qty = int(qty)
        normalised.append((num, description, article or "", unit, qty, marker))
    return normalised


class TKPTest(unittest.TestCase):
    def test_embedded_source_matches_editable_source(self):
        expected = (ROOT / "Module1.bas").read_text(encoding="utf-8").replace("\r\n", "\n").strip()
        with ExcelFile(WORKBOOK) as book:
            actual = book.get_module("Module1").replace("\r\n", "\n").strip()
            self.assertEqual(actual, expected)
            self.assertIn("Module12", book.module_names())  # existing macros preserved
        self.assertIn('"ТКП_шаблон.dotx"', expected)
        self.assertNotIn("ТКП_шаблон.docx", expected)
        self.assertNotIn("Коммерческое предложение_шаблон.docx", expected)
        self.assertNotIn("Коммерческое предложение.docx", expected)

    def test_bookmarks_are_in_their_destination_not_the_notes(self):
        self.assertTrue(TEMPLATE.exists())
        with ZipFile(TEMPLATE) as archive:
            document = etree.fromstring(archive.read("word/document.xml"))
            header = etree.fromstring(archive.read("word/header1.xml"))
        body = document.find(".//w:body", NS)
        self.assertIsNotNone(body)
        notes = next(
            p
            for p in body.findall("w:p", NS)
            if "Примечания к Приложению" in "".join(p.itertext())
        )
        target = {}
        for name in ("Num", "Customer", "Object", "Subject", "Spec_Table"):
            starts = document.xpath(
                ".//w:bookmarkStart[@w:name=$name]", namespaces=NS, name=name
            )
            self.assertEqual(len(starts), 1)
            self.assertIsNot(starts[0].getparent(), notes)
            target[name] = starts[0]
        self.assertLess(body.index(target["Spec_Table"].getparent()), body.index(notes))
        self.assertEqual(
            len(header.xpath(".//w:bookmarkStart[@w:name='HeaderNum']", namespaces=NS)),
            1,
        )
        for name in ("Num", "Customer", "Object", "Subject"):
            self.assertEqual(target[name].getparent().tag, f"{{{W}}}p")
            self.assertEqual(target[name].getparent().getparent().tag, f"{{{W}}}tc")

    def test_default_template_is_only_the_dotx_asset(self):
        self.assertEqual(find_default_template(ROOT), TEMPLATE)

        with TemporaryDirectory() as tmp:
            base_dir = Path(tmp)
            for name in (
                "ТКП_шаблон.docx",
                "Коммерческое предложение_шаблон.docx",
                "Коммерческое предложение.docx",
            ):
                (base_dir / name).touch()
            with self.assertRaises(FileNotFoundError):
                find_default_template(base_dir)

    def test_preview_has_only_the_standard_specification_columns(self):
        with ZipFile(WORKBOOK) as archive:
            sheet = etree.fromstring(archive.read("xl/worksheets/sheet26.xml"))
        header = sheet.xpath(".//m:c[@r='A1']//m:t/text()", namespaces=SHEET_NS)
        self.assertEqual(header, ["№ п/п"])
        helper_column = sheet.xpath(".//m:col[@min='6' and @max='6']", namespaces=SHEET_NS)
        self.assertEqual(len(helper_column), 1)
        self.assertEqual(helper_column[0].get("hidden"), "1")

        source = (ROOT / "Module1.bas").read_text(encoding="utf-8")
        self.assertNotIn('"Тип"', source)
        self.assertIn("For col = 1 To 5", source)  # Word receives visible A:E verbatim.

    def test_generation_groups_sums_skips_and_refreshes(self):
        app = ExcelApplication.open(WORKBOOK)
        rows = [
            ("SW(w.o.VAT)", 14, "Лицензия ПО", "MasterPLC.N1000", 2),
            ("1", 3, None, "NV-1", 2),  # broken lookup in C: fall back to article D
            ("1", 14, "Шкаф Эксара", "EX-1", 4),
            ("1", 20, "Модуль ввода", "X-123", 2),
            ("2", 20, "Модуль ввода", "X-123", 3),
            ("1", 21, "Не включать", "X-0", 0),
            ("Spare parts", 25, "Реле ЗИП", "ZIP-1", 1),
        ]
        for sheet, row, description, article, qty in rows:
            source = app.sheet(sheet)
            if description is not None:
                source.set_value(f"C{row}", description)
            source.set_value(f"D{row}", article)
            source.set_value(f"G{row}", qty)

        app.run("ОбновитьСпецификациюТКП")
        preview = app.sheet("Спец_ТКП")
        values = [
            tuple(preview.value(f"{letter}{row}") for letter in "ABCDEF")
            for row in range(2, 14)
        ]
        group_rows = [v for v in values if v[5] == "H"]
        item_rows = [v for v in values if v[5] == "I"]

        self.assertEqual(
            [v[1] for v in group_rows],
            [
                "Программное обеспечение",
                "Шкаф 1 ПТК",
                "Шкаф 1 телекоммуникационное оборудование",
                "Шкаф 1 шкафное оборудование",
                "Шкаф 2 шкафное оборудование",
                "ЗИП",
            ],
        )
        self.assertEqual([v[0] for v in group_rows], ["1.", "2.", "3.", "4.", "5.", "6."])
        self.assertEqual([v[3] for v in group_rows], ["Компл."] * 6)
        self.assertEqual([v[4] for v in group_rows], [1.0] * 6)

        self.assertEqual(
            [v[1] for v in item_rows],
            [
                "Лицензия ПО",
                "NV-1",
                "Шкаф Эксара",
                "Модуль ввода",
                "Модуль ввода",
                "Реле ЗИП",
            ],
        )
        self.assertEqual([v[0] for v in item_rows], ["1."] * 6)
        self.assertEqual(preview.value("A1"), "№ п/п")
        self.assertEqual(preview.value("A14"), None)

        # Verify Python generator produces the exact same rows on the same input data
        wb = openpyxl.load_workbook(WORKBOOK, data_only=False)
        for sheet, row, description, article, qty in rows:
            ws = wb[sheet]
            if description is not None:
                ws[f"C{row}"] = description
            ws[f"D{row}"] = article
            ws[f"G{row}"] = qty
        py_data = read_tkp_data(wb)
        py_rows = build_preview_rows(py_data.sections)
        vba_normalized = [
            (v[0], v[1], v[2] or "", v[3], int(v[4]), v[5]) for v in values
        ]
        self.assertEqual(py_rows, vba_normalized)

        # A second run must clear old rows: the preview is regenerated, not appended.
        for sheet, row, *_ in rows:
            app.sheet(sheet).set_value(f"G{row}", 0)
        app.run("ОбновитьСпецификациюТКП")
        self.assertEqual(preview.value("A1"), "№ п/п")
        self.assertEqual(preview.value("A2"), None)
        self.assertEqual(preview.value("B13"), None)

    def _run_both(self, edits, last_row=14):
        """Apply `edits` to both engines and return their normalised preview rows.

        Each edit is (sheet, row, description, article, quantity) written into the
        C/D/G columns. Neither engine writes to disk: pyopenvba interprets the VBA
        in memory and openpyxl only mutates the loaded workbook.
        """
        app = ExcelApplication.open(WORKBOOK)
        for sheet, row, description, article, qty in edits:
            source = app.sheet(sheet)
            source.set_value(f"C{row}", description)
            source.set_value(f"D{row}", article)
            source.set_value(f"G{row}", qty)
        app.run("ОбновитьСпецификациюТКП")
        preview = app.sheet("Спец_ТКП")
        values = [
            tuple(preview.value(f"{letter}{row}") for letter in "ABCDEF")
            for row in range(2, last_row)
        ]
        vba_rows = normalise_preview_rows([v for v in values if v[1] is not None])

        wb = openpyxl.load_workbook(WORKBOOK, data_only=False)
        for sheet, row, description, article, qty in edits:
            ws = wb[sheet]
            ws[f"C{row}"] = description
            ws[f"D{row}"] = article
            ws[f"G{row}"] = qty
        py_rows = normalise_preview_rows(build_preview_rows(read_tkp_data(wb).sections))
        return vba_rows, py_rows

    def test_total_and_margin_rows_are_skipped_in_both_engines(self):
        """«Итого»/«МАРЖА» in the description column must never reach Appendix 3.

        Regression (H1): the VBA implementation inspected the caption column (B)
        only, so a total typed into column C was emitted as a normal line item.
        """
        # A total ends its block, so the ordinary positions are placed before it,
        # matching the real sheet layout (rows 3-9, then «Итого…» on row 10).
        vba_rows, py_rows = self._run_both(
            [
                ("1", 4, "Модуль ввода", "ART-1", 2),
                ("1", 5, "Обычная позиция", "ART-2", 3),
                ("1", 6, "Итого по шкафу", "ART-TOTAL", 5),
                ("1", 7, "МАРЖА по шкафу", "ART-MARGIN", 7),
            ]
        )
        self.assertEqual(vba_rows, py_rows)
        descriptions = [row[1] for row in py_rows]
        self.assertIn("Шкаф 1 ПТК", descriptions)
        self.assertIn("Модуль ввода", descriptions)
        self.assertIn("Обычная позиция", descriptions)
        self.assertNotIn("Итого по шкафу", descriptions)
        self.assertNotIn("МАРЖА по шкафу", descriptions)

    def test_python_build_from_dotx_template(self):
        with TemporaryDirectory() as tmp:
            out_path = Path(tmp) / "out.docx"
            data = generate_tkp(
                excel_source=WORKBOOK,
                output_path=out_path,
                use_demo_if_empty=True,
            )
            self.assertTrue(out_path.exists())
            self.assertEqual(len(data.sections), 7)
            self.assertEqual(data.total_items, 16)

            doc = docx.Document(str(out_path))
            self.assertGreaterEqual(len(doc.tables), 3)
            spec_tbl = doc.tables[2]
            self.assertEqual(
                [c.text.strip() for c in spec_tbl.rows[0].cells],
                [
                    "№ п/п",
                    "Наименование и техническая характеристика",
                    "Артикул",
                    "Ед. изм.",
                    "Кол-во",
                ],
            )
            self.assertEqual(
                [c.text.strip() for c in spec_tbl.rows[1].cells], ["1", "2", "3", "4", "5"]
            )
            self.assertEqual(
                [c.text.strip() for c in spec_tbl.rows[2].cells],
                ["1.", "Программное обеспечение", "", "Компл.", "1"],
            )
            self.assertTrue(all(c.paragraphs[0].runs[0].bold for c in spec_tbl.rows[2].cells))
            self.assertFalse(any(c.paragraphs[0].runs[0].bold for c in spec_tbl.rows[3].cells))


if __name__ == "__main__":
    unittest.main()
