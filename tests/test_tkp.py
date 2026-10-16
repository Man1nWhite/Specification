"""Static packaging checks and a headless VBA preview test (no Office needed)."""

from pathlib import Path
import unittest
from zipfile import ZipFile

from lxml import etree
from pyopenvba import ExcelFile
from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"
TEMPLATE = ROOT / "ТКП_шаблон.dotx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"w": W}
SHEET_NS = {"m": M}


class TKPTest(unittest.TestCase):
    def test_embedded_source_matches_editable_source(self):
        expected = (ROOT / "Module1.bas").read_text(encoding="utf-8").replace("\r\n", "\n").strip()
        with ExcelFile(WORKBOOK) as book:
            actual = book.get_module("Module1").replace("\r\n", "\n").strip()
            self.assertEqual(actual, expected)
            self.assertIn("Module12", book.module_names())  # existing macros preserved

    def test_bookmarks_are_in_their_destination_not_the_notes(self):
        with ZipFile(TEMPLATE) as archive:
            document = etree.fromstring(archive.read("word/document.xml"))
            header = etree.fromstring(archive.read("word/header1.xml"))
        body = document.find(".//w:body", NS)
        self.assertIsNotNone(body)
        notes = next(p for p in body.findall("w:p", NS)
                     if "Примечания к Приложению" in "".join(p.itertext()))
        target = {}
        for name in ("Num", "Customer", "Object", "Spec_Table"):
            starts = document.xpath(".//w:bookmarkStart[@w:name=$name]", namespaces=NS, name=name)
            self.assertEqual(len(starts), 1)
            self.assertIsNot(starts[0].getparent(), notes)
            target[name] = starts[0]
        self.assertLess(body.index(target["Spec_Table"].getparent()), body.index(notes))
        self.assertEqual(
            len(header.xpath(".//w:bookmarkStart[@w:name='HeaderNum']", namespaces=NS)), 1
        )
        for name in ("Num", "Customer", "Object"):
            self.assertEqual(target[name].getparent().tag, f"{{{W}}}p")
            self.assertEqual(target[name].getparent().getparent().tag, f"{{{W}}}tc")

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
            ("1", 3, None, "NV-1", 2),  # broken lookup in C: fall back to article D
            ("1", 14, "Шкаф Эксара", "EX-1", 4),
            ("1", 20, "Модуль ввода", "X-123", 2),
            ("2", 20, "Модуль ввода", "X-123", 3),  # same item: sum 2 + 3
            ("1", 21, "Не включать", "X-0", 0),
            ("Spare parts", 25, "Реле ЗИП", "ZIP-1", 1),
            ("SW(w.o.VAT)", 14, "Лицензия ПО", "MasterPLC.N1000", 2),
        ]
        for sheet, row, description, article, qty in rows:
            source = app.sheet(sheet)
            if description is not None:
                source.set_value(f"C{row}", description)
            source.set_value(f"D{row}", article)
            source.set_value(f"G{row}", qty)

        app.run("ОбновитьСпецификациюТКП")
        preview = app.sheet("Спец_ТКП")
        values = [tuple(preview.value(f"{letter}{row}") for letter in "ABCDEF")
                  for row in range(2, 12)]
        group_rows = [v for v in values if v[5] == "H"]
        item_rows = [v for v in values if v[5] == "I"]
        self.assertEqual([v[1] for v in group_rows], [
            "Оборудование НВТ", "Оборудование Эксара",
            "Оборудование 3-х поставщиков", "ЗИП 3-х поставщиков", "ПО ИнСАТ",
        ])
        self.assertEqual([v[1] for v in item_rows], [
            "NV-1", "Шкаф Эксара", "Модуль ввода", "Реле ЗИП", "Лицензия ПО",
        ])
        self.assertEqual([v[0] for v in group_rows], ["1.", "2.", "3.", "4.", "5."])
        self.assertEqual([v[0] for v in item_rows], ["1."] * 5)
        self.assertEqual(next(v for v in item_rows if v[2] == "X-123")[4], 5)
        self.assertEqual(preview.value("A1"), "№ п/п")
        self.assertEqual(preview.value("A12"), None)

        # A second run must clear old rows: the preview is regenerated, not appended.
        for sheet, row, *_ in rows:
            app.sheet(sheet).set_value(f"G{row}", 0)
        app.run("ОбновитьСпецификациюТКП")
        self.assertEqual(preview.value("A1"), "№ п/п")
        self.assertEqual(preview.value("A2"), None)
        self.assertEqual(preview.value("B11"), None)


if __name__ == "__main__":
    unittest.main()
