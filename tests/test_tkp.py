"""Static packaging checks, headless VBA preview test, and Python generator tests."""

import re
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

import docx
import openpyxl
from lxml import etree
from pyopenvba import ExcelFile
from pyopenvba.apps.excel import ExcelApplication

from generate_tkp import (
    _eval_cell,
    build_preview_rows,
    clean_text,
    find_default_template,
    format_qty,
    generate_tkp,
    read_tkp_data,
)

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"
TEMPLATE = ROOT / "ТКП_шаблон.dotx"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}


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

    def test_template_has_no_latin_letters_inside_russian_words(self):
        """A Latin "e" typed instead of "ё" is invisible to the spellchecker.

        The template used to ship "Объeм", "выдаeтся", "приeмо-сдаточные",
        "надeжность", "счeт" and "трeхуровневой" — every one of them reaches the
        customer as a typo because Word does not flag a Latin letter inside a
        Russian word. Only the certificate number "RU.AЖ49.H04589" mixes scripts.
        """
        cyrillic = re.compile(r"[Ѐ-ӿ]")
        latin = re.compile(r"[A-Za-z]")
        mixed = set()
        for name in ("word/document.xml", "word/header1.xml", "word/header2.xml"):
            with ZipFile(TEMPLATE) as archive:
                root = etree.fromstring(archive.read(name))
            text = " ".join(t.text or "" for t in root.iter(f"{{{W}}}t"))
            for token in re.findall(r"[^\W\d_]+", text):
                if cyrillic.search(token) and latin.search(token):
                    mixed.add((name, token))
        self.assertEqual(mixed, {("word/document.xml", "AЖ")})

    def test_template_uses_the_proofread_wording(self):
        """The wording follows the reference ТКП: «ё», «работ», «Акта», «тайму»."""
        with ZipFile(TEMPLATE) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
        text = "".join(t.text or "" for t in root.iter(f"{{{W}}}t"))
        for broken in (
            "Оплата работы по пунктам",
            "стоимости работы по этапу",
            "Комплексного опробование",
            "третьих сторон и лиц",
            "оформлением Акт о завершении",
            "исходных данных, Спецификация",
            "Приемо-сдаточные испытания",
            "приема-передачи",
            "приемки-передачи",
            "3 (трех)",
            "собственности Заказчику",
        ):
            self.assertNotIn(broken, text)
        for fixed in (
            "Оплата работ по пунктам 5–8",
            "от стоимости работ по этапу",
            "Комплексное опробование",
            "для третьих лиц",
            "оформлением Акта о завершении ПНР",
            "При изменении исходных данных Спецификация",
            "Приёмо-сдаточные испытания",
            "Акта приёма-передачи документации",
            "Протокола приёма-передачи",
            "3 (трёх)",
            "Объём поставки",
            "выдаётся сертификат",
            "надёжность",
            "за счёт трёхуровневой",
            "собственности к Заказчику",
            "по созданию АСУ ТП осуществляется на склад Заказчика по адресу:",
        ):
            self.assertIn(fixed, text)

    def test_preview_has_only_the_standard_specification_columns(self):
        # Read the sheet by name: "xl/worksheets/sheet26.xml" would silently point
        # at an unrelated worksheet after any insert, removal or reordering.
        wb = openpyxl.load_workbook(WORKBOOK)
        self.assertIn("Спец_ТКП", wb.sheetnames)
        preview = wb["Спец_ТКП"]
        self.assertEqual(preview["A1"].value, "№ п/п")
        self.assertTrue(preview.column_dimensions["F"].hidden)

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
                "Оборудование ПТК (шкаф 1)",
                "Телекоммуникационное оборудование для АСУ ТП (шкаф 1)",
                "Шкафное оборудование ПТК (шкаф 1)",
                "Шкафное оборудование ПТК (шкаф 2)",
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
        self.assertEqual(
            normalise_preview_rows(py_rows), normalise_preview_rows(values)
        )

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
        self.assertIn("Оборудование ПТК (шкаф 1)", descriptions)
        self.assertIn("Модуль ввода", descriptions)
        self.assertIn("Обычная позиция", descriptions)
        self.assertNotIn("Итого по шкафу", descriptions)
        self.assertNotIn("МАРЖА по шкафу", descriptions)

    def test_hash_prefixed_values_survive_in_both_engines(self):
        """A value starting with "#" is data, not an Excel error.

        Regression (H2): clean_text() blanked every value beginning with "#", so
        an article such as "#A-100" was silently dropped from the specification.
        """
        vba_rows, py_rows = self._run_both(
            [
                ("1", 4, "Контроллер", "#A-100", 3),
                ("1", 5, "#1 Модуль", "ART-1", 2),
                ("1", 6, "Позиция", "ART-2", 4),
            ]
        )
        self.assertEqual(vba_rows, py_rows)
        self.assertIn("#A-100", [row[2] for row in py_rows])
        self.assertIn("#1 Модуль", [row[1] for row in py_rows])

    def test_excel_error_literals_are_still_blanked(self):
        """The error set must keep real Excel errors out of the document."""
        for error in ("#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NUM!"):
            self.assertEqual(clean_text(error), "")

    def test_format_qty_is_always_plain_decimal(self):
        """Quantities must never reach the document in scientific notation.

        Regression (M1): %g renders 1234567.5 as "1.23457e+06" and truncates to
        six significant digits.
        """
        self.assertEqual(format_qty(1), "1")
        self.assertEqual(format_qty(50), "50")
        self.assertEqual(format_qty(2.5), "2.5")
        self.assertEqual(format_qty(0.3), "0.3")
        self.assertEqual(format_qty(1234567.5), "1234567.5")
        self.assertEqual(format_qty(10000000.5), "10000000.5")
        for value in (2.5, 1 / 3, 1234567.5, 10000000.5, 0.1 + 0.2):
            self.assertNotIn("e", format_qty(value).lower())

    def test_circular_reference_resolves_to_empty(self):
        """A reference cycle degrades to a blank value instead of crashing.

        Regression (M3): _eval_cell() recursed without a guard, so two cells that
        point at each other raised an unhandled RecursionError.
        """
        wb = openpyxl.Workbook()
        wb.active.title = "Наценки-Скидки"
        first = wb.create_sheet("X")
        second = wb.create_sheet("Y")
        first["C1"] = "=Y!C1"
        second["C1"] = "=X!C1"
        self.assertEqual(_eval_cell(first, first, wb, wb, 1, 3), "")

        # A ordinary reference chain must still resolve.
        first["C1"] = "=Наценки-Скидки!C2"
        wb["Наценки-Скидки"]["C2"] = "Значение"
        self.assertEqual(_eval_cell(first, first, wb, wb, 1, 3), "Значение")

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
