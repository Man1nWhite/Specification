"""Differential coverage for the VBA and Python specification engines."""

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import docx
import openpyxl
from pyopenvba.apps.excel import ExcelApplication

from generate_tkp import (
    WorkbookError,
    build_preview_rows,
    demo_cells,
    format_qty,
    generate_tkp,
    populate_demo_workbook,
    read_tkp_data,
)

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"


def rows_in_both(cells):
    app = ExcelApplication.open(WORKBOOK)
    wb = openpyxl.load_workbook(WORKBOOK, data_only=False)
    for sheet, ref, value in cells:
        if isinstance(value, str):
            app.sheet(sheet).sheet.vba_get("Range", [ref]).vba_set("NumberFormat", "@")
        app.sheet(sheet).set_value(ref, value)
        wb[sheet][ref] = value
        if isinstance(value, str) and value.startswith("="):
            wb[sheet][ref].data_type = "s"
    app.run("ОбновитьСпецификациюТКП")
    preview = app.sheet("Спец_ТКП")
    vba = []
    for row in range(2, 1000):
        values = tuple(preview.value(f"{letter}{row}") for letter in "ABCDEF")
        if values[0] is None:
            break
        vba.append(values)
    data = read_tkp_data(wb)
    python = build_preview_rows(data.sections)

    def normalise(rows):
        return [(a, b, c or "", d, q, mark) for a, b, c, d, q, mark in rows]

    return normalise(vba), normalise(python), data


class EngineTests(unittest.TestCase):
    def test_complete_demo(self):
        cells = [cell for cell in demo_cells() if cell[0] != "Наценки-Скидки"]
        vba, python, data = rows_in_both(cells)
        self.assertEqual(data.issues, [])
        self.assertEqual(data.total_items, 16)
        self.assertEqual(vba, python)

    def test_caption_order_repeats_and_rows_below_48(self):
        cells = [
            ("1", "B1", "Спецификация оборудования 3-х поставщиков"),
            ("1", "B18", "Спецификация оборудования НВТ"),
            ("1", "B24", "Спецификация оборудования НВТ"),
            ("1", "B58", "Спецификация оборудования Эксара"),
            ("1", "C4", "First"), ("1", "D4", "A"), ("1", "G4", 1),
            ("1", "C20", "Second"), ("1", "D20", "B"), ("1", "G20", 2),
            ("1", "C25", "Third"), ("1", "D25", "C"), ("1", "G25", 3),
            ("1", "C59", "Fourth"), ("1", "D59", "D"), ("1", "G59", 4),
        ]
        vba, python, _ = rows_in_both(cells)
        self.assertEqual(vba, python)
        self.assertEqual([r[1] for r in python if r[5] == "H"], [
            "Шкафное оборудование ПТК (шкаф 1)", "Оборудование ПТК (шкаф 1)",
            "Оборудование ПТК (шкаф 1)", "Телекоммуникационное оборудование для АСУ ТП (шкаф 1)",
        ])
        self.assertEqual([r[1] for r in python if r[5] == "I"], ["First", "Second", "Third", "Fourth"])

    def test_text_quantities_and_articles(self):
        cells = []
        values = ("3,5", "3.5", " 4 ", "1_000", "inf", "1e3", "2,5,1", "１２", ".5")
        articles = ("0441504", "1-2", "=article")
        for row, value in zip(range(20, 29), values, strict=True):
            article = articles[(row - 20) % len(articles)]
            cells.extend([("1", f"C{row}", f"Item {row}"), ("1", f"D{row}", article),
                          ("1", f"G{row}", value)])
        vba, python, _ = rows_in_both(cells)
        self.assertEqual(vba, python)
        self.assertEqual([r[4] for r in python if r[5] == "I"], [3.5, 3.5, 4, 0.5])
        self.assertEqual([r[2] for r in python if r[5] == "I"],
                         ["0441504", "1-2", "=article", "=article"])

    def test_unknown_caption_clears_vba_preview_and_rejects_python(self):
        app = ExcelApplication.open(WORKBOOK)
        wb = openpyxl.load_workbook(WORKBOOK)
        cells = [("1", "C4", "first"), ("1", "D4", "A"), ("1", "G4", 2),
                 ("1", "B18", "Спецификация оборудования 3-х поставшиков")]
        for sheet, ref, value in cells:
            app.sheet(sheet).set_value(ref, value)
            wb[sheet][ref] = value
        app.run("ОбновитьСпецификациюТКП")
        self.assertIsNone(app.sheet("Спец_ТКП").value("A2"))
        self.assertIn("поставшиков", app.dialogs[-1][1])
        with self.assertRaisesRegex(WorkbookError, "поставшиков"):
            read_tkp_data(wb)

    def test_quantity_formatting_in_both_engines(self):
        app = ExcelApplication.open(WORKBOOK)
        for value in (1, 2.5, 0.3, 1 / 3, 1.0000005, 1.0000015, 1234567.5, 10000000.5):
            self.assertEqual(app.run("ТекстКоличества", value), format_qty(value))

    def test_word_quantity_uses_decimal_comma(self):
        wb = openpyxl.load_workbook(WORKBOOK)
        populate_demo_workbook(wb)
        wb["1"]["G4"] = 2.5
        with TemporaryDirectory() as tmp:
            output = Path(tmp) / "result.docx"
            generate_tkp(wb, output_path=output)
            table = docx.Document(output).tables[2]
            row = next(r for r in table.rows if r.cells[2].text.strip() == "АРКС400-ДВ-32")
            self.assertEqual(row.cells[4].text.strip(), "2,5")
