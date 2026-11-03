"""Regression checks for the workbook's cost links, using isolated test inputs.

These tests don't certify commercial prices or VAT; an estimator must recalculate
and review a real populated workbook in Excel before it is used.
"""

import unittest
import warnings
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile

import openpyxl
from pyopenvba.apps.excel import ExcelApplication

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "Gen_TKP.xlsm"


def load_formulas():
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        return openpyxl.load_workbook(WORKBOOK, data_only=False)


class WorkbookCostTests(unittest.TestCase):
    def test_work_costs_align_with_all_nine_input_rows(self):
        wb = load_formulas()
        ws = wb["Сводная спецификация"]
        for index in range(9):
            self.assertEqual(ws[f"J{55 + index * 4}"].value, f"=G{21 + index}")
            group = 52 + index * 4
            self.assertEqual(ws[f"I{group}"].value, f"=I{21 + index}")
            self.assertEqual(ws[f"J{group}"].value, f"=SUM(J{group + 1}:J{group + 3})")
        self.assertEqual(ws["L52"].value, "=SUM(L53:L55)")
        self.assertIn("J68", ws["J51"].value)
        self.assertIn("J51", ws["J88"].value)

    def test_discounted_cabinet_and_license_references(self):
        wb = load_formulas()
        self.assertEqual(wb["3"]["H55"].value, "=K48")
        license_sheet = wb["Конфигуратор лицензии"]
        self.assertEqual(license_sheet["G4"].value, "=G12*(1-$F$3)")
        for row in range(5, 11):
            self.assertEqual(license_sheet[f"G{row}"].value, f"=G$4*(1-$F{row})")

    def test_workbook_requests_full_recalculation(self):
        wb = load_formulas()
        self.assertEqual(wb.calculation.calcMode, "auto")
        self.assertTrue(wb.calculation.fullCalcOnLoad)
        self.assertTrue(wb.calculation.forceFullCalc)
        with ZipFile(WORKBOOK) as z:
            self.assertNotIn("xl/calcChain.xml", z.namelist())
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
            cached = openpyxl.load_workbook(WORKBOOK, data_only=True)
        # Old total is not presented to a non-recalculating reader as a current price.
        self.assertIsNone(cached["Сводная спецификация"]["J88"].value)
        self.assertIsNone(cached["3"]["H55"].value)

    def test_recalculate_cost_scenarios_with_actual_formulas(self):
        source = load_formulas()
        wb = openpyxl.Workbook()
        summary = wb.active
        summary.title = "Сводная спецификация"
        pricing = wb.create_sheet("Наценки-Скидки")
        pricing["F7"] = 0.35
        pricing["I9"] = 0
        costs = (100, 200, 300, 400, 300000, 600, 700, 800, 900)
        for index, cost in enumerate(costs):
            row = 21 + index
            summary[f"G{row}"] = cost
            summary[f"J{row}"] = 0
            summary[f"M{row}"] = 0
        for row in range(51, 89):
            for col in "JKLMN":
                addr = f"{col}{row}"
                if source[summary.title][addr].data_type == "f":
                    summary[addr] = source[summary.title][addr].value
        for addr in ("J37", "J47", "K37", "K47", "L37", "L47", "M37", "M47", "N37", "N47"):
            summary[addr] = 0
        cab = wb.create_sheet("3")
        cab["I48"], cab["K48"], cab["I55"] = 1000, 800, 0.22
        for addr in ("H55", "J55"):
            cab[addr] = source["3"][addr].value
        lic = wb.create_sheet("Конфигуратор лицензии")
        lic["F3"], lic["G12"] = 0.1, 100
        for row in range(5, 11):
            lic[f"F{row}"] = 0.2
        for row in range(4, 11):
            lic[f"G{row}"] = source[lic.title][f"G{row}"].value
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "calculation.xlsx"
            wb.save(path)
            app = ExcelApplication.open(path)
            result = app.sheet("Сводная спецификация")
            # openpyxl writes empty caches; changing inputs marks dependent formulas
            # dirty in pyopenvba, just as editing inputs does in Excel.
            for index, cost in enumerate(costs):
                result.set_value(f"G{21 + index}", cost)
                result.set_value(f"J{21 + index}", 0)
                result.set_value(f"M{21 + index}", 0)
            for addr in ("J37", "J47", "K37", "K47", "L37", "L47", "M37", "M47", "N37", "N47"):
                result.set_value(addr, 0)
            app.sheet("Наценки-Скидки").set_value("F7", 0.35)
            app.sheet("Наценки-Скидки").set_value("I9", 0)
            app.sheet("3").set_value("K48", 800)
            app.sheet("3").set_value("I55", 0.22)
            app.sheet(lic.title).set_value("F3", 0.1)
            app.sheet(lic.title).set_value("G12", 100)
            for row in range(5, 11):
                app.sheet(lic.title).set_value(f"F{row}", 0.2)
            app.application.Calculate()
            for index, cost in enumerate(costs):
                self.assertEqual(result.value(f"J{55 + 4 * index}"), cost)
                self.assertEqual(result.value(f"J{52 + 4 * index}"), cost)
            self.assertEqual(result.value("J51"), sum(costs))
            self.assertEqual(result.value("J88"), sum(costs))
            self.assertAlmostEqual(result.value("K68"), 461538.46, places=2)
            self.assertAlmostEqual(result.value("L52"), 53.85, places=2)
            self.assertEqual(app.sheet("3").value("H55"), 800)
            self.assertAlmostEqual(app.sheet("3").value("J55"), 976, places=2)
            self.assertAlmostEqual(app.sheet(lic.title).value("G4"), 90, places=2)
            for row in range(5, 11):
                self.assertAlmostEqual(app.sheet(lic.title).value(f"G{row}"), 72, places=2)
