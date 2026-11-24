"""Strukturtest og scenarier med genberegning i LibreOffice Calc."""

import hashlib
import os
import shutil
import subprocess
import tempfile
import unittest
from datetime import date
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula import Tokenizer
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from migrate_3_3 import TYPE_COLORS, migrate

ROOT = Path(__file__).resolve().parents[1]
BASELINE = ROOT / "Ressourceplan_3_2.xlsx"
RELEASE = ROOT / "Ressourceplan_3_3.xlsx"
CALC = shutil.which("libreoffice") or shutil.which("soffice")


class StructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = load_workbook(BASELINE)
        cls.result = load_workbook(RELEASE)

    def test_release_structure_and_formulas(self):
        with ZipFile(RELEASE) as package:
            self.assertIsNone(package.testzip())
            for entry in package.namelist():
                self.assertNotRegex(entry.lower(), r"vba|externallinks|connections")
                if entry.endswith((".xml", ".rels")):
                    root = ElementTree.fromstring(package.read(entry))
                    self.assertFalse(
                        any(node.get("TargetMode") == "External" for node in root.iter())
                    )
        self.assertEqual(self.result.sheetnames, [
            "Min ressourceplan", "Ressourceoverblik", "Projektregister",
            "Projektoversigt", "Medarbejdere",
        ])
        self.assertEqual(self.result.active.title, self.result.sheetnames[0])
        self.assertTrue(self.result.calculation.fullCalcOnLoad)
        self.assertEqual(self.result.calculation.calcMode, "auto")
        for sheet in self.result:
            for row in sheet:
                for cell in row:
                    self.assertNotEqual(cell.data_type, "e", f"{sheet.title}!{cell.coordinate}")
                    if cell.data_type == "f":
                        self.assertNotIn("#REF!", cell.value)
                        self.assertTrue(Tokenizer(cell.value).items)
            for table in sheet.tables.values():
                self.assertEqual(table.autoFilter.ref, table.ref)
                first, row, last, _ = range_boundaries(table.ref)
                headers = [sheet.cell(row, column).value for column in range(first, last + 1)]
                self.assertEqual(headers, [c.name for c in table.tableColumns])
                self.assertEqual(len(headers), len(set(headers)))
                self.assertTrue(all(column.showButton is not False and column.hiddenButton is not True
                                    for column in table.autoFilter.filterColumn))

    def test_original_and_input_data(self):
        self.assertEqual(hashlib.sha256(BASELINE.read_bytes()).hexdigest(),
                         "7eb4f9e1886c98c754a1f6c48f59f78bce1572fa5a4666ed3cf84d6ca82c9959")
        plan, updated = self.source.worksheets[0], self.result.worksheets[0]
        offers_preserved = offers_removed = 0
        keys = set()
        for row in range(10, 2170):
            for column in (1, 3, 5, *range(8, 62)):
                self.assertEqual(plan.cell(row, column).value, updated.cell(row, column).value)
            keys.add((updated[f"C{row}"].value, updated[f"A{row}"].value))
            before, after = plan[f"D{row}"].value, updated[f"D{row}"].value
            if before == "Tilbud":
                if row == 17:
                    self.assertEqual(after, "Tilbud")
                    offers_preserved += 1
                else:
                    self.assertIsNone(plan[f"E{row}"].value)
                    self.assertEqual(plan[f"H{row}"].value, 0.5)
                    self.assertTrue(all(cell.value is None for cell in plan[row][8:61]))
                    self.assertEqual(after, "Sag")
                    offers_removed += 1
            else:
                self.assertEqual(before, after)
        self.assertEqual((offers_preserved, offers_removed), (1, 179))
        self.assertEqual(len(keys), 30 * 6)
        for name in ("Projektregister", "Medarbejdere"):
            old, new = self.source[name], self.result[name]
            for row in old.iter_rows(min_row=6, max_col=6):
                for cell in row:
                    self.assertEqual(new[cell.coordinate].value, cell.value)

    def test_view_validations_and_dynamic_styles(self):
        plan, overview, _, report, _ = self.result.worksheets
        self.assertTrue(plan.column_dimensions["C"].hidden)
        self.assertEqual(plan.freeze_panes, "I10")
        self.assertEqual(plan.tables["PlanTimer"].ref, "A9:BI2169")
        filters = plan.tables["PlanTimer"].autoFilter.filterColumn
        self.assertEqual([(f.colId, f.filters.filter) for f in filters],
                         [(0, ["2026"]), (1, ["Medarbejder 1"])])
        self.assertEqual([r for r in range(10, 2170) if not plan.row_dimensions[r].hidden],
                         list(range(10, 22)))
        self.assertTrue(all(overview.row_dimensions[r].hidden for r in range(41, 250)))
        self.assertFalse(overview.row_dimensions[40].hidden)
        self.assertTrue(all(report.row_dimensions[r].hidden for r in range(41, 75)))
        by_range = {str(rule.sqref): rule for rule in plan.data_validations.dataValidation}
        self.assertEqual(by_range["E10:E2169"].formula1, "Sagsnumre")
        self.assertEqual((by_range["I10:BH2169"].formula1, by_range["I10:BH2169"].formula2),
                         ("0", "80"))
        self.assertIn("WEEKDAY", by_range["BI10:BI2169"].formula1)
        self.assertEqual(by_range["H10:H2169"].formula2, "1")
        self.assertTrue(all(rule.showErrorMessage for rule in by_range.values()))
        for addr in ("I5", "K5", "M5"):
            self.assertTrue(any(addr in rule for rule in overview.data_validations.dataValidation))
        self.assertEqual(report["H5"].fill.fgColor.rgb, "FFF04E23")
        for sheet, rng in ((plan, "I7:BI7"), (overview, "H7:BH7"), (report, "H7:BH7")):
            self.assertEqual(len(sheet.conditional_formatting[rng]), 4)
        styles = plan.conditional_formatting["A10:BI2169"]
        for activity, color in TYPE_COLORS.items():
            rule = next(rule for rule in styles if rule.formula == [f'$D10="{activity}"'])
            self.assertEqual(rule.dxf.fill.fgColor.rgb, "FF" + color)
        self.assertTrue(any('MOD(ROW(),2)' in rule.formula[0] for rule in styles))
        percentage = plan.conditional_formatting["H10:H2169"][0]
        self.assertEqual(percentage.formula, ['$D10<>"Tilbud"'])
        self.assertEqual(percentage.dxf.numFmt.formatCode, ";;;")

    def test_rebuild_and_refuse_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / RELEASE.name
            migrate(BASELINE, output)
            with self.assertRaises(ValueError):
                migrate(BASELINE, output)
            with self.assertRaises(ValueError):
                migrate(BASELINE, BASELINE)
            # ZIP timestamps and document modification time may differ.
            with ZipFile(output) as generated, ZipFile(RELEASE) as published:
                self.assertEqual(generated.namelist(), published.namelist())
                for entry in generated.namelist():
                    if entry == "docProps/core.xml":
                        continue
                    actual, expected = generated.read(entry), published.read(entry)
                    if entry.endswith((".xml", ".rels")):
                        actual = ElementTree.canonicalize(actual)
                        expected = ElementTree.canonicalize(expected)
                    self.assertEqual(actual, expected, entry)


@unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
class CalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        tmp = tempfile.TemporaryDirectory(prefix="ressourceplan-test-")
        cls.addClassCleanup(tmp.cleanup)
        directory = Path(tmp.name)
        output = directory / "beregnet"
        output.mkdir()

        def scenario(name, change=None):
            path = directory / f"{name}.xlsx"
            if change:
                workbook = load_workbook(RELEASE)
                change(workbook)
                workbook.save(path)
            else:
                shutil.copyfile(RELEASE, path)
            return path

        def additional_hours(workbook):
            plan, overview, projects, report, staff = workbook.worksheets
            overview["F5"] = "4 uger"
            overview["I5"], overview["K5"], overview["M5"] = "Øst", "EL", "7134.03"
            staff["B7"] = "Testnavn"
            staff["B12"], staff["C12"], staff["D12"], staff["E12"], staff["F12"] = (
                "Testperson 6", "Øst", "EL", "7134.03", 37
            )
            projects["B7"] = "Omdøbt projekt"
            plan["BD19"], plan["BD20"], plan["BD21"] = 13, 17, 5
            plan["H17"] = None
            plan["E17"] = "24015"
            plan["D13"], plan["H13"], plan["BD13"] = "Tilbud", 0, 8
            new_employee_row = 10 + 5 * 72
            plan[f"E{new_employee_row}"], plan[f"BD{new_employee_row}"] = "24015", 7

        def convert_offer(workbook):
            additional_hours(workbook)
            plan = workbook.worksheets[0]
            project, hours = plan["E17"].value, plan["BD17"].value
            plan["D17"], plan["E17"], plan["BD17"] = "Sag", None, None
            destination = 11 + 5 * 72
            plan[f"E{destination}"], plan[f"BD{destination}"] = project, hours
            plan[f"H{destination}"] = 0.25  # A stale percentage on an ordinary project.

        def iso_year(year, period):
            def configure(workbook):
                plan, overview, _, report, _ = workbook.worksheets
                filters = plan.tables["PlanTimer"].autoFilter.filterColumn
                filters[0].filters.filter = [str(year)]
                for row in range(10, 2170):
                    plan.row_dimensions[row].hidden = not (
                        plan[f"A{row}"].value == year and plan[f"C{row}"].value == "p01"
                    )
                for target in (overview, report):
                    target["B5"], target["D5"], target["F5"] = year, 53, period
                for row, hours in ((10, 3), (22, 65), (34, 67)):
                    plan[f"BI{row}"] = hours
                plan["I10"], plan["I22"], plan["I34"] = 5, 11, 12
            return configure

        def mixed_years(workbook):
            iso_year(2026, "Q4")(workbook)
            plan = workbook.worksheets[0]
            plan.tables["PlanTimer"].autoFilter.filterColumn.pop(0)
            for row in range(10, 82):
                plan.row_dimensions[row].hidden = False

        inputs = [
            scenario("baseline"), scenario("aktivitet", additional_hours),
            scenario("vundet", convert_offer), scenario("uge53", iso_year(2026, "Q4")),
            scenario("uge52", iso_year(2027, "Hele året")),
            scenario("aar2028", iso_year(2028, "4 uger")), scenario("flere-aar", mixed_years),
        ]
        result = subprocess.run(
            [CALC, f"-env:UserInstallation={directory.joinpath('profile').as_uri()}",
             "--headless", "--convert-to", "xlsx", "--outdir", str(output), *map(str, inputs)],
            capture_output=True, text=True, check=True, timeout=180,
            env=dict(os.environ, XDG_CACHE_HOME=str(directory / "cache")),
        )
        cls.books = {}
        for source in inputs:
            cached = output / source.name
            if not cached.is_file():
                raise AssertionError(f"Calc genererede ikke {source.name}\n{result.stdout}\n{result.stderr}")
            cls.books[source.stem] = load_workbook(cached, data_only=True)

    def test_no_formula_errors_in_all_scenarios(self):
        for scenario, workbook in self.books.items():
            with self.subTest(scenario=scenario):
                errors = [(sheet.title, cell.coordinate, cell.value) for sheet in workbook
                          for row in sheet for cell in row if cell.data_type == "e"]
                self.assertEqual(errors, [])

    def test_reference_calculation_and_filter_safe_totals(self):
        plan, overview, _, report, _ = self.books["baseline"].worksheets
        self.assertEqual(plan["G6"].value, 37)
        self.assertEqual(plan["BD6"].value, 37)
        self.assertEqual(plan["H8"].value, 2026)
        for address, value in {"BC45": 20, "BC80": 30, "BC115": 10, "BC150": 5,
                               "BC185": 10, "BC220": 0, "B2": 20, "J2": 10, "L2": 5}.items():
            self.assertEqual(overview[address].value, value, address)
        self.assertAlmostEqual(overview["BC10"].value, 2 / 3)
        self.assertEqual(report["BC40"].value, 20)

    def test_capacity_forecast_activation_and_project_name(self):
        plan, overview, _, report, _ = self.books["aktivitet"].worksheets
        self.assertEqual(plan["F10"].value, "Omdøbt projekt")
        self.assertEqual(plan["B10"].value, "Testnavn")
        self.assertEqual(overview["A15"].value, "Testperson 6")
        self.assertIsNone(overview["A11"].value)
        for address, value in {"BC45": 25, "BC80": 0, "BC115": 18, "BC150": 5, "BC10": "Over",
                               "BC190": 30, "BC220": 25, "F2": 252, "H2": 25}.items():
            self.assertEqual(overview[address].value, value, address)
        self.assertEqual(report["H6"].value, "Omdøbt projekt")
        self.assertEqual(report["BC10"].value, 30)
        self.assertEqual(report["BC15"].value, 7)
        self.assertEqual(report["F40"].value, 37)

    def test_won_offer_transfer_disregards_stored_probability(self):
        plan, overview, _, report, _ = self.books["vundet"].worksheets
        self.assertEqual(plan["H371"].value, 0.25)
        self.assertEqual(plan["BJ371"].value, 0)
        self.assertEqual(report["F40"].value, 37)
        self.assertEqual(report["BC10"].value, 20)
        self.assertEqual(report["BC15"].value, 17)
        self.assertEqual(overview["B2"].value, 42)
        self.assertEqual(overview["J2"].value, 8)
        self.assertEqual(overview["L2"].value, 0)

    def test_iso_weeks_quarters_and_multiple_years(self):
        for name, year in (("uge53", 2026), ("uge52", 2027), ("aar2028", 2028)):
            with self.subTest(year=year):
                plan, overview, _, report, _ = self.books[name].worksheets
                valid = date(year, 12, 28).isocalendar().week
                for week in range(1, valid + 1):
                    monday = date.fromisocalendar(year, week, 1)
                    quarter = f"Q{(monday.month - 1) // 3 + 1}"
                    for sheet, first, quarter_row in ((plan, 9, 7), (overview, 8, 7), (report, 8, 7)):
                        col = get_column_letter(first + week - 1)
                        self.assertEqual(sheet[f"{col}8"].value.date(), monday)
                        self.assertEqual(sheet[f"{col}{quarter_row}"].value, quarter)
                self.assertEqual(overview["D6"].value, valid)
                if valid == 53:
                    self.assertEqual(overview["H43"].value, 1)  # The Monday is in December 2025.
                    self.assertEqual(overview["I43"].value, 0)
                    self.assertEqual(overview["BH43"].value, 1)
                    self.assertEqual(plan["BI6"].value, 3)
                    self.assertEqual(report["BH40"].value, 3)
                    self.assertEqual(report["F40"].value, 28)
                else:
                    self.assertEqual(plan["BI6"].value, 0)
                    for sheet, col in ((plan, "BI"), (overview, "BH"), (report, "BH")):
                        self.assertIsNone(sheet[f"{col}7"].value)
                        self.assertIsNone(sheet[f"{col}8"].value)
                    row = 22 if year == 2027 else 34
                    hours = 11 if year == 2027 else 12
                    self.assertEqual(plan[f"G{row}"].value, hours)
                    self.assertEqual(report["G40"].value, hours)
                    self.assertEqual(report["F40"].value, hours if year == 2027 else 0)
        plan = self.books["flere-aar"].worksheets[0]
        self.assertEqual(plan["H8"].value, 0)
        self.assertIsNone(plan["I7"].value)
        self.assertIsNone(plan["I8"].value)
        self.assertEqual(plan["BI6"].value, 3)
        self.assertEqual(plan["G6"].value, 68)


if __name__ == "__main__":
    unittest.main()
