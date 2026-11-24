"""Kontrol af rækkevækst, nøgleopslag og Excel-pakken i Ressourceplan 3.5."""

import os
import subprocess
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook

from migrate_3_5 import FIELDS, PLAN_STYLE_LAST, keep_record, migrate
from test_ressourceplan import CALC, ROOT
from test_ressourceplan_3_3_1 import NS, digest
from test_ressourceplan_3_4 import RELEASE as STABLE

RELEASE = ROOT / "Ressourceplan_3_5.xlsx"
WEEK48 = 9 + 48


class DynamicPlanStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = load_workbook(STABLE)
        cls.release = load_workbook(RELEASE)

    def test_preserve_recorded_data_and_stable_reference(self):
        self.assertEqual(digest(STABLE), "f1eb3b88d34730cf08151ea438a6c6e57860a18149130b73c72876b47be798af")
        old, plan = self.source.worksheets[0], self.release.worksheets[0]
        meaningful = []
        for row, values in enumerate(old.iter_rows(min_row=10, max_col=61, values_only=True), 10):
            year, _, id_, kind, key, _, _, percent, *hours = values
            if keep_record(row, values):
                meaningful.append((year, id_, kind, key, percent, hours))
            else:
                self.assertIsNone(key)
                self.assertTrue(all(h is None for h in hours))
                self.assertIn(percent, (None, 0.5))
        self.assertEqual(len(meaningful), 32)
        migrated = list(plan.iter_rows(min_row=10, max_col=62, values_only=True))
        self.assertEqual(len(migrated), len(meaningful))
        for expected, values in zip(meaningful, migrated):
            year, _, _, kind, project, _, _, percent, employee, *hours = values
            project_code = project.rsplit(" · ", 1)[1] if project else None
            self.assertEqual((year, employee.rsplit(" · ", 1)[1], kind, project_code, percent, hours), expected)
        for title, rows, columns in (("Medarbejdere", 36, 6), ("Projektregister", 106, 2)):
            for row in self.source[title].iter_rows(min_row=7, max_row=rows, max_col=columns):
                for cell in row:
                    self.assertEqual(self.release[title][cell.coordinate].value, cell.value)

    def test_table_formulas_filters_validation_and_ooxml(self):
        plan = self.release["Min ressourceplan"]
        self.assertEqual(self.release.sheetnames, self.source.sheetnames)
        table = plan.tables["PlanTimer"]
        self.assertEqual(table.ref, table.autoFilter.ref)
        self.assertEqual([c.name for c in table.tableColumns], FIELDS)
        self.assertEqual([c.value for c in plan[9][:65]], FIELDS)
        self.assertEqual(plan.freeze_panes, "J10")
        self.assertEqual([(f.colId, f.filters.filter) for f in table.autoFilter.filterColumn],
                         [(0, ["2026"]), (1, ["Medarbejder 1"])])
        self.assertEqual([c.index for c in plan.row_dimensions.values()
                          if c.index >= 10 and not c.hidden], list(range(10, 15)))
        self.assertTrue(all(plan.column_dimensions[c].hidden for c in ("C", "BK", "BL", "BM")))
        self.assertFalse(plan.column_dimensions["BJ"].hidden)
        for address, target, kind in (
            ("A", '"2026,2027,2028,2029,2030,2031"', "list"),
            ("D", '"Sag,Tilbud,Ferie,Fravær,Kursus,Andet"', "list"),
            ("E", "Projektvisninger", "list"), ("I", "Medarbejdervisninger", "list"),
            ("J", "0", "decimal"), ("H", "0", "decimal"),
        ):
            rule = next(d for d in plan.data_validations.dataValidation if f"{address}{PLAN_STYLE_LAST}" in d)
            self.assertEqual((rule.formula1, rule.type), (target, kind))
            self.assertTrue(rule.showErrorMessage)
        self.assertIn(f"A10:BJ{PLAN_STYLE_LAST}", {str(r.sqref) for r in plan.conditional_formatting})
        self.assertEqual(self.release["Medarbejdere"].tables["PersonRegister"].ref, "A6:F106")
        self.assertEqual(self.release["Projektregister"].tables["SagRegister"].ref, "A6:B506")
        for name in ("Projektvisninger", "Medarbejdervisninger"):
            self.assertIn("INDEX(", self.release.defined_names[name].attr_text)
        with ZipFile(RELEASE) as archive:
            self.assertIsNone(archive.testzip())
            self.assertFalse(any(part in name.lower() for name in archive.namelist()
                                 for part in ("vba", "externallinks", "connections")))
            for name in archive.namelist():
                if name.endswith((".xml", ".rels")):
                    root = ET.fromstring(archive.read(name))
                    self.assertFalse(any(node.get("TargetMode") == "External" for node in root.iter()))
                    for formula in root.findall(".//s:f", NS):
                        self.assertNotIn("#REF!", formula.text)
                        self.assertLessEqual(len(formula.text), 8192)
            tbl = ET.fromstring(archive.read("xl/tables/table1.xml"))
            calculated = tbl.findall("s:tableColumns/s:tableColumn/s:calculatedColumnFormula", NS)
            self.assertEqual(len(calculated), 7)
            self.assertTrue(all("PlanTimer[[#This Row]" in f.text for f in calculated))
        for name in ("Ressourceoverblik", "Projektoversigt"):
            self.assertTrue(self.release[name].row_dimensions[250].hidden)
        self.assertIn("PlanTimer[ID]", self.release["Ressourceoverblik"]["BC370"].value)
        self.assertIn("PlanTimer[Sagsnr.]", self.release["Projektoversigt"]["G250"].value)

    def test_reproducible_generation_and_no_overwrite(self):
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary) / RELEASE.name
            migrate(STABLE, output)
            with self.assertRaises(ValueError):
                migrate(STABLE, output)
            with ZipFile(RELEASE) as reference, ZipFile(output) as rebuilt:
                self.assertEqual(reference.namelist(), rebuilt.namelist())
                for name in reference.namelist():
                    if name == "docProps/core.xml":
                        continue
                    left, right = reference.read(name), rebuilt.read(name)
                    if name.endswith((".xml", ".rels")):
                        left, right = ET.canonicalize(left), ET.canonicalize(right)
                    self.assertEqual(left, right, name)


def append_record(plan, year, employee, activity, project, hours, week=48):
    table = plan.tables["PlanTimer"]
    last = int(table.ref.split("BM")[1])
    row = last + 1
    for column in table.tableColumns:
        model = column.calculatedColumnFormula
        if model is not None:
            plan.cell(row, column.id, "=" + model.attr_text)
    for number, value in ((1, year), (4, activity), (5, project), (9, employee), (9 + week, hours)):
        plan.cell(row, number).value = value
    table.ref = table.autoFilter.ref = f"A9:BM{row}"
    return row


def emulate_filter(plan):
    """Set visibility for the saved p01 / 2026 native filter in the test copy."""
    for row in range(10, plan.max_row + 1):
        plan.row_dimensions[row].hidden = (
            plan[f"A{row}"].value != 2026 or plan[f"I{row}"].value != "Medarbejder 1 · p01"
        )


def sort_table(plan):
    rows = list(plan.iter_rows(min_row=10, max_col=65, values_only=True))
    rows.sort(key=lambda cells: (cells[8] or "", cells[0] or 0, cells[4] or ""), reverse=True)
    for row, data in enumerate(rows, 10):
        for column, value in enumerate(data, 1):
            plan.cell(row, column).value = value


@unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
class DynamicPlanCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-35-")
        cls.addClassCleanup(temporary.cleanup)
        root = Path(temporary.name)
        output = root / "beregnet"
        output.mkdir()
        inputs = []
        for scenario in ("reference", "vokset", "sorteret", "andet-projekt", "flytning"):
            workbook = load_workbook(RELEASE)
            if scenario != "reference":
                cls.grow(workbook, scenario)
            path = root / f"{scenario}.xlsx"
            workbook.save(path)
            inputs.append(path)
        process = subprocess.run(
            [CALC, f"-env:UserInstallation={(root / 'profile').as_uri()}",
             "--headless", "--convert-to", "xlsx", "--outdir", str(output), *map(str, inputs)],
            capture_output=True, text=True, check=True, timeout=300,
            env=dict(os.environ, XDG_CACHE_HOME=str(root / "cache")),
        )
        cls.books = {}
        for path in inputs:
            result = output / path.name
            if not result.is_file():
                raise AssertionError(process.stdout + process.stderr)
            cls.books[path.stem] = load_workbook(result, data_only=True)

    @staticmethod
    def grow(book, scenario):
        plan, overview, register, report, staff = book.worksheets
        for col, value in zip("ABCDEF", ("p31", "Oprettet person", "Vest", "EL", "QA35", 30)):
            staff[f"{col}37"] = value
        for index in range(25):
            key = {0: "00123", 1: "123", 2: "QA · 123"}.get(index, f"QA{index+1:03}")
            name = "Hospital Nord" if index < 2 else f"Testprojekt {index+1}"
            row = 108 + index
            register[f"A{row}"], register[f"B{row}"] = key, name
            append_record(plan, 2026, "Oprettet person · p31", "Tilbud" if index == 0 else "Sag",
                          f"{name} · {key}", index + 1)
        append_record(plan, 2026, "Medarbejder 1 · p01", "Sag", "Hospital Nord · 24015", 3)
        append_record(plan, 2027, "Medarbejder 1 · p01", "Sag", "Hospital Nord · 24015", 7, week=1)
        append_record(plan, None, None, "Sag", "Hospital Nord · 24015", 55)
        register["B7"] = "Omdøbt · Hospital Nord"
        register["B110"] = "Omdøbt projekt med udvidet nøgle"
        staff["B37"] = "Omdøbt person"
        if scenario == "flytning":
            row = append_record(plan, 2026, "Medarbejder 1 · p01", "Tilbud", "Hospital Nord · 24015", 8)
            # Reassign the original row, keeping its hours, and set it to Sag.
            plan[f"I{row}"], plan[f"D{row}"] = "Omdøbt person · p31", "Sag"
        if scenario in ("sorteret", "andet-projekt"):
            sort_table(plan)
        if scenario == "andet-projekt":
            overview["D6"], overview["E6"], overview["F6"] = "Vest", "EL", "QA35"
            report["A6"] = "Hospital Nord · 00123"
        emulate_filter(plan)

    def test_reference_totals_and_week_layout(self):
        plan, overview, _, report, _ = self.books["reference"].worksheets
        self.assertEqual(plan["G6"].value, 37)
        self.assertEqual(plan["H8"].value, 2026)
        self.assertEqual(plan["BJ7"].value, "Q4")
        self.assertTrue(plan["BJ8"].value)
        self.assertEqual([overview[f"{c}8"].value for c in "ABCDEFG"], [5, 20, 705, 685, 0, 10, 5])
        self.assertEqual([report[f"{c}14"].value for c in "GHIJK"], [48, 49, 50, 51, None])
        self.assertEqual(report["F16"].value, 20)

    def test_append_under_saved_filter_25_rows_and_new_employee(self):
        plan, overview, registry, report, staff = self.books["vokset"].worksheets
        self.assertEqual(plan["G6"].value, 40)
        self.assertEqual(plan["H8"].value, 2026)
        records = [(
            plan[f"A{r}"].value, plan[f"C{r}"].value, plan[f"BK{r}"].value,
            plan.cell(r, WEEK48).value,
        ) for r in range(10, plan.max_row + 1)]
        self.assertEqual(sum(id_ == "p31" and year == 2026 for year, id_, _, _ in records), 25)
        self.assertIn((2026, "p01", "24015", 3), records)
        self.assertIn((2027, "p01", "24015", None), records)
        self.assertIn((2026, "p31", "QA · 123", 3), records)
        for year, id_, key, hours in records:
            if hours == 55:
                self.assertIsNone(id_)
                self.assertIsNone(year)
        self.assertEqual(staff["L12"].value, "Omdøbt person · p31")
        self.assertEqual(registry["L12"].value, "Hospital Nord · 00123")
        self.assertEqual(registry["L13"].value, "Hospital Nord · 123")
        self.assertEqual(overview["A20"].value, "Omdøbt person")
        self.assertEqual(overview["F20"].value, 324)
        self.assertEqual(overview["A21"].value, "I ALT")
        self.assertEqual(overview["B8"].value, 347)
        self.assertEqual(overview["C8"].value, 825)
        self.assertEqual(overview["F8"].value, 11)
        self.assertEqual(overview["G8"].value, 5.5)
        self.assertEqual(report["A8"].value, "Omdøbt · Hospital Nord · 24015")
        self.assertEqual(report["F16"].value, 23)
        self.assertTrue(any(r[5].value == "Omdøbt · Hospital Nord" for r in list(plan.rows)[9:]))

    def test_sorting_and_identical_project_names_keep_keys_independent(self):
        left, right = (self.books[name] for name in ("vokset", "sorteret"))
        def records(book):
            plan = book.worksheets[0]
            return Counter(tuple(row[number - 1].value for number in (1, 3, 4, 63, *range(10, 63)))
                           for row in plan.iter_rows(min_row=10))
        self.assertEqual(records(left), records(right))
        for name, cells in (("Ressourceoverblik", ("A8", "B8", "D8", "E8", "F8", "G8", "F20", "F21")),
                            ("Projektoversigt", ("A8", "F15", "F16", "G16"))):
            self.assertEqual([left[name][c].value for c in cells], [right[name][c].value for c in cells])
        alternative = self.books["andet-projekt"]
        self.assertEqual(alternative["Ressourceoverblik"]["A15"].value, "Omdøbt person")
        self.assertEqual(alternative["Ressourceoverblik"]["F16"].value, 324)
        self.assertEqual(alternative["Projektoversigt"]["A15"].value, "Omdøbt person")
        self.assertEqual(alternative["Projektoversigt"]["BJ162"].value, "00123")
        self.assertEqual(alternative["Projektoversigt"]["G16"].value, 1)

    def test_reassignment_and_no_formula_errors(self):
        overview, project = (self.books["flytning"][name] for name in ("Ressourceoverblik", "Projektoversigt"))
        self.assertEqual((overview["F20"].value, overview["B8"].value), (332, 355))
        self.assertEqual(overview["F8"].value, 11)
        self.assertEqual((project["A16"].value, project["F16"].value, project["F17"].value),
                         ("Omdøbt person", 8, 31))
        for scenario, book in self.books.items():
            for sheet in book:
                with self.subTest(scenario=scenario, sheet=sheet.title):
                    self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"])


if __name__ == "__main__":
    unittest.main()
