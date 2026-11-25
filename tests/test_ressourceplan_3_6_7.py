"""Regressioner for filtre, tilbud og den nye 3.6.7-leverance."""

import json
import os
import runpy
import subprocess
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

from migrate_3_6_7 import NEW_BANNER, gross_formulas, migrate
from test_ressourceplan_3_6 import CALC, NODE, NAMES, ROOT, TESTS, TSC, pack, replay

OLD = ROOT / "Ressourceplan_3_6_3.xlsx"
BOOK = ROOT / "Ressourceplan_3_6_7.xlsx"
SCRIPTS = ROOT / "office-scripts" / "3.6.7"
NAMES = (*NAMES, "upgrade_workbook")


class ReleaseTests(unittest.TestCase):
    def test_only_gross_week_formulas_and_banners_change(self):
        old, new = load_workbook(OLD), load_workbook(BOOK)
        self.assertEqual(old.sheetnames, new.sheetnames)
        self.assertEqual(pack(old)["tables"], pack(new)["tables"])
        for before, after in zip(old.worksheets, new.worksheets):
            for a, b in zip(before.iter_rows(), after.iter_rows()):
                for left, right in zip(a, b):
                    if left.coordinate == "A1":
                        self.assertEqual(right.value, NEW_BANNER)
                    elif before.title == "Ressourceoverblik" and 590 <= left.row <= 689 and 8 <= left.column <= 60:
                        self.assertEqual((left.value, right.value), gross_formulas(left.row - 590, left.column - 8))
                    else:
                        self.assertEqual(left.value, right.value, f"{before.title}!{left.coordinate}")
            self.assertEqual(before.data_validations, after.data_validations)
            self.assertEqual(list(before.conditional_formatting), list(after.conditional_formatting))
            for area in before.conditional_formatting:
                self.assertEqual(before.conditional_formatting[area], after.conditional_formatting[area])
                self.assertEqual([rule.dxf for rule in before.conditional_formatting[area]],
                                 [rule.dxf for rule in after.conditional_formatting[area]])
        self.assertEqual(old["Min ressourceplan"].tables["PlanTimer"].autoFilter,
                         new["Min ressourceplan"].tables["PlanTimer"].autoFilter)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / BOOK.name
            migrate(OLD, output)
            self.assertEqual(pack(load_workbook(output))["tables"], pack(new)["tables"])
            with self.assertRaisesRegex(ValueError, "findes allerede"):
                migrate(OLD, output)

    def test_release_archive_and_sources(self):
        paths = [BOOK, SCRIPTS / "README.md", *(SCRIPTS / f"{name}.ts" for name in NAMES)]
        archive_path = SCRIPTS / "Ressourceplan_3_6_7.zip"
        before = {path: path.read_bytes() for path in [*paths, archive_path]}
        build = runpy.run_path(str(SCRIPTS / "build.py"))
        build["build"]()
        build["package"]()
        for path, contents in before.items():
            self.assertEqual(path.read_bytes(), contents, path.name)
        with ZipFile(archive_path) as archive:
            self.assertIsNone(archive.testzip())
            self.assertCountEqual(archive.namelist(), [path.name for path in paths])
            for path in paths:
                self.assertEqual(archive.read(path.name), before[path])

    @unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
    def test_enum_and_desktop_compatible_source(self):
        scripts = [SCRIPTS / f"{name}.ts" for name in NAMES]
        for path in scripts[:3]:
            code = path.read_text()
            self.assertIn("criteria.filterOn !== ExcelScript.FilterOn.custom", code)
            self.assertNotIn("String(criteria.filterOn)", code)
            self.assertNotIn("getActiveNamedSheetView", code)
            self.assertNotIn("view.getName()", code)
        result = subprocess.run(
            [NODE, str(TESTS / "check_office_scripts_compat.cjs"), TSC, str(TESTS / "excelscript_api.d.ts"),
             "--strict-office", *map(str, scripts)], text=True, capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(json.loads(result.stdout), [])


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
class ScriptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-367-")
        cls.addClassCleanup(temporary.cleanup)
        cls.folder = Path(temporary.name)
        compiled = cls.folder / "3.6.7"
        compiled.mkdir()
        for name in NAMES:
            result = subprocess.run(
                [TSC, "--strict", "--target", "ES2019", "--module", "none", "--outFile",
                 str(compiled / f"{name}.js"), str(TESTS / "excelscript_api.d.ts"), str(SCRIPTS / f"{name}.ts")],
                text=True, capture_output=True,
            )
            if result.returncode:
                raise AssertionError(result.stdout + result.stderr)
        book = load_workbook(OLD)
        fixture = pack(book)
        fixture["upgrade"] = {
            "formulas": list(book["Ressourceoverblik"].iter_rows(min_row=590, max_row=689, min_col=8,
                                                                 max_col=60, values_only=True)),
            "expected": [[gross_formulas(slot, week)[1] for week in range(53)] for slot in range(100)],
            "banners": {sheet.title: sheet["A1"].value for sheet in book.worksheets},
        }
        cls.fixture = cls.folder / "fixture.json"
        cls.fixture.write_text(json.dumps(fixture))

    def run_scenario(self, name):
        result = subprocess.run(
            [NODE, str(TESTS / "office_scripts_364_runner.cjs"), str(self.folder), str(self.fixture), name, "3.6.7"],
            text=True, capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_confirmation_feedback_and_preserved_hours(self):
        self.run_scenario("confirmed-367")

    def test_default_and_explicit_probability_and_format(self):
        self.run_scenario("probabilities-367")

    def test_upgrade_idempotent_and_refuses_changed_formulas_without_writing(self):
        self.run_scenario("upgrade-367")


@unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet")
class CalculationTests(unittest.TestCase):
    def test_gross_weighted_period_year_and_employee_filters(self):
        with tempfile.TemporaryDirectory(prefix="ressourceplan-367-calc-") as directory:
            folder = Path(directory)
            output = folder / "calculated"
            output.mkdir()
            expected = {
                "reference": (2026, 1, "Alle", "Alle", "Alle", (6, 33, 860, 827, 20, 10)),
                "mixed": (2026, 1, "Alle", "Alle", "Alle", (6, 33, 860, 827, 42, 21)),
                "combined": (2026, 1, "Øst", "EL", "7134.03", (1, 33, 148, 115, 25, 10)),
                "other": (2027, 5, "Vest", "EL", "7134.01", (1, 0, 120, 120, 23, 11.5)),
            }
            for label, (year, week, area, group, department, _) in expected.items():
                book = load_workbook(BOOK)
                state = pack(book)
                people = state["tables"]["PersonRegister"]["rows"]
                people[5] = ["p06", "Medarbejder 6", "Vest", "EL", "7134.03", 37]
                projects = state["tables"]["SagRegister"]["rows"]
                projects[5] = ["99999", "Testtilbud"]
                plan = state["tables"]["PlanTimer"]
                rows = plan["rows"]
                for row in rows:
                    row[9:62] = [""] * 53
                case = next(row for row in rows if row[0] == 2026 and row[1].endswith("p01") and row[3] == "Sag")
                case[9] = 33

                def offer(year, employee, week, hours, probability):
                    row = case.copy()
                    row[0], row[1], row[3], row[4], row[7] = year, employee, "Tilbud", "Testtilbud · 99999", probability
                    row[9:62] = [""] * 53
                    row[week + 8] = hours
                    rows.append(row)

                offer(2026, "Medarbejder 1 · p01", 1, 20, 0.5)
                if label != "reference":
                    offer(2026, "Medarbejder 1 · p01", 4, 5, 0)
                    offer(2026, "Medarbejder 2 · p02", 2, 5, 1)
                    offer(2026, "Medarbejder 5 · p05", 3, 12, "")
                    offer(2026, "Medarbejder 5 · p05", 5, 60, 1)
                    offer(2027, "Medarbejder 5 · p05", 5, 23, 0.5)
                replay(book, state)
                sheet = book["Ressourceoverblik"]
                for cell, value in zip(("A6", "B6", "D6", "E6", "F6"), (year, week, area, group, department)):
                    sheet[cell] = value
                book.save(folder / f"{label}.xlsx")
            run = subprocess.run(
                [CALC, f"-env:UserInstallation={(folder / 'profile').as_uri()}", "--headless", "--convert-to",
                 "xlsx", "--outdir", str(output), *(str(folder / f"{label}.xlsx") for label in expected)],
                check=True, capture_output=True, text=True, timeout=180,
                env=dict(os.environ, XDG_CACHE_HOME=str(folder / "cache")),
            )
            for label, (*_, totals) in expected.items():
                path = output / f"{label}.xlsx"
                self.assertTrue(path.is_file(), run.stdout + run.stderr)
                sheet = load_workbook(path, data_only=True)["Ressourceoverblik"]
                self.assertEqual(tuple(sheet[f"{column}8"].value for column in "ABCDFG"), totals, label)
                self.assertAlmostEqual(sheet["H8"].value, totals[1] / totals[2], places=9)
