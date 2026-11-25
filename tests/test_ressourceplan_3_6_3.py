"""Kildekompatibilitet, simulering og beregning af Ressourceplan 3.6.3."""

import json
import os
import runpy
import subprocess
import tempfile
import unittest
from collections import Counter
from copy import deepcopy
from itertools import groupby
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from migrate_3_6_3 import ACTIVITIES, PLAN_AREA, STRIPE_REFERENCE, migrate, stripe_rule
from test_ressourceplan_3_3_1 import digest
from test_ressourceplan_3_6 import CALC, NODE, NAMES, ROOT, TESTS, TSC, pack, replay
from test_ressourceplan_3_6_1 import PROTOTYPE, RELEASE as PREVIOUS

RELEASE = ROOT / "Ressourceplan_3_6_3.xlsx"
SCRIPTS = ROOT / "office-scripts" / "3.6.3"
SCRIPT_NAMES = (*NAMES, "upgrade_workbook")


class WorkbookTests(unittest.TestCase):
    def test_reference_preserves_data_input_validations_filters_and_outline(self):
        old, new = load_workbook(PREVIOUS), load_workbook(RELEASE)
        before, after = pack(old), pack(new)
        for name in ("PersonRegister", "SagRegister"):
            self.assertEqual(before["tables"][name], after["tables"][name])
        old_plan = before["tables"]["PlanTimer"]
        new_plan = after["tables"]["PlanTimer"]
        self.assertEqual(old_plan["headers"], new_plan["headers"])
        self.assertEqual(Counter(map(tuple, old_plan["rows"])), Counter(map(tuple, new_plan["rows"])))
        self.assertEqual(before["settings"][1:], after["settings"][1:])
        self.assertEqual(digest(PREVIOUS), "373220dcd439666fcdad4d9d36b98fa0b7dc24c4ac71f94fd49011be5d028331")
        self.assertEqual(digest(PROTOTYPE), "b51a94ea20c4d1834bd682d6cbe464dbd5ff619611a554ba632cddc17c58e2d3")
        for before, after in zip(old.worksheets, new.worksheets):
            self.assertEqual(before.title, after.title)
            self.assertEqual(before.data_validations, after.data_validations)
            self.assertEqual(list(before.tables.values()), list(after.tables.values()))
            self.assertEqual(before.sheet_state, after.sheet_state)
            for row in before:
                for cell in row:
                    if before.title == "Min ressourceplan" and cell.row >= 10:
                        continue
                    if cell.value is not None and cell.coordinate != "A1":
                        self.assertEqual(cell.value, after[cell.coordinate].value)
        plan = new["Min ressourceplan"]
        seen = set()
        for key, group in groupby(new_plan["rows"], key=lambda row: (row[0], row[1])):
            self.assertNotIn(key, seen)
            seen.add(key)
            kinds = [row[3] for row in group]
            self.assertEqual(kinds, sorted(kinds, key=ACTIVITIES.index))
        visible = [row for row in range(10, 10 + len(new_plan["rows"])) if not plan.row_dimensions[row].hidden]
        self.assertEqual(visible, list(range(10, 19)))
        employee = new["Medarbejdere"]
        self.assertTrue(all(cell.number_format == "@" for row in employee["A7:E106"] for cell in row))
        self.assertEqual(employee["E7"].data_type, "s")
        self.assertEqual(employee["F7"].data_type, "n")
        self.assertEqual(employee["F7"].number_format, old["Medarbejdere"]["F7"].number_format)
        self.assertEqual(new["Projektregister"]["A7"].data_type, "s")
        self.assertTrue(all(cell.number_format == "@" for row in new["Projektregister"]["A7:B506"] for cell in row))
        self.assertEqual(new["Ressourceoverblik"]["F6"].number_format, "@")

    def test_complete_package_has_current_excel_banners_and_button_spaces(self):
        workbook = load_workbook(RELEASE)
        for sheet in workbook:
            self.assertEqual(sheet["A1"].value, "COWI  |  Ressourceplan 3.6.3")
        for sheet, area, label in (
            ("Medarbejdere", "A5:B5", "Opret medarbejder"),
            ("Min ressourceplan", "A2:B2", "Tilføj projekt"),
            ("Min ressourceplan", "D2:F2", "Opret nyt planår"),
        ):
            self.assertIn(area, workbook[sheet].merged_cells)
            self.assertEqual(workbook[sheet][area.split(":")[0]].value, label)
        package = ROOT / "Ressourceplan_3_6_3.zip"
        expected = {RELEASE.name: RELEASE, "INSTALLATION.md": SCRIPTS / "README.md",
                    **{f"{name}.ts": SCRIPTS / f"{name}.ts" for name in SCRIPT_NAMES}}
        with ZipFile(package) as archive:
            self.assertIsNone(archive.testzip())
            self.assertEqual(set(archive.namelist()), set(expected))
            for name, path in expected.items():
                self.assertEqual(archive.read(name), path.read_bytes())
        with tempfile.TemporaryDirectory() as folder:
            rebuilt = Path(folder) / package.name
            runpy.run_path(str(SCRIPTS / "build.py"))["package"](rebuilt)
            self.assertEqual(digest(package), digest(rebuilt))

    @unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet")
    def test_reordered_packaged_workbook_preserves_baseline_overview_and_visible_hours(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            result = subprocess.run(
                [CALC, f"-env:UserInstallation={(folder / 'profile').as_uri()}", "--headless",
                 "--convert-to", "xlsx", "--outdir", str(folder), str(PREVIOUS), str(RELEASE)],
                capture_output=True, check=True, text=True, timeout=180,
                env=dict(os.environ, XDG_CACHE_HOME=str(folder / "cache")),
            )
            files = [folder / path.name for path in (PREVIOUS, RELEASE)]
            self.assertTrue(all(path.is_file() for path in files), result.stdout + result.stderr)
            before, after = (load_workbook(path, data_only=True) for path in files)
            for name, cells in (
                ("Min ressourceplan", ("G6", "H8", "BE6", "BF6", "BG6", "BH6")),
                ("Ressourceoverblik", tuple(f"{column}8" for column in "ABCDEFGH")),
                ("Projektoversigt", ("A15", "G15", "A16", "G16")),
            ):
                self.assertEqual([before[name][cell].value for cell in cells],
                                 [after[name][cell].value for cell in cells])
            self.assertEqual([after["Planstart"][f"R{row}"].value for row in range(10, 15)],
                             [1, 0, 1, 0, 0])
            for sheet in after:
                self.assertFalse([cell.coordinate for row in sheet for cell in row if cell.data_type == "e"],
                                 sheet.title)

    def test_rule_fills_borders_and_priorities_preserve_activity_calendar_and_overview(self):
        old, new = load_workbook(PREVIOUS), load_workbook(RELEASE)
        before, after = old["Min ressourceplan"], new["Min ressourceplan"]
        self.assertFalse(after.tables["PlanTimer"].tableStyleInfo.showRowStripes)
        self.assertEqual(new.defined_names["PlanSagsstribe"].attr_text, STRIPE_REFERENCE)
        self.assertEqual(new["Planstart"].sheet_state, "hidden")
        self.assertIn("SUBTOTAL(103,INDEX(PlanTimer[Type],ROW()-9))", new["Planstart"]["R10"].value)
        self.assertIn("N(R20008)", new["Planstart"]["R20009"].value)
        for area in before.conditional_formatting:
            for left, right in zip(before.conditional_formatting[area], after.conditional_formatting[area], strict=True):
                self.assertEqual(left.priority, right.priority)
                self.assertEqual(left.stopIfTrue, right.stopIfTrue)
                if str(area.sqref) == PLAN_AREA and "MOD(ROW(),2)" in left.formula[0]:
                    parity = int(left.formula[0][-2])
                    self.assertEqual(right.formula, [stripe_rule(parity)])
                    self.assertEqual(right.dxf.fill.bgColor.rgb, "FFEAF2FC" if parity else "FFFFFFFF")
                    self.assertEqual(right.dxf.border, left.dxf.border)
                else:
                    self.assertEqual(left, right)
                    self.assertEqual(left.dxf, right.dxf)
        for sheet in ("Ressourceoverblik", "Projektoversigt"):
            for area in old[sheet].conditional_formatting:
                old_rules = old[sheet].conditional_formatting[area]
                new_rules = new[sheet].conditional_formatting[area]
                self.assertEqual(old_rules, new_rules)
                self.assertEqual([rule.dxf for rule in old_rules], [rule.dxf for rule in new_rules])

    def test_reference_build_and_internal_structure(self):
        with tempfile.TemporaryDirectory() as folder:
            rebuilt = Path(folder) / "test.xlsx"
            migrate(PREVIOUS, rebuilt)
            with self.assertRaisesRegex(ValueError, "findes allerede"):
                migrate(PREVIOUS, rebuilt)
            with ZipFile(RELEASE) as saved, ZipFile(rebuilt) as built:
                self.assertIsNone(saved.testzip())
                self.assertEqual(saved.namelist(), built.namelist())
                for part in saved.namelist():
                    if part == "docProps/core.xml":
                        continue
                    left, right = saved.read(part), built.read(part)
                    if part.endswith((".xml", ".rels")):
                        left, right = ET.canonicalize(left), ET.canonicalize(right)
                    self.assertEqual(left, right, part)


class ScriptBuildTests(unittest.TestCase):
    def test_generated_sources(self):
        for name in SCRIPT_NAMES:
            parts = [name, "common"]
            if name in ("add_project", "initialize_plan_year"):
                parts.append("selection")
            if name in ("create_employee", "initialize_plan_year"):
                parts.append("standard")
            expected = "\n".join((SCRIPTS / "src" / f"{part}.ts").read_text(encoding="utf-8") for part in parts)
            actual = (SCRIPTS / f"{name}.ts").read_text(encoding="utf-8")
            self.assertEqual(actual.split("\n", 1)[1], expected)
            self.assertEqual(actual.count("function main("), 1)

    @unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
    def test_static_compatibility(self):
        command = [NODE, str(TESTS / "check_office_scripts_compat.cjs"), TSC,
                   str(TESTS / "excelscript_api.d.ts"), *[str(SCRIPTS / f"{name}.ts") for name in SCRIPT_NAMES]]
        result = subprocess.run(command, text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(json.loads(result.stdout), [])


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
class ScriptExecutionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-363-")
        cls.addClassCleanup(temporary.cleanup)
        cls.folder = Path(temporary.name)
        compiled = cls.folder / "compiled"
        compiled.mkdir()
        for name in SCRIPT_NAMES:
            command = [TSC, "--strict", "--target", "ES2019", "--module", "none",
                       "--outFile", str(compiled / f"{name}.js"),
                       str(TESTS / "excelscript_api.d.ts"), str(SCRIPTS / f"{name}.ts")]
            result = subprocess.run(command, capture_output=True, text=True)
            if result.returncode:
                raise AssertionError(result.stdout + result.stderr)
        fixture, packaged, output = (cls.folder / name for name in ("fixture.json", "packaged.json", "result.json"))
        cls.old = pack(load_workbook(PREVIOUS))
        fixture.write_text(json.dumps(cls.old))
        packaged.write_text(json.dumps(pack(load_workbook(RELEASE))))
        result = subprocess.run(
            [NODE, str(TESTS / "office_scripts_363_runner.cjs"), str(compiled), str(fixture), str(output), str(packaged)],
            text=True, capture_output=True,
        )
        if result.returncode:
            raise AssertionError(result.stdout + result.stderr)
        cls.state = json.loads(output.read_text())

    def test_mock_department_types_ordering_and_preserved_historical_records(self):
        people = self.state["tables"]["PersonRegister"]["rows"]
        self.assertEqual([entry[4] for entry in people[5:9]], ["7134.01", "1734.03", "7134.10", "01734.03"])
        self.assertTrue(all(type(entry[5]) is int for entry in people[5:9]))
        rows = self.state["tables"]["PlanTimer"]["rows"]
        original = self.old["tables"]["PlanTimer"]["rows"]
        self.assertEqual(rows[:len(original)], original)
        inserted = [row for row in rows if row[0] == 2026 and row[1] == "Testperson 6 · p06"]
        self.assertEqual([row[3] for row in inserted],
                         ["Sag"] * 5 + ["Tilbud"] * 2 + ["Ferie", "Fravær", "Kursus", "Andet"])
        self.assertEqual([row[56] for row in inserted], [17, "", "", "", "", 5, "", "", "", "", ""])

    @unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet")
    def test_calc_stripes_after_filter_sort_insert_and_delete_and_department_dropdown(self):
        original = self.state["tables"]["PlanTimer"]["rows"]
        alternate = deepcopy(self.state)
        rows = alternate["tables"]["PlanTimer"]["rows"]

        def order(row):
            return (row[0], row[1], ["Sag", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"].index(row[3]), row[4])

        rows.sort(key=order)
        # Delete a new zero-hour case and insert a fresh vacancy in the sorted table.
        sample = next(row for row in rows if row[0] == 2026 and row[1] == "Testperson 6 · p06"
                      and row[4] == "Datacenter · 24022" and row[3] == "Sag")
        position = rows.index(sample)
        rows.pop(position)
        copied = sample.copy()
        copied[4] = ""
        rows.insert(position, copied)
        inputs = []
        scenarios = []
        for label, state in (("filtreret", self.state), ("sorteret", alternate)):
            workbook = load_workbook(RELEASE)
            replay(workbook, state)
            plan = workbook["Min ressourceplan"]
            worker_cases = [
                i for i, row in enumerate(state["tables"]["PlanTimer"]["rows"], 10)
                if row[0] == 2026 and row[1] == "Testperson 6 · p06" and row[3] == "Sag"
            ]
            for i, row in enumerate(state["tables"]["PlanTimer"]["rows"], 10):
                plan.row_dimensions[i].hidden = (row[0] != 2026 or row[1] != "Testperson 6 · p06" or i == worker_cases[1])
            for i, parity in zip(worker_cases, [1, 1, 0, 1, 0], strict=True):
                plan[f"BO{i}"] = Translator("=" + stripe_rule(parity), origin="A10").translate_formula(f"A{i}")
            self.assertEqual(plan.tables["PlanTimer"].autoFilter.filterColumn[1].filters.filter, ["Medarbejder 1 · p01"])
            department = "7134.01" if label == "filtreret" else "01734.03"
            workbook["Ressourceoverblik"]["F6"] = department
            path = self.folder / f"{label}.xlsx"
            workbook.save(path)
            inputs.append(str(path))
            scenarios.append((path.name, worker_cases, department))
        self.assertEqual(len(rows), len(original))
        output = self.folder / "calc"
        output.mkdir()
        result = subprocess.run(
            [CALC, f"-env:UserInstallation={(self.folder / 'profile').as_uri()}", "--headless",
             "--convert-to", "xlsx", "--outdir", str(output), *inputs],
            capture_output=True, check=True, text=True, timeout=180,
            env=dict(os.environ, XDG_CACHE_HOME=str(self.folder / "cache")),
        )
        for name, cases, department in scenarios:
            self.assertTrue((output / name).is_file(), result.stdout + result.stderr)
            workbook = load_workbook(output / name, data_only=True)
            overview = workbook["Ressourceoverblik"]
            self.assertEqual(workbook["Medarbejdere"]["H12"].value, "7134.01")
            self.assertEqual(workbook["Medarbejdere"]["H15"].value, "01734.03")
            expected = [3, 17, 392, 375, 0, 5, 2.5] if department == "7134.01" else [1, 0, 124, 124, 0, 0, 0]
            self.assertEqual([overview[c].value for c in ("A8", "B8", "C8", "D8", "E8", "F8", "G8")], expected)
            self.assertEqual([workbook["Planstart"][f"R{row}"].value for row in cases], [1, 1, 0, 1, 0])
            self.assertTrue(all(workbook["Min ressourceplan"][f"BO{row}"].value for row in cases))
            self.assertEqual(workbook["Planstart"]["R20009"].value, 0)
            project = workbook["Projektoversigt"]
            self.assertEqual([project[f"G{row}"].value for row in (15, 16, 17)], [20, 17, 37])
            for sheet in workbook:
                self.assertFalse([cell.coordinate for row in sheet for cell in row if cell.data_type == "e"], sheet.title)
