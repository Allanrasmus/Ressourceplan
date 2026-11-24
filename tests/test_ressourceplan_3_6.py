"""Office Scripts compile against the documented API and act on workbook-shaped fixtures."""

import json
import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries

from migrate_3_6 import migrate
from test_ressourceplan import CALC, ROOT
from test_ressourceplan_3_3_1 import NS, digest
from test_ressourceplan_3_5_1 import RELEASE as PREVIOUS

RELEASE = ROOT / "Ressourceplan_3_6.xlsx"
SCRIPTS = ROOT / "office-scripts"
TESTS = ROOT / "tests"
TSC = os.environ.get("TSC") or shutil.which("tsc")
NODE = shutil.which("node")
NAMES = ("create_employee", "add_project", "initialize_plan_year")


class WorkbookTests(unittest.TestCase):
    def test_source_cells_tables_formulas_and_status(self):
        self.assertEqual(digest(PREVIOUS), "8c1ec6759d07baa4434dfd31c1f2f0192dc6215c9c30738ad4375831ffe93edf")
        old, new = load_workbook(PREVIOUS), load_workbook(RELEASE)
        self.assertEqual(old.sheetnames, new.sheetnames)
        self.assertEqual(new["Planstart"].sheet_state, "hidden")
        self.assertEqual(new.active.title, old.active.title)
        for name, start, cols in (("Min ressourceplan", 9, 65), ("Projektregister", 6, 13),
                                  ("Medarbejdere", 6, 13)):
            s, target = old[name], new[name]
            for row in s.iter_rows(min_row=start, max_col=cols):
                for cell in row:
                    self.assertEqual(cell.value, target[cell.coordinate].value)
            for table in s.tables.values():
                self.assertEqual(target.tables[table.name].ref, table.ref)
                self.assertEqual(target.tables[table.name].autoFilter, table.autoFilter)
            self.assertEqual(target.data_validations, s.data_validations)
        for name in ("Ressourceoverblik", "Projektoversigt"):
            for row in old[name].iter_rows(min_row=2):
                for cell in row:
                    self.assertEqual(cell.value, new[name][cell.coordinate].value)
        plan, config = new["Min ressourceplan"], new["Planstart"]
        self.assertEqual((plan["A2"].value, plan["D2"].value, new["Medarbejdere"]["A5"].value),
                         ("Tilføj projekt", "Opret nyt planår", "Opret medarbejder"))
        self.assertEqual(config["P6"].value, "PlanTimer-formler 3.6")
        expected = [(c.name, "=" + c.calculatedColumnFormula.attr_text)
                    for c in plan.tables["PlanTimer"].tableColumns if c.calculatedColumnFormula]
        self.assertEqual([(config[f"P{n}"].value, config[f"Q{n}"].value)
                          for n in range(7, 14)], expected)
        self.assertTrue(all(config[f"Q{n}"].data_type == "s" for n in range(7, 14)))

    def test_ooxml_week_53_and_reproducibility(self):
        with ZipFile(RELEASE) as release:
            self.assertIsNone(release.testzip())
            self.assertFalse(any(s in n.lower() for n in release.namelist()
                                 for s in ("vba", "externallinks", "connections", "drawings")))
            for n in release.namelist():
                if n.endswith((".xml", ".rels")):
                    root = ET.fromstring(release.read(n))
                    self.assertFalse(any(e.get("TargetMode") == "External" for e in root.iter()))
            root = ET.fromstring(release.read("xl/workbook.xml"))
            target = root.find("s:definedNames/s:definedName[@name='PlanSidsteRaekke']", NS)
            self.assertIn("PlanTimer[#Data]", target.text)
            sheet = ET.fromstring(release.read("xl/worksheets/sheet1.xml"))
            rule = sheet.find("s:conditionalFormatting[@sqref='BJ10:BJ20009']/s:cfRule/s:formula", NS)
            self.assertIn("ROW()<=PlanSidsteRaekke", rule.text)
            self.assertIn("ISNUMBER($A10)", rule.text)
            self.assertIn("=364", rule.text)
            validation = sheet.find("s:dataValidations/s:dataValidation[@sqref='BJ10:BJ20009']/s:formula1", NS)
            self.assertIn("=371", validation.text)
        with tempfile.TemporaryDirectory() as d:
            output = Path(d) / RELEASE.name
            migrate(PREVIOUS, output)
            with self.assertRaises(ValueError):
                migrate(PREVIOUS, output)
            with ZipFile(RELEASE) as release, ZipFile(output) as rebuilt:
                self.assertEqual(release.namelist(), rebuilt.namelist())
                for part in release.namelist():
                    if part == "docProps/core.xml":
                        continue
                    left, right = release.read(part), rebuilt.read(part)
                    if part.endswith((".xml", ".rels")):
                        left, right = ET.canonicalize(left), ET.canonicalize(right)
                    self.assertEqual(left, right, part)

    def test_script_build_matches_checked_in_sources(self):
        shared = (SCRIPTS / "src/common.ts").read_text(encoding="utf-8")
        standard = (SCRIPTS / "src/standard.ts").read_text(encoding="utf-8")
        for name in NAMES:
            parts = [(SCRIPTS / "src" / f"{name}.ts").read_text(encoding="utf-8"), shared]
            if name != "add_project":
                parts.append(standard)
            actual = (SCRIPTS / f"{name}.ts").read_text(encoding="utf-8")
            self.assertEqual(actual.split("\n", 1)[1], "\n".join(parts))
            self.assertEqual(actual.count("function main("), 1)


def pack(workbook):
    tables = {}
    for name, sheet in (("PlanTimer", "Min ressourceplan"), ("PersonRegister", "Medarbejdere"),
                        ("SagRegister", "Projektregister")):
        tab = workbook[sheet].tables[name]
        first, header, last, end = range_boundaries(tab.ref)
        tables[name] = {
            "header": header - 1,
            "headers": [column.name for column in tab.tableColumns],
            "rows": [[c if c is not None else "" for c in row]
                     for row in workbook[sheet].iter_rows(min_row=header+1, max_row=end,
                                                          min_col=first, max_col=last, values_only=True)],
        }
    config = workbook["Planstart"]
    settings = [[v if v is not None else "" for v in row]
                for row in config.iter_rows(max_row=26, max_col=17, values_only=True)]
    return {"tables": tables, "settings": settings}


def replay(workbook, state):
    # Repeat only the writes that the tested scripts performed. Calc evaluates
    # those resulting Excel formulas on a separate workbook, not on the release.
    old = pack(workbook)
    for name, sheet in (("PlanTimer", "Min ressourceplan"), ("PersonRegister", "Medarbejdere"),
                        ("SagRegister", "Projektregister")):
        table = workbook[sheet].tables[name]
        start = state["tables"][name]["header"] + 2
        rows = state["tables"][name]["rows"]
        previous = old["tables"][name]["rows"]
        for idx, row in enumerate(rows):
            for col, value in enumerate(row, 1):
                if idx >= len(previous) or value != previous[idx][col-1]:
                    workbook[sheet].cell(start+idx, col).value = value if value != "" else None
        if name == "PlanTimer":
            table.ref = table.autoFilter.ref = f"A9:BM{start + len(rows) - 1}"
    for row in range(7, 27):
        value = state["settings"][row - 1][5]
        if value != old["settings"][row - 1][5]:
            workbook["Planstart"][f"F{row}"] = value


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet til Office Scripts-test")
class ScriptTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-36-")
        cls.addClassCleanup(temporary.cleanup)
        cls.folder = Path(temporary.name)
        compiled = cls.folder / "compiled"
        compiled.mkdir()
        for name in NAMES:
            subprocess.run([TSC, "--strict", "--target", "ES2019", "--module", "none",
                            "--outFile", str(compiled / f"{name}.js"),
                            str(TESTS / "excelscript_api.d.ts"), str(SCRIPTS / f"{name}.ts")],
                           check=True, capture_output=True, text=True)
        fixture = cls.folder / "input.json"
        result = cls.folder / "result.json"
        cls.source = load_workbook(RELEASE)
        fixture.write_text(json.dumps(pack(cls.source)))
        subprocess.run([NODE, str(TESTS / "office_scripts_runner.cjs"), str(compiled),
                        str(fixture), str(result)], check=True, text=True)
        cls.state = json.loads(result.read_text())

    def test_three_scripts_compile_and_run_validation_retry_and_duplicate_cases(self):
        rows = self.state["tables"]["PlanTimer"]["rows"]
        original = pack(self.source)["tables"]["PlanTimer"]["rows"]
        self.assertEqual(rows[:len(original)], original)
        expected = ["Sag", "Ferie", "Fravær", "Kursus", "Andet"]
        self.assertEqual([row[3] for row in rows[len(original):len(original)+5]], expected)
        self.assertEqual([row[3] for row in rows[len(original)+5:len(original)+10]], expected)
        self.assertTrue(all(row[2].startswith("=") and row[62].startswith("=")
                            for row in rows[len(original):]))

    @unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
    def test_recalculate_script_writes_and_week53_rule(self):
        workbook = load_workbook(RELEASE)
        replay(workbook, self.state)
        plan = workbook["Min ressourceplan"]
        table = plan.tables["PlanTimer"]
        last = plan.max_row
        new_year_row = next(r for r in range(10, last + 1)
                            if plan[f"A{r}"].value == 2027 and "p31" in plan[f"B{r}"].value)
        gray = plan.conditional_formatting["BJ10:BJ20009"][0].formula[0]
        plan[f"A{last+1}"] = 2027
        for row in (10, new_year_row, last+1, last+2):
            plan[f"BO{row}"] = Translator("=" + gray, origin="BJ10").translate_formula(f"BJ{row}")
        plan["BO1"] = "=PlanSidsteRaekke"
        self.assertEqual(table.autoFilter.filterColumn[1].filters.filter, ["Medarbejder 1 · p01"])
        folder = self.folder
        source = folder / "test.xlsx"
        output = folder / "calc"
        output.mkdir()
        workbook.save(source)
        tool = subprocess.run(
            [CALC, f"-env:UserInstallation={(folder / 'profile').as_uri()}", "--headless",
             "--convert-to", "xlsx", "--outdir", str(output), str(source)], check=True,
            capture_output=True, text=True, timeout=120,
            env=dict(os.environ, XDG_CACHE_HOME=str(folder / "cache")),
        )
        self.assertTrue((output / source.name).exists(), tool.stdout + tool.stderr)
        rebuilt = load_workbook(output / source.name, data_only=True)
        plan, overview, _, project, members, _ = rebuilt.worksheets
        self.assertEqual(plan["BO1"].value, last)
        self.assertEqual([plan[f"BO{r}"].value for r in (10, new_year_row, last+1, last+2)],
                         [False, True, False, False])
        self.assertEqual(plan[f"BJ{new_year_row}"].value, 20)
        self.assertEqual(plan[f"BM{new_year_row}"].value, 0)
        self.assertEqual(members["A37"].value, "p31")
        self.assertEqual(members["B37"].value, "Omdøbt Testperson")
        self.assertEqual([overview[f"{col}8"].value for col in "ABCDEFG"], [6, 28, 825, 797, 0, 22, 8])
        self.assertEqual(overview["A20"].value, "Omdøbt Testperson")
        self.assertAlmostEqual(overview["BH20"].value, 9/30)
        self.assertEqual([project[f"G{r}"].value for r in (15, 16, 17)], [20, 20, 40])
        for sheet in rebuilt:
            self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"],
                             sheet.title)


if __name__ == "__main__":
    unittest.main()
