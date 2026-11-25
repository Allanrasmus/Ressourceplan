"""Kompatibilitet og regression for scriptrettelsen. Excel-filen er 3.6.1."""

import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from test_ressourceplan import CALC
from test_ressourceplan_3_3_1 import digest
from test_ressourceplan_3_6 import NODE, NAMES, ROOT, TESTS, TSC, pack, replay
from test_ressourceplan_3_6_1 import RELEASE, PROTOTYPE

SCRIPTS = ROOT / "office-scripts" / "3.6.2"
OLD = ROOT / "office-scripts" / "3.6.1"


class ScriptReleaseTests(unittest.TestCase):
    def test_workbooks_and_built_script_sources(self):
        self.assertEqual(digest(RELEASE), "373220dcd439666fcdad4d9d36b98fa0b7dc24c4ac71f94fd49011be5d028331")
        self.assertEqual(digest(PROTOTYPE), "b51a94ea20c4d1834bd682d6cbe464dbd5ff619611a554ba632cddc17c58e2d3")
        for script in NAMES:
            parts = [script, "common"]
            if script != "create_employee":
                parts.append("selection")
            if script != "add_project":
                parts.append("standard")
            expected = "\n".join((SCRIPTS / "src" / f"{name}.ts").read_text(encoding="utf-8")
                                 for name in parts)
            self.assertEqual((SCRIPTS / f"{script}.ts").read_text(encoding="utf-8").split("\n", 1)[1],
                             expected)


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet til Office Scripts-test")
class CompatibilityTests(unittest.TestCase):
    def check_sources(self, root):
        result = subprocess.run(
            [NODE, str(TESTS / "check_office_scripts_compat.cjs"), TSC,
             str(TESTS / "excelscript_api.d.ts"), *[str(root / f"{name}.ts") for name in NAMES]],
            capture_output=True, text=True, check=False,
        )
        self.assertIn(result.returncode, (0, 1), result.stderr)
        return result.returncode, json.loads(result.stdout)

    def test_static_guard_detects_three_reported_legacy_spots(self):
        code, issues = self.check_sources(OLD)
        self.assertEqual(code, 1)
        points = {(Path(issue["file"]).name, issue["line"]) for issue in issues
                  if issue["rule"] == "api-helper-argument"}
        self.assertTrue({("add_project.ts", 32), ("add_project.ts", 282),
                         ("initialize_plan_year.ts", 244)}.issubset(points), points)

    def test_all_three_scripts_avoid_native_handles_in_helpers_and_data(self):
        code, issues = self.check_sources(SCRIPTS)
        self.assertEqual((code, issues), (0, []))


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet til Office Scripts-test")
class ScriptExecutionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temp = tempfile.TemporaryDirectory(prefix="ressourceplan-362-")
        cls.addClassCleanup(temp.cleanup)
        cls.folder = Path(temp.name)
        cls.compiled = cls.folder / "compiled"
        cls.compiled.mkdir()
        for script in NAMES:
            completed = subprocess.run(
                [TSC, "--strict", "--target", "ES2019", "--module", "none",
                 "--outFile", str(cls.compiled / f"{script}.js"),
                 str(TESTS / "excelscript_api.d.ts"), str(SCRIPTS / f"{script}.ts")],
                capture_output=True, text=True,
            )
            if completed.returncode:
                raise AssertionError(completed.stdout + completed.stderr)
        cls.fixture = cls.folder / "fixture.json"
        cls.fixture.write_text(json.dumps(pack(load_workbook(RELEASE))))
        result = cls.run_simulator("office_scripts_361_runner.cjs", "result.json")
        cls.rows = result

    @classmethod
    def run_simulator(cls, runner, output):
        target = cls.folder / output
        subprocess.run([NODE, str(TESTS / runner), str(cls.compiled),
                        str(cls.fixture), str(target)], check=True, text=True)
        return json.loads(target.read_text())

    def test_explicit_inputs_markings_duplicates_and_independent_historical_rows(self):
        original = pack(load_workbook(RELEASE))["tables"]["PlanTimer"]["rows"]
        saved = self.rows["tables"]["PlanTimer"]["rows"]
        self.assertEqual(saved[:len(original)], original)
        new = [row for row in saved if row[1] == "Testperson 361 · p06"]
        self.assertEqual([row[3] for row in new if row[0] == 2032],
                         ["Sag", "Ferie", "Fravær", "Kursus", "Andet"])
        self.assertEqual([row[7] for row in new if row[3] == "Tilbud"], ["", 0, .25, .5, 1])
        self.assertTrue(all(row[7] == "" for row in new if row[3] == "Sag"))

    def test_large_table_batch_limits_and_race_before_append(self):
        metrics = self.run_simulator("office_scripts_362_runner.cjs", "metrics.json")
        self.assertEqual(metrics["largeMetadataReads"], 5)
        self.assertLessEqual(metrics["largeHourReads"], 41)
        self.assertEqual(metrics["smallFormulaWrites"], 4)
        self.assertEqual(metrics["selectedRowReads"], 1)

    @unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
    def test_calc_on_separate_copy_preserves_weekly_totals_and_filter_definition(self):
        workbook = load_workbook(RELEASE)
        replay(workbook, self.rows)
        path = self.folder / "scriptresultat.xlsx"
        workbook.save(path)
        output = self.folder / "calc"
        output.mkdir()
        result = subprocess.run(
            [CALC, f"-env:UserInstallation={(self.folder / 'profile').as_uri()}", "--headless",
             "--convert-to", "xlsx", "--outdir", str(output), str(path)],
            capture_output=True, check=True, text=True, timeout=120,
            env=dict(os.environ, XDG_CACHE_HOME=str(self.folder / "cache")),
        )
        self.assertTrue((output / path.name).is_file(), result.stderr)
        calculated = load_workbook(output / path.name, data_only=True)
        overview = calculated["Ressourceoverblik"]
        self.assertEqual([overview[f"{c}8"].value for c in "ABCDEFG"], [6,28,825,797,0,50,23])
        self.assertEqual(calculated["Projektoversigt"]["G17"].value, 36)
        plan_filter = workbook["Min ressourceplan"].tables["PlanTimer"].autoFilter
        self.assertEqual(plan_filter.filterColumn[1].filters.filter, ["Medarbejder 1 · p01"])
        for sheet in calculated:
            self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"],
                             sheet.title)
