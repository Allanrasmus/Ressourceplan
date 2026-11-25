"""Regressioner for Excels filteraliasing og Office Scripts 3.6.5-arkivet."""

import json
import runpy
import subprocess
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

from test_ressourceplan_3_6 import NODE, NAMES, ROOT, TESTS, TSC, pack

SCRIPTS = ROOT / "office-scripts" / "3.6.5"
ALL_SCRIPTS = (*NAMES, "filter_probe")
DECLARATION = TESTS / "excelscript_api.d.ts"
PACKAGE = "Ressourceplan_Office_Scripts_3_6_5.zip"


def compatibility(*sources):
    result = subprocess.run(
        [NODE, str(TESTS / "check_office_scripts_compat.cjs"), TSC, str(DECLARATION),
         "--strict-office", *map(str, sources)],
        text=True, capture_output=True,
    )
    return result, json.loads(result.stdout)


class PackageTests(unittest.TestCase):
    def test_reproducible_scripts_and_package(self):
        names = [f"{name}.ts" for name in ALL_SCRIPTS]
        before = {name: (SCRIPTS / name).read_bytes() for name in [*names, PACKAGE]}
        build = runpy.run_path(str(SCRIPTS / "build.py"))
        build["build"]()
        build["package"]()
        for name, contents in before.items():
            self.assertEqual((SCRIPTS / name).read_bytes(), contents, name)
        with ZipFile(SCRIPTS / PACKAGE) as archive:
            self.assertIsNone(archive.testzip())
            self.assertEqual(sorted(archive.namelist()), sorted(["README.md", *names[:3]]))
            for name in archive.namelist():
                self.assertEqual(archive.read(name), (SCRIPTS / name).read_bytes(), name)


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
class CompatibilityTests(unittest.TestCase):
    def test_known_364_property_alias_in_all_three_scripts(self):
        sources = [SCRIPTS.parent / "3.6.4" / f"{name}.ts" for name in NAMES]
        result, issues = compatibility(*sources)
        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)
        by_file = {Path(item["file"]).name: item for item in issues if item["rule"] == "api-property-argument"}
        self.assertEqual(set(by_file), {file.name for file in sources})
        for file in sources:
            issue = by_file[file.name]
            self.assertIn("filter.criteria", file.read_text().splitlines()[issue["line"] - 1])
        self.assertTrue({"api-in-data", "api-in-data-type", "api-array-write"} <=
                        {issue["rule"] for issue in issues})

    def test_nested_api_values_assignments_and_callback_rules(self):
        code = """
function main(workbook: ExcelScript.Workbook) {
  const table = workbook.getTable("PlanTimer");
  if (!table) return;
  const autoFilter = table.getAutoFilter();
  const read = autoFilter.getCriteria();
  const range = table.getRange();
  const nested: { data: [ExcelScript.FilterCriteria] } = { data: [read[0]] };
  let copied = nested;
  copied = nested;
  autoFilter.apply(range, 0, copied.data[0]);
  [1].map(function (value) { return value; });
}
"""
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "alias.ts"
            path.write_text(code)
            result, issues = compatibility(path)
        self.assertEqual(result.returncode, 1)
        self.assertTrue({"api-in-data", "api-alias", "api-assignment", "api-property-argument",
                         "array-callback"} <= {item["rule"] for item in issues}, issues)

    def test_new_scripts_and_probe_pass_stricter_guard(self):
        result, issues = compatibility(*(SCRIPTS / f"{name}.ts" for name in ALL_SCRIPTS))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(issues, [])


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
class FilterExecutionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-365-")
        cls.addClassCleanup(temporary.cleanup)
        cls.folder = Path(temporary.name)
        cls.fixture = cls.folder / "fixture.json"
        cls.fixture.write_text(json.dumps(pack(load_workbook(ROOT / "Ressourceplan_3_6_3.xlsx"))))
        for version, names in (("3.6.5", ALL_SCRIPTS), ("3.6.3", ("add_project",))):
            destination = cls.folder / version
            destination.mkdir()
            for name in names:
                subprocess.run(
                    [TSC, "--strict", "--target", "ES2019", "--module", "none",
                     "--outFile", str(destination / f"{name}.js"),
                     str(DECLARATION), str(SCRIPTS.parent / version / f"{name}.ts")],
                    check=True, text=True, capture_output=True,
                )

    def run_scenario(self, scenario):
        result = subprocess.run(
            [NODE, str(TESTS / "office_scripts_364_runner.cjs"), str(self.folder),
             str(self.fixture), scenario, "3.6.5"],
            text=True, capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_preserves_filter_and_insert_regressions(self):
        for scenario in ("project", "standard", "bypass", "view", "clear", "insert", "restore"):
            with self.subTest(scenario=scenario):
                self.run_scenario(scenario)

    def test_copies_complete_criteria_and_round_trips_in_probe(self):
        self.run_scenario("snapshot")

    def test_restore_setup_failure_preserves_original_insert_failure(self):
        self.run_scenario("restore-setup")
