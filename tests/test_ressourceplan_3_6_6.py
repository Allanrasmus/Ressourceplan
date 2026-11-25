"""Regressioner for Excels filteraliasing og Office Scripts 3.6.6-arkivet."""

import json
import runpy
import subprocess
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from openpyxl import load_workbook

from test_ressourceplan_3_6 import NODE, NAMES, ROOT, TESTS, TSC, pack

SCRIPTS = ROOT / "office-scripts" / "3.6.6"
ALL_SCRIPTS = NAMES
DECLARATION = TESTS / "excelscript_api.d.ts"
PACKAGE = "Ressourceplan_Office_Scripts_3_6_6.zip"


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
    def test_new_scripts_pass_stricter_guard(self):
        result, issues = compatibility(*(SCRIPTS / f"{name}.ts" for name in ALL_SCRIPTS))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(issues, [])


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
class FilterExecutionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-366-")
        cls.addClassCleanup(temporary.cleanup)
        cls.folder = Path(temporary.name)
        cls.fixture = cls.folder / "fixture.json"
        cls.fixture.write_text(json.dumps(pack(load_workbook(ROOT / "Ressourceplan_3_6_3.xlsx"))))
        for version, names in (("3.6.6", ALL_SCRIPTS),):
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
             str(self.fixture), scenario, "3.6.6"],
            text=True, capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_custom_filters_single_combined_and_unfiltered_with_formulas(self):
        self.run_scenario("custom-columns")

    def test_rejects_unsupported_filters_before_plan_mutation(self):
        self.run_scenario("custom-refusal")

    def test_attempts_each_restore_and_preserves_insert_diagnostics(self):
        self.run_scenario("custom-failure")

    def test_both_standard_script_entry_points_use_column_filters(self):
        self.run_scenario("custom-standard")
