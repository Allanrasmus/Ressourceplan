"""Regressionstest af filtreret PlanTimer-indsættelse og scriptleverancen."""

import json
import runpy
import subprocess
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from test_ressourceplan_3_6 import NODE, NAMES, ROOT, TESTS, TSC, pack

SCRIPTS = ROOT / "office-scripts" / "3.6.4"
BASELINE = ROOT / "office-scripts" / "3.6.3"


class ScriptBuildTests(unittest.TestCase):
    def test_generated_sources(self):
        before = {name: (SCRIPTS / f"{name}.ts").read_bytes() for name in NAMES}
        runpy.run_path(str(SCRIPTS / "build.py"))["build"]()
        for name, source in before.items():
            self.assertEqual(source, (SCRIPTS / f"{name}.ts").read_bytes(), name)
            self.assertEqual(source.count(b"function main("), 1, name)

    @unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
    def test_static_compatibility(self):
        command = [NODE, str(TESTS / "check_office_scripts_compat.cjs"), TSC,
                   str(TESTS / "excelscript_api.d.ts"), *[str(SCRIPTS / f"{name}.ts") for name in NAMES]]
        result = subprocess.run(command, text=True, capture_output=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(json.loads(result.stdout), [])


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet")
class FilterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temporary = tempfile.TemporaryDirectory(prefix="ressourceplan-364-")
        cls.addClassCleanup(temporary.cleanup)
        cls.folder = Path(temporary.name)
        for version, names in ((SCRIPTS, NAMES), (BASELINE, ("add_project",))):
            destination = cls.folder / version.name
            destination.mkdir()
            for name in names:
                result = subprocess.run(
                    [TSC, "--strict", "--target", "ES2019", "--module", "none",
                     "--outFile", str(destination / f"{name}.js"),
                     str(TESTS / "excelscript_api.d.ts"), str(version / f"{name}.ts")],
                    text=True, capture_output=True,
                )
                if result.returncode:
                    raise AssertionError(result.stdout + result.stderr)
        cls.fixture = cls.folder / "fixture.json"
        cls.fixture.write_text(json.dumps(pack(load_workbook(ROOT / "Ressourceplan_3_6_3.xlsx"))))

    def run_scenario(self, name):
        result = subprocess.run(
            [NODE, str(TESTS / "office_scripts_364_runner.cjs"), str(self.folder), str(self.fixture), name],
            text=True, capture_output=True,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_filtered_insert_delete_conflict_regression(self):
        self.run_scenario("project")

    def test_standard_rows_through_both_entry_points(self):
        self.run_scenario("standard")

    def test_reuse_append_and_unfiltered_insert_do_not_change_filters(self):
        self.run_scenario("bypass")

    def test_unconfirmed_personal_view_or_criteria_prevents_filter_writes(self):
        self.run_scenario("view")

    def test_clear_failure_still_restores_attempted_columns(self):
        self.run_scenario("clear")

    def test_insert_and_verification_failures_still_restore_filters(self):
        self.run_scenario("insert")

    def test_restore_failure_preserves_other_criteria_and_reports_both_errors(self):
        self.run_scenario("restore")
