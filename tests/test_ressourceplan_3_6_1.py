"""Kvalitetssikring for 3.6.1 og den selvstændige outline-prototype."""

import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter as col

from migrate_3_6_1 import migrate
from prototype_outline_3_6_1 import build, quarter_segments
from test_ressourceplan import CALC
from test_ressourceplan_3_3_1 import digest
from test_ressourceplan_3_6 import NODE, NAMES, ROOT, TESTS, TSC, pack, replay

PREVIOUS = ROOT / "Ressourceplan_3_6.xlsx"
RELEASE = ROOT / "Ressourceplan_3_6_1.xlsx"
PROTOTYPE = ROOT / "Ressourceplan_3_6_1_Gruppeprototype.xlsx"
SCRIPTS = ROOT / "office-scripts" / "3.6.1"


def headers(sheet):
    return {cell.value: cell.column for cell in sheet[9] if cell.value}


def employee_summary(sheet, name):
    return next(row for row in range(10, sheet.max_row + 1) if sheet[f"A{row}"].value == name)


def detail(sheet, summary, label):
    return next(row for row in range(summary + 1, sheet.max_row + 1)
                if sheet[f"E{row}"].value == label)


class ReleaseStructureTests(unittest.TestCase):
    def test_old_file_and_production_inputs_calc_and_filters_preserved(self):
        self.assertEqual(digest(PREVIOUS), "adb100331c7ae3ee3101a7be497fa06b8ddf8444a2cb3eaece913e8e4a4631bc")
        old, new = load_workbook(PREVIOUS), load_workbook(RELEASE)
        self.assertEqual(new.sheetnames, old.sheetnames)
        self.assertEqual(new["Planstart"].sheet_state, "hidden")
        for before, after in zip(old, new):
            for row in before.iter_rows(min_row=5):
                for cell in row:
                    if before.title == "Planstart" and cell.coordinate == "P6":
                        self.assertEqual(after[cell.coordinate].value, "PlanTimer-formler 3.6.1")
                    else:
                        self.assertEqual(after[cell.coordinate].value, cell.value)
            self.assertEqual(after.data_validations, before.data_validations)
            for table in before.tables.values():
                self.assertEqual(after.tables[table.name], table)
        sheet = new["Min ressourceplan"]
        self.assertIn("PlanSidsteRaekke", sheet.conditional_formatting["BJ10:BJ20009"][0].formula[0])
        self.assertTrue(all(d.outlineLevel == 0 for d in sheet.column_dimensions.values()))
        with ZipFile(RELEASE) as package:
            self.assertIsNone(package.testzip())
            self.assertFalse(any(token in part.lower() for part in package.namelist()
                                 for token in ("vba", "externallinks", "connections")))
            for part in package.namelist():
                if part.endswith((".xml", ".rels")):
                    ET.fromstring(package.read(part))

    def test_build_is_reproducible_and_scripts_match_standalone_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            dest = Path(tmp) / RELEASE.name
            migrate(PREVIOUS, dest)
            with self.assertRaises(ValueError):
                migrate(PREVIOUS, dest)
            with ZipFile(RELEASE) as original, ZipFile(dest) as rebuilt:
                self.assertEqual(original.namelist(), rebuilt.namelist())
                for entry in original.namelist():
                    if entry != "docProps/core.xml":
                        a, b = original.read(entry), rebuilt.read(entry)
                        if entry.endswith((".xml", ".rels")):
                            a, b = ET.canonicalize(a), ET.canonicalize(b)
                        self.assertEqual(a, b, entry)
        for script in NAMES:
            parts = [script, "common"]
            if script != "create_employee":
                parts.append("selection")
            if script != "add_project":
                parts.append("standard")
            expected = "\n".join((SCRIPTS / "src" / f"{name}.ts").read_text(encoding="utf-8") for name in parts)
            self.assertEqual((SCRIPTS / f"{script}.ts").read_text(encoding="utf-8").split("\n", 1)[1], expected)


class OutlineTests(unittest.TestCase):
    def test_quarter_ranges_years_and_unfoldable_metadata(self):
        first = [(y, q, len(days)) for y, q, days in quarter_segments(2026)]
        self.assertEqual(first, [(2025,4,1), (2026,1,13), (2026,2,13), (2026,3,13), (2026,4,13)])
        book = load_workbook(PROTOTYPE)
        for name in ("Kvartalsprototype", "Medarbejderprototype"):
            sheet = book[name]
            outline = [(d.min, d.max) for d in sheet.column_dimensions.values() if d.outlineLevel > 0]
            self.assertEqual([end-start+1 for start, end in outline], [1,13,13,13,13])
            self.assertTrue(all(start >= 7 and "i alt" in sheet.cell(9, end + 1).value
                                for start, end in outline))
            self.assertTrue(all(left[1] < right[0] - 1 for left, right in zip(outline, outline[1:])))
            self.assertEqual(sheet.freeze_panes, "G10")
        summary_sheet = book["Medarbejderprototype"]
        self.assertEqual(summary_sheet.cell(9, headers(summary_sheet)["Uge 48"]).fill.fgColor.rgb, "FFF04E23")
        self.assertEqual(summary_sheet.cell(9, headers(summary_sheet)["Uge 47"]).fill.fgColor.rgb, "FF1C2B54")
        with tempfile.TemporaryDirectory() as directory:
            copy = Path(directory) / "vest.xlsx"
            build(RELEASE, copy, 2027, ("Vest", "", ""))
            test = load_workbook(copy)
            self.assertNotIn("Uge 53", headers(test["Kvartalsprototype"]))
            self.assertEqual(len([c for c in headers(test["Kvartalsprototype"]) if c.startswith("Uge ")]), 52)
            self.assertEqual([row[0].value for row in test["Medarbejderprototype"].iter_rows(min_row=10)
                              if row[4].value == "Belægning"], ["Medarbejder 4", "Medarbejder 5"])
            self.assertTrue(all(d.outlineLevel == 0 or d.hidden
                                for d in test["Medarbejderprototype"].row_dimensions.values()))

    def test_compact_membership_and_details_are_bound_to_stable_keys(self):
        book = load_workbook(PROTOTYPE)
        sheet = book["Medarbejderprototype"]
        summaries = [r for r in range(10, sheet.max_row+1) if sheet[f"E{r}"].value == "Belægning"]
        self.assertEqual(len(summaries), 5)
        for person, row in enumerate(summaries, 1):
            end = summaries[person] - 1 if person < len(summaries) else sheet.max_row - 1
            self.assertTrue(sheet.row_dimensions[row].collapsed)
            self.assertTrue(all(sheet.row_dimensions[i].hidden and sheet.row_dimensions[i].outlineLevel == 1
                                for i in range(row+1, end + 1)))
            week = col(headers(sheet)["Uge 48"])
            booked = detail(sheet, row, "Bookede timer")
            self.assertIn(f'p{person:02}', sheet[f"{week}{booked}"].value)
        self.assertEqual(sheet[f"A{sheet.max_row}"].value, "I alt for udsnittet")


@unittest.skipUnless(TSC and NODE, "TypeScript og Node er påkrævet til Office Scripts-test")
class ScriptAndCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temp = tempfile.TemporaryDirectory(prefix="ressourceplan-361-")
        cls.addClassCleanup(temp.cleanup)
        cls.dir = Path(temp.name)
        compiled = cls.dir / "compiled"
        compiled.mkdir()
        for script in NAMES:
            subprocess.run([TSC, "--strict", "--target", "ES2019", "--module", "none",
                            "--outFile", str(compiled / f"{script}.js"),
                            str(TESTS / "excelscript_api.d.ts"), str(SCRIPTS / f"{script}.ts")],
                           capture_output=True, text=True, check=True)
        fixture, result = cls.dir / "fixture.json", cls.dir / "result.json"
        source = load_workbook(RELEASE)
        fixture.write_text(json.dumps(pack(source)))
        subprocess.run([NODE, str(TESTS / "office_scripts_361_runner.cjs"), str(compiled),
                        str(fixture), str(result)], check=True, text=True)
        cls.changed = json.loads(result.read_text())

    def test_profile_proposals_repeated_initialization_and_offers(self):
        rows = self.changed["tables"]["PlanTimer"]["rows"]
        staff = self.changed["tables"]["PersonRegister"]["rows"]
        self.assertEqual(staff[5][0:2], ["p06", "Testperson 361"])
        new = [row for row in rows if row[1] == "Testperson 361 · p06"]
        self.assertEqual([row[3] for row in new if row[0] == 2032],
                         ["Sag", "Ferie", "Fravær", "Kursus", "Andet"])
        self.assertEqual([row[7] for row in new if row[3] == "Tilbud"], ["", 0, 0.25, 0.5, 1])
        self.assertTrue(all(row[7] == "" for row in new if row[3] == "Sag"))

    @unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
    def test_recalculate_project_writes_and_outline_with_filters(self):
        workbook = load_workbook(RELEASE)
        replay(workbook, self.changed)
        changed = self.dir / "scriptresultat.xlsx"
        workbook.save(changed)

        prototype = load_workbook(PROTOTYPE)
        # Fold the group containing a week with recorded hours in both demos.
        for name in ("Kvartalsprototype", "Medarbejderprototype"):
            view = prototype[name]
            week = headers(view)["Uge 48"]
            group = next(d for d in view.column_dimensions.values()
                         if d.outlineLevel and d.min <= week <= d.max)
            group.hidden = True
            view.column_dimensions[col(group.max + 1)].collapsed = True
        demo = prototype["Kvartalsprototype"]
        demo.tables["Kvartalsprøve"].autoFilter.add_filter_column(1, ["Medarbejder 2 · p02"])
        for row in range(10, demo.max_row + 1):
            demo.row_dimensions[row].hidden = demo[f"B{row}"].value != "Medarbejder 2 · p02"
        plan = prototype["Min ressourceplan"]
        plan.tables["PlanTimer"].autoFilter.filterColumn[1].filters.filter = ["Medarbejder 2 · p02"]
        for row in range(10, plan.max_row + 1):
            plan.row_dimensions[row].hidden = (plan[f"A{row}"].value, plan[f"B{row}"].value) != (2026, "Medarbejder 2 · p02")
        outlined = self.dir / "foldet.xlsx"
        prototype.save(outlined)

        output = self.dir / "beregnet"
        output.mkdir()
        profile = (self.dir / "profile").as_uri()
        run = subprocess.run([CALC, f"-env:UserInstallation={profile}", "--headless", "--convert-to",
                              "xlsx", "--outdir", str(output), str(changed), str(outlined)],
                             check=True, capture_output=True, text=True, timeout=160,
                             env=dict(os.environ, XDG_CACHE_HOME=str(self.dir / "cache")))
        self.assertTrue((output / changed.name).is_file() and (output / outlined.name).is_file(), run.stderr)
        new, group = (load_workbook(output / p.name, data_only=True) for p in (changed, outlined))
        totals = new["Ressourceoverblik"]
        self.assertEqual([totals[f"{c}8"].value for c in "ABCDEFG"], [6, 28, 825, 797, 0, 50, 23])
        self.assertEqual(new["Projektoversigt"]["G17"].value, 36)
        self.assertEqual(group["Kvartalsprototype"]["F6"].value, 0)
        self.assertEqual(group["Min ressourceplan"]["G6"].value, 0)
        self.assertEqual(group["Ressourceoverblik"]["B8"].value, 20)
        detail_sheet = group["Medarbejderprototype"]
        week = col(headers(detail_sheet)["Uge 48"])
        p01 = employee_summary(detail_sheet, "Medarbejder 1")
        self.assertAlmostEqual(detail_sheet[f"{week}{p01}"].value, 2/3)
        self.assertEqual(detail_sheet[f"{week}{detail(detail_sheet, p01, 'Disponibel kapacitet')}"].value, 30)
        self.assertEqual(detail_sheet[f"{week}{detail(detail_sheet, p01, 'Bookede timer')}"].value, 20)
        self.assertEqual(detail_sheet["F6"].value, 20)
        for book in (new, group):
            for sheet in book:
                self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"],
                                 sheet.title)


if __name__ == "__main__":
    unittest.main()
