"""Regressionstests for 3.4-visningerne og den uændrede beregningsmodel."""

import os
import subprocess
import tempfile
import unittest
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook

from migrate_3_4 import FIRST, LAST, migrate
from test_ressourceplan import CALC, ROOT
from test_ressourceplan_3_3_1 import NS, RELEASE as STABLE, digest

RELEASE = ROOT / "Ressourceplan_3_4.xlsx"


class LayoutStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = load_workbook(STABLE)
        cls.book = load_workbook(RELEASE)

    def test_plan_and_registers_preserve_inputs_and_metadata(self):
        self.assertEqual(digest(STABLE), "b2d89c00c344b16c56c5c0d7e40420ab144e77109cc03fd7184107cb1a2fd8b5")
        for name, first, last, columns in (
            ("Min ressourceplan", 10, 2169, 63),
            ("Projektregister", 7, 106, 2),
            ("Medarbejdere", 7, 36, 6),
        ):
            old, updated = self.source[name], self.book[name]
            for row in old.iter_rows(min_row=first, max_row=last, max_col=columns):
                for cell in row:
                    self.assertEqual(updated[cell.coordinate].value, cell.value, f"{name}!{cell.coordinate}")
            for table in old.tables.values():
                self.assertEqual(updated.tables[table.name].ref, table.ref)
                self.assertEqual(updated.tables[table.name].autoFilter, table.autoFilter)

        plan = self.book["Min ressourceplan"]
        self.assertEqual(plan.freeze_panes, "I10")
        self.assertTrue(plan.column_dimensions["C"].hidden)
        self.assertEqual([r for r in range(10, 2170) if not plan.row_dimensions[r].hidden],
                         list(range(10, 22)))
        self.assertEqual(plan.data_validations, self.source["Min ressourceplan"].data_validations)
        self.assertTrue(any(cell.value is None and cell.data_type != "f" for cell in plan["I"][9:21]))

    def test_dropdowns_and_model_are_separate_from_the_public_layout(self):
        self.assertEqual(self.source.sheetnames, self.book.sheetnames)
        for sheet in self.book:
            self.assertFalse(sheet.sheet_view.showGridLines)
            self.assertEqual(sheet["A1"].fill.fgColor.rgb, "FF1C2B54")
            self.assertEqual(sheet["A3"].font.name, "Segoe UI")
        for name, end, pane, inputs in (
            ("Ressourceoverblik", 309, "H15", ("A6", "B6", "C6", "D6", "E6", "F6")),
            ("Projektoversigt", 134, "G15", ("A6", "B6", "C6", "D6")),
        ):
            sheet = self.book[name]
            self.assertEqual(sheet.freeze_panes, pane)
            self.assertTrue(all(sheet.row_dimensions[r].hidden for r in range(61, end + 1)))
            self.assertTrue(all(sheet.column_dimensions[c].hidden for c in ("BI", "BJ", "BK")))
            self.assertIsNone(sheet["A46"].value)
            for cell in inputs:
                self.assertTrue(any(cell in d and not d.showDropDown
                                    for d in sheet.data_validations.dataValidation))
                self.assertEqual(sheet[cell].fill.fgColor.rgb, "FFEAF2FC")
                self.assertEqual(sheet[cell].border.bottom.style, "thin")
                self.assertIn("▾", sheet[cell.replace("6", "5")].value)
                self.assertNotEqual(sheet[cell].data_type, "f")
            visible = {c.value for row in sheet.iter_rows(max_row=LAST) for c in row if c.data_type != "f"}
            for title in ("Fra uge", "Uger i året", "Uger valgt"):
                self.assertNotIn(title, visible)
        registry = self.book["Projektregister"]
        self.assertTrue(all(registry.column_dimensions[c].hidden for c in "JKLM"))
        report = self.book["Projektoversigt"]
        rule = next(d for d in report.data_validations.dataValidation if "A6" in d)
        self.assertEqual(rule.formula1, "Projektvisninger")
        self.assertIn("INDEX(", self.book.defined_names["Projektvisninger"].attr_text)
        self.assertEqual(report["A6"].value, "Hospital Nord · 24015")

    def test_excel_xml_dxf_and_rebuild(self):
        with ZipFile(RELEASE) as package:
            self.assertIsNone(package.testzip())
            self.assertFalse(any(k in n.lower() for k in ("vba", "externallinks", "connections")
                                 for n in package.namelist()))
            parts = {n: ET.fromstring(package.read(n)) for n in package.namelist()
                     if n.endswith((".xml", ".rels"))}
        for root in parts.values():
            self.assertFalse(any(el.attrib.get("TargetMode") == "External" for el in root.iter()))
        dxfs = parts["xl/styles.xml"].findall("s:dxfs/s:dxf", NS)

        for number in range(1, 6):
            sheet = parts[f"xl/worksheets/sheet{number}.xml"]
            rules = sheet.findall("s:conditionalFormatting/s:cfRule", NS)
            self.assertEqual(len({r.get("priority") for r in rules}), len(rules))
            for rule in rules:
                pattern = dxfs[int(rule.get("dxfId"))].find("s:fill/s:patternFill", NS)
                if pattern is not None:
                    self.assertEqual(pattern.attrib, {})
                    self.assertIsNone(pattern.find("s:fgColor", NS))
                    self.assertIsNotNone(pattern.find("s:bgColor", NS))
            for f in sheet.findall(".//s:c/s:f", NS):
                self.assertNotIn("#REF!", f.text)
                self.assertNotIn("_xlfn.", f.text)
                self.assertNotEqual(f.get("t"), "array")
            self.assertIsNone(sheet.find(".//s:c[@t='e']", NS))

        for name, region, desired in (
            ("Ressourceoverblik", "A15:G45", ["D5F0E9", "FFFFFF", "D2EAF7"]),
            ("Ressourceoverblik", "H15:BH45", ["EAF2FC", "DFF2E9", "FFF1CF", "F9DADD"]),
            ("Projektoversigt", "A15:F45", ["D5F0E9", "FFFFFF", "D2EAF7"]),
            ("Projektoversigt", "G15:BG45", ["D5F0E9", "FFFFFF", "D2EAF7"]),
        ):
            rules = self.book[name].conditional_formatting[region]
            self.assertEqual([r.dxf.fill.bgColor.rgb[2:] for r in rules], desired)
            self.assertTrue(all(r.dxf.border.bottom.style == "thin" for r in rules))

        with tempfile.TemporaryDirectory() as directory:
            rebuilt = Path(directory) / RELEASE.name
            migrate(STABLE, rebuilt)
            with self.assertRaises(ValueError):
                migrate(STABLE, rebuilt)
            with ZipFile(RELEASE) as released, ZipFile(rebuilt) as generated:
                self.assertEqual(released.namelist(), generated.namelist())
                for name in released.namelist():
                    if name != "docProps/core.xml":
                        a, b = released.read(name), generated.read(name)
                        if name.endswith((".xml", ".rels")):
                            a, b = ET.canonicalize(a), ET.canonicalize(b)
                        self.assertEqual(a, b, name)


@unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
class CompactCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        tmp = tempfile.TemporaryDirectory(prefix="ressourceplan-34-")
        cls.addClassCleanup(tmp.cleanup)
        directory = Path(tmp.name)
        output = directory / "calc"
        output.mkdir()
        paths = []
        for scenario in ("spredt", "nul", "en", "alle", "aarsskifte"):
            book = cls.scenario(scenario)
            path = directory / f"{scenario}.xlsx"
            book.save(path)
            paths.append(path)
        process = subprocess.run(
            [CALC, f"-env:UserInstallation={(directory / 'profile').as_uri()}",
             "--headless", "--convert-to", "xlsx", "--outdir", str(output), *map(str, paths)],
            capture_output=True, check=True, timeout=250,
            env=dict(os.environ, XDG_CACHE_HOME=str(directory / "cache")),
        )
        cls.result = {}
        for path in paths:
            calculated = output / path.name
            if not calculated.is_file():
                raise AssertionError(process.stdout.decode() + process.stderr.decode())
            cls.result[path.stem] = load_workbook(calculated, data_only=True)

    @classmethod
    def scenario(cls, scenario):
        book = load_workbook(RELEASE)
        plan, resource, registry, reports, people = book.worksheets
        if scenario == "aarsskifte":
            reports["B6"], reports["C6"], reports["D6"] = 2027, 50, "8 uger"
            plan["BH22"], plan["BI22"], plan["I34"] = 5, 65, 12
            return book
        if scenario == "alle":
            for index in range(30):
                cls.member(people, plan, index + 1, index + 1, 37)
            return book

        for person, hours, norm in ((1, 20, 37), (5, 55, 40), (9, 10, 25)):
            cls.member(people, plan, person, hours, norm)
        resource["D6"], resource["E6"], resource["F6"] = "Øst", "EL", "QA.34"
        registry["B7"] = "Omdøbt · Hospital Nord"
        registry["A31"], registry["B31"] = "88001", "Nyt testprojekt"
        if scenario in ("nul", "en"):
            reports["A6"] = "Nyt testprojekt · 88001"
        if scenario == "nul":
            resource["D6"] = "Vest"
        if scenario == "en":
            people["D11"] = "Belysning"
            resource["E6"] = "Belysning"
            registry["A32"], registry["B32"] = "QA · 88001", "Omdøbt testprojekt"
            reports["A6"] = "Tidligere testprojekt · QA · 88001"
            plan["E587"], plan["BD587"] = "QA · 88001", 19
        return book

    @staticmethod
    def member(people, plan, person, hours, norm):
        employee, row = 6 + person, 10 + (person - 1) * 72
        for column, value in zip("BCDEF", (f"Testperson {person}", "Øst", "EL", "QA.34", norm)):
            people[f"{column}{employee}"] = value
        plan[f"E{row}"], plan[f"BD{row}"] = "24015", hours

    def assert_compact(self, sheet, active):
        values = [sheet[f"A{r}"].value for r in range(FIRST, LAST + 1)]
        self.assertEqual(values, [*active, "I ALT"] + [None] * (30 - len(active)))

    def test_nonconsecutive_matches_totals_and_current_project(self):
        _, overview, registry, project, _ = self.result["spredt"].worksheets
        for sheet in (overview, project):
            self.assert_compact(sheet, ["Testperson 1", "Testperson 5", "Testperson 9"])
            self.assertEqual([sheet[f"F{row}"].value for row in range(15, 19)], [20, 55, 10, 85])
        self.assertEqual([overview[f"{col}8"].value for col in "ABCDEFG"], [3, 85, 401, 331, 15, 10, 5])
        self.assertAlmostEqual(overview["H8"].value, 85 / 401)
        self.assertAlmostEqual(overview["BC18"].value, 85 / 95)
        self.assertEqual(project["A6"].value, "Hospital Nord · 24015")
        self.assertEqual(project["A8"].value, "Omdøbt · Hospital Nord · 24015")
        self.assertEqual(project["B62"].value, "24015")
        self.assertEqual(project["G18"].value, 85)
        self.assertEqual(registry["L7"].value, "Omdøbt · Hospital Nord · 24015")
        self.assertEqual(registry["L8"].value, "Datacenter · 24022")
        self.assertEqual(registry["L12"].value, "Nyt testprojekt · 88001")
        self.assertIsNone(registry["L13"].value)

    def test_zero_one_and_all_slots(self):
        for scenario, resource_names, project_names in (
            ("nul", [], []),
            ("en", ["Testperson 5"], ["Testperson 9"]),
            ("alle", [f"Testperson {p}" for p in range(1, 31)], [f"Testperson {p}" for p in range(1, 31)]),
        ):
            with self.subTest(scenario=scenario):
                resource = self.result[scenario]["Ressourceoverblik"]
                project = self.result[scenario]["Projektoversigt"]
                self.assert_compact(resource, resource_names)
                self.assert_compact(project, project_names)
                self.assertEqual(resource["A8"].value, len(resource_names))
                self.assertEqual(project["B8"].value, len(project_names))
                expected_resource, expected_project = {"nul": (0, 0), "en": (55, 19), "alle": (465, 465)}[scenario]
                self.assertEqual(resource[f"F{FIRST + len(resource_names)}"].value, expected_resource)
                self.assertEqual(project[f"F{FIRST + len(project_names)}"].value, expected_project)
                if scenario == "en":
                    self.assertEqual(project["B62"].value, "QA · 88001")
                    self.assertEqual(project["A8"].value, "Omdøbt testprojekt · QA · 88001")
                    self.assertEqual(project["G15"].value, 19)

    def test_calendar_reference_and_recalculation(self):
        default = self.result["spredt"]["Projektoversigt"]
        self.assertEqual([default[f"{c}14"].value for c in "GHIJK"], [48, 49, 50, 51, None])
        report = self.result["aarsskifte"]["Projektoversigt"]
        self.assertEqual([report[f"{c}14"].value for c in "GHIJKLMN"], [50, 51, 52, 1, 2, 3, 4, 5])
        self.assertEqual(report["J13"].value.date(), date.fromisocalendar(2028, 1, 1))
        self.assertEqual((report["I16"].value, report["J16"].value, report["F16"].value), (5, 12, 17))
        for name, book in self.result.items():
            for sheet in book:
                with self.subTest(scenario=name, sheet=sheet.title):
                    self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"])


if __name__ == "__main__":
    unittest.main()
