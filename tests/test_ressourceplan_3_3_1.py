"""Regressionskontrol af 3.3.1-pakken og beregning på separate Calc-kopier."""

import hashlib
import os
import re
import subprocess
import tempfile
import unittest
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter as col
from openpyxl.utils.cell import range_boundaries

from migrate_3_3_1 import migrate
from test_ressourceplan import BASELINE, CALC, RELEASE as PREVIOUS, ROOT

RELEASE = ROOT / "Ressourceplan_3_3_1.xlsx"
NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
COLORS = {
    "Sag": "D2EAF7", "Tilbud": "FCE3E3", "Ferie": "FFF4D9",
    "Fravær": "FFF4D9", "Kursus": "E3F0DF", "Andet": "EDE5F5",
}
QUARTERS = ("D2EAF7", "D5F0E9", "EFE1D6", "FCE3E3")
OCCUPANCY = ("EAF2FC", "DFF2E9", "FFF1CF", "F9DADD")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def conditional_backgrounds(pkg, sheet, area):
    nodes = pkg[sheet].findall(f"s:conditionalFormatting[@sqref='{area}']/s:cfRule", NS)
    return [
        pkg[f"xl/styles.xml"].findall("s:dxfs/s:dxf", NS)[int(node.attrib["dxfId"])]
        .find("s:fill/s:patternFill", NS)
        for node in nodes
    ]


class PatchStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = load_workbook(PREVIOUS)
        cls.release = load_workbook(RELEASE)

    def test_source_and_all_input_cells_preserved(self):
        self.assertEqual(digest(BASELINE),
                         "7eb4f9e1886c98c754a1f6c48f59f78bce1572fa5a4666ed3cf84d6ca82c9959")
        self.assertEqual(digest(PREVIOUS),
                         "ccdc0770a50d36250178479157484ff069721420c3ddbad02e50bd9df2105f81")
        self.assertEqual(self.source.sheetnames, self.release.sheetnames)
        for name in ("Min ressourceplan", "Projektregister", "Medarbejdere"):
            old, new = self.source[name], self.release[name]
            for row in old.iter_rows(min_row=2):
                for cell in row:
                    self.assertEqual(cell.value, new[cell.coordinate].value, f"{name}!{cell.coordinate}")
        for sheet in ("Ressourceoverblik", "Projektoversigt"):
            for addr in ("B5", "D5", "F5"):
                self.assertEqual(self.source[sheet][addr].value, self.release[sheet][addr].value)
        self.assertEqual(self.source["Projektoversigt"]["I5"].value,
                         self.release["Projektoversigt"]["B2"].value)
        plan = self.release.worksheets[0]
        self.assertEqual([r for r in range(10, 2170) if not plan.row_dimensions[r].hidden],
                         list(range(10, 22)))

    def test_excel_dxf_xml_and_priorities(self):
        with ZipFile(RELEASE) as archive:
            self.assertIsNone(archive.testzip())
            parts = {name: ET.fromstring(archive.read(name)) for name in archive.namelist()
                     if name.endswith((".xml", ".rels"))}
            self.assertFalse(any(re.search(r"vba|externalLinks|connections", name, re.I)
                                 for name in archive.namelist()))
        for root in parts.values():
            self.assertFalse(any(node.get("TargetMode") == "External" for node in root.iter()))

        styles = parts["xl/styles.xml"]
        counts = {}
        for collection in ("numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs"):
            node = styles.find(f"s:{collection}", NS)
            counts[collection] = len(node)
            self.assertEqual(int(node.get("count")), counts[collection])
        for xf in styles.findall("s:cellXfs/s:xf", NS):
            for attribute, collection in (("fontId", "fonts"), ("fillId", "fills"),
                                          ("borderId", "borders"), ("xfId", "cellStyleXfs")):
                self.assertLess(int(xf.get(attribute, "0")), counts[collection])

        worksheets = [f"xl/worksheets/sheet{i}.xml" for i in range(1, 6)]
        for part in worksheets:
            sheet = parts[part]
            rules = sheet.findall("s:conditionalFormatting/s:cfRule", NS)
            priorities = [int(rule.attrib["priority"]) for rule in rules]
            self.assertEqual(len(priorities), len(set(priorities)), part)
            for rule in rules:
                self.assertLess(int(rule.get("dxfId")), counts["dxfs"])
            for cell in sheet.findall("s:sheetData/s:row/s:c", NS):
                self.assertLess(int(cell.get("s", "0")), counts["cellXfs"])
                self.assertNotEqual(cell.get("t"), "e", f"{part}!{cell.get('r')}")
            for formula in sheet.findall("s:sheetData/s:row/s:c/s:f", NS):
                self.assertNotIn("#REF!", formula.text)
                self.assertLessEqual(len(formula.text), 8192)

        plan, overview, projects = (worksheets[i] for i in (0, 1, 3))
        for part, area, expected in [
            (plan, "A10:BI2169",
             [COLORS[key] for key in ("Tilbud", "Ferie", "Fravær", "Kursus", "Andet", "Sag")] + ["FFFFFF"]),
            (plan, "I7:BI7", QUARTERS),
            (overview, "H7:BH7", QUARTERS),
            (overview, "H9:BH9", ("F04E23",)),
            (overview, "H10:BH40", OCCUPANCY),
            (projects, "G7:BG7", QUARTERS),
        ]:
            fills = conditional_backgrounds(parts, part, area)
            self.assertEqual(len(fills), len(expected))
            for pattern, color in zip(fills, expected):
                self.assertEqual(pattern.attrib, {})
                self.assertIsNone(pattern.find("s:fgColor", NS))
                self.assertEqual(pattern.find("s:bgColor", NS).attrib, {"rgb": "FF" + color})

        table_style = parts["xl/tables/table1.xml"].find("s:tableStyleInfo", NS)
        self.assertIsNone(table_style.get("name"))
        self.assertEqual(table_style.get("showRowStripes"), "0")
        plan_sheet = self.release.worksheets[0]
        number_rule = plan_sheet.conditional_formatting["H10:H2169"][0]
        self.assertFalse(number_rule.stopIfTrue)
        self.assertEqual(number_rule.dxf.numFmt.formatCode, ";;;")
        for rule in plan_sheet.conditional_formatting["A10:BI2169"]:
            self.assertGreater(rule.priority, number_rule.priority)
            self.assertTrue(rule.stopIfTrue)
        for sheet in self.release:
            for table in sheet.tables.values():
                start, row, end, _ = range_boundaries(table.ref)
                self.assertEqual(table.autoFilter.ref, table.ref)
                self.assertEqual([c.name for c in table.tableColumns],
                                 [sheet.cell(row, c).value for c in range(start, end + 1)])

        views = parts["xl/workbook.xml"].findall("s:bookViews/s:workbookView", NS)
        self.assertEqual(len(views), 1)
        self.assertEqual(views[0].get("activeTab"), "0")
        for path in worksheets:
            view = parts[path].find("s:sheetViews/s:sheetView", NS)
            self.assertEqual(view.get("workbookViewId"), "0")
            pane = view.find("s:pane", NS)
            if pane is not None:
                self.assertIsNotNone(view.find(f"s:selection[@pane='{pane.get('activePane')}']", NS))
            for table in parts[path].findall("s:tableParts/s:tablePart", NS):
                relpath = path.replace("worksheets/", "worksheets/_rels/") + ".rels"
                self.assertTrue(any(rel.get("Id") == table.get(f"{{{REL_NS}}}id") for rel in parts[relpath]))

    def test_usable_views_and_rebuild(self):
        plan, overview, _, projects, _ = self.release.worksheets
        self.assertTrue(plan.column_dimensions["C"].hidden)
        self.assertEqual(plan.freeze_panes, "I10")
        self.assertEqual([(rule.colId, rule.filters.filter)
                          for rule in plan.tables["PlanTimer"].autoFilter.filterColumn],
                         [(0, ["2026"]), (1, ["Medarbejder 1"])])
        self.assertEqual(projects.freeze_panes, "G10")
        self.assertTrue(any("B2" in d and d.formula1 == "Sagsnumre"
                            for d in projects.data_validations.dataValidation))
        self.assertEqual(projects["F9"].value, "Periodetimer (t)")
        self.assertEqual(overview["H40"].number_format, "0%")
        self.assertTrue(all(overview.row_dimensions[row].hidden for row in range(41, 250)))
        self.assertTrue(all(projects.row_dimensions[row].hidden for row in range(41, 75)))
        self.assertGreaterEqual(overview.column_dimensions["G"].width, 23)
        self.assertGreaterEqual(overview.row_dimensions[9].height, 30)
        with tempfile.TemporaryDirectory() as directory:
            rebuilt = Path(directory) / RELEASE.name
            migrate(PREVIOUS, rebuilt)
            with self.assertRaises(ValueError):
                migrate(PREVIOUS, rebuilt)
            with ZipFile(RELEASE) as released, ZipFile(rebuilt) as generated:
                self.assertEqual(released.namelist(), generated.namelist())
                for name in released.namelist():
                    if name != "docProps/core.xml":
                        a, b = released.read(name), generated.read(name)
                        if name.endswith((".xml", ".rels")):
                            a, b = ET.canonicalize(a), ET.canonicalize(b)
                        self.assertEqual(a, b, name)


def add_probes(sheet, area, targets, first_probe_col):
    """Evaluate copied CF expressions with the same relative references as Excel."""
    left, top, _, _ = range_boundaries(area)
    origin = f"{col(left)}{top}"
    rules = sheet.conditional_formatting[area]
    for ordinal, address in enumerate(targets):
        row = sheet[address].row
        for offset, rule in enumerate(rules):
            formula = Translator("=" + rule.formula[0], origin=origin).translate_formula(address)
            color = rule.dxf.fill.bgColor.rgb[2:] if rule.dxf.fill else "skjul"
            sheet.cell(row, first_probe_col + ordinal * len(rules) + offset,
                       f'=IF({formula[1:]},"{color}","")')


@unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
class PatchCalculationTests(unittest.TestCase):
    choices = {
        "uge48": (2026, 48, "4 uger"), "tilbud": (2026, 48, "4 uger"),
        "sag_igen": (2027, 50, "8 uger"), "tolv": (2028, 1, "12 uger"),
        "seksogtyve": (2026, 48, "26 uger"), "q1": (2027, 48, "Q1"),
        "q2": (2028, 48, "Q2"), "q3": (2026, 48, "Q3"), "q4": (2026, 48, "Q4"),
        "helt53": (2026, 1, "Hele året"), "helt52": (2027, 1, "Hele året"),
    }

    @classmethod
    def setUpClass(cls):
        tmp = tempfile.TemporaryDirectory(prefix="ressourceplan-331-")
        cls.addClassCleanup(tmp.cleanup)
        cls.directory = Path(tmp.name)
        cls.output = cls.directory / "beregnet"
        cls.output.mkdir()
        inputs = []
        for name, choice in cls.choices.items():
            source = cls.directory / f"{name}.xlsx"
            cls.scenario(name, choice).save(source)
            inputs.append(source)
        process = subprocess.run(
            [CALC, f"-env:UserInstallation={(cls.directory / 'profile').as_uri()}",
             "--headless", "--convert-to", "xlsx", "--outdir", str(cls.output), *map(str, inputs)],
            capture_output=True, text=True, check=True, timeout=300,
            env=dict(os.environ, XDG_CACHE_HOME=str(cls.directory / "cache")),
        )
        for source in inputs:
            if not (cls.output / source.name).exists():
                raise AssertionError(f"Calc gemte ikke {source.name}\n{process.stdout}\n{process.stderr}")

    @staticmethod
    def scenario(name, choice):
        book = load_workbook(RELEASE)
        plan, overview, registry, project, people = book.worksheets
        project["B5"], project["D5"], project["F5"] = choice
        people["F7"], people["F8"] = 50, 0
        for week, hours in enumerate((0, 25, 42.5, 50, 59), 1):
            plan.cell(10, week + 8, hours)
        for row, kind, hours in ((11, "Tilbud", 10), (12, "Ferie", 7),
                                 (13, "Fravær", 10), (14, "Kursus", 8), (15, "Andet", 10)):
            plan[f"D{row}"], plan[f"N{row}"] = kind, hours
        plan["D19"], plan["H10"] = "Sag", 0.5
        plan["I82"], plan["BD82"] = 6, 11
        plan["E154"], plan["BD154"], plan["BE154"] = "24015", 8, 9
        plan["E226"], plan["R226"] = "24015", 6
        plan["BI10"], plan["BI22"], plan["BH22"], plan["I34"] = 3, 70, 5, 12
        registry["B7"] = "Testprojekt A"
        if name in ("tilbud", "sag_igen"):
            plan["D10"] = "Tilbud"
            if name == "sag_igen":
                plan["D10"] = "Sag"
            else:
                project["B2"] = "24048"
        if name in ("uge48", "tilbud", "sag_igen"):
            for row in (10, 11, 12, 13, 14, 15, 19):
                add_probes(plan, "A10:BI2169", [f"A{row}"], 67)
                add_probes(plan, "H10:H2169", [f"H{row}"], 74)
            add_probes(overview, "H10:BH40", [f"{c}10" for c in "HIJKL"] + ["H11"], 63)
            add_probes(overview, "H9:BH9", ["BB9", "BC9", "BF9", "BG9"], 63)
            add_probes(project, "G7:BG7", ["G7"], 63)
        return book

    def read_case(self, name):
        return load_workbook(self.output / f"{name}.xlsx", data_only=True)

    def test_types_switches_heat_and_orange_predicates(self):
        expected_rows = {11: "Tilbud", 12: "Ferie", 13: "Fravær", 14: "Kursus", 15: "Andet", 19: "Sag"}
        for name in ("uge48", "tilbud", "sag_igen"):
            plan, overview, _, _, _ = self.read_case(name).worksheets
            for row, kind in expected_rows.items():
                self.assertEqual([c.value for c in plan[row][66:73] if c.value], [COLORS[kind]])
            if name == "tilbud":
                self.assertEqual(plan.cell(10, 67).value, COLORS["Tilbud"])
                self.assertIsNone(plan.cell(10, 74).value)
                self.assertEqual(overview["I45"].value, 0)
                self.assertEqual(overview["I150"].value, 12.5)
            else:
                self.assertEqual([c.value for c in plan[10][66:73] if c.value], ["FFFFFF"])
                self.assertEqual(plan.cell(10, 74).value, "skjul")
                self.assertEqual(overview["I45"].value, 25)
                self.assertEqual(overview["I150"].value, 0)
        overview = self.read_case("uge48")["Ressourceoverblik"]
        for idx, (week, ratio, expected) in enumerate(zip(
            "HIJKL", (0, 0.5, 0.85, 1, 1.18),
            ("EAF2FC", "DFF2E9", "FFF1CF", "FFF1CF", "F9DADD"),
        )):
            self.assertAlmostEqual(overview[f"{week}10"].value, ratio)
            self.assertEqual([c.value for c in overview[10][62+idx*4:66+idx*4] if c.value], [expected])
        self.assertEqual(overview["H11"].value, "Over")
        self.assertEqual([c.value for c in overview[11][82:86] if c.value], ["F9DADD"])
        self.assertEqual(overview["M45"].value, 10)  # Andet
        self.assertEqual(overview["M80"].value, 25)  # Ferie, Fravær og Kursus
        self.assertEqual(overview["M115"].value, 10)
        self.assertAlmostEqual(overview["BC40"].value,
                               sum(overview[f"BC{r}"].value for r in range(45, 75)) /
                               sum(overview[f"BC{r}"].value for r in range(80, 110)))
        self.assertEqual([overview.cell(9, c).value for c in range(63, 67)],
                         [None, "F04E23", "F04E23", None])

    def test_projects_people_and_three_total_levels(self):
        book = self.read_case("uge48")
        project = book["Projektoversigt"]
        self.assertEqual(project["C2"].value, "Testprojekt A")
        self.assertEqual([project[f"A{r}"].value for r in range(10, 13)],
                         ["Medarbejder 1", "Medarbejder 3", None])
        self.assertEqual([project[f"F{r}"].value for r in (10, 11, 40)], [20, 17, 37])
        self.assertEqual([project[f"{c}40"].value for c in "GHIJK"], [28, 9, 0, 0, None])
        self.assertEqual(project.cell(7, 66).value, QUARTERS[3])
        alternate = self.read_case("tilbud")["Projektoversigt"]
        self.assertEqual(alternate["A10"].value, "Medarbejder 2")
        self.assertIsNone(alternate["A11"].value)
        self.assertEqual(alternate["F40"].value, 11)
        across = self.read_case("sag_igen")["Projektoversigt"]
        self.assertEqual([across[f"{c}9"].value for c in "GHIJKLMN"], [50, 51, 52, 1, 2, 3, 4, 5])
        self.assertEqual((across["I40"].value, across["J40"].value, across["F40"].value), (5, 12, 17))
        no_hours = self.read_case("q2")["Projektoversigt"]
        self.assertEqual(no_hours["A7"].value, "Ingen projekttimer i perioden")
        self.assertIsNone(no_hours["A10"].value)

    def test_periods_iso_mondays_quarters_and_no_formula_errors(self):
        for name, (year, start, period) in self.choices.items():
            with self.subTest(name=name):
                book = self.read_case(name)
                for sheet in book:
                    self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"],
                                     sheet.title)
                project = book["Projektoversigt"]
                weeks = [(year, w) for w in range(1, date(year, 12, 28).isocalendar().week + 1)]
                if period.startswith("Q"):
                    weeks = [(y, w) for y, w in weeks
                             if (date.fromisocalendar(y, w, 1).month - 1) // 3 + 1 == int(period[1:])]
                elif period != "Hele året":
                    n = int(period.split()[0])
                    weeks = (weeks[start-1:] + [
                        (year + 1, w) for w in range(1, date(year + 1, 12, 28).isocalendar().week + 1)
                    ])[:n]
                self.assertEqual(project["F6"].value, len(weeks))
                for slot in range(53):
                    letter = col(7 + slot)
                    if slot >= len(weeks):
                        self.assertEqual([project[f"{letter}{r}"].value for r in (7, 8, 9, 10, 40)],
                                         [None] * 5)
                        continue
                    y, w = weeks[slot]
                    monday = date.fromisocalendar(y, w, 1)
                    self.assertEqual(project[f"{letter}9"].value, w)
                    self.assertEqual(project[f"{letter}8"].value.date(), monday)
                    self.assertEqual(project[f"{letter}7"].value, f"Q{(monday.month - 1)//3+1}")
        self.assertEqual(self.read_case("helt53")["Projektoversigt"]["BG40"].value, 3)
        self.assertIsNone(self.read_case("helt52")["Projektoversigt"]["BG40"].value)

    def test_calc_export_is_separate_from_excel_delivery(self):
        source = self.directory / "uge48.xlsx"
        exported = self.output / source.name
        self.assertNotEqual(digest(source), digest(exported))
        with ZipFile(source) as excel, ZipFile(exported) as calc:
            self.assertNotEqual(excel.read("xl/styles.xml"), calc.read("xl/styles.xml"))
            excel_tables, calc_tables = [
                {ET.fromstring(package.read(part)).get("name"): part for part in package.namelist()
                 if part.startswith("xl/tables/")}
                for package in (excel, calc)
            ]
            self.assertEqual(excel_tables.keys(), calc_tables.keys())
            for package, parts in ((excel, excel_tables), (calc, calc_tables)):
                plan = ET.fromstring(package.read(parts["PlanTimer"]))
                self.assertEqual(plan.get("ref"), "A9:BI2169")
        self.assertEqual(load_workbook(RELEASE).worksheets[0].tables["PlanTimer"].ref, "A9:BI2169")


if __name__ == "__main__":
    unittest.main()
