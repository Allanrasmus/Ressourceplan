"""Test Excel-valideringsregressionen og oprettelse af standardrækker."""

import os
import subprocess
import tempfile
import unittest
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries

from migrate_3_5_1 import ACTIVITIES, YEARS, YEAR_RANGE, append_standard_rows, initialise_existing, migrate
from test_ressourceplan import CALC, ROOT
from test_ressourceplan_3_3_1 import NS, digest
from test_ressourceplan_3_4 import RELEASE as ACCEPTED
from test_ressourceplan_3_5 import RELEASE as PREVIOUS

RELEASE = ROOT / "Ressourceplan_3_5_1.xlsx"


def xml_part(path, n):
    with ZipFile(path) as workbook:
        return ET.fromstring(workbook.read(n))


def validations(path, number):
    parent = xml_part(path, f"xl/worksheets/sheet{number}.xml").find("s:dataValidations", NS)
    return parent.findall("s:dataValidation", NS)


class RepairStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = load_workbook(PREVIOUS)
        cls.workbook = load_workbook(RELEASE)

    def test_specific_excel_validation_regression_on_sheet3_and_sheet5(self):
        accepted = validations(ACCEPTED, 3)[0]
        self.assertEqual(accepted.findtext("s:formula1", namespaces=NS),
                         'SUMPRODUCT(--($A$7:$A$106&""=A7&""))=1')
        for number, sheet, end, old_token in (
            (3, "Projektregister", 506, "SagRegister[Sagsnr.]"),
            (5, "Medarbejdere", 106, "PersonRegister[ID]"),
        ):
            self.assertEqual(self.workbook.sheetnames[number - 1], sheet)
            old = {r.get("sqref"): r for r in validations(PREVIOUS, number)}
            new = {r.get("sqref"): r for r in validations(RELEASE, number)}
            self.assertEqual(old.keys(), new.keys())
            key = f"A7:A{end}"
            self.assertIn(old_token, old[key].findtext("s:formula1", namespaces=NS))
            self.assertEqual(new[key].findtext("s:formula1", namespaces=NS),
                             f'SUMPRODUCT(--($A$7:$A${end}&""=A7&""))=1')
            self.assertEqual((new[key].get("type"), new[key].get("errorStyle"),
                              new[key].get("showErrorMessage")), ("custom", "stop", "1"))
            for area in new.keys() - {key}:
                self.assertEqual(ET.tostring(new[area]), ET.tostring(old[area]))
            regions = [range_boundaries(area) for area in new]
            for i, (left, top, right, bottom) in enumerate(regions):
                for other_left, other_top, other_right, other_bottom in regions[i+1:]:
                    self.assertTrue(right < other_left or other_right < left
                                    or bottom < other_top or other_bottom < top)

    def test_dropdowns_hidden_identity_filters_and_package(self):
        with ZipFile(RELEASE) as archive:
            self.assertIsNone(archive.testzip())
            self.assertFalse(any(s in n.lower() for n in archive.namelist()
                                 for s in ("vba", "externallinks", "connections")))
            for path in archive.namelist():
                if path.endswith((".xml", ".rels")):
                    part = ET.fromstring(archive.read(path))
                    self.assertFalse(any(n.get("TargetMode") == "External" for n in part.iter()))
                    self.assertFalse(any("#REF!" in (n.text or "") for n in part.findall(".//s:f", NS)))
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            self.assertIsNone(workbook.find("s:fileSharing", NS))
            self.assertIsNone(workbook.find("s:workbookProtection", NS))
        for sheet, area, source in (
            ("Min ressourceplan", "A10:A20009", YEAR_RANGE),
            ("Min ressourceplan", "B10:B20009", "Medarbejdervisninger"),
            ("Min ressourceplan", "E10:E20009", "Projektvisninger"),
            ("Ressourceoverblik", "A6", YEAR_RANGE), ("Projektoversigt", "B6", YEAR_RANGE),
            ("Planstart", "A6", YEAR_RANGE), ("Planstart", "B6", "Medarbejdervisninger"),
        ):
            rule = next(r for r in self.workbook[sheet].data_validations.dataValidation
                        if str(r.sqref) == area)
            self.assertEqual((rule.type, rule.formula1, rule.showDropDown), ("list", source, False))
            self.assertIn(source, self.workbook.defined_names)
        plan = self.workbook["Min ressourceplan"]
        table = plan.tables["PlanTimer"]
        self.assertEqual(table.ref, table.autoFilter.ref)
        self.assertEqual([(f.colId, f.filters.filter) for f in table.autoFilter.filterColumn],
                         [(0, ["2026"]), (1, ["Medarbejder 1 · p01"])])
        self.assertTrue(all(plan.column_dimensions[c].hidden for c in ("C", "I", "BK")))
        self.assertIsNone(table.tableColumns[1].calculatedColumnFormula)
        self.assertIn("[Medarbejder]", table.tableColumns[2].calculatedColumnFormula.attr_text)
        self.assertEqual(table.tableColumns[8].name, "Aktuelt medarbejdernavn")
        self.assertEqual(self.workbook["Projektregister"].tables["SagRegister"].ref, "A6:B506")
        self.assertEqual(self.workbook["Medarbejdere"].tables["PersonRegister"].ref, "A6:F106")

    def test_preserve_existing_records_seed_examples_and_repeat_without_duplicates(self):
        self.assertEqual(digest(PREVIOUS), "ee698ff30d5bca481f061214a1fc0bcc51c67179a13c8da6ad2855efa03a061b")
        plan, old = self.workbook["Min ressourceplan"], self.source["Min ressourceplan"]
        for row in range(10, old.max_row + 1):
            self.assertEqual(plan[f"B{row}"].value, old[f"I{row}"].value)
            for col in (1, 4, 5, 8, *range(10, 63)):
                self.assertEqual(plan.cell(row, col).value, old.cell(row, col).value)
        for name, cols in (("Projektregister", 2), ("Medarbejdere", 6)):
            for row in self.source[name].iter_rows(min_row=7, max_col=cols):
                for cell in row:
                    self.assertEqual(self.workbook[name][cell.coordinate].value, cell.value)
        counts = Counter()
        for row in plan.iter_rows(min_row=10, max_col=62, values_only=True):
            kind = row[3]
            if kind == "Sag" and row[4] is not None:
                continue
            counts[(row[0], row[1], kind)] += 1
        for person in range(1, 6):
            for year in YEARS:
                for kind in ACTIVITIES:
                    self.assertEqual(counts[(year, f"Medarbejder {person} · p{person:02}", kind)], 1)
        self.assertEqual(plan.max_row - old.max_row, 149)
        replica = load_workbook(RELEASE)
        staff = replica["Medarbejdere"].iter_rows(min_row=7, max_col=2, values_only=True)
        self.assertEqual(initialise_existing(replica["Min ressourceplan"], staff, YEARS), 0)
        self.assertEqual(replica["Min ressourceplan"].max_row, plan.max_row)

    def test_reproducible_generation_and_overwrite_guard(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / RELEASE.name
            migrate(PREVIOUS, target)
            with self.assertRaises(ValueError):
                migrate(PREVIOUS, target)
            with ZipFile(RELEASE) as released, ZipFile(target) as rebuilt:
                self.assertEqual(released.namelist(), rebuilt.namelist())
                for part in released.namelist():
                    if part == "docProps/core.xml":
                        continue
                    l, r = released.read(part), rebuilt.read(part)
                    if part.endswith((".xml", ".rels")):
                        l, r = ET.canonicalize(l), ET.canonicalize(r)
                    self.assertEqual(l, r, part)


def register_employee(workbook):
    staff = workbook["Medarbejdere"]
    for column, value in zip("BCDEF", ("Medarbejder 5", "Vest", "EL", "QA351", 30)):
        staff[f"{column}12"] = value


def chosen_standard_rows(workbook):
    init = workbook["Planstart"]
    return [(init[f"A{row}"].value, init[f"B{row}"].value, init[f"D{row}"].value)
            for row in range(10, 15) if init[f"D{row}"].value]


def set_filtered_plan(plan, year, employee):
    table = plan.tables["PlanTimer"]
    table.autoFilter.filterColumn[0].filters.filter = [str(year)]
    table.autoFilter.filterColumn[1].filters.filter = [employee]
    for row in range(10, plan.max_row + 1):
        plan.row_dimensions[row].hidden = (plan[f"A{row}"].value, plan[f"B{row}"].value) != (year, employee)


@unittest.skipUnless(CALC, "LibreOffice Calc er påkrævet til beregningstest")
class PlanStartCalculationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        temp = tempfile.TemporaryDirectory(prefix="ressourceplan-351-")
        cls.addClassCleanup(temp.cleanup)
        folder = Path(temp.name)
        output = folder / "beregnet"
        output.mkdir()

        def convert(paths):
            result = subprocess.run(
                [CALC, f"-env:UserInstallation={(folder / 'profile').as_uri()}",
                 "--headless", "--convert-to", "xlsx", "--outdir", str(output), *map(str, paths)],
                capture_output=True, text=True, check=True, timeout=240,
                env=dict(os.environ, XDG_CACHE_HOME=str(folder / "cache")),
            )
            if not all((output / p.name).exists() for p in paths):
                raise AssertionError(result.stdout + result.stderr)
            return {p.stem: load_workbook(output / p.name, data_only=True) for p in paths}

        inputs = []
        for name in ("reference", "ny-person", "nyt-aar", "delvis"):
            book = load_workbook(RELEASE)
            init = book["Planstart"]
            if name == "reference":
                init["B6"] = "Medarbejder 5 · p05"
                cls.validation_probes(book)
            elif name == "nyt-aar":
                init["F13"], init["A6"], init["B6"] = 2032, 2032, "Medarbejder 5 · p05"
                report = book["Projektoversigt"]
                report["B6"], report["C6"], report["D6"] = 2031, 52, "4 uger"
            else:
                register_employee(book)
                init["B6"] = "Medarbejder 5 · p06"
                if name == "delvis":
                    plan = book["Min ressourceplan"]
                    append_standard_rows(plan, [(2026, "Medarbejder 5 · p06", "Ferie"),
                                                 (2026, "Medarbejder 5 · p06", "Sag")])
                    plan[f"BE{plan.max_row}"] = 1   # Sag with hours is not a free project row.
            path = folder / f"{name}.xlsx"
            book.save(path)
            inputs.append(path)
        cls.results = convert(inputs)

        complete = load_workbook(RELEASE)
        register_employee(complete)
        plan, init = complete["Min ressourceplan"], complete["Planstart"]
        init["F13"] = 2032
        append_standard_rows(plan, chosen_standard_rows(cls.results["ny-person"]) +
                             chosen_standard_rows(cls.results["nyt-aar"]))
        cases = [(2026, "Medarbejder 5 · p05", 6), (2026, "Medarbejder 5 · p06", 8),
                 (2026, "Medarbejder 5 · p05", 4)]
        set_filtered_plan(plan, 2026, "Medarbejder 5 · p06")
        for year, label, hours in cases:
            append_standard_rows(plan, [(year, label, "Sag")])
            row = plan.max_row
            plan[f"E{row}"], plan[f"BE{row}"] = "Hospital Nord · 24015", hours
        plan[f"B{plan.max_row}"] = "Medarbejder 5 · p06"
        complete["Medarbejdere"]["B12"] = "Medarbejder 6"
        init["B6"] = "Medarbejder 6 · p06"
        unsorted = folder / "indsat.xlsx"
        set_filtered_plan(plan, 2026, "Medarbejder 5 · p06")
        complete.save(unsorted)
        records = list(plan.iter_rows(min_row=10, max_col=65, values_only=True))
        records.sort(key=lambda row: (row[1], row[0], row[4] or ""), reverse=True)
        for number, row in enumerate(records, 10):
            for index, value in enumerate(row, 1):
                plan.cell(number, index).value = value
        set_filtered_plan(plan, 2026, "Medarbejder 5 · p06")
        sorted_path = folder / "sorteret.xlsx"
        complete.save(sorted_path)
        cls.results.update(convert([unsorted, sorted_path]))

    @staticmethod
    def validation_probes(book):
        register = book["Projektregister"]
        for row, value in enumerate(("123", "00123", "DUPL", "dupl"), 500):
            register[f"A{row}"] = value
        staff = book["Medarbejdere"]
        staff["A100"], staff["A101"], staff["A102"] = "Q1", "q1", "Q2"
        for sheet, addresses in ((register, [500, 501, 502, 503]), (staff, [100, 101, 102])):
            rule = next(rule for rule in sheet.data_validations.dataValidation if "A7" in rule)
            for i, row in enumerate(addresses, 1):
                sheet[f"Q{i}"] = Translator("=" + rule.formula1, origin="A7").translate_formula(f"A{row}")

    def test_rules_keep_distinct_keys_and_block_duplicates(self):
        baseline = self.results["reference"]
        self.assertEqual([baseline["Projektregister"][f"Q{r}"].value for r in range(1, 5)],
                         [True, True, False, False])
        self.assertEqual([baseline["Medarbejdere"][f"Q{r}"].value for r in range(1, 4)],
                         [False, False, True])
        self.assertEqual([baseline["Ressourceoverblik"][f"{col}8"].value for col in "ABCDEFG"],
                         [5, 20, 705, 685, 0, 10, 5])

    def test_initializer_uses_registry_not_table_filter_and_only_lists_missing_rows(self):
        old = self.results["reference"]
        self.assertFalse(chosen_standard_rows(old))
        self.assertEqual(old["Planstart"]["D6"].value, "Alle standardrækker findes")
        new = self.results["ny-person"]
        self.assertEqual(new["Medarbejdere"]["L11"].value, "Medarbejder 5 · p05")
        self.assertEqual(new["Medarbejdere"]["L12"].value, "Medarbejder 5 · p06")
        self.assertEqual(chosen_standard_rows(new), [(2026, "Medarbejder 5 · p06", a) for a in ACTIVITIES])
        self.assertEqual([kind for _, _, kind in chosen_standard_rows(self.results["delvis"])],
                         ["Sag", "Fravær", "Kursus", "Andet"])

    def test_real_values_can_be_copied_for_new_year_and_existing_history_kept(self):
        new = self.results["nyt-aar"]
        self.assertEqual(new["Planstart"]["K13"].value, 2032)
        self.assertEqual(chosen_standard_rows(new), [(2032, "Medarbejder 5 · p05", a) for a in ACTIVITIES])
        self.assertEqual([new["Projektoversigt"][f"{c}14"].value for c in "GHIJ"], [52, 1, 2, 3])
        done = self.results["indsat"]
        plan = done["Min ressourceplan"]
        self.assertEqual([(row[3].value) for row in list(plan.rows)[9:]
                          if row[0].value == 2032 and row[2].value == "p05"], list(ACTIVITIES))
        self.assertTrue(any(row[0].value == 2026 and row[2].value == "p05"
                            for row in list(plan.rows)[9:]))
        self.assertFalse(chosen_standard_rows(done))
        self.assertEqual(done["Planstart"]["D6"].value, "Alle standardrækker findes")

    def test_filtered_add_sort_rename_and_reassignment_totals(self):
        reference = None
        for name in ("indsat", "sorteret"):
            book = self.results[name]
            plan, overview, _, project, staff, _ = book.worksheets
            self.assertEqual(staff["L12"].value, "Medarbejder 6 · p06")
            self.assertEqual(plan["G6"].value, 12)
            self.assertEqual(plan["H8"].value, 2026)
            records = Counter((row[0].value, row[2].value, row[3].value, row[62].value,
                               tuple(cell.value for cell in row[9:62]))
                              for row in list(plan.rows)[9:])
            if reference is None:
                reference = records
            else:
                self.assertEqual(reference, records)
            self.assertTrue(any(row[1].value == "Medarbejder 5 · p06" and row[2].value == "p06"
                                and row[8].value == "Medarbejder 6" for row in list(plan.rows)[9:]))
            self.assertEqual([overview[f"{c}8"].value for c in "ABC"], [6, 38, 825])
            self.assertEqual((overview["A19"].value, overview["F19"].value), ("Medarbejder 5", 6))
            self.assertEqual((overview["A20"].value, overview["F20"].value), ("Medarbejder 6", 12))
            self.assertEqual([project[f"F{r}"].value for r in range(15, 19)], [20, 6, 12, 38])
        for name, book in self.results.items():
            for sheet in book:
                with self.subTest(scenario=name, sheet=sheet.title):
                    self.assertFalse([(c.coordinate, c.value) for row in sheet for c in row if c.data_type == "e"])


if __name__ == "__main__":
    unittest.main()
