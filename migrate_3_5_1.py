"""Byg 3.5.1 fra den dynamiske plantabel i Ressourceplan 3.5."""

import argparse
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import TableFormula
from openpyxl.worksheet.views import Selection

from migrate_3_3 import NAVY, validation
from migrate_3_4 import BLUE, SEPARATOR, WHITE, controls, font, format_label, header
from migrate_3_5 import lookup_key, table_field, this, workbook_name

ACTIVITIES = ("Sag", "Ferie", "Fravær", "Kursus", "Andet")
YEARS = tuple(range(2026, 2032))
YEAR_RANGE = "Planårsvalg"
LAST_YEAR_ROW = 26


def fix_register_validation(workbook):
    for title, end in (("Projektregister", 506), ("Medarbejdere", 106)):
        rule = next(rule for rule in workbook[title].data_validations.dataValidation
                    if str(rule.sqref) == f"A7:A{end}")
        # Excel's custom validation parser does not accept the direct table
        # reference used in 3.5. The equivalent local range keeps text keys
        # such as 00123 distinct from 123, following the accepted 3.4 rule.
        rule.formula1 = f'SUMPRODUCT(--($A$7:$A${end}&""=A7&""))=1'
    workbook_name(workbook, "Sagsnumre", "'Projektregister'!$A$7:$A$506")


def combine_employee_columns(plan):
    table = plan.tables["PlanTimer"]
    old_live_name = table.tableColumns[1].calculatedColumnFormula
    table.tableColumns[1].calculatedColumnFormula = None
    table.tableColumns[8].name = "Aktuelt medarbejdernavn"
    table.tableColumns[8].calculatedColumnFormula = old_live_name
    table.tableColumns[2].calculatedColumnFormula = TableFormula(
        attr_text=lookup_key(this("Medarbejder"), "Medarbejder", "PersonRegister[ID]")
    )
    plan["I9"] = table.tableColumns[8].name
    for row in range(10, plan.max_row + 1):
        plan[f"B{row}"], plan[f"I{row}"] = plan[f"I{row}"].value, "=" + old_live_name.attr_text
        if plan[f"B{row}"].value is not None:
            plan[f"B{row}"].data_type = "s"
        plan[f"C{row}"] = "=" + table.tableColumns[2].calculatedColumnFormula.attr_text
    plan.column_dimensions["B"].width = 33
    plan.column_dimensions["B"].number_format = "@"
    plan.column_dimensions["I"].hidden = True
    table.autoFilter = AutoFilter(ref=table.ref)
    year, employee = plan["A10"].value, plan["B10"].value
    if year is not None and employee:
        table.autoFilter.add_filter_column(0, [str(year)])
        table.autoFilter.add_filter_column(1, [employee])
        for row in range(10, plan.max_row + 1):
            plan.row_dimensions[row].hidden = (plan[f"A{row}"].value, plan[f"B{row}"].value) != (year, employee)
    for rule in plan.data_validations.dataValidation:
        area = str(rule.sqref)
        if area.startswith("I10:"):
            rule.sqref = area.replace("I", "B")
        elif area.startswith("A10:"):
            rule.formula1 = YEAR_RANGE
    plan["A4"] = ("Vælg og filtrér medarbejder i samme kolonne. Udfyld år, medarbejder og type før timer. "
                  "Brug Planstart til at finde og tilføje manglende standardrækker.")


def missing_activities(items):
    present = set()
    for row in items:
        kind = row[3]
        if kind != "Sag" or (row[4] in (None, "") and all(v in (None, "", 0) for v in row[9:62])):
            present.add(kind)
    return [kind for kind in ACTIVITIES if kind not in present]


def append_standard_rows(plan, records):
    table = plan.tables["PlanTimer"]
    formulas = {c.id: "=" + c.calculatedColumnFormula.attr_text for c in table.tableColumns
                if c.calculatedColumnFormula is not None}
    last = plan.max_row
    for row, (year, employee, kind) in enumerate(records, last + 1):
        for column in range(1, 66):
            target = plan.cell(row, column, formulas.get(column))
            target._style = copy(plan.cell(10, column)._style)
        plan[f"A{row}"], plan[f"B{row}"], plan[f"D{row}"] = year, employee, kind
        plan[f"B{row}"].data_type = "s"
        plan.row_dimensions[row].hidden = bool(table.autoFilter.filterColumn) and (
            year != plan["A10"].value or employee != plan["B10"].value
        )
        last = row
    table.ref = table.autoFilter.ref = f"A9:BM{last}"


def initialise_existing(plan, employees, years):
    workers = [(str(key), str(name)) for key, name in employees if key not in (None, "") and name]
    keys = sorted((key for key, _ in workers), key=len, reverse=True)
    records = {}
    for record in plan.iter_rows(min_row=10, max_col=62, values_only=True):
        if record[1]:
            caption = str(record[1]).casefold()
            key = next((key.casefold() for key in keys
                        if caption.endswith((SEPARATOR + key).casefold())), None)
            if key is not None:
                records.setdefault((str(record[0]), key), []).append(record)
    additions = [(year, f"{name}{SEPARATOR}{key}", kind)
                 for year in years for key, name in workers
                 for kind in missing_activities(records.get((str(year), key.casefold()), ()))]
    append_standard_rows(plan, additions)
    return len(additions)


def year_options(workbook, sheet):
    sheet["F5"] = "Tilgængelige planår"
    sheet["F5"].font = font(NAVY, 10, True)
    sheet["F6"] = "Tilføj år i ledigt felt"
    sheet["F6"].font = font(NAVY, 10)
    for row in range(7, LAST_YEAR_ROW + 1):
        cell = sheet[f"F{row}"]
        if row - 7 < len(YEARS):
            cell.value = YEARS[row - 7]
        format_label(cell, "EAF2FC", centered=True, bold=False)
        sheet[f"J{row}"] = f'=IF(ISNUMBER(F{row}),ROW()-6,"")'
        sheet[f"K{row}"] = (
            f'=IFERROR(INDEX($F$7:$F${LAST_YEAR_ROW},SMALL($J$7:$J${LAST_YEAR_ROW},ROW()-6)),"")'
        )
    workbook_name(workbook, YEAR_RANGE,
                  f"'Planstart'!$K$7:INDEX('Planstart'!$K$7:$K${LAST_YEAR_ROW},"
                  f"MAX(1,COUNT('Planstart'!$J$7:$J${LAST_YEAR_ROW})))")
    validation(sheet, f"F7:F{LAST_YEAR_ROW}", "custom",
               f"AND(ISNUMBER(F7),F7=INT(F7),F7>=1901,F7<=9998,COUNTIF($F$7:$F${LAST_YEAR_ROW},F7)=1)",
               message="Indtast et unikt helt planår mellem 1901 og 9998.")


def initializer_sheet(workbook):
    sheet = workbook.create_sheet("Planstart")
    for column, width in {"A": 14, "B": 36, "C": 4, "D": 25, "E": 5, "F": 27}.items():
        sheet.column_dimensions[column].width = width
    header(sheet, 6, "Vælg år og medarbejder. Tabellen viser kun standardrækker, "
           "der endnu mangler i Min ressourceplan.")
    year_options(workbook, sheet)
    validation(sheet, "A6", "list", YEAR_RANGE)
    validation(sheet, "B6", "list", "Medarbejdervisninger")
    inputs = list(sheet.data_validations.dataValidation)[-2:]
    for rule in inputs:
        sheet.data_validations.dataValidation.remove(rule)
    controls(sheet, [("Planår", YEARS[0], inputs[0]), ("Medarbejder", None, inputs[1])])

    sheet["H6"] = "=" + lookup_key("B6", "Medarbejder", "PersonRegister[ID]")
    sheet["I6"] = f'=AND(H6<>"",ISNUMBER(A6),COUNTIF({YEAR_RANGE},A6)=1)'
    sheet["D6"] = '=IF(NOT($I$6),"Vælg år og medarbejder",IF(COUNT(H10:H14)=0,"Alle standardrækker findes",COUNT(H10:H14)&" rækker mangler"))'
    sheet["D6"].font = font(NAVY, 10, True)
    sheet["D6"].alignment = Alignment(wrap_text=True, vertical="center")
    for address, label in (("A9", "Planår"), ("B9", "Medarbejder"), ("D9", "Type")):
        format_label(sheet[address], NAVY, color=WHITE)
        sheet[address] = label

    for row, kind in enumerate(ACTIVITIES, 10):
        sheet[f"I{row}"] = kind
        match = (f'--({table_field("ID")}&""=$H$6),'
                 f'--({table_field("Planår")}&""=$A$6&""),'
                 f'--({table_field("Type")}=I{row})')
        if kind == "Sag":
            match += (f',--({table_field("Projektvalg")}=""),'
                      f'--({table_field("Timer i alt")}=0),--({table_field("Uge 53")}=0)')
        sheet[f"H{row}"] = f'=IF($I$6,IF(SUMPRODUCT({match})=0,ROW()-9,""),"")'
        sheet[f"D{row}"] = f'=IFERROR(INDEX($I$10:$I$14,SMALL($H$10:$H$14,ROW()-9)),"")'
        sheet[f"A{row}"] = f'=IF(D{row}="","",$A$6)'
        sheet[f"B{row}"] = f'=IF(D{row}="","",$B$6)'
        sheet.row_dimensions[row].height = 20
        for letter in "ABD":
            format_label(sheet[f"{letter}{row}"], WHITE if row % 2 == 0 else BLUE, bold=False)

    for row, note in (
        (16, "1. Gør plads til det viste antal tomme datarækker i PlanTimer. Brug din egen arkvisning."),
        (17, "2. Kopiér A og B fra de udfyldte hjælperækker, og indsæt som værdier i planens A og B."),
        (18, "3. Kopiér D for de samme hjælperækker, og indsæt som værdier i planens D. Spring skjult C over."),
        (19, "4. Genanvend filtrene i Min ressourceplan. Rækkerne bør nu vises for det valgte år og navn."),
    ):
        sheet.merge_cells(start_row=row, end_row=row, start_column=1, end_column=5)
        sheet[f"A{row}"] = note
        sheet[f"A{row}"].font = font(NAVY, 10)
        sheet[f"A{row}"].alignment = Alignment(vertical="center", wrap_text=True)
        sheet.row_dimensions[row].height = 22
    for letter in "HIJK":
        sheet.column_dimensions[letter].hidden = True
    sheet.sheet_view.selection = [Selection(activeCell="B6", sqref="B6")]


def update_calendar_controls(workbook):
    overview, project = workbook["Ressourceoverblik"], workbook["Projektoversigt"]
    for sheet, address in ((overview, "A6"), (project, "B6")):
        next(rule for rule in sheet.data_validations.dataValidation if address in rule).formula1 = YEAR_RANGE
    for number in range(7, 60):
        cell = project[f"{col(number)}221"]
        cell.value = cell.value.replace("$B$185<2031", f"COUNTIF({YEAR_RANGE},$B$185+1)>0")


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    workbook = load_workbook(source)
    plan = workbook["Min ressourceplan"]
    if not plan["A1"].value.endswith("3.5"):
        raise ValueError("Migrationen kræver Ressourceplan 3.5.")
    fix_register_validation(workbook)
    combine_employee_columns(plan)
    employees = workbook["Medarbejdere"]
    initialise_existing(plan, list(employees.iter_rows(min_row=7, max_col=2, values_only=True)), YEARS)
    initializer_sheet(workbook)
    update_calendar_controls(workbook)
    employees["A4"] = "Udfyld en ledig række med varigt ID og medarbejderdata. Åbn derefter Planstart."
    for sheet in workbook:
        sheet["A1"] = "COWI  |  Ressourceplan 3.5.1"
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / destination.name
        workbook.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_5.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_5_1.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
