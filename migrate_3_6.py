"""Byg 3.6 fra 3.5.1 med aktionspladser, scripteksemplarer og afgrænset uge 53."""

import argparse
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from migrate_3_3 import iso_start
from migrate_3_4 import FILTER, NAVY, format_label
from migrate_3_5 import PLAN_STYLE_LAST, workbook_name


def button_placeholder(sheet, area, label):
    sheet.merge_cells(area)
    cell = sheet[area.split(":")[0]]
    cell.value = label
    format_label(cell, FILTER, centered=True, color=NAVY)
    sheet.row_dimensions[cell.row].height = 27


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    workbook = load_workbook(source)
    plan = workbook["Min ressourceplan"]
    if not plan["A1"].value.endswith("3.5.1"):
        raise ValueError("Migrationen kræver Ressourceplan 3.5.1.")
    workbook_name(workbook, "PlanSidsteRaekke",
                  "IFERROR(ROW(INDEX(PlanTimer[#Data],ROWS(PlanTimer[#Data]),1)),9)")
    gray = plan.conditional_formatting[f"BJ10:BJ{PLAN_STYLE_LAST}"][0]
    span = f"{iso_start('$A10+1')}-{iso_start('$A10')}"
    gray.formula = [f"AND(ROW()<=PlanSidsteRaekke,ISNUMBER($A10),$A10>=1901,$A10<=9998,{span}=364)"]

    button_placeholder(plan, "A2:B2", "Tilføj projekt")
    button_placeholder(plan, "D2:F2", "Opret nyt planår")
    button_placeholder(workbook["Medarbejdere"], "A5:B5", "Opret medarbejder")
    plan["A4"] = ("Vælg medarbejder og planår med filterpilene. Brug Tilføj projekt eller Opret nyt "
                  "planår, når scripts er tilknyttet. Indtast timer direkte i ugefelterne.")
    workbook["Medarbejdere"]["A4"] = ("Brug Opret medarbejder, når scriptet er tilknyttet. "
                                     "Behold altid medarbejderens ID ved ændring af navn.")
    settings = workbook["Planstart"]
    settings.sheet_state = "hidden"
    settings["P6"] = "PlanTimer-formler 3.6"
    columns = [column for column in plan.tables["PlanTimer"].tableColumns
               if column.calculatedColumnFormula is not None]
    for row, column in enumerate(columns, 7):
        settings[f"P{row}"] = column.name
        settings[f"Q{row}"] = "=" + column.calculatedColumnFormula.attr_text
        settings[f"Q{row}"].data_type = "s"
    for sheet in workbook:
        sheet["A1"] = "COWI  |  Ressourceplan 3.6"
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / destination.name
        workbook.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_5_1.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_6.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
