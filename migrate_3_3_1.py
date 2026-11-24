"""Migrér Excel 3.3 til en ny 3.3.1-fil med de eksisterende plan- og registerdata."""

import argparse
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.views import BookView
from openpyxl.worksheet.views import Selection

from migrate_3_3 import (
    NAVY, PEOPLE, PLAN, QUARTERS, YEARS, FIRST_PLAN_ROW, col, fill, hide_calculations,
    iso_start, project_name,
)

HOURS = '0.0;-0.0;"–"'
FIRST_WEEK_COLUMN = 7
LAST_WEEK_COLUMN = FIRST_WEEK_COLUMN + 52
LAST_WEEK_LETTER = col(LAST_WEEK_COLUMN)
PROJECT = "$B$2"


def normalize_dxf(workbook):
    """Match Excel's serialized background fill for conditional formatting."""
    for style in workbook._differential_styles.styles:
        if style.fill is None:
            continue
        pattern = style.fill
        if pattern.patternType not in (None, "solid"):
            raise ValueError("Ukendt mønster i Excels differentialformatering.")
        rgb = pattern.fgColor if pattern.fgColor.rgb != "00000000" else pattern.bgColor
        style.fill = PatternFill(bgColor=rgb)
    percentage_format = next(
        style.numFmt for style in workbook._differential_styles.styles
        if style.numFmt is not None and style.numFmt.formatCode == ";;;"
    )
    # Give the differential format its own non-conflicting custom number format ID.
    percentage_format.numFmtId = 164 + workbook._number_formats.add(";;;")


def set_panes(sheet, column, row):
    first_week = f"{column}{row}"
    sheet.freeze_panes = None
    sheet.freeze_panes = first_week
    sheet.sheet_view.selection = [
        Selection(pane="topRight", activeCell=f"{column}9", sqref=f"{column}9"),
        Selection(pane="bottomLeft", activeCell=f"A{row}", sqref=f"A{row}"),
        Selection(pane="bottomRight", activeCell=first_week, sqref=first_week),
    ]


def fix_plan(sheet):
    # A table without an assigned style still has its AutoFilter and Excel table behaviour.
    sheet.tables["PlanTimer"].tableStyleInfo.name = None
    white = FormulaRule(formula=['AND($D10="Sag",MOD(ROW(),2)=0)'],
                        fill=PatternFill(bgColor="FFFFFFFF"), stopIfTrue=True)
    white.priority = max(rule.priority for area in sheet.conditional_formatting
                         for rule in sheet.conditional_formatting[area]) + 1
    sheet.conditional_formatting.add("A10:BI2169", white)
    for cf, rules in sheet.conditional_formatting._cf_rules.items():
        if str(cf.sqref) in {"BI10:BI2169", "A10:BI2169"}:
            for rule in rules:
                rule.stopIfTrue = True
                if str(cf.sqref) == "BI10:BI2169":
                    rule.dxf.font.color = "FF64748B"
        if str(cf.sqref) == "H10:H2169":
            for rule in rules:
                rule.stopIfTrue = False
    set_panes(sheet, "I", 10)


def fix_resource_overview(sheet):
    sheet["F9"], sheet["G9"] = "Booket i perioden (t)", "Kapacitet i perioden (t)"
    for addr in ("F9", "G9"):
        sheet[addr].alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[9].height = 32
    for letter, width in {"B": 16, "D": 16, "E": 15, "F": 23, "G": 25}.items():
        sheet.column_dimensions[letter].width = width
    # The filter labels, selector contents and KPI captions share these columns.
    for number in range(8, 15):
        sheet.column_dimensions[col(number)].width = 18
    for address in ("A2", "C2", "E2", "G2", "I2", "K2"):
        sheet[address] = sheet[address].value + " (t)"
    for week in range(1, 54):
        letter = col(week + 7)
        booked = f"SUM({letter}45:{letter}74)"
        capacity = f"SUM({letter}80:{letter}109)"
        sheet[f"{letter}40"] = (
            f'=IF({letter}9="","",IF({capacity}=0,IF({booked}>0,"Over",0),{booked}/{capacity}))'
        )
        sheet[f"{letter}40"].number_format = "0%"
        sheet[f"{letter}40"].font = Font(name="Segoe UI", color="FF" + NAVY, bold=True)
        sheet[f"{letter}40"].fill = fill("EAF2FC")
    rules = sheet.conditional_formatting["H10:BH39"]
    del sheet.conditional_formatting["H10:BH39"]
    for rule in rules:
        rule.stopIfTrue = True
        sheet.conditional_formatting.add("H10:BH40", rule)
    set_panes(sheet, "H", 10)


def project_controls(sheet):
    sheet["A4"] = ("Vælg projekt, planår og periode. Ugerne starter i første ugekolonne. "
                   "Kun medarbejdere med projekttimer i perioden vises.")
    sheet["A2"] = "Vælg projekt"
    sheet["A2"]._style = copy(sheet["H5"]._style)
    sheet["B2"] = sheet["I5"].value
    sheet["B2"]._style = copy(sheet["I5"]._style)
    sheet.row_dimensions[2].height = 24
    sheet["C2"] = '=IF(B2="","",' + project_name("B2") + ")"
    sheet["C2"]._style = copy(sheet["H6"]._style)
    sheet.merge_cells("C2:F2")
    sheet.unmerge_cells("H6:M6")
    for address in ("H5", "I5", "H6"):
        sheet[address].value = None
        sheet[address]._style = copy(sheet["N5"]._style)
    selector = next(rule for rule in sheet.data_validations.dataValidation if str(rule.sqref) == "I5")
    selector.sqref = "B2"
    sheet["F9"] = "Periodetimer (t)"
    sheet.column_dimensions["F"].width = 20
    sheet["F6"] = f"=COUNT(G9:{LAST_WEEK_LETTER}9)"
    sheet["A7"] = '=IF(AND(B2<>"",COUNT($BI$45:$BI$74)=0),"Ingen projekttimer i perioden","Kvartal")'


def project_weeks(sheet):
    last = LAST_WEEK_LETTER
    raw_start = iso_start("$B$5")
    quarter_dates = f"MONTH({raw_start}+7*({{week}}-1))"
    sheet.conditional_formatting._cf_rules.clear()
    for index, color in enumerate(QUARTERS, 1):
        sheet.conditional_formatting.add(
            f"G7:{last}7", FormulaRule(formula=[f'G7="Q{index}"'],
                                      fill=PatternFill(bgColor="FF" + color), stopIfTrue=True),
        )
    # Style the first week like subsequent week columns, while keeping controls fixed.
    for row in (7, 8, 9, *range(10, 41)):
        sheet.cell(row, 7)._style = copy(sheet.cell(row, 8)._style)
        sheet[f"BH{row}"].value = None
    for offset in range(53):
        week, letter = offset + 1, col(FIRST_WEEK_COLUMN + offset)
        sheet.column_dimensions[letter].width = 11
        quarter = quarter_dates.format(week=week)
        sheet[f"{letter}43"] = (
            f'=IF(AND({week}<=$D$6,OR($F$5="Hele året",'
            f'$F$5="Q"&ROUNDUP({quarter}/3,0))),{week},"")'
        )
        # For numbered periods, use a continuous ISO week sequence, including year rollover.
        sheet[f"{letter}41"] = (
            f'=IF(OR($F$5="Hele året",LEFT($F$5,1)="Q"),'
            f'IFERROR(SMALL($G$43:${last}$43,{week}),""),'
            f'IF({week}<=IFERROR(VALUE(LEFT($F$5,FIND(" ",$F$5)-1)),0),'
            f'IF(OR($B$6+{offset}<=$D$6,$B$5<{YEARS[-1]}),$B$6+{offset},""),""))'
        )
        sheet[f"{letter}42"] = (
            f'=IF({letter}41="","",$B$5+--({letter}41>$D$6))'
        )
        sheet[f"{letter}9"] = f'=IF({letter}41="","",{letter}41-IF({letter}42>$B$5,$D$6,0))'
        sheet[f"{letter}8"] = (
            f'=IF({letter}9="","",{iso_start(f"{letter}42")}+7*({letter}9-1))'
        )
        sheet[f"{letter}7"] = f'=IF({letter}8="","","Q"&ROUNDUP(MONTH({letter}8)/3,0))'
        sheet[f"{letter}9"].number_format = '"Uge "00'
        for row in range(10, 41):
            sheet[f"{letter}{row}"].number_format = HOURS
    sheet["BH43"].value = None


def project_hours(sheet):
    last = LAST_WEEK_LETTER
    for index, employee in enumerate(PEOPLE, 1):
        shown, calc = index + 9, index + 44
        start = FIRST_PLAN_ROW + (index - 1) * 72
        finish = start + 71
        ranges = {c: f"{PLAN}${c}${start}:${c}${finish}" for c in ("A", "C", "D", "E")}
        plan_hours = f"{PLAN}$I${start}:$BI${finish}"
        sheet[f"F{calc}"] = f"=SUM(G{calc}:{last}{calc})"
        sheet[f"BI{calc}"] = f'=IF(F{calc}>0,{index},"")'
        sheet[f"BI{shown}"] = f'=IFERROR(SMALL($BI$45:$BI$74,{index}),"")'
        for target, source in zip("ABCDE", "BCDEF"):
            reference = f"INDEX(Medarbejdere!${source}$7:${source}$36,$BI{shown})"
            sheet[f"{target}{shown}"] = f'=IF($BI{shown}="","",IF({reference}="","",{reference}))'
        sheet[f"F{shown}"] = f'=IF($BI{shown}="","",INDEX($F$45:$F$74,$BI{shown}))'
        sheet[f"F{shown}"].number_format = HOURS
        sheet[f"BH{calc}"].value = None
        for number in range(FIRST_WEEK_COLUMN, LAST_WEEK_COLUMN + 1):
            week = col(number)
            sheet[f"{week}{calc}"] = (
                f'=IF(OR({week}$9="",{PROJECT}="",Medarbejdere!$B${employee}=""),0,'
                f'SUMPRODUCT(INDEX({plan_hours},0,{week}$9),--({ranges["A"]}={week}$42),'
                f'--({ranges["C"]}=Medarbejdere!$A${employee}),--({ranges["E"]}&""={PROJECT}&""),'
                f'(({ranges["D"]}="Sag")+({ranges["D"]}="Tilbud"))))'
            )
            sheet[f"{week}{shown}"] = (
                f'=IF(OR($BI{shown}="",{week}$9=""),"",INDEX({week}$45:{week}$74,$BI{shown}))'
            )
    sheet["F40"] = f'=IF({PROJECT}="","",SUM(F10:F39))'
    sheet["F40"].number_format = HOURS
    for number in range(FIRST_WEEK_COLUMN, LAST_WEEK_COLUMN + 1):
        week = col(number)
        sheet[f"{week}40"] = f'=IF(OR({week}9="",{PROJECT}=""),"",SUM({week}10:{week}39))'
    sheet.column_dimensions["BI"].hidden = True
    hide_calculations(sheet, 74)


def fix_project_overview(sheet):
    project_controls(sheet)
    project_weeks(sheet)
    project_hours(sheet)
    set_panes(sheet, "G", 10)


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    workbook = load_workbook(source)
    plan = workbook["Min ressourceplan"]
    if plan.tables["PlanTimer"].ref != "A9:BI2169" or not plan["A1"].value.endswith("3.3"):
        raise ValueError("Migrationen kræver Ressourceplan 3.3 med den oprindelige tabelstruktur.")

    if not workbook.views:
        workbook.views = [BookView(activeTab=workbook.index(workbook.active))]
    for sheet in workbook:
        if isinstance(sheet["A1"].value, str) and sheet["A1"].value.endswith("3.3"):
            sheet["A1"] = sheet["A1"].value[:-3] + "3.3.1"
    normalize_dxf(workbook)
    fix_plan(plan)
    fix_resource_overview(workbook["Ressourceoverblik"])
    fix_project_overview(workbook["Projektoversigt"])
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / destination.name
        workbook.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_3.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_3_1.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
