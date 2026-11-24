"""Byg en ny Excel 3.3-fil fra en kopi af den oprindelige 3.2-struktur."""

import argparse
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.utils import get_column_letter as col
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import AutoFilter

YEARS = range(2026, 2032)
FIRST_PLAN_ROW = 10
LAST_PLAN_ROW = 2169
PEOPLE = range(7, 37)
PLAN = "'Min ressourceplan'!"
NAVY = "1C2B54"
ORANGE = "F04E23"
QUARTERS = ("D2EAF7", "D5F0E9", "EFE1D6", "FCE3E3")
TYPE_COLORS = {
    "Tilbud": "FCE3E3",
    "Ferie": "FFF4D9",
    "Fravær": "FFF4D9",
    "Kursus": "E3F0DF",
    "Andet": "EDE5F5",
}
HOURS = "0.0;-0.0;–"


def fill(color):
    return PatternFill("solid", fgColor="FF" + color)


def highlight(sheet, cells, formula, color, font=None):
    sheet.conditional_formatting.add(
        cells, FormulaRule(formula=[formula], fill=fill(color), font=font)
    )


def validation(sheet, cells, kind, formula, upper=None, *,
               message="Vælg eller indtast en gyldig værdi.", allow_blank=True):
    rule = DataValidation(
        type=kind, formula1=formula, formula2=upper,
        operator="between" if upper is not None else None,
        allowBlank=allow_blank, showErrorMessage=True, errorStyle="stop",
        errorTitle="Ugyldig indtastning", error=message,
    )
    sheet.add_data_validation(rule)
    rule.add(cells)


def iso_start(year):
    jan4 = f"DATE({year},1,4)"
    return f"({jan4}-WEEKDAY({jan4},2)+1)"


def show_header(cell, template=None):
    if template is not None:
        cell._style = copy(template._style)
    cell.alignment = Alignment(horizontal="center", vertical="center")


def quarter_header(sheet, quarter_row, first, last):
    cells = f"{first}{quarter_row}:{last}{quarter_row}"
    for quarter, color in enumerate(QUARTERS, 1):
        highlight(sheet, cells, f'{first}{quarter_row}="Q{quarter}"', color)


def project_name(key):
    # Exact MATCH/VLOOKUP with escaped wildcard characters in text keys.
    escaped = f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({key},"~","~~"),"*","~*"),"?","~?")'
    lookup = f"IF(ISTEXT({key}),{escaped},{key})"
    return f'IFERROR(VLOOKUP({lookup},Projektregister!$A$7:$B$106,2,FALSE),"Ukendt sag")'


def migrate_plan(sheet, people):
    sheet["A4"] = "Vælg ét planår og din medarbejder i tabelpilene. Datoer vises kun for ét synligt planår."
    sheet["A5"] = "Indtast timer direkte. Teams: Vis > Arkvisning > Ny. Filtre giver ingen adgangsbegrænsning. Undgå sortering."
    sheet["A6"] = "Synlige timer i alt / pr. uge (alle typer)"
    sheet["A7"] = "Kvartal (efter mandagens måned)"
    sheet["A8"] = "Mandag"
    sheet["H6"] = sheet["H6"].value.replace("STOP: sortér efter ID og år", "STOP")
    sheet["H8"].number_format = '0;0;""'
    sheet.column_dimensions["C"].hidden = True
    sheet.column_dimensions["BJ"].hidden = True
    sheet.column_dimensions["BK"].hidden = True
    sheet["BJ9"] = "Effektiv tilbudsprocent"
    sheet["BK9"] = "Timer i gyldig ISO-uge 53"

    table = sheet.tables["PlanTimer"]
    table.tableStyleInfo.showRowStripes = False
    initial_year = YEARS[0]
    active_people = {people[f"A{row}"].value: people[f"B{row}"].value for row in PEOPLE}
    employee_name = next((name for name in active_people.values() if name), None)
    table.autoFilter.add_filter_column(0, [str(initial_year)])
    if employee_name is not None:
        table.autoFilter.add_filter_column(1, [employee_name])

    # Retain existing rows and inputs. Count zero and formulas as actual inputs.
    last = LAST_PLAN_ROW
    for row in range(FIRST_PLAN_ROW, last + 1):
        year, person = sheet[f"A{row}"].value, sheet[f"C{row}"].value
        activity = sheet[f"D{row}"]
        if (
            activity.value == "Tilbud"
            and sheet[f"E{row}"].value in (None, "")
            and sheet[f"H{row}"].value in (None, 0.5)
            and all(cell.value is None for cell in sheet[row][8:61])
        ):
            activity.value = "Sag"
        sheet.row_dimensions[row].hidden = (
            year != initial_year or
            (employee_name is not None and active_people.get(person) != employee_name)
        )
        sheet[f"F{row}"] = (
            f'=IF(OR(D{row}="Sag",D{row}="Tilbud"),'
            f'IF(E{row}="","",{project_name(f"E{row}")}),D{row})'
        )
        weeks_53 = f"{iso_start(f'A{row}+1')}-{iso_start(f'A{row}')}=371"
        sheet[f"BK{row}"] = f"=IF({weeks_53},SUM(BI{row}),0)"
        sheet[f"G{row}"] = f"=SUM(I{row}:BH{row})+BK{row}"
        sheet[f"BJ{row}"] = (
            f'=IF(D{row}="Tilbud",IF(H{row}="",0.5,'
            f'IF(ISNUMBER(H{row}),MAX(0,MIN(1,H{row})),0.5)),0)'
        )
        sheet[f"H{row}"].number_format = "0%"

    sheet.conditional_formatting._cf_rules.clear()
    # The number format and whole-row colour can apply together.
    sheet.conditional_formatting.add(
        f"H10:H{last}",
        Rule(type="expression", formula=['$D10<>"Tilbud"'],
             dxf=DifferentialStyle(numFmt=NumberFormat(numFmtId=206, formatCode=";;;"))),
    )
    highlight(sheet, f"BI10:BI{last}", f"{iso_start('$A10+1')}-{iso_start('$A10')}=364",
              "E1E5EC", Font(color="64748B"))
    for activity, color in TYPE_COLORS.items():
        highlight(sheet, f"A10:BI{last}", f'$D10="{activity}"', color)
    highlight(sheet, f"A10:BI{last}", 'AND($D10="Sag",MOD(ROW(),2)=1)', "D2EAF7")

    sheet.data_validations.dataValidation.clear()
    validation(sheet, f"D10:D{last}", "list", '"Sag,Tilbud,Ferie,Fravær,Kursus,Andet"')
    validation(sheet, f"E10:E{last}", "list", "Sagsnumre")
    validation(sheet, f"H10:H{last}", "decimal", "0", "1",
               message="Tilbudssandsynlighed skal være mellem 0 % og 100 %.")
    validation(sheet, f"I10:BH{last}", "decimal", "0", "80",
               message="Indtast mellem 0 og 80 timer.")
    validation(
        sheet, f"BI10:BI{last}", "custom",
        f"AND(ISNUMBER(BI10),BI10>=0,BI10<=80,{iso_start('$A10+1')}-{iso_start('$A10')}=371)",
        message="Indtast mellem 0 og 80 timer i år med ISO-uge 53.",
    )

    for week in range(1, 54):
        letter = col(week + 8)
        # Count only valid week 53 rows, including when multiple years are shown.
        hours_column = "BK" if week == 53 else letter
        sheet[f"{letter}6"] = f"=SUBTOTAL(109,{hours_column}10:{hours_column}{last})"
        sheet[f"{letter}6"].number_format = HOURS
        sheet[f"{letter}7"] = f'=IF(ISNUMBER({letter}8),"Q"&ROUNDUP(MONTH({letter}8)/3,0),"")'
        for row in (6, 7, 8):
            show_header(sheet[f"{letter}{row}"])
    quarter_header(sheet, 7, "I", "BI")


def annual_period(sheet):
    """Use the same week membership for totals, headers and period highlights."""
    sheet.conditional_formatting._cf_rules.clear()
    sheet["B6"] = '=IF($F$5="Hele året",1,IF(LEFT($F$5,1)="Q","",MAX(1,MIN($D$5,$D$6))))'
    sheet["E6"] = "Uger valgt"
    sheet["F6"] = "=SUM(H43:BH43)"
    sheet["A7"] = "Kvartal"
    sheet["A8"] = "Mandag"
    for week in range(1, 54):
        letter = col(week + 7)
        sheet[f"{letter}9"] = f'=IF({week}<=$D$6,{week},"")'
        sheet[f"{letter}43"] = (
            f'=IF({letter}9="",0,IF($F$5="Hele året",1,IF(LEFT($F$5,1)="Q",'
            f'--({letter}7=$F$5),--AND({letter}9>=$B$6,'
            f'{letter}9<$B$6+IFERROR(VALUE(LEFT($F$5,FIND(" ",$F$5)-1)),0)))))'
        )
        for row in (7, 8, 9):
            show_header(sheet[f"{letter}{row}"])
    quarter_header(sheet, 7, "H", "BH")
    highlight(sheet, "H9:BH9", "H$43=1", ORANGE, Font(color="FFFFFFFF", bold=True))
    for rule in sheet.data_validations.dataValidation:
        rule.showErrorMessage = True
        rule.allowBlank = False


def period_sum(row):
    return f"SUMPRODUCT($H$43:$BH$43,H{row}:BH{row})"


def plan_ranges(person_index, week):
    start = FIRST_PLAN_ROW + person_index * 72
    return [f"{PLAN}${letter}${start}:${letter}${start + 71}"
            for letter in (col(week + 8), "A", "C", "D", "E", "BJ")]


def sum_weeks(sheet, row, cache_row, blank_if):
    sheet[f"F{row}"] = f'=IF({blank_if},"",{period_sum(cache_row)})'


def resource_overview(sheet):
    annual_period(sheet)
    sheet["A4"] = "Vælg planår, periode og medarbejderfiltre. Orange markerer de valgte uger i årsoversigten."
    sheet["H5"], sheet["J5"], sheet["L5"] = "Område", "Faggruppe", "Afdeling"
    sheet.row_dimensions[5].height = 30
    for addr, formula in (("I5", '"Alle,Øst,Vest"'), ("K5", '"Alle,EL,Belysning,BMS/CTS,Sikring"'),
                          ("M5", "Afdelingsvalg")):
        sheet[addr] = "Alle"
        sheet[addr].fill = fill("EAF2FC")
        validation(sheet, addr, "list", formula)
    for addr in ("H5", "J5", "L5"):
        sheet[addr].alignment = Alignment(wrap_text=True, vertical="center")
    sheet.column_dimensions["BI"].hidden = True
    sheet["BI9"] = "Med i visning"

    for index, reg in enumerate(PEOPLE):
        visible = index + 10
        booked, capacity, gross = index + 45, index + 80, index + 115
        weighted, free, overload = index + 150, index + 185, index + 220
        sheet[f"BI{visible}"] = (
            f'=AND(Medarbejdere!B{reg}<>"",OR($I$5="Alle",$I$5="",Medarbejdere!C{reg}=$I$5),'
            f'OR($K$5="Alle",$K$5="",Medarbejdere!D{reg}=$K$5),'
            f'OR($M$5="Alle",$M$5="",Medarbejdere!E{reg}&""=$M$5&""))'
        )
        for dest, source in zip("ABCDE", "BCDEF"):
            sheet[f"{dest}{visible}"] = (
                f'=IF($BI{visible},IF(Medarbejdere!{source}{reg}="","",Medarbejdere!{source}{reg}),"")'
            )
        sum_weeks(sheet, visible, booked, f'$A{visible}=""')
        sheet[f"G{visible}"] = f'=IF(A{visible}="","",{period_sum(capacity)})'
        for week in range(1, 54):
            letter = col(week + 7)
            hours, years, ids, _, _, probabilities = plan_ranges(index, week)
            on = f'AND($BI{visible},{letter}$9<>"")'
            sheet[f"{letter}{weighted}"] = (
                f"=IF({on},SUMPRODUCT({hours},--({years}=$B$5),"
                f"--({ids}=Medarbejdere!$A${reg}),{probabilities}),0)"
            )
            # Keep the original booked, capacity and gross offer formulas and their links.
            # A missing norm is zero, while the visible norm field can stay blank.
            cell = sheet[f"{letter}{capacity}"]
            cell.value = cell.value.replace(f"$E{visible}-", f"N($E{visible})-")
            sheet[f"{letter}{free}"] = f"=MAX(0,{letter}{capacity}-{letter}{booked})"
            sheet[f"{letter}{overload}"] = f"=MAX(0,{letter}{booked}-{letter}{capacity})"
    for row, label in ((149, "Vægtede tilbud"), (184, "Ledige timer"), (219, "Overbelastning")):
        sheet[f"A{row}"] = label
    for row, label, first in ((44, "Bookede timer", 45), (79, "Disponibel kapacitet", 80),
                              (114, "Bruttotilbud", 115), (149, "Vægtede tilbud", 150),
                              (184, "Ledige timer", 185), (219, "Overbelastning", 220)):
        sheet[f"A{row}"] = label
        for offset in range(30):
            sheet[f"F{first + offset}"] = "=" + period_sum(first + offset)
    for column, label, formula in (
        (1, "Booket", "=F40"), (3, "Kapacitet", "=G40"),
        (5, "Ledig", "=SUM(F185:F214)"), (7, "Overbelastning", "=SUM(F220:F249)"),
        (9, "Tilbud brutto", "=SUM(F115:F144)"), (11, "Tilbud vægtet", "=SUM(F150:F179)"),
        (13, "Belægning", '=IF(G40=0,IF(F40>0,"Over",0),F40/G40)'),
    ):
        title = sheet.cell(2, column, label)
        title.fill = fill(NAVY)
        title.font = Font(name="Segoe UI", color="FFFFFFFF", bold=True)
        title.alignment = Alignment(vertical="center", wrap_text=True)
        value = sheet.cell(2, column + 1, formula)
        value.fill = fill("EAF2FC")
        value.font = Font(name="Segoe UI", color=NAVY, bold=True)
        value.number_format = "0.0%" if column == 13 else HOURS
    sheet.row_dimensions[2].height = 30
    for week in range(1, 54):
        letter = col(week + 7)
        sheet[f"{letter}40"] = f'=IF({letter}9="","",SUM({letter}45:{letter}74))'
        sheet[f"{letter}40"].number_format = HOURS
    for predicate, color in (
        ("AND(ISNUMBER(H10),H10=0)", "EAF2FC"),
        ("AND(ISNUMBER(H10),H10>0,H10<0.85)", "DFF2E9"),
        ("AND(ISNUMBER(H10),H10>=0.85,H10<=1)", "FFF1CF"),
        ('OR(AND(ISNUMBER(H10),H10>1),H10="Over")', "F9DADD"),
    ):
        highlight(sheet, "H10:BH39", predicate, color)
    hide_calculations(sheet, 249)


def hide_calculations(sheet, last):
    for row in range(41, last + 1):
        sheet.row_dimensions[row].hidden = True


def project_overview(sheet):
    annual_period(sheet)
    sheet["A4"] = "Vælg projekt, planår og periode. Ugekolonner og Timer i periode følger valget. Årstotal viser hele planåret."
    sheet["H5"] = "Vælg projekt"
    sheet.row_dimensions[5].height = 30
    sheet["H5"].alignment = Alignment(wrap_text=True, vertical="center")
    sheet["H5"].fill = fill(ORANGE)
    sheet["H5"].font = Font(name="Segoe UI", color="FFFFFFFF", bold=True)
    sheet["I5"].fill = fill("EAF2FC")
    sheet["H6"] = '=IF(I5="","",'+project_name("I5")+")"
    sheet.merge_cells("H6:M6")
    sheet["H6"].font = Font(name="Segoe UI", color=NAVY, bold=True)
    sheet["G9"] = "Årstotal"
    rule = next(rule for rule in sheet.data_validations.dataValidation if str(rule.sqref) == "I5")
    rule.formula1 = "Sagsnumre"
    rule.allowBlank = True
    for index, reg in enumerate(PEOPLE):
        visible, calc = index + 10, index + 45
        blank = f'OR($A{visible}="",$I$5="")'
        sum_weeks(sheet, visible, calc, blank)
        sheet[f"G{visible}"] = f'=IF({blank},"",SUM(H{calc}:BH{calc}))'
        for week in range(1, 54):
            letter = col(week + 7)
            hours, years, ids, types, projects, _ = plan_ranges(index, week)
            sheet[f"{letter}{calc}"] = (
                f'=IF(OR({blank},{letter}$9=""),0,SUMPRODUCT({hours},--({years}=$B$5),'
                f'--({ids}=Medarbejdere!$A${reg}),--({projects}&""=$I$5&""),'
                f'(({types}="Sag")+({types}="Tilbud"))))'
            )
            sheet[f"{letter}{visible}"] = f'=IF(OR({blank},{letter}$43=0),"",{letter}{calc})'
    for week in range(1, 54):
        letter = col(week + 7)
        sheet[f"{letter}40"] = f'=IF(OR($I$5="",{letter}$43=0),"",SUM({letter}10:{letter}39))'
        sheet[f"{letter}40"].number_format = HOURS
    hide_calculations(sheet, 74)


def registers(workbook):
    people = workbook["Medarbejdere"]
    people["A4"] = "Udfyld en ledig plads. Behold det tildelte ID, og undgå sortering. Alle seks planår er klar."
    people["H6"] = "Alle"
    for row in PEOPLE:
        people[f"H{row}"] = f'=IF(E{row}="","",E{row})'
    people.column_dimensions["H"].hidden = True
    workbook.defined_names.add(DefinedName("Afdelingsvalg", attr_text="'Medarbejdere'!$H$6:$H$36"))
    for rule in people.data_validations.dataValidation:
        rule.showErrorMessage = True
    validation(people, "F7:F36", "decimal", "0", "80",
               message="Indtast en ugentlig normtid mellem 0 og 80 timer.")

    projects = workbook["Projektregister"]
    workbook.defined_names.add(DefinedName("Sagsnumre", attr_text="'Projektregister'!$A$7:$A$106"))
    for row in range(7, 107):
        projects[f"A{row}"].number_format = "@"
    validation(projects, "A7:A106", "custom", 'SUMPRODUCT(--($A$7:$A$106&""=A7&""))=1',
               message="Sagsnummeret skal være unikt. Behold nummeret ved ændring af projektnavn.")
    for row in range(FIRST_PLAN_ROW, LAST_PLAN_ROW + 1):
        workbook["Min ressourceplan"][f"E{row}"].number_format = "@"
    workbook["Projektoversigt"]["I5"].number_format = "@"


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    workbook = load_workbook(source)
    plan = workbook["Min ressourceplan"]
    if plan.tables["PlanTimer"].ref != "A9:BI2169" or not plan["A1"].value.endswith("3.2"):
        raise ValueError("Migrationen kræver Ressourceplan 3.2 med den oprindelige tabelstruktur.")

    for sheet in workbook:
        if isinstance(sheet["A1"].value, str):
            sheet["A1"] = sheet["A1"].value.replace("3.2", "3.3")
        sheet.sheet_view.showGridLines = False
        for table in sheet.tables.values():
            if table.autoFilter is None:
                table.autoFilter = AutoFilter(ref=table.ref)
    registers(workbook)
    migrate_plan(plan, workbook["Medarbejdere"])
    resource_overview(workbook["Ressourceoverblik"])
    project_overview(workbook["Projektoversigt"])
    workbook.active = plan
    workbook.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    with TemporaryDirectory(dir=destination.parent) as temporary:
        staged = Path(temporary) / destination.name
        workbook.save(staged)
        # Publish the complete file atomically without replacing an existing destination.
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_2.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_3.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
