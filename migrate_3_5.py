"""Byg Ressourceplan 3.5 med voksende plantabel og nøglebaserede oversigter."""

import argparse
import re
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as col
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import TableColumn, TableFormula
from openpyxl.worksheet.views import Selection

from migrate_3_3 import iso_start, validation
from migrate_3_4 import FIRST, HOURS, SEPARATOR, TABLE_BORDER, escaped_key, font, rebase

PEOPLE = 100
PROJECTS = 500
PLAN_STYLE_LAST = 20009
LAST = FIRST + PEOPLE
WEEKS = [f"Uge {week:02}" for week in range(1, 54)]
FIELDS = [
    "Planår", "Medarbejder", "ID", "Type", "Projektvalg", "Projekt / aktivitet",
    "Timer i alt", "Tilbud %", "Medarbejdervalg", *WEEKS, "Sagsnr.",
    "Effektiv tilbudsprocent", "Timer gyldig uge 53",
]
BASE_TYPES = ("Sag",) * 8 + ("Ferie", "Fravær", "Kursus", "Andet")
SLOT_START = 250
SLOT_TOTAL = SLOT_START + PEOPLE
BOOKED, CAPACITY, GROSS, WEIGHTED = 370, 480, 590, 700
MASK = "$H$223:$BH$223"
YEAR = "$B$185"


def table_field(name):
    return f"PlanTimer[{name}]"


def this(name):
    return f"PlanTimer[[#This Row],[{name}]]"


def workbook_name(book, name, formula):
    book.defined_names.add(DefinedName(name, attr_text=formula))


def lookup_key(value, prefix, raw_keys=None):
    labels, keys = f"{prefix}Etiketter", f"{prefix}Koder"
    suffix_keys = raw_keys or keys
    # Match a live caption first. The longest key suffix recovers a stored caption
    # after a rename, including a key that contains the display separator itself.
    empty_safe = f"({suffix_keys}&\"\")"
    length = f"LEN({empty_safe})"
    matches = f'(RIGHT({value},{length}+{len(SEPARATOR)})="{SEPARATOR}"&{empty_safe})*{length}'
    # SUMPRODUCT supplies array evaluation to MAX and MATCH in both Excel and Calc.
    longest = f"SUMPRODUCT(MAX({matches}))"
    position = f"SUMPRODUCT(MATCH({longest},{matches},0))"
    fallback = f'IFERROR(IF({longest}=0,"",INDEX({suffix_keys},{position})&""),"")'
    return (f'IF({value}="","",IFERROR(INDEX({keys},MATCH({escaped_key(value)},{labels},0)),'
            f'{fallback}))')


def options(book, sheet, prefix, count):
    end = count + 6
    for row in range(7, end + 1):
        index = f"SMALL($J$7:$J${end},ROW()-6)"
        present = f'AND(A{row}<>"",B{row}<>"")' if prefix == "Medarbejder" else f'A{row}<>""'
        sheet[f"J{row}"] = f'=IF({present},ROW()-6,"")'
        sheet[f"K{row}"] = f'=IFERROR(INDEX($A$7:$A${end},{index})&"","")'
        name = f"INDEX($B$7:$B${end},{index})"
        sheet[f"L{row}"] = (
            f'=IF(K{row}="","",IF({name}="","Uden projektnavn",{name})&"{SEPARATOR}"&K{row})'
        )
        if prefix == "Projekt":
            sheet[f"M{row}"] = None
    for letter in "JKLM":
        sheet.column_dimensions[letter].hidden = True
    for letter, suffix in (("K", "Koder"), ("L", "Etiketter")):
        address = f"'{sheet.title}'!${letter}$7:${letter}${end}"
        workbook_name(book, prefix + suffix, address)
    workbook_name(book, prefix + "visninger",
                  f"'{sheet.title}'!$L$7:INDEX({prefix}Etiketter,MAX(1,COUNT('{sheet.title}'!$J$7:$J${end})))")


def extend_register(book, sheet, table_name, prefix, size):
    table = sheet.tables[table_name]
    old = int(re.search(r"(\d+)$", table.ref)[1])
    end = size + 6
    last_column = "F" if prefix == "Medarbejder" else "B"
    table.ref = f"A6:{last_column}{end}"
    table.autoFilter.ref = table.ref
    for row in range(old + 1, end + 1):
        sheet.row_dimensions[row].height = 20
        for number in range(1, 1 + len(table.tableColumns)):
            sheet.cell(row, number)._style = copy(sheet.cell(7 + (row + 1) % 2, number)._style)
        sheet[f"A{row}"].number_format = "@"
    sheet.data_validations.dataValidation = [
        rule for rule in sheet.data_validations.dataValidation if not str(rule.sqref).startswith("A7:")
    ]
    for rule in sheet.data_validations.dataValidation:
        rule.sqref = re.sub(rf":([A-Z]+){old}$", rf":\g<1>{end}", str(rule.sqref))
    key = table.tableColumns[0].name
    validation(sheet, f"A7:A{end}", "custom",
               f'SUMPRODUCT(--({table_name}[{key}]&""=A7&""))=1',
               message="Indtast en unik, varig nøgle. Behold nøglen ved navneændring.")
    if prefix == "Medarbejder":
        for row in range(7, end + 1):
            sheet[f"H{row}"] = f'=IF(E{row}="","",E{row})'
        workbook_name(book, "Afdelingsvalg", f"'Medarbejdere'!$H$6:$H${end}")
        sheet["A4"] = "Udfyld en ledig række med et unikt ID, navn, område, faggruppe, afdeling og normtid."
    else:
        workbook_name(book, "Sagsnumre", f"{table_name}[{key}]")
    options(book, sheet, prefix, size)


def plan_formulas():
    person = this("Medarbejdervalg")
    employee_id = this("ID")
    # INDEX materializes the text-converted key array without a legacy CSE formula.
    employee_match = f'MATCH({escaped_key(employee_id)},INDEX(PersonRegister[ID]&"",0),0)'
    employee_name = f"INDEX(PersonRegister[Medarbejder],{employee_match})"
    year, activity, project = (this(field) for field in ("Planår", "Type", "Sagsnr."))
    project_match = f"MATCH({escaped_key(project)},ProjektKoder,0)"
    actual_name = f"INDEX(SagRegister[Projektnavn],MATCH({escaped_key(project)},INDEX(SagRegister[Sagsnr.]&\"\",0),0))"
    percent = this("Tilbud %")
    types_with_project = f'OR({activity}="Sag",{activity}="Tilbud")'
    isoweeks = f"{iso_start(year+'+1')}-{iso_start(year)}"
    return {
        2: f'=IF({employee_id}="","",IFERROR(IF({employee_name}="","",{employee_name}),""))',
        3: "=" + lookup_key(person, "Medarbejder", "PersonRegister[ID]"),
        6: (f'=IF(AND(SUM({this("Timer i alt")})>0,OR({employee_id}="",'
            f'{year}="",{activity}="")),"Udfyld år, medarbejder og type",'
            f'IF({types_with_project},IF({project}="","",IFERROR(IF({actual_name}="","",{actual_name}),'
            f'IFERROR(INDEX(ProjektEtiketter,{project_match}),"Ukendt projekt"))),IF({activity}="","",{activity})))'),
        7: f'=SUM(PlanTimer[[#This Row],[Uge 01]:[Uge 52]])+{this("Timer gyldig uge 53")}',
        63: f'=IF({types_with_project},{lookup_key(this("Projektvalg"), "Projekt")},"")',
        64: (f'=IF({activity}="Tilbud",IF({percent}="",0.5,'
             f'IF(ISNUMBER({percent}),MAX(0,MIN(1,{percent})),0.5)),0)'),
        65: f'=IF({year}="",0,IF({isoweeks}=371,SUM({this("Uge 53")}),0))',
    }


def keep_record(index, row):
    year, _, employee_id, activity, project, _, _, percent, *hours = row
    slot = (index - 10) % 12
    expected_id = f"p{(index - 10) // 72 + 1:02}"
    expected_year = 2026 + (index - 10) // 12 % 6
    seeded_percentage = slot == 7 and percent == 0.5
    template = (employee_id == expected_id and year == expected_year
                and activity == BASE_TYPES[slot] and project is None
                and all(h is None for h in hours) and (percent is None or seeded_percentage))
    return not template


def resize_plan_table(sheet, last):
    table = sheet.tables["PlanTimer"]
    table.ref = f"A9:BM{last}"
    table.autoFilter = AutoFilter(ref=table.ref)
    year, selected = sheet["A10"].value, sheet["I10"].value
    if year is not None and selected:
        table.autoFilter.add_filter_column(0, [str(year)])
        table.autoFilter.add_filter_column(1, [selected.rsplit(SEPARATOR, 1)[0]])
    table.tableColumns = []
    formulas = plan_formulas()
    for number, name in enumerate(FIELDS, 1):
        column = TableColumn(id=number, name=name)
        if number in formulas:
            column.calculatedColumnFormula = TableFormula(attr_text=formulas[number][1:])
        table.tableColumns.append(column)


def plan_validations(sheet):
    last = PLAN_STYLE_LAST
    sheet.data_validations.dataValidation.clear()
    for cells, choices in (
        (f"A10:A{last}", '"2026,2027,2028,2029,2030,2031"'),
        (f"D10:D{last}", '"Sag,Tilbud,Ferie,Fravær,Kursus,Andet"'),
        (f"E10:E{last}", "Projektvisninger"), (f"I10:I{last}", "Medarbejdervisninger"),
    ):
        validation(sheet, cells, "list", choices)
    validation(sheet, f"H10:H{last}", "decimal", "0", "1")
    validation(sheet, f"J10:BI{last}", "decimal", "0", "80")
    span = f"{iso_start('$A10+1')}-{iso_start('$A10')}"
    validation(sheet, f"BJ10:BJ{last}", "custom", f"AND(ISNUMBER(BJ10),BJ10>=0,BJ10<=80,{span}=371)")


def plan_layout(sheet):
    original = sheet.conditional_formatting
    sheet.conditional_formatting = ConditionalFormattingList()
    areas = {
        "H10:H2169": f"H10:H{PLAN_STYLE_LAST}",
        "BI10:BI2169": f"BJ10:BJ{PLAN_STYLE_LAST}",
        "A10:BI2169": f"A10:BJ{PLAN_STYLE_LAST}", "I7:BI7": "J7:BJ7",
    }
    for area, target in areas.items():
        for rule in original[area]:
            if area == "I7:BI7":
                rule.formula = [rule.formula[0].replace("I7", "J7")]
            sheet.conditional_formatting.add(target, rule)
    sheet["A4"] = ("Filtrér på Planår og Medarbejder. Ved ny tabelrække: vælg planår, medarbejder og type "
                   "før timer. Se README for indsættelse under filtre og personlige arkvisninger.")
    sheet["A6"] = "Timer i synlig plan"
    sheet["G6"] = "=SUBTOTAL(109,PlanTimer[Timer i alt])"
    sheet["H6"] = None
    sheet["H8"] = ('=IF(AND(SUBTOTAL(102,PlanTimer[Planår])>0,'
                   'SUBTOTAL(105,PlanTimer[Planår])=SUBTOTAL(104,PlanTimer[Planår])),'
                   'SUBTOTAL(105,PlanTimer[Planår]),0)')
    for number in range(62, 9, -1):
        letter = col(number)
        for row in (6, 7, 8, 9):
            sheet.cell(row, number)._style = copy(sheet.cell(row, number-1)._style)
        sheet.cell(6, number, f"=SUBTOTAL(109,{table_field(WEEKS[number-10] if number < 62 else FIELDS[-1])})")
        first = iso_start("$H$8")
        monday = f"{first}+7*{number-10}"
        sheet.cell(8, number, f'=IF($H$8=0,"",IF({monday}>={iso_start("$H$8+1")},"",{monday}))')
        sheet.cell(7, number, f'=IF(ISNUMBER({letter}8),"Q"&ROUNDUP(MONTH({letter}8)/3,0),"")')
        sheet.column_dimensions[letter].width = 10
    for row in (6, 7, 8):
        sheet.cell(row, 9).value = None
    sheet["I9"]._style = copy(sheet["E9"]._style)
    for number, field in enumerate(FIELDS, 1):
        sheet.cell(9, number, field)
    for letter, width in {"A": 10, "B": 22, "D": 12, "E": 30, "F": 25, "G": 12, "H": 11, "I": 25}.items():
        sheet.column_dimensions[letter].width = width
    for letter in ("C", "BK", "BL", "BM"):
        sheet.column_dimensions[letter].hidden = True
    sheet.column_dimensions["BJ"].hidden = False
    sheet.column_dimensions["H"].number_format = "0%"
    sheet.column_dimensions["E"].number_format = "@"
    sheet.column_dimensions["I"].number_format = "@"
    sheet.sheet_view.zoomScale = 90
    sheet.freeze_panes = "J10"
    sheet.sheet_view.selection = [
        Selection(pane="topRight", activeCell="J9", sqref="J9"),
        Selection(pane="bottomLeft", activeCell="A10", sqref="A10"),
        Selection(pane="bottomRight", activeCell="J10", sqref="J10"),
    ]


def employee_plan(sheet, people, projects):
    rows = [(n, values) for n, values in enumerate(
        sheet.iter_rows(min_row=10, max_col=61, values_only=True), 10) if keep_record(n, values)]
    sheet.delete_rows(10, sheet.max_row - 9)
    for number in list(sheet.row_dimensions):
        if number >= 10:
            del sheet.row_dimensions[number]
    plan_layout(sheet)
    plan_validations(sheet)
    formulas = plan_formulas()
    body = font()
    if not rows:
        rows = [(10, (None,) * 61)]
    for row, (_, values) in enumerate(rows, 10):
        year, _, identifier, activity, project, _, _, percent, *hours = values
        name = people.get(str(identifier)) or "Ukendt medarbejder"
        assigned = None if identifier in (None, "") else f"{name}{SEPARATOR}{identifier}"
        chosen = f"{projects.get(str(project)) or 'Uden projektnavn'}{SEPARATOR}{project}" if project is not None else None
        for number, value in enumerate((year, None, None, activity, chosen, None, None,
                                        percent, assigned, *hours), 1):
            cell = sheet.cell(row, number, formulas.get(number, value))
            if number <= 62:
                cell.font = body
                cell.alignment = Alignment(horizontal="right" if number >= 10 or number in (7, 8) else "left",
                                           vertical="center", indent=1)
                cell.border = TABLE_BORDER
            if number in (5, 9) and value is not None:
                cell.data_type = "s"
        for number in (63, 64, 65):
            sheet.cell(row, number, formulas[number])
        sheet[f"H{row}"].number_format = "0%"
        sheet.row_dimensions[row].hidden = (year, assigned) != (sheet["A10"].value, sheet["I10"].value)
    resize_plan_table(sheet, len(rows) + 9)


def advance_calendar(sheet, *, annual):
    start, finish = (8, 60) if annual else (7, 59)
    saved = {address: sheet[address].value for address in ("B66", "D66")}
    for number in range(start, finish + 1):
        letter = col(number)
        for row in (67, 68, 69, 103) if annual else (67, 68, 69, 101, 102, 103):
            address = f"{letter}{row}"
            saved[address] = sheet[address].value
        for row in (12, 13, 14):
            cell = sheet[f"{letter}{row}"]
            cell.value = rebase(rebase(cell.value))
    sheet.delete_rows(FIRST, sheet.max_row - FIRST + 1)
    for row in list(sheet.row_dimensions):
        if row >= FIRST:
            del sheet.row_dimensions[row]
    for address, value in saved.items():
        sheet[address[:-len(str(sheet[address].row))] + str(sheet[address].row + 120)] = rebase(rebase(value))
    controls = ("A6", "B6", "C6") if annual else ("B6", "C6", "D6")
    for address, control in zip(("B185", "D185", "F185"), controls):
        sheet[address] = "=" + control
    original = sheet.conditional_formatting
    sheet.conditional_formatting = ConditionalFormattingList()
    for area in original:
        address = re.sub(r"([A-Z]+)45$", rf"\g<1>{LAST}", str(area.sqref))
        for rule in original[area]:
            rule.formula = [formula.replace("$103", "$223") for formula in rule.formula]
            sheet.conditional_formatting.add(address, rule)


def worker(sheet, row, employee, *, annual):
    ident = f"Medarbejdere!A{employee}"
    name = f"Medarbejdere!B{employee}"
    ok = [f'{ident}<>""', f'{name}<>""',
          f'SUMPRODUCT(--(PersonRegister[ID]&""={ident}&""))=1']
    if annual:
        for registry, choice in zip("CDE", ("D6", "E6", "F6")):
            ok.append(f'OR(${choice[0]}$6="Alle",${choice[0]}$6="",'
                      f'Medarbejdere!{registry}{employee}&""=${choice[0]}$6&"")')
    sheet[f"BI{row}"] = f'=IF(AND({",".join(ok)}),{ident}&"","")'
    for visible, registered in zip("ABCDE", "BCDEF"):
        reference = f"Medarbejdere!{registered}{employee}"
        sheet[f"{visible}{row}"] = f'=IF(OR($BI{row}="",{reference}=""),"",{reference})'


def activity_mask(*activities):
    types = table_field("Type")
    return "(" + "+".join(f'({types}="{name}")' for name in activities) + ")"


def key_masks(row, year):
    return f'--({table_field("ID")}&""=$BI{row}),--({table_field("Planår")}&""={year}&"")'


def ratio(booked, cap):
    return f'IF({cap}=0,IF({booked}>0,"Over",0),{booked}/{cap})'


def period_sum(row):
    return f"SUMPRODUCT({MASK},H{row}:BH{row})"


def annual_calculations(sheet):
    for index in range(PEOPLE):
        row, employee = SLOT_START + index, 7 + index
        worker(sheet, row, employee, annual=True)
        booked, available, gross, weighted = [block + index for block in (BOOKED, CAPACITY, GROSS, WEIGHTED)]
        sheet[f"BJ{row}"] = f'=IF(BI{row}="","",{index+1})'
        sheet[f"F{row}"] = "=" + period_sum(booked)
        sheet[f"G{row}"] = "=" + period_sum(available)
        hr_range, cap_range = f"H{booked}:BH{booked}", f"H{available}:BH{available}"
        sheet[f"F{gross}"] = "=" + period_sum(gross)
        sheet[f"F{weighted}"] = "=" + period_sum(weighted)
        sheet[f"BK{row}"] = f"=SUMPRODUCT({MASK},--({cap_range}>{hr_range}),{cap_range}-{hr_range})"
        sheet[f"BL{row}"] = f"=SUMPRODUCT({MASK},--({hr_range}>{cap_range}),{hr_range}-{cap_range})"
        for week, field in enumerate(WEEKS, 8):
            letter = col(week)
            hours = table_field(field)
            masks = key_masks(row, YEAR)
            invalid = f'OR($BI{row}="",{letter}$189="")'
            prefix = f"SUMPRODUCT({hours},{masks},"
            sheet[f"{letter}{booked}"] = f'=IF({invalid},0,{prefix}{activity_mask("Sag", "Andet")}))'
            sheet[f"{letter}{available}"] = (
                f'=IF({invalid},0,MAX(0,N($E{row})-'
                f'{prefix}{activity_mask("Ferie", "Fravær", "Kursus")})))'
            )
            sheet[f"{letter}{gross}"] = f'=IF({invalid},0,{prefix}{activity_mask("Tilbud")}))'
            sheet[f"{letter}{weighted}"] = (
                f"=IF({invalid},0,{prefix}{table_field('Effektiv tilbudsprocent')}))"
            )
            sheet[f"{letter}{row}"] = (
                f'=IF({invalid},"",{ratio(f"{letter}{booked}", f"{letter}{available}")})'
            )
    sheet["BJ161"] = f"=COUNT(BJ{SLOT_START}:BJ{SLOT_TOTAL - 1})"
    sheet["A8"] = "=BJ161"
    sheet["B8"] = f"=F{SLOT_TOTAL}"
    sheet["C8"] = f"=G{SLOT_TOTAL}"
    for label, colname, first in (("D8", "BK", SLOT_START), ("E8", "BL", SLOT_START),
                                  ("F8", "F", GROSS), ("G8", "F", WEIGHTED)):
        sheet[label] = f"=SUM({colname}{first}:{colname}{first+PEOPLE-1})"
    sheet["H8"] = "=" + ratio("B8", "C8")
    sheet["A11"] = '=IF($BJ$161=0,"Ingen medarbejdere matcher filtrene","Årsoversigt")'
    for letter in "FG":
        sheet[f"{letter}{SLOT_TOTAL}"] = f"=SUM({letter}{SLOT_START}:{letter}{SLOT_TOTAL-1})"
    for week in range(8, 61):
        letter = col(week)
        booked = f"SUM({letter}{BOOKED}:{letter}{BOOKED+PEOPLE-1})"
        cap = f"SUM({letter}{CAPACITY}:{letter}{CAPACITY+PEOPLE-1})"
        sheet[f"{letter}{SLOT_TOTAL}"] = f'=IF({letter}189="","",{ratio(booked, cap)})'


def project_calculations(sheet):
    sheet["BJ161"] = f"=COUNT(BJ{SLOT_START}:BJ{SLOT_TOTAL-1})"
    sheet["BJ162"] = "=" + lookup_key("A6", "Projekt")
    key = "$BJ$162"
    sheet["A8"] = (f'=IF({key}="","",IFERROR(INDEX(ProjektEtiketter,'
                   f'MATCH({escaped_key(key)},ProjektKoder,0)),{key}))')
    sheet["B8"] = "=BJ161"
    sheet["C8"] = f'=IF({key}="","",F{SLOT_TOTAL})'
    sheet["A11"] = f'=IF({key}="","Vælg et projekt",IF(BJ161=0,"Ingen projekttimer i perioden","Periodens projektplan"))'
    for index in range(PEOPLE):
        row, employee = SLOT_START + index, 7 + index
        worker(sheet, row, employee, annual=False)
        sheet[f"F{row}"] = f"=SUM(G{row}:BG{row})"
        sheet[f"BJ{row}"] = f'=IF(F{row}>0,{index+1},"")'
        for week in range(7, 60):
            letter = col(week)
            calendar_week = f"{letter}$189"
            hours = f"INDEX(PlanTimer[[Uge 01]:[Uge 53]],0,{calendar_week})"
            masks = key_masks(row, f"{letter}$222")
            sheet[f"{letter}{row}"] = (
                f'=IF(OR({calendar_week}="",{key}="",$BI{row}=""),0,'
                f'SUMPRODUCT({hours},{masks},--({table_field("Sagsnr.")}={key})))'
            )
    for number in range(6, 60):
        letter = col(number)
        blank = f'{key}=""' if letter == "F" else f'OR({key}="",{letter}$189="")'
        sheet[f"{letter}{SLOT_TOTAL}"] = f'=IF({blank},"",SUM({letter}{SLOT_START}:{letter}{SLOT_TOTAL-1}))'


def report_rows(sheet, *, annual):
    end = 60 if annual else 59
    body = font()
    for row in range(FIRST, LAST + 1):
        slot = row - FIRST + 1
        tag = f"$BJ{row}"
        sheet[f"BJ{row}"] = f'=IF({slot}<=$BJ$161,1,IF({slot}=$BJ$161+1,2,0))'
        sheet[f"BK{row}"] = f'=IF({tag}=1,SMALL($BJ${SLOT_START}:$BJ${SLOT_TOTAL-1},{slot}),0)'
        sheet.row_dimensions[row].height = 20
        for number in range(1, end + 1):
            letter = col(number)
            total = '"I ALT"' if number == 1 else f'{letter}${SLOT_TOTAL}' if number >= 6 else '""'
            employee = f"INDEX({letter}${SLOT_START}:{letter}${SLOT_TOTAL-1},$BK{row})"
            if not annual and number >= 7:
                employee = f'IF({letter}$14="","",{employee})'
            cell = sheet.cell(row, number, f'=IF({tag}=1,{employee},IF({tag}=2,{total},""))')
            cell.font = body
            cell.alignment = Alignment(horizontal="left" if number < 5 else "right",
                                       vertical="center", indent=1)
            if number >= 5:
                cell.number_format = "0%" if annual and number >= 8 else HOURS
    for letter in ("BI", "BJ", "BK", "BL"):
        sheet.column_dimensions[letter].hidden = True
    end_model = WEIGHTED + PEOPLE - 1 if annual else SLOT_TOTAL
    for row in range(161, end_model + 1):
        sheet.row_dimensions[row].hidden = True


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    book = load_workbook(source)
    if not book["Min ressourceplan"]["A1"].value.endswith("3.4"):
        raise ValueError("Migrationen kræver Ressourceplan 3.4.")
    people = book["Medarbejdere"]
    registry = book["Projektregister"]
    employee_plan(book["Min ressourceplan"],
                  {str(row[0]): row[1] for row in people.iter_rows(min_row=7, max_col=2, values_only=True)},
                  {str(row[0]): row[1] for row in registry.iter_rows(min_row=7, max_col=2, values_only=True)})
    extend_register(book, people, "PersonRegister", "Medarbejder", PEOPLE)
    extend_register(book, registry, "SagRegister", "Projekt", PROJECTS)
    annual, project = book["Ressourceoverblik"], book["Projektoversigt"]
    advance_calendar(annual, annual=True)
    annual_calculations(annual)
    report_rows(annual, annual=True)
    advance_calendar(project, annual=False)
    project_calculations(project)
    report_rows(project, annual=False)
    for sheet in book:
        sheet["A1"] = "COWI  |  Ressourceplan 3.5"
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / destination.name
        book.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_4.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_5.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
