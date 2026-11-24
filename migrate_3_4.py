"""Byg Ressourceplan 3.4 ved at migrere den godkendte 3.3.1-fil."""

import argparse
import re
from copy import copy
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formula import Tokenizer
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.views import Selection

from migrate_3_3 import NAVY, ORANGE, QUARTERS, TYPE_COLORS, fill

BLUE = "D2EAF7"
WHITE = "FFFFFF"
MINT = "D5F0E9"
FILTER = "EAF2FC"
LINE = "E1E5EC"
TEXT = "192339"
MUTED = "53627A"
CANVAS = "F6F7F8"
OFFSET = 60
FIRST = 15
LAST = FIRST + 30
HOURS = '0.0;-0.0;"–"'
WEEK = '"Uge "00'
EDGE = Side(style="thin", color="FF" + LINE)
TABLE_BORDER = Border(left=EDGE, right=EDGE, top=EDGE, bottom=EDGE)
INPUT_EDGE = Side(style="thin", color="FF879AB2")
INPUT_BORDER = Border(left=INPUT_EDGE, right=INPUT_EDGE, top=INPUT_EDGE, bottom=INPUT_EDGE)
LOCAL_RANGE = re.compile(r'(\$?[A-Z]{1,3}\$?)(\d+)(?=[:]|$)')
SEPARATOR = " · "


def font(color=TEXT, size=11, bold=False):
    return Font(name="Segoe UI", size=size, color="FF" + color, bold=bold)


def band(sheet, row, end, value, background, color, size, height, *, bold=False, orange=False):
    sheet.merge_cells(start_row=row, end_row=row, start_column=1, end_column=end)
    cell = sheet.cell(row, 1, value)
    cell.fill = fill(background)
    cell.font = font(color, size, bold)
    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    if orange:
        cell.border = Border(bottom=Side(style="medium", color="FF" + ORANGE))
    sheet.row_dimensions[row].height = height


def header(sheet, end, note, *, note_height=26):
    for merged in list(sheet.merged_cells.ranges):
        if merged.min_row in (1, 3, 4):
            sheet.unmerge_cells(str(merged))
    band(sheet, 1, end, "COWI  |  Ressourceplan 3.4", NAVY, WHITE, 16, 33,
         bold=True, orange=True)
    band(sheet, 3, end, sheet.title, CANVAS, NAVY, 19, 33, bold=True)
    band(sheet, 4, end, note, CANVAS, MUTED, 10, note_height)
    for cell in sheet[2][:10]:
        cell.fill = PatternFill()
    sheet.row_dimensions[2].height = 7
    sheet.sheet_view.showGridLines = False


def format_label(cell, background, *, centered=False, color=TEXT, bold=True):
    cell.fill = fill(background)
    cell.font = font(color, 10, bold)
    cell.alignment = Alignment(horizontal="center" if centered else "left",
                               vertical="center", wrap_text=True)
    cell.border = TABLE_BORDER


def controls(sheet, items):
    for position, (label, value, rule) in enumerate(items, 1):
        caption = sheet.cell(5, position, label + "  ▾")
        caption.font = font(NAVY, 10, True)
        caption.alignment = Alignment(vertical="center", wrap_text=True)
        entry = sheet.cell(6, position, value)
        entry.fill = fill(FILTER)
        entry.font = font(NAVY, 11, True)
        entry.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        entry.border = INPUT_BORDER
        rule.sqref = entry.coordinate
        rule.showDropDown = False
        sheet.add_data_validation(rule)
    sheet.row_dimensions[5].height = 23
    sheet.row_dimensions[6].height = 25


def cards(sheet, items):
    for position, (label, value, number_format) in enumerate(items, 1):
        format_label(sheet.cell(7, position, label), CANVAS, color=MUTED)
        cell = sheet.cell(8, position, value)
        cell.fill = fill(WHITE)
        cell.font = font(NAVY, 13, True)
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        cell.border = TABLE_BORDER
        cell.number_format = number_format
    sheet.row_dimensions[7].height = 27
    sheet.row_dimensions[8].height = 28


def legend(sheet, row, entries):
    for column, (label, color) in entries.items():
        cell = sheet.cell(row, column, label)
        format_label(cell, color, centered=True, bold=False)
    sheet.row_dimensions[row].height = 22


def stripes_and_totals(sheet, first_col, last_col, *, formula="TRUE", color_total=True):
    start, end = col(first_col), col(last_col)
    area = f"{start}{FIRST}:{end}{LAST}"
    if color_total:
        sheet.conditional_formatting.add(
            area, FormulaRule(formula=[f'AND($BJ{FIRST}=2,{formula})'],
                              fill=PatternFill(bgColor="FF" + MINT), border=TABLE_BORDER,
                              font=font(NAVY, 11, True), stopIfTrue=True))
    for parity, rgb in ((0, WHITE), (1, BLUE)):
        sheet.conditional_formatting.add(
            area, FormulaRule(formula=[f'AND($BJ{FIRST}=1,MOD(ROW()-{FIRST},2)={parity},{formula})'],
                              fill=PatternFill(bgColor="FF" + rgb), border=TABLE_BORDER,
                              stopIfTrue=True))


def rebase(formula):
    # Move only this sheet's row references, including absolute references.
    # Registry and plan references continue to use their original coordinates.
    tokens = Tokenizer(formula)
    for item in tokens.items:
        if item.type == "OPERAND" and item.subtype == "RANGE" and "!" not in item.value:
            item.value = LOCAL_RANGE.sub(lambda match: match[1] + str(int(match[2]) + OFFSET),
                                         item.value)
    return "=" + "".join(item.value for item in tokens.items)


def retire_view(sheet):
    # Store the accepted model below the visible report. Only the report input
    # cells will be rewired to the controls; the existing formula graph stays intact.
    validations = {str(rule.sqref): copy(rule) for rule in sheet.data_validations.dataValidation}
    sheet.data_validations.dataValidation.clear()
    sheet.conditional_formatting = ConditionalFormattingList()
    for merged in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged))
    snapshot = [(c.row, c.column, c.value, c.data_type) for c in sheet._cells.values()
                if c.value is not None]
    last = sheet.max_row
    sheet.delete_rows(1, last)
    sheet.row_dimensions.clear()
    for row, column, value, datatype in snapshot:
        if datatype == "f":
            sheet.cell(OFFSET + row, column, rebase(value))
        elif not isinstance(value, str) or row in (2, 5):
            sheet.cell(OFFSET + row, column, value)
    for row in range(OFFSET + 1, OFFSET + last + 1):
        sheet.row_dimensions[row].hidden = True
    for column in ("BI", "BJ", "BK"):
        sheet.column_dimensions[column].hidden = True
    sheet.sheet_view.topLeftCell = "A1"
    return validations


def freeze(sheet, column):
    sheet.freeze_panes = f"{column}{FIRST}"
    sheet.sheet_view.selection = [
        Selection(pane="topRight", activeCell=f"{column}14", sqref=f"{column}14"),
        Selection(pane="bottomLeft", activeCell=f"A{FIRST}", sqref=f"A{FIRST}"),
        Selection(pane="bottomRight", activeCell=f"{column}{FIRST}", sqref=f"{column}{FIRST}"),
    ]


def calendar(sheet, start, end, *, annual):
    for row, label in ((12, "Kvartal"), (13, "Mandag")):
        sheet.cell(row, 1, label)
        format_label(sheet.cell(row, 1), CANVAS, color=MUTED)
        sheet.row_dimensions[row].height = 19
    for number in range(start, end + 1):
        letter = col(number)
        sheet.column_dimensions[letter].width = 11
        for public, model, fmt in ((12, 67, "General"), (13, 68, "dd.mm.yy"), (14, 69, WEEK)):
            cell = sheet.cell(public, number, f'=IF({letter}69="","",{letter}{model})')
            format_label(cell, NAVY if public == 14 else CANVAS, centered=True,
                         color=WHITE if public == 14 else NAVY)
            cell.number_format = fmt
    first, last = col(start), col(end)
    for index, rgb in enumerate(QUARTERS, 1):
        sheet.conditional_formatting.add(
            f"{first}12:{last}12", FormulaRule(formula=[f'{first}12="Q{index}"'],
                                               fill=PatternFill(bgColor="FF" + rgb), stopIfTrue=True))
    if annual:
        sheet.conditional_formatting.add(
            f"{first}14:{last}14",
            FormulaRule(formula=[f"{first}$103=1"], font=font(WHITE, 10, True),
                        fill=PatternFill(bgColor="FF" + ORANGE), stopIfTrue=True))


def report_rows(sheet, headers, last, *, annual):
    for number, label in enumerate(headers, 1):
        format_label(sheet.cell(14, number, label), NAVY, color=WHITE)
    sheet.row_dimensions[14].height = 29
    for row in range(FIRST, LAST + 1):
        slot = row - FIRST + 1
        sheet.row_dimensions[row].height = 20
        tag = f"$BJ{row}"
        # Tags 1 and 2 identify a record and the immediately following total.
        sheet[f"BJ{row}"] = f'=IF({slot}<=$BJ$61,1,IF({slot}=$BJ$61+1,2,0))'
        model = slot + 69
        if annual:
            sheet[f"BK{row}"] = f'=IF({tag}=1,SMALL($BJ$70:$BJ$99,{slot}),0)'
        for number in range(1, last + 1):
            letter = col(number)
            source = (f"INDEX({letter}$70:{letter}$99,$BK{row})" if annual
                      else f"{letter}{model}")
            total = ('"I ALT"' if number == 1 else
                     f'{letter}$100' if number >= 6 else '""')
            # Annual capacity and project helper values can be the empty string.
            # Let that value pass through without coercion.
            value = f'=IF({tag}=1,{source},IF({tag}=2,{total},""))'
            cell = sheet.cell(row, number, value)
            cell.font = font()
            cell.alignment = Alignment(horizontal="left" if number < 5 else "right",
                                       vertical="center", indent=1)
            if number >= 5:
                cell.number_format = "0%" if annual and number >= 8 else HOURS
    stripes_and_totals(sheet, 1, len(headers))
    first_week = col(len(headers) + 1)
    if annual:
        heat = (
            (f"AND(ISNUMBER({first_week}{FIRST}),{first_week}{FIRST}=0)", "EAF2FC"),
            (f"AND(ISNUMBER({first_week}{FIRST}),{first_week}{FIRST}>0,{first_week}{FIRST}<0.85)", "DFF2E9"),
            (f"AND(ISNUMBER({first_week}{FIRST}),{first_week}{FIRST}>=0.85,{first_week}{FIRST}<=1)", "FFF1CF"),
            (f'OR({first_week}{FIRST}="Over",AND(ISNUMBER({first_week}{FIRST}),{first_week}{FIRST}>1))', "F9DADD"),
        )
        for condition, rgb in heat:
            sheet.conditional_formatting.add(
                f"{first_week}{FIRST}:{col(last)}{LAST}",
                FormulaRule(formula=[condition], fill=PatternFill(bgColor="FF" + rgb),
                            border=TABLE_BORDER, stopIfTrue=True))
    else:
        stripes_and_totals(sheet, len(headers) + 1, last, formula=f'{first_week}$14<>""')


def overview(sheet):
    rules = retire_view(sheet)
    header(sheet, 7, "Filtrér medarbejderne og vælg en periode. Orange viser periodens uger i årsoversigten.")
    labels = ("Planår", "Startuge", "Periode", "Område", "Faggruppe", "Afdeling")
    original = ("B5", "D5", "F5", "I5", "K5", "M5")
    model = ("B65", "D65", "F65", "I65", "K65", "M65")
    controls(sheet, [(label, sheet[address].value, rules[old])
                     for label, address, old in zip(labels, model, original)])
    for index, address in enumerate(model, 1):
        sheet[address] = "=" + col(index) + "6"
    for row in range(70, 100):
        sheet[f"BJ{row}"] = f'=IF(BI{row},{row - 69},"")'
    sheet["BJ61"] = "=COUNT(BJ70:BJ99)"
    cards(sheet, [(label, "=" + address, fmt) for label, address, fmt in (
        ("Medarbejdere", "BJ61", "0"), ("Bookede timer\n(t)", "B62", HOURS),
        ("Disponibel\nkapacitet (t)", "D62", HOURS), ("Ledig kapacitet (t)", "F62", HOURS),
        ("Overbelastning (t)", "H62", HOURS), ("Bruttotilbud (t)", "J62", HOURS),
        ("Vægtede tilbud (t)", "L62", HOURS), ("Samlet\nbelægning", "N62", "0.0%"),
    )])
    legend(sheet, 10, {1: ("0 %", "EAF2FC"), 2: ("Under 85 %", "DFF2E9"),
                       3: ("85–100 %", "FFF1CF"), 4: ("Over 100 % / Over", "F9DADD"),
                       5: ("Valgt periode", ORANGE)})
    sheet["E10"].font = font(WHITE, 10, True)
    calendar(sheet, 8, 60, annual=True)
    report_rows(sheet, ("Medarbejder", "Område", "Faggruppe", "Afdeling", "Normtid (t)",
                        "Booket i perioden (t)", "Kapacitet i perioden (t)"), 60, annual=True)
    for index, width in enumerate((25, 15, 19, 20, 18, 21, 23, 16), 1):
        sheet.column_dimensions[col(index)].width = width
    sheet["A11"] = '=IF($BJ$61=0,"Ingen medarbejdere matcher filtrene","Årsoversigt")'
    sheet["A11"].font = font(NAVY, 11, True)
    freeze(sheet, "H")


def escaped_key(address):
    return f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({address},"~","~~"),"*","~*"),"?","~?")'


def project_options(workbook):
    registry = workbook["Projektregister"]
    selected = "'Projektoversigt'!$A$6"
    for row in range(7, 107):
        registry[f"J{row}"] = f'=IF(A{row}="","",ROW()-6)'
        registry[f"K{row}"] = f'=IFERROR(INDEX($A$7:$A$106,SMALL($J$7:$J$106,ROW()-6))&"","")'
        name = f"INDEX($B$7:$B$106,SMALL($J$7:$J$106,ROW()-6))"
        registry[f"L{row}"] = (
            f'=IF(K{row}="","",IF({name}="","Uden projektnavn",{name})&"{SEPARATOR}"&K{row})'
        )
        suffix = f'"{SEPARATOR}"&K{row}'
        registry[f"M{row}"] = f'=IF(AND(K{row}<>"",RIGHT({selected},LEN({suffix}))={suffix}),LEN(K{row}),0)'
    for letter in "JKLM":
        registry.column_dimensions[letter].hidden = True
    workbook.defined_names.add(DefinedName(
        "Projektvisninger", attr_text="'Projektregister'!$L$7:"
        "INDEX('Projektregister'!$L$7:$L$106,MAX(1,COUNT('Projektregister'!$J$7:$J$106)))"
    ))


def project_report(sheet, registry):
    rules = retire_view(sheet)
    header(sheet, 6, "Vælg projekt og periode. Oversigten viser de medarbejdere, der har timer på projektet i perioden.")
    original = sheet["B62"].value
    initial = next((str(registry[f"B{r}"].value or "Uden projektnavn") + SEPARATOR + str(original)
                    for r in range(7, 107)
                    if str(registry[f"A{r}"].value) == str(original)), original)
    project_rule = rules["B2"]
    project_rule.formula1 = "Projektvisninger"
    controls(sheet, [("Projekt", initial, project_rule),
                     ("Planår", sheet["B65"].value, rules["B5"]),
                     ("Startuge", sheet["D65"].value, rules["D5"]),
                     ("Periode", sheet["F65"].value, rules["F5"])])
    if initial is not None:
        sheet["A6"].data_type = "s"
    sheet["A6"].number_format = "@"
    codes = "Projektregister!$K$7:$K$106"
    labels = "Projektregister!$L$7:$L$106"
    # Exact lookup handles live selections. For a retained old caption after
    # renaming, prefer the longest registered key matching its suffix.
    matches = "Projektregister!$M$7:$M$106"
    suffix = f'IF(MAX({matches})=0,A6,INDEX({codes},MATCH(MAX({matches}),{matches},0)))'
    sheet["B62"] = f'=IF(A6="","",IFERROR(INDEX({codes},MATCH({escaped_key("A6")},{labels},0)),{suffix}))'
    for address, value in (("B65", "B6"), ("D65", "C6"), ("F65", "D6")):
        sheet[address] = "=" + value
    sheet["BJ61"] = "=COUNT(BI105:BI134)"
    current = f'=IF(B62="","",IFERROR(INDEX({labels},MATCH({escaped_key("B62")},{codes},0)),B62))'
    cards(sheet, [("Aktuelt projekt", current, "General"),
                  ("Medarbejdere", "=BJ61", "0"), ("Periodetimer (t)", '=IF(B62="","",F100)', HOURS)])
    sheet["A11"] = ('=IF(B62="","Vælg et projekt",IF(BJ61=0,'
                    '"Ingen projekttimer i perioden","Periodens projektplan"))')
    sheet["A11"].font = font(NAVY, 11, True)
    calendar(sheet, 7, 59, annual=False)
    report_rows(sheet, ("Medarbejder", "Område", "Faggruppe", "Afdeling", "Normtid (t)",
                        "Periodetimer (t)"), 59, annual=False)
    for index, width in enumerate((42, 15, 18, 18, 14, 21), 1):
        sheet.column_dimensions[col(index)].width = width
    freeze(sheet, "G")


def employee_plan(sheet):
    header(sheet, 8,
           "Filtrér ét planår og én medarbejder i tabeloverskriften. Indtast timer i ugefelterne. "
           "I Teams kan du vælge Vis > Arkvisning > Ny. Undgå sortering.", note_height=28)
    sheet.unmerge_cells("A5:J5")
    sheet["A5"] = None
    legend(sheet, 5, {1: ("Sag", BLUE), 2: ("Tilbud", TYPE_COLORS["Tilbud"]),
                      4: ("Ferie / fravær", TYPE_COLORS["Ferie"]),
                      5: ("Kursus", TYPE_COLORS["Kursus"]),
                      6: ("Andet", TYPE_COLORS["Andet"])})
    sheet["H8"].number_format = ";;;"
    sheet["H6"] = sheet["H6"].value.replace(',"OK","STOP")', ',"","STOP")')
    sheet["A6"] = "Timer i synlig plan"
    sheet.sheet_format.defaultRowHeight = 19
    for row in range(6, 10):
        sheet.row_dimensions[row].height = 21
    for address in ("A6", "G6", "A7", "A8"):
        sheet[address].font = font(NAVY, 10, True)
    sheet["H6"].font = font(ORANGE, 10, True)
    for number in range(1, 62):
        format_label(sheet.cell(9, number), NAVY, centered=number >= 9, color=WHITE)
        if number < 9:
            sheet.cell(9, number).alignment = Alignment(vertical="center", indent=1)
        else:
            sheet.column_dimensions[col(number)].width = 10
            for row in (6, 7, 8):
                sheet.cell(row, number).font = font(NAVY, 10, True)
    body_font = font()
    caption_alignment = Alignment(horizontal="left", vertical="center", indent=1)
    number_alignment = Alignment(horizontal="right", vertical="center", indent=1)
    for row in sheet.iter_rows(min_row=10, max_row=2169, max_col=61):
        for cell in row:
            cell.font = body_font
            cell.alignment = number_alignment if cell.column == 7 or cell.column >= 8 else caption_alignment
    # Retain the accepted fill rules and add thin borders without changing their precedence.
    for rule in sheet.conditional_formatting["A10:BI2169"]:
        rule.dxf.border = TABLE_BORDER
    # Ugyldig uge 53 has an earlier stopping rule, so give its cells the same border.
    sheet.conditional_formatting["BI10:BI2169"][0].dxf.border = TABLE_BORDER


def registry_layout(sheet, name, widths, note):
    header(sheet, 6, note)
    sheet.row_dimensions[6].height = 24
    table = sheet.tables[name]
    table.tableStyleInfo.name = None
    table.tableStyleInfo.showRowStripes = False
    for index, width in enumerate(widths, 1):
        sheet.column_dimensions[col(index)].width = width
        format_label(sheet.cell(6, index), NAVY, color=WHITE)
    for row in range(7, sheet.max_row + 1):
        sheet.row_dimensions[row].height = 20
        for index in range(1, len(widths) + 1):
            cell = sheet.cell(row, index)
            cell.fill = fill(WHITE if row % 2 == 1 else BLUE)
            cell.border = TABLE_BORDER
            cell.font = font()
            cell.alignment = Alignment(vertical="center", indent=1)
    sheet.freeze_panes = "A7"
    sheet.sheet_view.selection = [Selection(pane="bottomLeft", activeCell="A7", sqref="A7")]


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    workbook = load_workbook(source)
    if not workbook["Min ressourceplan"]["A1"].value.endswith("3.3.1"):
        raise ValueError("Migrationen kræver Ressourceplan 3.3.1.")
    employee_plan(workbook["Min ressourceplan"])
    overview(workbook["Ressourceoverblik"])
    registry = workbook["Projektregister"]
    registry_layout(registry, "SagRegister", (20, 47),
                    "Opret projekter i ledige rækker. Behold sagsnummeret, når du ændrer et projektnavn.")
    project_options(workbook)
    project_report(workbook["Projektoversigt"], registry)
    registry_layout(workbook["Medarbejdere"], "PersonRegister", (12, 29, 15, 19, 18, 15),
                    "Udfyld en ledig medarbejderplads. Behold det tildelte ID og rækkefølgen.")
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / destination.name
        workbook.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_3_1.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_4.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
