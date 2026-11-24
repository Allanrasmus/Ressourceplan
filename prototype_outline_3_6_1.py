"""Byg en særskilt, lokal outline-prøve fra Ressourceplan 3.6.1."""

import argparse
from collections import defaultdict
from datetime import date
from itertools import groupby
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter as col
from openpyxl.worksheet.table import Table, TableStyleInfo

from migrate_3_3 import NAVY, ORANGE, QUARTERS, TYPE_COLORS, fill, validation
from migrate_3_3_1 import set_panes
from migrate_3_4 import BLUE, MINT, TABLE_BORDER, WHITE, escaped_key, font, format_label, header
from migrate_3_5 import table_field

YEARLY_COL = 6


def quarter_segments(year):
    mondays = [date.fromisocalendar(year, week, 1)
               for week in range(1, date(year, 12, 28).isocalendar().week + 1)]
    return [(calendar_year, quarter, days)
            for (calendar_year, quarter), dates in groupby(
                mondays, key=lambda monday: (monday.year, (monday.month - 1) // 3 + 1))
            if (days := list(dates))]


def find_key(caption, labels):
    if not caption:
        return ""
    value = str(caption).casefold()
    exact = next((key for key, name in labels if f"{name} · {key}".casefold() == value), "")
    return exact or max((key for key, _ in labels if value.endswith(f" · {key}".casefold())),
                        key=len, default="")


def records(book, year, filters):
    staff = [(str(key), name, area, group, department, hours)
             for key, name, area, group, department, hours in book["Medarbejdere"].iter_rows(
                 min_row=7, max_col=6, max_row=106, values_only=True) if key and name]
    projects = [(str(key), name) for key, name in book["Projektregister"].iter_rows(
                min_row=7, max_col=2, max_row=506, values_only=True) if key]
    employees = [(key, name) for key, name, *_ in staff]
    selected = [row for row in staff if all(
        not restriction or str(value).casefold() == restriction.casefold()
        for value, restriction in zip(row[2:5], filters)
    )]
    active = {row[0] for row in selected}
    rows = defaultdict(list)
    for row in book["Min ressourceplan"].iter_rows(min_row=10, max_col=65, values_only=True):
        if str(row[0]) == str(year):
            worker = find_key(row[1], employees)
            if worker in active:
                rows[worker].append((row, find_key(row[4], projects)))
    return selected, dict(projects), rows


def calendar(sheet, year, selected_weeks=()):
    """Separate adjacent outline ranges with permanent quarter-total columns."""
    weeks, subtotals = {}, {}
    column = YEARLY_COL + 1
    for calendar_year, quarter, dates in quarter_segments(year):
        start = column
        for monday in dates:
            week = monday.isocalendar().week
            weeks[week] = column
            for row, value in ((7, f"Q{quarter}"), (8, monday), (9, f"Uge {week:02}")):
                cell = sheet.cell(row, column, value)
                format_label(cell, QUARTERS[quarter - 1] if row == 7 else NAVY if row == 9 else "F6F7F8",
                             centered=True, color=WHITE if row == 9 else NAVY)
                if row == 8:
                    cell.number_format = "dd.mm.yy"
                elif row == 9 and week in selected_weeks:
                    format_label(cell, ORANGE, centered=True, color=WHITE)
            sheet.column_dimensions[col(column)].width = 11
            column += 1
        sheet.column_dimensions.group(col(start), col(column - 1), outline_level=1, hidden=False)
        subtotals[column] = (start, column - 1)
        format_label(sheet.cell(9, column, f"Q{quarter} {calendar_year} i alt"), MINT, centered=True)
        sheet.column_dimensions[col(column)].width = 16
        column += 1
    sheet.sheet_properties.outlinePr.summaryRight = True
    sheet.sheet_format.outlineLevelCol = 1
    sheet.sheet_view.showOutlineSymbols = True
    set_panes(sheet, "G", 10)
    return weeks, subtotals, column - 1


def hourly_style(cell, color, format_="0.##"):
    cell.fill = fill(color)
    cell.font = font()
    cell.border = TABLE_BORDER
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.number_format = format_


def write_sums(sheet, row, subtotals, color):
    for target, (first, last) in subtotals.items():
        hourly_style(sheet.cell(row, target, f"=SUM({col(first)}{row}:{col(last)}{row})"), color)
    cells = ",".join(f"{col(i)}{row}" for i in subtotals)
    hourly_style(sheet.cell(row, YEARLY_COL, f"=SUM({cells})"), color)


def quarter_demo(book, year, staff, plan_rows):
    sheet = book.create_sheet("Kvartalsprototype")
    widths = [12, 31, 13, 32, 14, 16]
    labels = ["Planår", "Medarbejder", "Type", "Projektvalg", "Tilbud %", "Året i alt"]
    header(sheet, 6, "Lokal prøve. Ugefelterne er redigerbare kopier. Fold grupper med plus og minus.")
    sheet["A1"] = "COWI  |  Gruppeprototype 3.6.1"
    weeks, subtotals, last = calendar(sheet, year)
    for column, (width, label) in enumerate(zip(widths, labels), 1):
        sheet.column_dimensions[col(column)].width = width
        format_label(sheet.cell(9, column, label), NAVY, color=WHITE)
    row = 10
    for worker, *_ in staff:
        for record, _ in plan_rows.get(worker, ()):
            activity = record[3]
            color = TYPE_COLORS.get(activity, WHITE if row % 2 == 0 else BLUE)
            for number, value in enumerate((record[0], record[1], activity, record[4], record[7]), 1):
                target = sheet.cell(row, number, value)
                if isinstance(value, str):
                    target.data_type = "s"
                format_label(target, color, bold=False)
            sheet[f"E{row}"].number_format = "0%"
            for week, number in weeks.items():
                hours = record[week + 8]
                if not isinstance(hours, (int, float)):
                    hours = None
                hourly_style(sheet.cell(row, number, hours), color)
            write_sums(sheet, row, subtotals, color)
            row += 1
    if row == 10:
        sheet["A6"] = "Ingen planrækker for dette udsnit"
    else:
        table = Table(displayName="Kvartalsprøve", ref=f"A9:{col(last)}{row - 1}")
        table.tableStyleInfo = TableStyleInfo(showRowStripes=False)
        sheet.add_table(table)
        sheet["A6"] = "Timer i synlige rækker"
        sheet["F6"] = "=SUBTOTAL(109,Kvartalsprøve[Året i alt])"
        for first, last_week in subtotals.values():
            validation(sheet, f"{col(first)}10:{col(last_week)}{row-1}", "decimal", "0", "80")
    sheet.row_dimensions[6].height = 20


def excel_text(value):
    return '"' + str(value).replace('"', '""') + '"'


def project_hours(worker, year, week, types, key=None):
    filters = [
        f'--({table_field("ID")}&""={excel_text(worker)})',
        f'--({table_field("Planår")}&""="{year}")',
        "(" + "+".join(f'({table_field("Type")}={excel_text(kind)})' for kind in types) + ")",
    ]
    if key is not None:
        filters.append(f'--({table_field("Sagsnr.")}&""={excel_text(key)})')
    return f'SUMPRODUCT({table_field(f"Uge {week:02}")},{",".join(filters)})'


def ratio(booked, capacity):
    return f'IF({capacity}=0,IF({booked}>0,"Over",0),{booked}/{capacity})'


def period_highlight(book, year):
    overview = book["Ressourceoverblik"]
    period = str(overview["C6"].value or "")
    count = date(year, 12, 28).isocalendar().week
    if period == "Hele året":
        return set(range(1, count + 1))
    if period in ("Q1", "Q2", "Q3", "Q4"):
        return {day.isocalendar().week for _, quarter, days in quarter_segments(year)
                if period == f"Q{quarter}" for day in days}
    length = {"4 uger": 4, "8 uger": 8, "12 uger": 12, "26 uger": 26}.get(period, 0)
    start = min(int(overview["B6"].value or 1), count)
    return set(range(start, min(count + 1, start + length)))


def employee_demo(book, year, staff, projects, plan_rows):
    sheet = book.create_sheet("Medarbejderprototype")
    # Keep the filtered list stable. Detail formulas still match the source by
    # key when PlanTimer is filtered, sorted or augmented.
    header(sheet, 6, "Udsnittet vælges ved bygning. Medarbejdergrupperne følger faste ID'er og det angivne planår.")
    sheet["A1"] = "COWI  |  Gruppeprototype 3.6.1"
    selected_weeks = period_highlight(book, year)
    weeks, subtotals, last = calendar(sheet, year, selected_weeks)
    period_column = col(last + 1)
    sheet.column_dimensions[period_column].hidden = True
    for column, (width, label) in enumerate(zip((27, 13, 17, 17, 39, 15),
                                              ("Medarbejder", "Område", "Faggruppe", "Afdeling", "Detalje", "Året i alt")), 1):
        sheet.column_dimensions[col(column)].width = width
        format_label(sheet.cell(9, column, label), NAVY, color=WHITE)
    sheet["A6"] = f"Planår {year} · {len(staff)} medarbejdere"
    sheet["A6"].font = font(NAVY, 10, True)
    sheet.sheet_properties.outlinePr.summaryBelow = False
    sheet.sheet_format.outlineLevelRow = 1
    row = 10
    blocks = []
    for key, name, area, group, department, norm in staff:
        summary = row
        relevant_projects = {(record[3], code) for record, code in plan_rows.get(key, ())
                             if record[3] in ("Sag", "Tilbud") and (code or any(record[9:62]))}
        details = [("Normtid", "Normtid", ""), ("Disponibel kapacitet", "Kapacitet", "")]
        details.extend((
            f"{kind} · {projects.get(code, code) if code else 'uden projekt'}"
            + (f" · {code}" if code and projects.get(code) else ""), kind, code)
            for kind, code in sorted(relevant_projects))
        details.extend((kind, kind, "") for kind in ("Ferie", "Fravær", "Kursus", "Andet",
                                                      "Bookede timer", "Ledig kapacitet", "Overbelastning"))
        capacity, booked = summary + 2, summary + len(details) - 2
        blocks.append((booked, capacity))
        norm_lookup = (f'IFERROR(N(INDEX(PersonRegister[Normtid],'
                       f'MATCH({escaped_key(excel_text(key))},INDEX(PersonRegister[ID]&"",0),0))),0)')
        for number, value in enumerate((name, area, group, department, "Belægning"), 1):
            target = sheet.cell(summary, number, value)
            if isinstance(value, str):
                target.data_type = "s"
            format_label(target, MINT)
        for index, (label, kind, code) in enumerate(details, summary + 1):
            color = TYPE_COLORS.get(kind, WHITE if index % 2 else BLUE)
            format_label(sheet.cell(index, 5, label), color, bold=False)
            for week, column in weeks.items():
                letter = col(column)
                if kind == "Normtid":
                    formula = norm_lookup
                elif kind == "Kapacitet":
                    formula = f'MAX(0,{letter}{summary+1}-{project_hours(key, year, week, ("Ferie","Fravær","Kursus"))})'
                elif kind == "Bookede timer":
                    formula = project_hours(key, year, week, ("Sag", "Andet"))
                elif kind == "Ledig kapacitet":
                    formula = f"MAX(0,{letter}{capacity}-{letter}{booked})"
                elif kind == "Overbelastning":
                    formula = f"MAX(0,{letter}{booked}-{letter}{capacity})"
                else:
                    formula = project_hours(key, year, week, (kind,), code if kind in ("Sag", "Tilbud") else None)
                hourly_style(sheet.cell(index, column, "=" + formula), color)
            write_sums(sheet, index, subtotals, color)
        for column in [YEARLY_COL, *weeks.values(), *subtotals]:
            letter = col(column)
            cell = sheet.cell(summary, column, "=" + ratio(f"{letter}{booked}", f"{letter}{capacity}"))
            hourly_style(cell, MINT, "0%")
            if column in weeks.values():
                addr = cell.coordinate
                for predicate, color in (
                    (f"AND(ISNUMBER({addr}),{addr}=0)", "EAF2FC"),
                    (f"AND(ISNUMBER({addr}),{addr}>0,{addr}<0.85)", "DFF2E9"),
                    (f"AND(ISNUMBER({addr}),{addr}>=0.85,{addr}<=1)", "FFF1CF"),
                    (f'OR({addr}="Over",AND(ISNUMBER({addr}),{addr}>1))', "F9DADD"),
                ):
                    sheet.conditional_formatting.add(addr, FormulaRule(
                        formula=[predicate], fill=PatternFill(bgColor="FF" + color), stopIfTrue=True))
        row = summary + len(details) + 1
        sheet.row_dimensions.group(summary + 1, row - 1, hidden=True)
        sheet.row_dimensions[summary].collapsed = True
        period_cells = [f"{col(column)}{booked}" for week, column in weeks.items()
                        if week in selected_weeks]
        sheet[f"{period_column}{summary}"] = "=" + ("SUM(" + ",".join(period_cells) + ")"
                                                  if period_cells else "0")
    sheet[f"{period_column}9"] = "Booket i perioden"
    format_label(sheet.cell(row, 1, "I alt for udsnittet"), MINT)
    for column in [YEARLY_COL, *weeks.values(), *subtotals]:
        letter = col(column)
        sums = ["SUM(" + ",".join(f"{letter}{block[i]}" for block in blocks) + ")" if blocks else "0"
                for i in (0, 1)]
        hourly_style(sheet.cell(row, column, "=" + ratio(*sums)), MINT, "0%")
    format_label(sheet["E6"], "F6F7F8")
    sheet["E6"] = "Booket i orange uger"
    hourly_style(sheet.cell(6, 6, f"=SUM({period_column}10:{period_column}{row})"), MINT)


def build(source, destination, year, filters=("", "", "")):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny prototypefil.")
    book = load_workbook(source)
    supported = {book["Planstart"][f"F{row}"].value for row in range(7, 27)}
    if year not in supported:
        raise ValueError("Planåret skal være registreret i kildefilen.")
    staff, projects, plan_rows = records(book, year, filters)
    quarter_demo(book, year, staff, plan_rows)
    employee_demo(book, year, staff, projects, plan_rows)
    book.active = book["Kvartalsprototype"]
    with TemporaryDirectory(dir=destination.parent) as folder:
        staged = Path(folder) / destination.name
        book.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_6_1.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_6_1_Gruppeprototype.xlsx"))
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--region", default="")
    parser.add_argument("--group", default="")
    parser.add_argument("--department", default="")
    args = parser.parse_args()
    build(args.source, args.output, args.year, (args.region, args.group, args.department))


if __name__ == "__main__":
    main()
