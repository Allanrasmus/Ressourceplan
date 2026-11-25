"""Byg den komplette 3.6.3-projektmappe fra demonstrationsfilen i repositoryet."""

from argparse import ArgumentParser
from copy import copy, deepcopy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.workbook.defined_name import DefinedName

from migrate_3_4 import SEPARATOR
from migrate_3_5 import PLAN_STYLE_LAST

PLAN_AREA = f"A10:BJ{PLAN_STYLE_LAST}"
STRIPE_NAME = "PlanSagsstribe"
STRIPE_REFERENCE = f"'Planstart'!$R$10:$R${PLAN_STYLE_LAST}"
BLUE = "EAF2FC"
ACTIVITIES = ("Sag", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet")


def stripe_rule(parity):
    return f'AND($D10="Sag",INDEX({STRIPE_NAME},ROW()-9)={parity})'


def stripe_formula(row):
    activity = "INDEX(PlanTimer[Type],ROW()-9)"
    return (f'=IF(ROW()>PlanSidsteRaekke,0,MOD(N(R{row-1})+'
            f'IF({activity}="Sag",SUBTOTAL(103,{activity}),0),2))')


def order_demo_plan(sheet):
    """Gruppér demonstrationsrækker uden at ændre celleværdier eller filtre."""
    table = sheet.tables["PlanTimer"]
    left, header, right, last = range_boundaries(table.ref)
    formulas = {column.id: "=" + column.calculatedColumnFormula.attr_text
                for column in table.tableColumns if column.calculatedColumnFormula}
    groups = {}
    rows = sheet.iter_rows(min_row=header + 1, max_row=last, min_col=left, max_col=right)
    for row in rows:
        year, employee, _, activity = (cell.value for cell in row[:4])
        if not isinstance(year, int) or not isinstance(employee, str) or activity not in ACTIVITIES:
            raise ValueError("Demonstrationsplanen indeholder en ukendt type, medarbejder eller et ugyldigt år.")
        key = employee.rpartition(SEPARATOR)[2]
        if any(cell.hyperlink or cell.comment or
               (cell.data_type == "f" and cell.value != formulas.get(cell.column))
               for cell in row):
            raise ValueError("Planen indeholder tilpassede celler. Brug Office Script til arbejdsfiler.")
        cells = [(cell.value, cell.data_type, cell._style) for cell in row]
        groups.setdefault((year, key), []).append((activity, cells, copy(sheet.row_dimensions[row[0].row])))

    target = header + 1
    for group in groups.values():
        for _, cells, dimension in sorted(group, key=lambda row: ACTIVITIES.index(row[0])):
            for number, (value, data_type, style) in enumerate(cells, left):
                cell = sheet.cell(target, number)
                cell.value, cell.data_type, cell._style = value, data_type, style
            dimension.index = target
            sheet.row_dimensions[target] = dimension
            target += 1


def migrate(source, output):
    if output.exists():
        raise ValueError(f"Outputfilen findes allerede: {output}")
    workbook = load_workbook(source)
    config = workbook["Planstart"]
    if config["P6"].value not in ("PlanTimer-formler 3.6", "PlanTimer-formler 3.6.1"):
        raise ValueError("Migrationen kræver Ressourceplan 3.6 eller 3.6.1.")
    plan = workbook["Min ressourceplan"]
    if STRIPE_NAME in workbook.defined_names or any(
        config.cell(row, 18).value is not None for row in range(9, PLAN_STYLE_LAST + 1)
    ):
        raise ValueError("Planstart-kolonnen eller planstribernes navn er allerede i brug.")
    order_demo_plan(plan)
    rules = plan.conditional_formatting[PLAN_AREA]
    for parity, color in ((1, BLUE), (0, "FFFFFF")):
        old_formula = f'AND($D10="Sag",MOD(ROW(),2)={parity})'
        matches = [rule for rule in rules if rule.formula == [old_formula]]
        if len(matches) != 1:
            raise ValueError("De oprindelige sagsregler dækker ikke det forventede område.")
        rule = matches[0]
        rule.formula = [stripe_rule(parity)]
        rule.dxf = deepcopy(rule.dxf)
        rule.dxfId = None
        rule.dxf.fill.bgColor = "FF" + color

    workbook.defined_names.add(DefinedName(STRIPE_NAME, attr_text=STRIPE_REFERENCE))
    for row in range(10, PLAN_STYLE_LAST + 1):
        config.cell(row, 18, stripe_formula(row))

    for sheet_name, table_name, last_text_col in (
        ("Medarbejdere", "PersonRegister", 5), ("Projektregister", "SagRegister", 2)
    ):
        sheet = workbook[sheet_name]
        left, header, _, last = range_boundaries(sheet.tables[table_name].ref)
        for row in sheet.iter_rows(min_row=header + 1, max_row=last,
                                   min_col=left, max_col=last_text_col):
            for cell in row:
                cell.number_format = "@"
    workbook["Ressourceoverblik"]["F6"].number_format = "@"
    if plan["A5"].value == "Sag":
        fill = copy(plan["A5"].fill)
        fill.fgColor = "FF" + BLUE
        plan["A5"].fill = fill
    for sheet in workbook.worksheets:
        if str(sheet["A1"].value).startswith("COWI  |  Ressourceplan "):
            sheet["A1"] = "COWI  |  Ressourceplan 3.6.3"
    workbook.save(output)


if __name__ == "__main__":
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_6_1.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_6_3.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)
