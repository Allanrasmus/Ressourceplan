"""Byg en ny 3.6.7-kopi af demonstrationsfilen 3.6.3."""

from argparse import ArgumentParser
from pathlib import Path

from openpyxl import load_workbook

from migrate_3_5 import GROSS, PEOPLE, SLOT_START, WEEKS, YEAR, col, key_masks

OLD_BANNER = "COWI  |  Ressourceplan 3.6.3"
NEW_BANNER = "COWI  |  Ressourceplan 3.6.7"


def gross_formulas(slot, week):
    field = WEEKS[week]
    column = col(week + 8)
    worker = SLOT_START + slot
    masks = key_masks(worker, YEAR)
    prefix = f'=IF(OR($BI{worker}="",{column}$189=""),0,SUMPRODUCT(PlanTimer[{field}],{masks},'
    condition = '(PlanTimer[Type]="Tilbud")'
    return prefix + f"({condition})))", prefix + f"--{condition}))"


def migrate(source, output):
    if output.exists():
        raise ValueError(f"Outputfilen findes allerede: {output}")
    book = load_workbook(source)
    overview = book["Ressourceoverblik"]
    if any(sheet["A1"].value != OLD_BANNER for sheet in book.worksheets):
        raise ValueError("Migrationen kræver demonstrationsprojektmappen 3.6.3.")
    for slot in range(PEOPLE):
        for week in range(len(WEEKS)):
            cell = overview.cell(GROSS + slot, week + 8)
            old, new = gross_formulas(slot, week)
            if cell.value != old:
                raise ValueError(f"Bruttotilbudsformlen i {cell.coordinate} afviger fra 3.6.3.")
            cell.value = new
    for sheet in book.worksheets:
        sheet["A1"] = NEW_BANNER
    book.save(output)


if __name__ == "__main__":
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_6_3.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_6_7.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)
