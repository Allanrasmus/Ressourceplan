"""Byg produktionsfilen 3.6.1 uden at ændre beregningsmodellen fra 3.6."""

import argparse
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook


def migrate(source, destination):
    if source.resolve() == destination.resolve() or destination.exists():
        raise ValueError("Vælg en ny uddatafil. Eksisterende filer overskrives ikke.")
    book = load_workbook(source)
    plan = book["Min ressourceplan"]
    if not plan["A1"].value.endswith("3.6"):
        raise ValueError("Migrationen kræver Ressourceplan 3.6.")
    for sheet in book:
        sheet["A1"] = "COWI  |  Ressourceplan 3.6.1"
    plan["A4"] = ("Tilføj projekt eller Opret nyt planår via Automatiser. Markér Medarbejder "
                  "i en planrække for at genbruge navn og år. Indtast timer direkte i ugefelterne.")
    book["Planstart"]["P6"] = "PlanTimer-formler 3.6.1"
    with TemporaryDirectory(dir=destination.parent) as directory:
        staged = Path(directory) / destination.name
        book.save(staged)
        destination.hardlink_to(staged)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, default=Path("Ressourceplan_3_6.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("Ressourceplan_3_6_1.xlsx"))
    args = parser.parse_args()
    migrate(args.source, args.output)


if __name__ == "__main__":
    main()
