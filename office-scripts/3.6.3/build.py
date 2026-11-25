"""Saml selvstændige TypeScript-filer til Excels Office Scripts-editor."""

from argparse import ArgumentParser
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parent
PROJECT = ROOT.parents[1]
SCRIPTS = ("create_employee", "add_project", "initialize_plan_year", "upgrade_workbook")


def build():
    shared = (ROOT / "src/common.ts").read_text(encoding="utf-8")
    standard = (ROOT / "src/standard.ts").read_text(encoding="utf-8")
    selection = (ROOT / "src/selection.ts").read_text(encoding="utf-8")
    for name in SCRIPTS:
        parts = [(ROOT / "src" / f"{name}.ts").read_text(encoding="utf-8"), shared]
        if name in ("add_project", "initialize_plan_year"):
            parts.append(selection)
        if name in ("create_employee", "initialize_plan_year"):
            parts.append(standard)
        (ROOT / f"{name}.ts").write_text(
            "// Genereret af office-scripts/3.6.3/build.py. Kopiér hele filen til Excel.\n"
            + "\n".join(parts), encoding="utf-8"
        )


def package(output):
    inputs = {
        "Ressourceplan_3_6_3.xlsx": PROJECT / "Ressourceplan_3_6_3.xlsx",
        "INSTALLATION.md": ROOT / "README.md",
        **{f"{name}.ts": ROOT / f"{name}.ts" for name in SCRIPTS},
    }
    contents = {name: path.read_bytes() for name, path in inputs.items()}
    with ZipFile(output, "w") as archive:
        for name, data in contents.items():
            archive.writestr(ZipInfo(name), data, compress_type=ZIP_DEFLATED)


if __name__ == "__main__":
    parser = ArgumentParser(description=__doc__)
    parser.add_argument("--package", action="store_true", help="Byg også Ressourceplan_3_6_3.zip.")
    args = parser.parse_args()
    build()
    if args.package:
        package(PROJECT / "Ressourceplan_3_6_3.zip")
