"""Saml scripts til Excel og byg et reproducerbart distributionsarkiv."""

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

ROOT = Path(__file__).resolve().parent
SCRIPTS = {
    "create_employee": ("common", "filters", "standard"),
    "add_project": ("common", "filters", "selection"),
    "initialize_plan_year": ("common", "filters", "selection", "standard"),
}


def build():
    for name, modules in SCRIPTS.items():
        parts = [(ROOT / "src" / f"{module}.ts").read_text(encoding="utf-8")
                 for module in (name, *modules)]
        (ROOT / f"{name}.ts").write_text(
            "// Genereret af office-scripts/3.6.6/build.py. Kopiér hele filen til Excel.\n"
            + "\n".join(parts), encoding="utf-8"
        )


def package():
    paths = ("README.md", *(f"{name}.ts" for name in SCRIPTS))
    with ZipFile(ROOT / "Ressourceplan_Office_Scripts_3_6_6.zip", "w") as archive:
        for name in paths:
            archive.writestr(ZipInfo(name), (ROOT / name).read_bytes(), compress_type=ZIP_DEFLATED)


if __name__ == "__main__":
    build()
    package()
