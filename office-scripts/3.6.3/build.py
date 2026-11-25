"""Saml selvstændige TypeScript-filer til Excels Office Scripts-editor."""

from pathlib import Path

ROOT = Path(__file__).resolve().parent


def build():
    shared = (ROOT / "src/common.ts").read_text(encoding="utf-8")
    standard = (ROOT / "src/standard.ts").read_text(encoding="utf-8")
    selection = (ROOT / "src/selection.ts").read_text(encoding="utf-8")
    for name in ("create_employee", "add_project", "initialize_plan_year", "upgrade_workbook"):
        parts = [(ROOT / "src" / f"{name}.ts").read_text(encoding="utf-8"), shared]
        if name in ("add_project", "initialize_plan_year"):
            parts.append(selection)
        if name in ("create_employee", "initialize_plan_year"):
            parts.append(standard)
        (ROOT / f"{name}.ts").write_text(
            "// Genereret af office-scripts/3.6.3/build.py. Kopiér hele filen til Excel.\n"
            + "\n".join(parts), encoding="utf-8"
        )


if __name__ == "__main__":
    build()
