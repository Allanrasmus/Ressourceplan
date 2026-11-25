type PlanCellSelection = { sheet: string; row: number; column: number };

function selectedCell(workbook: ExcelScript.Workbook): PlanCellSelection {
  try {
    const selected = workbook.getSelectedRange();
    if (selected.getRowCount() !== 1 || selected.getColumnCount() !== 1 ||
        selected.getRowHidden() === true) {
      throw new Error("Markeringen er ikke én synlig celle.");
    }
    return {
      sheet: selected.getWorksheet().getName(), row: selected.getRowIndex(),
      column: selected.getColumnIndex(),
    };
  } catch (error: unknown) {
    throw new Error("Markér præcis én synlig celle i tabellen, eller angiv alle valg i dialogen.");
  }
}

function selectedRecord(workbook: ExcelScript.Workbook, name: string, row: number, width: number): Cell[] | undefined {
  const table = workbook.getTable(name);
  if (!table) throw new Error("Tabellen " + name + " mangler.");
  const first = table.getHeaderRowRange().getRowIndex() + 1;
  const index = row - first;
  if (index < 0 || index >= table.getRowCount()) return undefined;
  return table.getRangeBetweenHeaderAndTotal().getCell(index, 0).getResizedRange(0, width - 1).getValues()[0];
}

function selectedPlanRow(workbook: ExcelScript.Workbook, cell: PlanCellSelection | undefined): Cell[] | undefined {
  if (cell && cell.sheet === "Min ressourceplan" && (cell.column === 1 || cell.column === 4)) {
    return selectedRecord(workbook, "PlanTimer", cell.row, 5);
  }
  return undefined;
}

function planTarget(people: Entity[], employee: string, year: number | undefined,
  row: Cell[] | undefined): { person: Entity; year: number } {
  if ((!employee.trim() || year === undefined) && !row) {
    throw new Error("Angiv medarbejder og planår, eller markér Medarbejder eller Projektvalg i en synlig planrække.");
  }
  const key = row ? recordKey(row[1], people) : "";
  const person = findEntity(employee.trim() ? employee : key, people, "Medarbejder");
  if (year === undefined) {
    const rowYear = row ? row[0] : "";
    if (!same(key, person.key) || rowYear === "" || !Number.isInteger(Number(rowYear))) {
      throw new Error("Året kan kun hentes fra en række for den valgte medarbejder. Angiv planår i dialogen.");
    }
    return { person, year: Number(rowYear) };
  }
  return { person, year };
}

function optionalYear(input: string | undefined): number | undefined {
  const raw = input == null ? "" : String(input).trim();
  if (!raw) return undefined;
  if (!/^\d{4}$/.test(raw) || Number(raw) < 1901 || Number(raw) > 9998) {
    throw new Error("Planår skal være et helt år mellem 1901 og 9998.");
  }
  return Number(raw);
}
