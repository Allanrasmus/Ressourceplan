type PlanCellSelection = { sheet: string; row: number; column: number };

function selectedCell(workbook: ExcelScript.Workbook): PlanCellSelection {
  let selected: ExcelScript.Range;
  try {
    selected = workbook.getSelectedRange();
  } catch (error: unknown) {
    throw new Error("Markér præcis én celle i tabellen, eller angiv alle valg i dialogen.");
  }
  if (selected.getRowCount() !== 1 || selected.getColumnCount() !== 1 ||
      selected.getRowHidden() === true) {
    throw new Error("Markér præcis én synlig celle i tabellen, eller angiv alle valg i dialogen.");
  }
  return {
    sheet: selected.getWorksheet().getName(), row: selected.getRowIndex(),
    column: selected.getColumnIndex(),
  };
}

function selectedRecord(table: ExcelScript.Table, cell: PlanCellSelection): Cell[] | undefined {
  const first = table.getHeaderRowRange().getRowIndex() + 1;
  const index = cell.row - first;
  if (index < 0 || index >= table.getRowCount()) return undefined;
  return table.getRangeBetweenHeaderAndTotal().getRow(index).getValues()[0];
}

function planTarget(plan: Plan, people: Entity[], employee: string, year: number | undefined,
  cell: PlanCellSelection | undefined): { person: Entity; year: number } {
  const row = cell && cell.sheet === "Min ressourceplan" &&
    (cell.column === field(plan, "Medarbejder") || cell.column === field(plan, "Projektvalg"))
    ? selectedRecord(plan.table, cell) : undefined;
  if ((!employee.trim() || year === undefined) && !row) {
    throw new Error("Angiv medarbejder og planår, eller markér Medarbejder eller Projektvalg i en synlig planrække.");
  }
  const key = row ? recordKey(row[field(plan, "Medarbejder")], people) : "";
  const person = findEntity(employee.trim() ? employee : key, people, "Medarbejder");
  if (year === undefined) {
    const rowYear = row ? row[field(plan, "Planår")] : "";
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
