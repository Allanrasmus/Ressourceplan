// Genereret af office-scripts/3.6.2/build.py. Kopiér hele filen til Excel.
/**
 * Tilføj Sag eller Tilbud til én medarbejder.
 * @param medarbejder Skriv et entydigt navn, navn · ID eller ID. Tomt henter medarbejderen fra markeret planrække.
 * @param planaar Tomt henter år fra den markerede planrække for samme medarbejder.
 * @param projekt Skriv navn, navn · sagsnummer eller sagsnummer. Tomt henter markeret Projektvalg eller registerrække.
 * @param tilbudsprocent Kun til Tilbud. Skriv 0–100. Tom tekst giver standarden på 50 %.
 */
function main(workbook: ExcelScript.Workbook, medarbejder: string = "", planaar?: string,
  projekt: string = "", aktivitetstype: "Sag" | "Tilbud" = "Sag", tilbudsprocent?: string): string {
  const activity = ["Sag", "Tilbud"].find(value => same(value, aktivitetstype.trim()));
  if (!activity) throw new Error("Aktivitetstype skal være Sag eller Tilbud.");
  const probability = activity === "Tilbud" ? offerProbability(tilbudsprocent) : "";
  const plan = planSchema(workbook);
  const people = entities(readRegister(workbook, "PersonRegister"));
  const projects = entities(readRegister(workbook, "SagRegister"));
  const inputYear = optionalYear(planaar);
  const personInput = medarbejder ?? "";
  const projectInput = projekt ?? "";
  const selection = !personInput.trim() || inputYear === undefined || !projectInput.trim()
    ? selectedCell(workbook) : undefined;
  const planRow = selectedPlanRow(workbook, selection);
  const { person, year } = planTarget(people, personInput, inputYear, planRow);
  let projectKey = projectInput;
  if (!projectKey.trim() && selection) {
    if (selection.sheet === "Projektregister" && (selection.column === 0 || selection.column === 1)) {
      const row = selectedRecord(workbook, "SagRegister", selection.row, 1);
      projectKey = row ? String(row[0]) : "";
    } else if (selection.sheet === "Min ressourceplan" && selection.column === field(plan, "Projektvalg")) {
      projectKey = planRow ? recordKey(planRow[4], projects) : "";
    }
    if (!projectKey.trim()) throw new Error("Vælg en projektcelle i tabellen, eller angiv projektet i dialogen.");
  }
  const caseRecord = findEntity(projectKey, projects, "Projekt");
  const newYear = yearSlot(workbook, year);
  let entries = planInputs(workbook);
  const isDuplicate = (row: Cell[]): boolean => same(row[0], year) &&
    same(row[3], activity) &&
    matchesRecord(row[1], person, people) &&
    matchesRecord(row[4], caseRecord, projects);
  const duplicateMessage = "Projektet findes allerede for medarbejderen, året og aktivitetstypen.";
  if (entries.some(row => isDuplicate(row))) return report(duplicateMessage);
  checkCapacity(entries.length, 1);
  if (newYear >= 0) {
    ensureYear(workbook, year);
    entries = planInputs(workbook);
    if (entries.some(row => isDuplicate(row))) return report(duplicateMessage);
  }
  append(workbook, plan, [{ year, employee: caption(person), type: activity,
    project: caption(caseRecord), probability }], entries.length);
  return report(caption(caseRecord) + " tilføjet som " + activity + " for " + caption(person) + " · " + year + ".");
}

function offerProbability(input: string | undefined): Cell {
  if (input == null) return "";
  const raw = String(input).trim();
  if (!raw) return "";
  const value = Number(raw.replace(",", "."));
  if (!/^\d+(?:[,.]\d+)?$/.test(raw) || !Number.isFinite(value) || value < 0 || value > 100) {
    throw new Error("Tilbudsprocent skal være 0–100. Lad feltet stå tomt for standardværdien.");
  }
  return value / 100;
}

type Cell = string | number | boolean;
type Entity = { key: string; name: string };
type Plan = {
  headers: string[];
  formulas: { index: number; formula: string }[];
};
type RecordInput = { year: number; employee: string; type: string; project: string; probability: Cell };

const separator = " · ";
const planLimit = 20000;

function required(text: string, field: string): string {
  const value = text.trim();
  if (!value) throw new Error(field + " skal udfyldes.");
  return value;
}

function same(left: Cell, right: Cell): boolean {
  return String(left).toLowerCase() === String(right).toLowerCase();
}

function literal(text: string): string {
  // Range/table writes can interpret a leading formula character as a formula.
  return /^[=+\-@']/.test(text) ? "'" + text : text;
}

function caption(entity: Entity): string {
  return entity.name + separator + entity.key;
}

function report(message: string): string {
  console.log(message);
  return message;
}

function readRegister(workbook: ExcelScript.Workbook, name: string): Cell[][] {
  const table = workbook.getTable(name);
  if (!table) throw new Error("Tabellen " + name + " mangler.");
  return table.getRowCount() ? table.getRangeBetweenHeaderAndTotal().getValues() : [];
}

function entities(rows: Cell[][]): Entity[] {
  return rows.filter(row => String(row[0]).trim() !== "" && String(row[1]).trim() !== "")
    .map(row => ({ key: String(row[0]), name: String(row[1]) }));
}

function findEntity(input: string, records: Entity[], label: string): Entity {
  const value = required(input, label);
  const exact = records.filter(entity => same(entity.key, value) || same(caption(entity), value));
  const matches = exact.length ? exact : records.filter(entity => same(entity.name, value));
  if (matches.length !== 1 ||
      records.filter(entity => same(entity.key, matches[0].key)).length !== 1) {
    throw new Error(label + " kunne ikke identificeres entydigt. Brug ID/sagsnummer eller det fulde valg.");
  }
  return matches[0];
}

function matchesRecord(label: Cell, entity: Entity, records: Entity[]): boolean {
  // A previously selected name may be stale. Resolve its longest known suffix,
  // matching the Excel formula even when a key contains the separator itself.
  return same(recordKey(label, records), entity.key);
}

function recordKey(label: Cell, records: Entity[]): string {
  const saved = String(label);
  const exact = records.find(record => same(caption(record), saved));
  if (exact) return exact.key;
  const candidates = records.filter(record => saved.toLowerCase().endsWith((separator + record.key).toLowerCase()));
  const longest = candidates.reduce((winner, record) =>
    record.key.length > winner.length ? record.key : winner, "");
  return longest;
}

function yearSlot(workbook: ExcelScript.Workbook, year: number): number {
  if (!Number.isInteger(year) || year < 1901 || year > 9998) {
    throw new Error("Planår skal være et helt år mellem 1901 og 9998.");
  }
  const sheet = workbook.getWorksheet("Planstart");
  if (!sheet) throw new Error("Planstart mangler.");
  const values = sheet.getRange("F7:F26").getValues();
  if (values.some(row => Number(row[0]) === year)) return -1;
  const slot = values.findIndex(row => row[0] === "");
  if (slot < 0) throw new Error("Planårslisten har ingen ledige pladser.");
  return slot;
}

function ensureYear(workbook: ExcelScript.Workbook, year: number): void {
  const slot = yearSlot(workbook, year);
  if (slot < 0) return;
  const sheet = workbook.getWorksheet("Planstart");
  if (!sheet) throw new Error("Planstart mangler.");
  sheet.getRange("F7:F26").getCell(slot, 0).setValues([[year]]);
}

function planSchema(workbook: ExcelScript.Workbook): Plan {
  const config = workbook.getWorksheet("Planstart");
  if (!config) throw new Error("Planstart mangler.");
  const marker = config.getRange("P6").getValue();
  if (marker !== "PlanTimer-formler 3.6.1" && marker !== "PlanTimer-formler 3.6") {
    throw new Error("Konfigurationen for Ressourceplan 3.6 eller 3.6.1 mangler.");
  }
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const header = table.getHeaderRowRange();
  const headers = header.getValues()[0].map(value => String(value));
  const fields: [number, string][] = [
    [0, "Planår"], [1, "Medarbejder"], [2, "ID"], [3, "Type"], [4, "Projektvalg"],
    [7, "Tilbud %"], [9, "Uge 01"], [61, "Uge 53"],
  ];
  const calculated = ["ID", "Projekt / aktivitet", "Timer i alt", "Aktuelt medarbejdernavn",
    "Sagsnr.", "Effektiv tilbudsprocent", "Timer gyldig uge 53"];
  const formulas = config.getRange("P7:Q13").getValues().map(row => ({
    index: headers.indexOf(String(row[0])), formula: String(row[1])
  }));
  if (header.getRowIndex() !== 8 || header.getColumnIndex() !== 0 || headers.length !== 65 ||
      fields.some(([position, name]) => headers[position] !== name) ||
      calculated.some(field => !formulas.some(item =>
        item.index === headers.indexOf(field) && item.index >= 0 && item.formula.startsWith("=")))) {
    throw new Error("PlanTimer eller dens formelskabelon stemmer ikke med version 3.6.");
  }
  return { headers, formulas };
}

function field(plan: Plan, name: string): number {
  return plan.headers.indexOf(name);
}

function planInputs(workbook: ExcelScript.Workbook): Cell[][] {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const count = table.getRowCount();
  const rows: Cell[][] = [];
  if (count === 0) return rows;
  const range = table.getRangeBetweenHeaderAndTotal();
  // Batch only A:E. The blocks bound each response for plans up to 20,000 rows.
  for (let first = 0; first < count; first += 4000) {
    const length = Math.min(4000, count - first);
    const values = range.getCell(first, 0).getResizedRange(length - 1, 4).getValues();
    values.forEach(row => { rows.push(row); });
  }
  return rows;
}

function checkCapacity(existing: number, additions: number): void {
  if (existing + additions > planLimit) {
    throw new Error("Handlingen overskrider de 20.000 planrækker med formatering og validering.");
  }
}

function formulaBlocks(plan: Plan): { index: number; formulas: string[] }[] {
  const blocks: { index: number; formulas: string[] }[] = [];
  const columns = plan.formulas.slice().sort((a, b) => a.index - b.index);
  for (const column of columns) {
    const last = blocks[blocks.length - 1];
    if (last && last.index + last.formulas.length === column.index) {
      last.formulas.push(column.formula);
    } else {
      blocks.push({ index: column.index, formulas: [column.formula] });
    }
  }
  return blocks;
}

function append(workbook: ExcelScript.Workbook, plan: Plan, inputs: RecordInput[], expectedRows: number): void {
  if (!inputs.length) return;
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const first = table.getRowCount();
  if (first !== expectedRows) throw new Error("Planens rækker ændrede sig. Prøv igen.");
  checkCapacity(first, inputs.length);
  const year = field(plan, "Planår");
  const worker = field(plan, "Medarbejder");
  const type = field(plan, "Type");
  const project = field(plan, "Projektvalg");
  const probability = field(plan, "Tilbud %");
  const payload = inputs.map(input => {
    const row: Cell[] = plan.headers.map(() => "");
    row[year] = input.year;
    row[worker] = literal(input.employee);
    row[type] = input.type;
    row[project] = literal(input.project);
    row[probability] = input.probability;
    // Include formulas in the initial write, then set them explicitly below.
    plan.formulas.forEach(item => { row[item.index] = item.formula; });
    return row;
  });
  try {
    table.addRows(-1, payload);
    if (table.getRowCount() !== first + inputs.length) {
      throw new Error("Planens rækker ændrede sig under oprettelsen.");
    }
    const range = table.getRangeBetweenHeaderAndTotal();
    const target = range.getCell(first, 0).getResizedRange(inputs.length - 1, plan.headers.length - 1);
    const actual = target.getCell(0, 0).getResizedRange(inputs.length - 1, probability).getValues();
    if (inputs.some((input, i) => !same(actual[i][year], input.year) ||
        !same(actual[i][worker], input.employee) || !same(actual[i][type], input.type) ||
        !same(actual[i][project], input.project) || !same(actual[i][probability], input.probability))) {
      throw new Error("Planens rækker ændrede sig under oprettelsen.");
    }
    for (const block of formulaBlocks(plan)) {
      target.getCell(0, block.index)
        .getResizedRange(inputs.length - 1, block.formulas.length - 1)
        .setFormulas(inputs.map(() => block.formulas));
    }
    target.getColumn(probability).setNumberFormat("0%");
  } catch (error: unknown) {
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error("PlanTimer kunne ikke fuldføres. Kontrollér de nye rækker før genkørsel. " + detail);
  }
}

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
