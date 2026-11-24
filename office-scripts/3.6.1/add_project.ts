// Genereret af office-scripts/3.6.1/build.py. Kopiér hele filen til Excel.
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
  const config = settings(workbook);
  const plan = planSchema(workbook, config);
  const personTable = getTable(workbook, "PersonRegister");
  const projectTable = getTable(workbook, "SagRegister");
  const people = entities(body(personTable));
  const projects = entities(body(projectTable));
  const inputYear = optionalYear(planaar);
  const personInput = medarbejder ?? "";
  const projectInput = projekt ?? "";
  const selection = !personInput.trim() || inputYear === undefined || !projectInput.trim()
    ? selectedCell(workbook) : undefined;
  const { person, year } = planTarget(plan, people, personInput, inputYear, selection);
  let projectKey = projectInput;
  if (!projectKey.trim() && selection) {
    if (selection.sheet === "Projektregister" && (selection.column === 0 || selection.column === 1)) {
      const row = selectedRecord(projectTable, selection);
      projectKey = row ? String(row[0]) : "";
    } else if (selection.sheet === "Min ressourceplan" && selection.column === field(plan, "Projektvalg")) {
      const row = selectedRecord(plan.table, selection);
      projectKey = row ? recordKey(row[field(plan, "Projektvalg")], projects) : "";
    }
    if (!projectKey.trim()) throw new Error("Vælg en projektcelle i tabellen, eller angiv projektet i dialogen.");
  }
  const caseRecord = findEntity(projectKey, projects, "Projekt");
  const newYear = yearSlot(config, year);
  const exists = (): boolean => planInputs(plan).some(row => same(row[field(plan, "Planår")], year) &&
    same(row[field(plan, "Type")], activity) &&
    matchesRecord(row[field(plan, "Medarbejder")], person, people) &&
    matchesRecord(row[field(plan, "Projektvalg")], caseRecord, projects));
  if (exists()) return report("Projektet findes allerede for medarbejderen, året og aktivitetstypen.");
  checkCapacity(plan, 1);
  if (newYear) {
    const fresh = yearSlot(config, year);
    if (fresh) fresh.setValues([[year]]);
    if (exists()) return report("Projektet findes allerede for medarbejderen, året og aktivitetstypen.");
  }
  append(plan, [{ year, employee: caption(person), type: activity, project: caption(caseRecord), probability }]);
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
  table: ExcelScript.Table;
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

function getTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error("Tabellen " + name + " mangler. Åbn Ressourceplan 3.6.");
  return table;
}

function body(table: ExcelScript.Table): Cell[][] {
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

function settings(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet("Planstart");
  const marker = sheet ? sheet.getRange("P6").getValue() : "";
  if (!sheet || (marker !== "PlanTimer-formler 3.6.1" && marker !== "PlanTimer-formler 3.6")) {
    throw new Error("Konfigurationen for Ressourceplan 3.6 eller 3.6.1 mangler.");
  }
  return sheet;
}

function yearSlot(sheet: ExcelScript.Worksheet, year: number): ExcelScript.Range | undefined {
  if (!Number.isInteger(year) || year < 1901 || year > 9998) {
    throw new Error("Planår skal være et helt år mellem 1901 og 9998.");
  }
  const range = sheet.getRange("F7:F26");
  const values = range.getValues();
  if (values.some(row => Number(row[0]) === year)) return undefined;
  const slot = values.findIndex(row => row[0] === "");
  if (slot < 0) throw new Error("Planårslisten har ingen ledige pladser.");
  return range.getCell(slot, 0);
}

function planSchema(workbook: ExcelScript.Workbook, config: ExcelScript.Worksheet): Plan {
  const table = getTable(workbook, "PlanTimer");
  const header = table.getHeaderRowRange();
  const headers = header.getValues()[0].map(value => String(value));
  const fields: [number, string][] = [
    [0, "Planår"], [1, "Medarbejder"], [3, "Type"], [4, "Projektvalg"],
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
  return { table, headers, formulas };
}

function field(plan: Plan, name: string): number {
  return plan.headers.indexOf(name);
}

function planInputs(plan: Plan): Cell[][] {
  const count = plan.table.getRowCount();
  const rows: Cell[][] = [];
  if (count === 0) return rows;
  const range = plan.table.getRangeBetweenHeaderAndTotal();
  const width = field(plan, "Uge 01");
  // Read only the metadata, in blocks below the range request size limit.
  for (let first = 0; first < count; first += 1000) {
    const length = Math.min(1000, count - first);
    const values = range.getCell(first, 0).getResizedRange(length - 1, width - 1).getValues();
    values.forEach(row => { rows.push(row); });
  }
  return rows;
}

function checkCapacity(plan: Plan, count: number): void {
  if (plan.table.getRowCount() + count > planLimit) {
    throw new Error("Handlingen overskrider de 20.000 planrækker med formatering og validering.");
  }
}

function append(plan: Plan, inputs: RecordInput[]): void {
  if (!inputs.length) return;
  checkCapacity(plan, inputs.length);
  const first = plan.table.getRowCount();
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
    plan.table.addRows(-1, payload);
    const range = plan.table.getRangeBetweenHeaderAndTotal();
    const target = range.getCell(first, 0).getResizedRange(inputs.length - 1, plan.headers.length - 1);
    const actual = target.getValues();
    if (plan.table.getRowCount() !== first + inputs.length ||
        inputs.some((input, i) => !same(actual[i][year], input.year) ||
          !same(actual[i][worker], input.employee) || !same(actual[i][type], input.type) ||
          !same(actual[i][project], input.project) || !same(actual[i][probability], input.probability))) {
      throw new Error("Planens rækker ændrede sig under oprettelsen.");
    }
    for (const item of plan.formulas) {
      target.getColumn(item.index).setFormulas(inputs.map(() => [item.formula]));
    }
    target.getColumn(probability).setNumberFormat("0%");
  } catch (error: unknown) {
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error("PlanTimer kunne ikke fuldføres. Kontrollér de nye rækker før genkørsel. " + detail);
  }
}

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
