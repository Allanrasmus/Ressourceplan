// Genereret af office-scripts/3.6.6/build.py. Kopiér hele filen til Excel.
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
      projectKey = planRow ? recordKey(planRow[4], projects) || String(planRow[4]) : "";
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
  if (newYear >= 0) checkCapacity(entries.length, 1);
  if (newYear >= 0) {
    readPlanFilters(workbook, plan.headers, false);
    ensureYear(workbook, year);
    entries = planInputs(workbook);
    if (entries.some(row => isDuplicate(row))) return report(duplicateMessage);
  }
  confirmRegisterEntry(entities(readRegister(workbook, "PersonRegister")), person);
  confirmRegisterEntry(entities(readRegister(workbook, "SagRegister")), caseRecord);
  const input = { year, employee: caption(person), type: activity, project: caption(caseRecord), probability };
  const group = planGroup(entries, person, people, year);
  if (activity !== "Sag" || !reuseCase(workbook, entries, group, input, plan.headers)) {
    insertRows(workbook, plan, [input], entries, insertionIndex(entries, group, activityRank(activity)));
  }
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
const activityOrder = ["sag", "tilbud", "ferie", "fravær", "kursus", "andet"];

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
  if (!matches.length && label === "Projekt") {
    throw new Error("Projektet findes ikke. Få det først oprettet i Projektregister.");
  }
  if (matches.length !== 1 ||
      records.filter(entity => same(entity.key, matches[0].key)).length !== 1) {
    throw new Error(label + " kunne ikke identificeres entydigt. Brug ID/sagsnummer eller det fulde valg.");
  }
  return matches[0];
}

function confirmRegisterEntry(records: Entity[], selected: Entity): void {
  const matches = records.filter(record => same(record.key, selected.key));
  if (matches.length !== 1 || matches[0].name !== selected.name) {
    throw new Error("En valgt registerpost blev ændret. Kontrollér valget, og prøv igen.");
  }
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

function activityRank(type: Cell): number {
  return activityOrder.indexOf(String(type).toLowerCase());
}

function planGroup(entries: Cell[][], person: Entity, people: Entity[], year: number): number[] {
  const group: number[] = [];
  entries.forEach((row, index) => {
    if (same(row[0], year) && matchesRecord(row[1], person, people)) group.push(index);
  });
  return group;
}

function insertionIndex(entries: Cell[][], group: number[], rank: number): number {
  const firstLater = group.find(index => activityRank(entries[index][3]) > rank);
  // Keep the first higher-priority boundary even in an old, unsorted group.
  // Existing out-of-order rows can be sorted separately in an exclusive session.
  if (firstLater !== undefined) {
    const previous = group.filter(index => index < firstLater).pop();
    return previous === undefined ? firstLater : previous + 1;
  }
  return group.length ? group[group.length - 1] + 1 : entries.length;
}

function reuseCase(workbook: ExcelScript.Workbook, entries: Cell[][], group: number[],
  input: RecordInput, headers: string[]): boolean {
  const boundary = insertionIndex(entries, group, 0);
  const index = group.filter(position => position < boundary && activityRank(entries[position][3]) === 0).pop();
  if (index === undefined || entries[index][4] !== "") return false;
  readPlanFilters(workbook, headers, false);
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  if (table.getRowCount() !== entries.length) {
    throw new Error("Planens rækker ændrede sig. Prøv igen.");
  }
  const row = table.getRangeBetweenHeaderAndTotal().getRow(index);
  const stored = row.getCell(0, 0).getResizedRange(0, 61).getFormulas()[0];
  // The editable cells must be genuinely blank, including formulas returning "" and raw week 53.
  if (stored[4] !== "" || stored[7] !== "" || stored.slice(9, 62).some(value => value !== "")) return false;
  const current = row.getCell(0, 0).getResizedRange(0, 4).getValues()[0];
  if (entries[index].some((value, column) => value !== current[column]) ||
      table.getRowCount() !== entries.length) {
    throw new Error("Planens rækker ændrede sig. Prøv igen.");
  }
  // Change just the project cell, retaining its validation, the formulas and all other row data.
  const project = row.getCell(0, 4);
  project.setValues([[literal(input.project)]]);
  if (project.getValue() !== input.project) {
    throw new Error("Projektvalget kunne ikke bekræftes. Kontrollér rækken før genkørsel.");
  }
  return true;
}

function insertRows(workbook: ExcelScript.Workbook, plan: Plan, inputs: RecordInput[],
  entries: Cell[][], index: number): void {
  if (!inputs.length) return;
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const count = entries.length;
  checkCapacity(count, inputs.length);
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
  if (table.getRowCount() !== count) throw new Error("Planens rækker ændrede sig. Prøv igen.");
  const localInsert = index < count;
  const validatedFilters = readPlanFilters(workbook, plan.headers, localInsert);
  const savedFilters = localInsert ? validatedFilters : [];
  const autoFilter = table.getAutoFilter();

  let attemptedClears = 0;
  let operationError: string | undefined;
  const restoreErrors: string[] = [];
  try {
    // Clearing individual columns preserves the table's sort state.
    for (const filter of savedFilters) {
      attemptedClears++;
      clearPlanFilter(workbook, filter.column);
    }
    if (attemptedClears && autoFilter.getIsDataFiltered()) {
      throw new Error("PlanTimers filtre kunne ikke ryddes før indsættelse.");
    }
    if (count) {
      const start = Math.max(0, index - 1);
      const end = Math.min(index, count - 1);
      const adjacent = table.getRangeBetweenHeaderAndTotal().getCell(start, 0)
        .getResizedRange(end - start, 4).getValues();
      if (adjacent.some((row, i) => row.some((value, column) => value !== entries[start + i][column]))) {
        throw new Error("Rækkerne ved indsættelsesstedet ændrede sig. Prøv igen.");
      }
    }
    if (table.getRowCount() !== count) throw new Error("Planens rækker ændrede sig. Prøv igen.");
    table.addRows(index === count ? -1 : index, payload);
    if (table.getRowCount() !== count + inputs.length) {
      throw new Error("Planens rækker ændrede sig under oprettelsen.");
    }
    const range = table.getRangeBetweenHeaderAndTotal();
    const target = range.getCell(index, 0).getResizedRange(inputs.length - 1, plan.headers.length - 1);
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
    // Keep the in-memory snapshot aligned for any other inserts in the same operation.
    entries.splice(index, 0, ...actual.map(row => row.slice(0, 5)));
  } catch (error: unknown) {
    operationError = errorDetail(error);
  } finally {
    // Include the last attempted clear in case Excel reported an error after changing it.
    for (let i = 0; i < attemptedClears; i++) {
      const saved = savedFilters[i];
      try {
        restorePlanFilter(workbook, saved.column, saved.criterion);
      } catch (error: unknown) {
        restoreErrors.push(saved.column + " (" + errorDetail(error) + ")");
      }
    }
  }
  if (operationError !== undefined || restoreErrors.length) {
    const restoreDetail = restoreErrors.length
      ? " Genskabelse af PlanTimer-filtre mislykkedes for: " + restoreErrors.join(", ") + ". Genskab disse filtre manuelt."
      : "";
    throw new Error("PlanTimer kunne ikke fuldføres. Ændringer kan være delvist gennemført. Kontrollér de nye rækker før genkørsel. " +
      (operationError ?? "") + restoreDetail);
  }
}

type SavedFilter = { column: string; criterion: string };

function readPlanFilters(workbook: ExcelScript.Workbook, headers: string[], requireView: boolean): SavedFilter[] {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const autoFilter = table.getAutoFilter();
  if (!autoFilter.getIsDataFiltered()) return [];

  const saved: SavedFilter[] = [];
  for (const name of headers) {
    const column = table.getColumnByName(name);
    if (!column) throw new Error("PlanTimer-kolonnen " + name + " mangler.");
    const filter = column.getFilter();
    const criteria = filter.getCriteria();
    if (!criteria || !criteria.filterOn) continue;
    const criterion = criteria.criterion1;
    if ((name !== "Medarbejder" && name !== "Planår") ||
        String(criteria.filterOn).toLowerCase() !== "custom" ||
        !criterion || !/^=[^=<>*?~][^*?~]*$/.test(criterion) ||
        !criterion.slice(1).trim() || criteria.criterion2 || criteria.subField) {
      throw new Error("PlanTimer-filteret på " + name + " kan ikke genskabes sikkert. " +
        "Brug et enkelt lighedsfilter på Planår eller Medarbejder i din personlige arkvisning.");
    }
    saved.push({ column: name, criterion });
  }
  if (!saved.length) {
    throw new Error("PlanTimers aktive filtre kunne ikke aflæses sikkert. Kontrollér filtrene før genkørsel.");
  }

  if (!requireView) return saved;

  const sheet = table.getWorksheet();
  try {
    const view = sheet.getActiveNamedSheetView();
    // En midlertidig arkvisning kan have tomt navn.
    if (!view || typeof view.getName() !== "string") {
      throw new Error("Excel returnerede ingen aktiv arkvisning.");
    }
  } catch (error: unknown) {
    throw new Error("Vælg en personlig arkvisning på Min ressourceplan før indsættelse med aktive filtre. " +
      errorDetail(error));
  }
  return saved;
}

function clearPlanFilter(workbook: ExcelScript.Workbook, name: string): void {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const column = table.getColumnByName(name);
  if (!column) throw new Error("PlanTimer-kolonnen " + name + " mangler.");
  const filter = column.getFilter();
  filter.clear();
}

function restorePlanFilter(workbook: ExcelScript.Workbook, name: string, criterion: string): void {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const column = table.getColumnByName(name);
  if (!column) throw new Error("PlanTimer-kolonnen " + name + " mangler.");
  const filter = column.getFilter();
  filter.applyCustomFilter(criterion);
}

function errorDetail(error: unknown): string {
  return error instanceof Error && error.message ? error.message : String(error);
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
