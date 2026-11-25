// Genereret af office-scripts/3.6.3/build.py. Kopiér hele filen til Excel.
/**
 * Opret medarbejder og manglende standardrækker. Brug samme profil ved genkørsel.
 * @param afdeling Afdelingsnummer som tekst. Behold punktum og alle nuller.
 * @param normtid Ugentlig normtid i timer, mellem 0 og 80.
 * @param planaar Året for medarbejderens første plan.
 */
function main(workbook: ExcelScript.Workbook, medarbejdernavn: string,
  omraade: "Øst" | "Vest", faggruppe: "EL" | "Belysning" | "BMS/CTS" | "Sikring",
  afdeling: string, normtid: number, planaar: number): string {
  const name = required(medarbejdernavn, "Medarbejdernavn");
  const department = required(afdeling, "Afdeling");
  const region = ["Øst", "Vest"].find(value => same(value, omraade.trim()));
  const group = ["EL", "Belysning", "BMS/CTS", "Sikring"].find(value => same(value, faggruppe.trim()));
  if (!region || !group) throw new Error("Vælg et gyldigt område og en gyldig faggruppe.");
  if (!Number.isFinite(normtid) || normtid < 0 || normtid > 80) {
    throw new Error("Normtid skal være mellem 0 og 80 timer.");
  }
  const plan = planSchema(workbook);
  const year = yearSlot(workbook, planaar);
  const register = workbook.getTable("PersonRegister");
  if (!register) throw new Error("Medarbejderregisteret mangler.");
  const header = register.getHeaderRowRange();
  const values = readRegister(workbook, "PersonRegister");
  if (header.getRowIndex() !== 5 || header.getColumnIndex() !== 0 ||
      values.length > 100 || values.some(row => row.length !== 6)) {
    throw new Error("Medarbejderregisteret stemmer ikke med version 3.6.");
  }
  const people = entities(values);
  const profile: Cell[] = [name, region, group, department, normtid];
  // Compare complete profiles so two distinct employees may share a name.
  const existing = values.filter(row => row[0] !== "" &&
    profile.every((value, index) => same(row[index + 1], value)));
  if (existing.length > 1) throw new Error("Flere medarbejdere har samme profil. Brug medarbejder-ID i Opret nyt planår.");
  const entries = planInputs(workbook);
  let person: Entity;
  let slot = -1;
  if (existing.length === 1) {
    person = findEntity(String(existing[0][0]), people, "Medarbejder");
  } else {
    const used = values.map(row => String(row[0]));
    const assigned = (key: string): boolean => hasHistory(entries, key);
    const reserved = Array.from({ length: 25 }, (_, i) => "p" + String(i + 6).padStart(2, "0"));
    slot = values.findIndex(row => row.every(value => value === "") ||
      (reserved.some(key => same(key, row[0])) && row.slice(1).every(value => value === "") &&
       !assigned(String(row[0])) && used.filter(value => same(value, row[0])).length === 1));
    if (slot < 0) throw new Error("Medarbejderregisteret har ingen sikkert ledige pladser.");
    let key = String(values[slot][0]);
    if (key === "") {
      let number = 1;
      do {
        key = "p" + String(number++).padStart(2, "0");
      } while (used.some(value => same(value, key)) || assigned(key));
    }
    person = { key, name };
  }
  const rows = standardRows(workbook, entries, person, slot < 0 ? people : [...people, person], planaar);
  checkCapacity(entries.length, rows.length);
  if (slot >= 0) {
    const currentEntries = planInputs(workbook);
    const fresh = readRegister(workbook, "PersonRegister");
    if (hasHistory(currentEntries, person.key) || fresh.length !== values.length ||
        fresh[slot].some((value, index) => value !== values[slot][index]) ||
        fresh.some((row, index) => (index !== slot && same(row[0], person.key)) ||
          (row[0] !== "" && profile.every((value, i) => same(row[i+1], value))))) {
      throw new Error("Medarbejderregisteret blev ændret. Prøv igen.");
    }
    const target = register.getRangeBetweenHeaderAndTotal().getCell(slot, 0).getResizedRange(0, 5);
    target.getCell(0, 0).getResizedRange(0, 4).setNumberFormat("@");
    target.setValues([[person.key, literal(name), region, group, literal(department), normtid]]);
    const stored = target.getValues()[0];
    if (stored[0] !== person.key || profile.some((value, index) => stored[index + 1] !== value)) {
      throw new Error("Stamdata kunne ikke bekræftes. Kontrollér posten i Medarbejdere før genkørsel.");
    }
  }
  if (year >= 0) ensureYear(workbook, planaar);
  const currentPeople = entities(readRegister(workbook, "PersonRegister"));
  confirmRegisterEntry(currentPeople, person);
  const finalEntries = planInputs(workbook);
  const currentRows = standardRows(workbook, finalEntries, person, currentPeople, planaar);
  insertStandardRows(workbook, plan, currentRows, finalEntries, person, currentPeople, planaar);
  return report(caption(person) + " · " + planaar + ": " + currentRows.length + " standardrækker oprettet.");
}

function hasHistory(entries: Cell[][], key: string): boolean {
  return entries.some(row => same(row[2], key) ||
    String(row[1]).toLowerCase().endsWith((separator + key).toLowerCase()));
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
  input: RecordInput): boolean {
  const boundary = insertionIndex(entries, group, 0);
  const index = group.filter(position => position < boundary && activityRank(entries[position][3]) === 0).pop();
  if (index === undefined || entries[index][4] !== "") return false;
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
  try {
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
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error("PlanTimer kunne ikke fuldføres. Kontrollér de nye rækker før genkørsel. " + detail);
  }
}

function standardRows(workbook: ExcelScript.Workbook, entries: Cell[][], person: Entity,
  people: Entity[], year: number): RecordInput[] {
  const present: string[] = [];
  const candidates: number[] = [];
  planGroup(entries, person, people, year).forEach(index => {
    const row = entries[index];
    const type = String(row[3]).toLowerCase();
    if (type !== "sag") {
      present.push(type);
    } else if (row[4] === "") {
      candidates.push(index);
    }
  });
  if (candidates.length && hasFreeCase(workbook, entries, candidates)) present.push("sag");
  return ["Sag", "Ferie", "Fravær", "Kursus", "Andet"].filter(type => present.indexOf(type.toLowerCase()) < 0)
    .map(type => ({ year, employee: caption(person), type, project: "", probability: "" }));
}

function hasFreeCase(workbook: ExcelScript.Workbook, entries: Cell[][], candidates: number[]): boolean {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  if (table.getRowCount() !== entries.length) {
    throw new Error("Planens rækker ændrede sig. Prøv igen.");
  }
  const range = table.getRangeBetweenHeaderAndTotal();
  let start = 0;
  while (start < candidates.length) {
    const first = candidates[start];
    let end = start + 1;
    while (end < candidates.length && candidates[end] - first < 500) end++;
    // Read the candidate window once, including keys and raw hours A:BJ.
    const values = range.getCell(first, 0)
      .getResizedRange(candidates[end - 1] - first, 61).getValues();
    for (let i = start; i < end; i++) {
      const index = candidates[i];
      const row = values[index - first];
      if (entries[index].some((value, column) => value !== row[column])) {
        throw new Error("Planens rækker ændrede sig. Prøv igen.");
      }
      if (row.slice(9, 62).every(value => value === "" || value === 0)) return true;
    }
    start = end;
  }
  return false;
}

function insertStandardRows(workbook: ExcelScript.Workbook, plan: Plan, inputs: RecordInput[],
  entries: Cell[][], person: Entity, people: Entity[], year: number): void {
  const group = planGroup(entries, person, people, year);
  const insertions: { index: number; rows: RecordInput[] }[] = [];
  for (const input of inputs) {
    const index = insertionIndex(entries, group, activityRank(input.type));
    const found = insertions.find(item => item.index === index);
    if (found) found.rows.push(input);
    else insertions.push({ index, rows: [input] });
  }
  insertions.sort((a, b) => b.index - a.index);
  for (const entry of insertions) {
    insertRows(workbook, plan, entry.rows, entries, entry.index);
  }
}
