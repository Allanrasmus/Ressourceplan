// Genereret af office-scripts/3.6.3/build.py. Kopiér hele filen til Excel.
/**
 * Opgradér tekstformater og sagsstriber i den eksisterende Ressourceplan.
 * Brug først en acceptkopi. Kørsel ændrer ingen planrækker eller registerværdier.
 */
function main(workbook: ExcelScript.Workbook): string {
  planSchema(workbook);
  const config = workbook.getWorksheet("Planstart");
  const plan = workbook.getWorksheet("Min ressourceplan");
  const overview = workbook.getWorksheet("Ressourceoverblik");
  const people = workbook.getTable("PersonRegister");
  const projects = workbook.getTable("SagRegister");
  if (!config || !plan || !overview || !people || !projects || !workbook.getNamedItem("PlanSidsteRaekke")) {
    throw new Error("Projektmappen mangler et standardark, et register eller PlanSidsteRaekke.");
  }
  const peopleFields = people.getHeaderRowRange().getValues()[0];
  const projectFields = projects.getHeaderRowRange().getValues()[0];
  if (peopleFields.join("|") !== "ID|Medarbejder|Område|Faggruppe|Afdeling|Normtid" ||
      projectFields.join("|") !== "Sagsnr.|Projektnavn" ||
      people.getRowCount() !== 100 || projects.getRowCount() !== 500) {
    throw new Error("Registerstrukturen stemmer ikke med Ressourceplan 3.6.");
  }

  const stripeName = "PlanSagsstribe";
  const reference = "'Planstart'!$R$10:$R$20009";
  const formula = '=IF(ROW()>PlanSidsteRaekke,0,MOD(N(R[-1]C)+IF(INDEX(PlanTimer[Type],ROW()-9)="Sag",SUBTOTAL(103,INDEX(PlanTimer[Type],ROW()-9)),0),2))';
  const formulas = config.getRange("R9:R20009").getFormulasR1C1();
  if (formulas[0][0] !== "" || formulas.slice(1).some(row => row[0] !== "" && !same(row[0], formula))) {
    throw new Error("Planstart!R9:R20009 indeholder allerede andre oplysninger. Ingen indstillinger er ændret.");
  }
  const named = workbook.getNamedItem(stripeName);
  if (named && !same(named.getFormula().replace(/'/g, ""), "=" + reference.replace(/'/g, ""))) {
    throw new Error("Navnet PlanSagsstribe er allerede i brug. Ingen indstillinger er ændret.");
  }

  const area = plan.getRange("A10:BJ20009");
  const rules = area.getConditionalFormats();
  const stripes: { id: string; parity: number }[] = [];
  for (let i = 0; i < rules.length; i++) {
    const custom = rules[i].getCustom();
    if (!custom) continue;
    const saved = custom.getRule().getFormula().toUpperCase().replace(/\s/g, "").replace(/^=/, "");
    const legacy = /^AND\(\$D(\d+)="SAG",MOD\(ROW\(\),2\)=([01])\)$/.exec(saved);
    const migrated = /^AND\(\$D(\d+)="SAG",INDEX\(PLANSAGSSTRIBE,ROW\(\)-9\)=([01])\)$/.exec(saved);
    const match = legacy ?? migrated;
    if (!match) continue;
    const scope = rules[i].getRanges();
    const address = scope.getAddress().split("!").pop() ?? "";
    const fullArea = /^A10:BJ(\d+)$/.exec(address.replace(/\$/g, ""));
    if (match[1] !== "10" || scope.getAreaCount() !== 1 || !fullArea || Number(fullArea[1]) < 20009) {
      throw new Error("En sagsregel er opdelt eller afgrænset til et andet område. Ingen indstillinger er ændret.");
    }
    stripes.push({ id: rules[i].getId(), parity: Number(match[2]) });
  }
  if (stripes.length !== 2 || stripes[0].parity === stripes[1].parity) {
    throw new Error("De to oprindelige sagsregler dækker ikke længere planområdet som forventet. Ingen indstillinger er ændret.");
  }

  if (formulas.slice(1).some(row => row[0] === "")) {
    config.getRange("R10:R20009").setFormulaR1C1(formula);
  }
  if (!named) workbook.addNamedItem(stripeName, "=" + reference);
  for (const stripe of stripes) {
    const rule = area.getConditionalFormat(stripe.id);
    if (!rule) throw new Error("En sagsregel ændrede sig under opgraderingen. Kontrollér på en acceptkopi.");
    const custom = rule.getCustom();
    if (!custom) throw new Error("En sagsregel ændrede type under opgraderingen.");
    custom.getRule().setFormula(stripeRule(stripe.parity));
    custom.getFormat().getFill().setColor(stripe.parity === 1 ? "#EAF2FC" : "#FFFFFF");
  }

  people.getRangeBetweenHeaderAndTotal().getCell(0, 0).getResizedRange(99, 4).setNumberFormat("@");
  projects.getRangeBetweenHeaderAndTotal().setNumberFormat("@");
  overview.getRange("F6").setNumberFormat("@");
  if (plan.getRange("A5").getValue() === "Sag") plan.getRange("A5").getFormat().getFill().setColor("#EAF2FC");
  return report("Tekstformater og sagsstriber er opdateret. Timer, registre, filtre og knapper er bevaret.");
}

function stripeRule(parity: number): string {
  return '=AND($D10="Sag",INDEX(PlanSagsstribe,ROW()-9)=' + parity + ')';
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
