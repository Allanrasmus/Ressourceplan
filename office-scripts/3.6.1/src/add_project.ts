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
