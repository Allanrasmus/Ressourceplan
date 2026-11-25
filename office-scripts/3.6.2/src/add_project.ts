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
