/**
 * Tilføj Sag eller Tilbud.
 * @param medarbejder Medarbejder-ID, navn · ID eller et entydigt navn.
 * @param projekt Sagsnummer som tekst, projektnavn · sagsnummer eller et entydigt projektnavn.
 * @param tilbudsprocent Angiv 0–100 for Tilbud. Lad feltet stå tomt for standarden på 50 %.
 */
function main(workbook: ExcelScript.Workbook, medarbejder: string, planaar: number,
  projekt: string, aktivitetstype: "Sag" | "Tilbud", tilbudsprocent?: number): string {
  const activity = ["Sag", "Tilbud"].find(value => same(value, aktivitetstype.trim()));
  if (!activity) throw new Error("Aktivitetstype skal være Sag eller Tilbud.");
  if (tilbudsprocent !== undefined &&
      (!Number.isFinite(tilbudsprocent) || tilbudsprocent < 0 || tilbudsprocent > 100 || activity !== "Tilbud")) {
    throw new Error("Tilbudsprocent skal være 0–100 for Tilbud. Lad feltet stå tomt for standardværdien.");
  }
  const config = settings(workbook);
  const plan = planSchema(workbook, config);
  const personTable = getTable(workbook, "PersonRegister");
  const projectTable = getTable(workbook, "SagRegister");
  const people = entities(body(personTable));
  const projects = entities(body(projectTable));
  const person = findEntity(medarbejder, people, "Medarbejder");
  const caseRecord = findEntity(projekt, projects, "Projekt");
  const year = yearSlot(config, planaar);
  const exists = (): boolean => planInputs(plan).some(row => same(row[field(plan, "Planår")], planaar) &&
    same(row[field(plan, "Type")], activity) &&
    matchesRecord(row[field(plan, "Medarbejder")], person, people) &&
    matchesRecord(row[field(plan, "Projektvalg")], caseRecord, projects));
  if (exists()) {
    return report("Projektet findes allerede for medarbejderen, året og aktivitetstypen.");
  }
  checkCapacity(plan, 1);
  if (year) {
    const fresh = yearSlot(config, planaar);
    if (fresh) fresh.setValues([[planaar]]);
    if (exists()) return report("Projektet findes allerede for medarbejderen, året og aktivitetstypen.");
  }
  append(plan, [{ year: planaar, employee: caption(person), type: activity,
    project: caption(caseRecord), probability: tilbudsprocent === undefined ? "" : tilbudsprocent / 100 }]);
  return report(caption(caseRecord) + " tilføjet som " + activity + " for " + caption(person) + ".");
}
