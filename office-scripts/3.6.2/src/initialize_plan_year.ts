/**
 * Opret kun manglende standardrækker for én medarbejder og ét planår.
 * @param medarbejder Skriv et entydigt navn, navn · ID eller ID. Tomt henter medarbejderen fra markeret planrække.
 * @param planaar Angiv et nyt år, eller lad feltet stå tomt for året i den markerede planrække.
 */
function main(workbook: ExcelScript.Workbook, medarbejder: string = "", planaar?: string): string {
  const plan = planSchema(workbook);
  const people = entities(readRegister(workbook, "PersonRegister"));
  const employee = medarbejder ?? "";
  const inputYear = optionalYear(planaar);
  const selection = !employee.trim() || inputYear === undefined ? selectedCell(workbook) : undefined;
  const row = selectedPlanRow(workbook, selection);
  const { person, year } = planTarget(people, employee, inputYear, row);
  const newYear = yearSlot(workbook, year);
  let entries = planInputs(workbook);
  let rows = standardRows(workbook, entries, person, people, year);
  checkCapacity(entries.length, rows.length);
  if (newYear >= 0) {
    ensureYear(workbook, year);
    entries = planInputs(workbook);
    rows = standardRows(workbook, entries, person, people, year);
  }
  append(workbook, plan, rows, entries.length);
  return report(caption(person) + " · " + year + ": " + rows.length + " standardrækker oprettet.");
}
