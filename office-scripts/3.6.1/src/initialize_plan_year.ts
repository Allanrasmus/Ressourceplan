/**
 * Opret kun manglende standardrækker for én medarbejder og ét planår.
 * @param medarbejder Skriv et entydigt navn, navn · ID eller ID. Tomt henter medarbejderen fra markeret planrække.
 * @param planaar Angiv et nyt år, eller lad feltet stå tomt for året i den markerede planrække.
 */
function main(workbook: ExcelScript.Workbook, medarbejder: string = "", planaar?: string): string {
  const config = settings(workbook);
  const plan = planSchema(workbook, config);
  const people = entities(body(getTable(workbook, "PersonRegister")));
  const employee = medarbejder ?? "";
  const inputYear = optionalYear(planaar);
  const selection = !employee.trim() || inputYear === undefined ? selectedCell(workbook) : undefined;
  const { person, year } = planTarget(plan, people, employee, inputYear, selection);
  const newYear = yearSlot(config, year);
  let rows = standardRows(plan, person, people, year);
  checkCapacity(plan, rows.length);
  if (newYear) {
    const fresh = yearSlot(config, year);
    if (fresh) fresh.setValues([[year]]);
    rows = standardRows(plan, person, people, year);
  }
  append(plan, rows);
  return report(caption(person) + " · " + year + ": " + rows.length + " standardrækker oprettet.");
}
