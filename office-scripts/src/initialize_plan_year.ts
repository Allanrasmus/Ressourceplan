/**
 * Opret kun manglende standardrækker for én medarbejder og ét planår.
 * @param medarbejder Medarbejder-ID, navn · ID eller et entydigt navn.
 * @param planaar Året som skal initialiseres, også hvis det er nyt i projektmappen.
 */
function main(workbook: ExcelScript.Workbook, medarbejder: string, planaar: number): string {
  const config = settings(workbook);
  const year = yearSlot(config, planaar);
  const plan = planSchema(workbook, config);
  const people = entities(body(getTable(workbook, "PersonRegister")));
  const person = findEntity(medarbejder, people, "Medarbejder");
  let rows = standardRows(plan, person, people, planaar);
  checkCapacity(plan, rows.length);
  if (year) {
    const fresh = yearSlot(config, planaar);
    if (fresh) fresh.setValues([[planaar]]);
    rows = standardRows(plan, person, people, planaar);
  }
  append(plan, rows);
  return report(caption(person) + " · " + planaar + ": " + rows.length + " standardrækker oprettet.");
}
