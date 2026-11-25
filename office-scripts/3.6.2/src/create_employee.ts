/**
 * Opret medarbejder og manglende standardrækker. Brug samme profil ved genkørsel.
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
    register.getRangeBetweenHeaderAndTotal().getCell(slot, 0).getResizedRange(0, 5)
      .setValues([[person.key, literal(name), region, group, literal(department), normtid]]);
  }
  if (year >= 0) ensureYear(workbook, planaar);
  const currentPeople = entities(readRegister(workbook, "PersonRegister"));
  const finalEntries = planInputs(workbook);
  const currentRows = standardRows(workbook, finalEntries, person, currentPeople, planaar);
  append(workbook, plan, currentRows, finalEntries.length);
  return report(caption(person) + " · " + planaar + ": " + currentRows.length + " standardrækker oprettet.");
}

function hasHistory(entries: Cell[][], key: string): boolean {
  return entries.some(row => same(row[2], key) ||
    String(row[1]).toLowerCase().endsWith((separator + key).toLowerCase()));
}
