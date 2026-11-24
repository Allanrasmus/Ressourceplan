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
  const config = settings(workbook);
  const year = yearSlot(config, planaar);
  const plan = planSchema(workbook, config);
  const register = getTable(workbook, "PersonRegister");
  const values = body(register);
  if (register.getHeaderRowRange().getRowIndex() !== 5 || register.getHeaderRowRange().getColumnIndex() !== 0 ||
      values.length > 100 || values.some(row => row.length !== 6)) {
    throw new Error("Medarbejderregisteret stemmer ikke med version 3.6.");
  }
  const people = entities(values);
  const profile: Cell[] = [name, region, group, department, normtid];
  // Compare complete profiles so two distinct employees may share a name.
  const existing = values.filter(row => row[0] !== "" &&
    profile.every((value, index) => same(row[index + 1], value)));
  if (existing.length > 1) throw new Error("Flere medarbejdere har samme profil. Brug medarbejder-ID i Opret nyt planår.");
  let person: Entity;
  let slot = -1;
  if (existing.length === 1) {
    person = findEntity(String(existing[0][0]), people, "Medarbejder");
  } else {
    const used = values.map(row => String(row[0]));
    const entries = planInputs(plan);
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
  const rows = standardRows(plan, person, slot < 0 ? people : [...people, person], planaar);
  checkCapacity(plan, rows.length);
  if (slot >= 0) {
    const currentEntries = planInputs(plan);
    const fresh = body(register);
    if (hasHistory(currentEntries, person.key) || fresh.length !== values.length ||
        fresh[slot].some((value, index) => value !== values[slot][index]) ||
        fresh.some((row, index) => (index !== slot && same(row[0], person.key)) ||
          (row[0] !== "" && profile.every((value, i) => same(row[i+1], value))))) {
      throw new Error("Medarbejderregisteret blev ændret. Prøv igen.");
    }
    register.getRangeBetweenHeaderAndTotal().getCell(slot, 0).getResizedRange(0, 5)
      .setValues([[person.key, literal(name), region, group, literal(department), normtid]]);
  }
  if (year) {
    // Confirm the slot again after any register write.
    const freshYear = yearSlot(config, planaar);
    if (freshYear) freshYear.setValues([[planaar]]);
  }
  const currentRows = standardRows(plan, person, entities(body(register)), planaar);
  append(plan, currentRows);
  return report(caption(person) + " · " + planaar + ": " + currentRows.length + " standardrækker oprettet.");
}

function hasHistory(entries: Cell[][], key: string): boolean {
  return entries.some(row => same(row[2], key) ||
    String(row[1]).toLowerCase().endsWith((separator + key).toLowerCase()));
}
