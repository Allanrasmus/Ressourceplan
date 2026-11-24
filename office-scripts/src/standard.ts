function standardRows(plan: Plan, person: Entity, people: Entity[], year: number): RecordInput[] {
  const present: string[] = [];
  const firstWeek = field(plan, "Uge 01");
  const lastWeek = field(plan, "Uge 53");
  const entries = planInputs(plan);
  for (let index = 0; index < entries.length; index++) {
    const row = entries[index];
    if (!same(row[field(plan, "Planår")], year) ||
        !matchesRecord(row[field(plan, "Medarbejder")], person, people)) continue;
    const type = String(row[field(plan, "Type")]).toLowerCase();
    if (type !== "sag") {
      present.push(type);
    } else if (row[field(plan, "Projektvalg")] === "" && present.indexOf(type) < 0) {
      const hours = plan.table.getRangeBetweenHeaderAndTotal().getCell(index, firstWeek)
        .getResizedRange(0, lastWeek - firstWeek).getValues()[0];
      if (hours.every(value => value === "" || value === 0)) present.push(type);
    }
  }
  return ["Sag", "Ferie", "Fravær", "Kursus", "Andet"].filter(type => present.indexOf(type.toLowerCase()) < 0)
    .map(type => ({ year, employee: caption(person), type, project: "", probability: "" }));
}
