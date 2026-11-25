function standardRows(workbook: ExcelScript.Workbook, entries: Cell[][], person: Entity,
  people: Entity[], year: number): RecordInput[] {
  const present: string[] = [];
  const candidates: number[] = [];
  planGroup(entries, person, people, year).forEach(index => {
    const row = entries[index];
    const type = String(row[3]).toLowerCase();
    if (type !== "sag") {
      present.push(type);
    } else if (row[4] === "") {
      candidates.push(index);
    }
  });
  if (candidates.length && hasFreeCase(workbook, entries, candidates)) present.push("sag");
  return ["Sag", "Ferie", "Fravær", "Kursus", "Andet"].filter(type => present.indexOf(type.toLowerCase()) < 0)
    .map(type => ({ year, employee: caption(person), type, project: "", probability: "" }));
}

function hasFreeCase(workbook: ExcelScript.Workbook, entries: Cell[][], candidates: number[]): boolean {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  if (table.getRowCount() !== entries.length) {
    throw new Error("Planens rækker ændrede sig. Prøv igen.");
  }
  const range = table.getRangeBetweenHeaderAndTotal();
  let start = 0;
  while (start < candidates.length) {
    const first = candidates[start];
    let end = start + 1;
    while (end < candidates.length && candidates[end] - first < 500) end++;
    // Read the candidate window once, including keys and raw hours A:BJ.
    const values = range.getCell(first, 0)
      .getResizedRange(candidates[end - 1] - first, 61).getValues();
    for (let i = start; i < end; i++) {
      const index = candidates[i];
      const row = values[index - first];
      if (entries[index].some((value, column) => value !== row[column])) {
        throw new Error("Planens rækker ændrede sig. Prøv igen.");
      }
      if (row.slice(9, 62).every(value => value === "" || value === 0)) return true;
    }
    start = end;
  }
  return false;
}

function insertStandardRows(workbook: ExcelScript.Workbook, plan: Plan, inputs: RecordInput[],
  entries: Cell[][], person: Entity, people: Entity[], year: number): void {
  const group = planGroup(entries, person, people, year);
  const insertions: { index: number; rows: RecordInput[] }[] = [];
  for (const input of inputs) {
    const index = insertionIndex(entries, group, activityRank(input.type));
    const found = insertions.find(item => item.index === index);
    if (found) found.rows.push(input);
    else insertions.push({ index, rows: [input] });
  }
  insertions.sort((a, b) => b.index - a.index);
  for (const entry of insertions) {
    insertRows(workbook, plan, entry.rows, entries, entry.index);
  }
}
