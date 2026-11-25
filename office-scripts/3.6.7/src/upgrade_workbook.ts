/** Opgradér Bruttotilbud og versionsoverskrifter på en kopi af en 3.6.3-arbejdsfil. */
function main(workbook: ExcelScript.Workbook): string {
  const sheets = ["Min ressourceplan", "Ressourceoverblik", "Projektregister",
    "Projektoversigt", "Medarbejdere", "Planstart"];
  const oldVersion = "COWI  |  Ressourceplan 3.6.3";
  const newVersion = "COWI  |  Ressourceplan 3.6.7";
  const overview = workbook.getWorksheet("Ressourceoverblik");
  if (!overview) throw new Error("Ressourceoverblik mangler.");
  for (const name of sheets) {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) throw new Error("Arket " + name + " mangler.");
    const version = sheet.getRange("A1").getValue();
    if (version !== oldVersion && version !== newVersion) {
      throw new Error("Versionsoverskriften i " + name + " afviger fra 3.6.3 og 3.6.7.");
    }
  }
  if (overview.getRange("F8").getFormulas()[0][0] !== "=SUM(F590:F689)") {
    throw new Error("Formlen i Ressourceoverblik!F8 afviger. Kontrollér arbejdsfilen før opgradering.");
  }
  const block = overview.getRange("H590:BH689");
  const formulas = block.getFormulas();
  const condition = '(PlanTimer[Type]="Tilbud")';
  let repair = false;
  for (let slot = 0; slot < 100; slot++) {
    for (let week = 0; week < 53; week++) {
      const column = weekColumn(week + 8);
      const field = "Uge " + String(week + 1).padStart(2, "0");
      const worker = 250 + slot;
      const prefix = '=IF(OR($BI' + worker + '="",'+ column + '$189=""),0,SUMPRODUCT(PlanTimer[' +
        field + '],--(PlanTimer[ID]&""=$BI' + worker + '),--(PlanTimer[Planår]&""=$B$185&""),';
      const previous = prefix + "(" + condition + ")))";
      const updated = prefix + "--" + condition + "))";
      const actual = formulas[slot][week];
      if (actual !== previous && actual !== updated) {
        throw new Error("Bruttotilbudsformlen i Ressourceoverblik!" + column + (590 + slot) +
          " afviger. Ingen ændringer er foretaget.");
      }
      if (actual === previous) {
        formulas[slot][week] = updated;
        repair = true;
      }
    }
  }
  if (repair) block.setFormulas(formulas);
  for (const name of sheets) {
    const sheet = workbook.getWorksheet(name);
    if (!sheet) throw new Error("Arket " + name + " mangler.");
    const banner = sheet.getRange("A1");
    if (banner.getValue() === oldVersion) banner.setValues([[newVersion]]);
  }
  const message = "Bruttotilbud og versionsoverskrifter er kontrolleret for Ressourceplan 3.6.7.";
  console.log(message);
  return message;
}

function weekColumn(index: number): string {
  const first = Math.floor((index - 1) / 26);
  const last = String.fromCharCode(65 + (index - 1) % 26);
  return first ? String.fromCharCode(64 + first) + last : last;
}
