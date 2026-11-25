/**
 * Opgradér tekstformater og sagsstriber i den eksisterende Ressourceplan.
 * Brug først en acceptkopi. Kørsel ændrer ingen planrækker eller registerværdier.
 */
function main(workbook: ExcelScript.Workbook): string {
  planSchema(workbook);
  const config = workbook.getWorksheet("Planstart");
  const plan = workbook.getWorksheet("Min ressourceplan");
  const overview = workbook.getWorksheet("Ressourceoverblik");
  const people = workbook.getTable("PersonRegister");
  const projects = workbook.getTable("SagRegister");
  if (!config || !plan || !overview || !people || !projects || !workbook.getNamedItem("PlanSidsteRaekke")) {
    throw new Error("Projektmappen mangler et standardark, et register eller PlanSidsteRaekke.");
  }
  const peopleFields = people.getHeaderRowRange().getValues()[0];
  const projectFields = projects.getHeaderRowRange().getValues()[0];
  if (peopleFields.join("|") !== "ID|Medarbejder|Område|Faggruppe|Afdeling|Normtid" ||
      projectFields.join("|") !== "Sagsnr.|Projektnavn" ||
      people.getRowCount() !== 100 || projects.getRowCount() !== 500) {
    throw new Error("Registerstrukturen stemmer ikke med Ressourceplan 3.6.");
  }

  const stripeName = "PlanSagsstribe";
  const reference = "'Planstart'!$R$10:$R$20009";
  const formula = '=IF(ROW()>PlanSidsteRaekke,0,MOD(N(R[-1]C)+IF(INDEX(PlanTimer[Type],ROW()-9)="Sag",SUBTOTAL(103,INDEX(PlanTimer[Type],ROW()-9)),0),2))';
  const formulas = config.getRange("R9:R20009").getFormulasR1C1();
  if (formulas[0][0] !== "" || formulas.slice(1).some(row => row[0] !== "" && !same(row[0], formula))) {
    throw new Error("Planstart!R9:R20009 indeholder allerede andre oplysninger. Ingen indstillinger er ændret.");
  }
  const named = workbook.getNamedItem(stripeName);
  if (named && !same(named.getFormula().replace(/'/g, ""), "=" + reference.replace(/'/g, ""))) {
    throw new Error("Navnet PlanSagsstribe er allerede i brug. Ingen indstillinger er ændret.");
  }

  const area = plan.getRange("A10:BJ20009");
  const rules = area.getConditionalFormats();
  const stripes: { id: string; parity: number }[] = [];
  for (let i = 0; i < rules.length; i++) {
    const custom = rules[i].getCustom();
    if (!custom) continue;
    const saved = custom.getRule().getFormula().toUpperCase().replace(/\s/g, "").replace(/^=/, "");
    const legacy = /^AND\(\$D(\d+)="SAG",MOD\(ROW\(\),2\)=([01])\)$/.exec(saved);
    const migrated = /^AND\(\$D(\d+)="SAG",INDEX\(PLANSAGSSTRIBE,ROW\(\)-9\)=([01])\)$/.exec(saved);
    const match = legacy ?? migrated;
    if (!match) continue;
    const scope = rules[i].getRanges();
    const address = scope.getAddress().split("!").pop() ?? "";
    const fullArea = /^A10:BJ(\d+)$/.exec(address.replace(/\$/g, ""));
    if (match[1] !== "10" || scope.getAreaCount() !== 1 || !fullArea || Number(fullArea[1]) < 20009) {
      throw new Error("En sagsregel er opdelt eller afgrænset til et andet område. Ingen indstillinger er ændret.");
    }
    stripes.push({ id: rules[i].getId(), parity: Number(match[2]) });
  }
  if (stripes.length !== 2 || stripes[0].parity === stripes[1].parity) {
    throw new Error("De to oprindelige sagsregler dækker ikke længere planområdet som forventet. Ingen indstillinger er ændret.");
  }

  if (formulas.slice(1).some(row => row[0] === "")) {
    config.getRange("R10:R20009").setFormulaR1C1(formula);
  }
  if (!named) workbook.addNamedItem(stripeName, "=" + reference);
  for (const stripe of stripes) {
    const rule = area.getConditionalFormat(stripe.id);
    if (!rule) throw new Error("En sagsregel ændrede sig under opgraderingen. Kontrollér på en acceptkopi.");
    const custom = rule.getCustom();
    if (!custom) throw new Error("En sagsregel ændrede type under opgraderingen.");
    custom.getRule().setFormula(stripeRule(stripe.parity));
    custom.getFormat().getFill().setColor(stripe.parity === 1 ? "#EAF2FC" : "#FFFFFF");
  }

  people.getRangeBetweenHeaderAndTotal().getCell(0, 0).getResizedRange(99, 4).setNumberFormat("@");
  projects.getRangeBetweenHeaderAndTotal().setNumberFormat("@");
  overview.getRange("F6").setNumberFormat("@");
  if (plan.getRange("A5").getValue() === "Sag") plan.getRange("A5").getFormat().getFill().setColor("#EAF2FC");
  return report("Tekstformater og sagsstriber er opdateret. Timer, registre, filtre og knapper er bevaret.");
}

function stripeRule(parity: number): string {
  return '=AND($D10="Sag",INDEX(PlanSagsstribe,ROW()-9)=' + parity + ')';
}
