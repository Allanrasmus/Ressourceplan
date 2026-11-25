type SavedFilter = { column: string; criterion: string };

function readPlanFilters(workbook: ExcelScript.Workbook, headers: string[], requireView: boolean): SavedFilter[] {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const autoFilter = table.getAutoFilter();
  if (!autoFilter.getIsDataFiltered()) return [];

  const saved: SavedFilter[] = [];
  for (const name of headers) {
    const column = table.getColumnByName(name);
    if (!column) throw new Error("PlanTimer-kolonnen " + name + " mangler.");
    const filter = column.getFilter();
    const criteria = filter.getCriteria();
    if (!criteria || !criteria.filterOn) continue;
    const criterion = criteria.criterion1;
    if ((name !== "Medarbejder" && name !== "Planår") ||
        String(criteria.filterOn).toLowerCase() !== "custom" ||
        !criterion || !/^=[^=<>*?~][^*?~]*$/.test(criterion) ||
        !criterion.slice(1).trim() || criteria.criterion2 || criteria.subField) {
      throw new Error("PlanTimer-filteret på " + name + " kan ikke genskabes sikkert. " +
        "Brug et enkelt lighedsfilter på Planår eller Medarbejder i din personlige arkvisning.");
    }
    saved.push({ column: name, criterion });
  }
  if (!saved.length) {
    throw new Error("PlanTimers aktive filtre kunne ikke aflæses sikkert. Kontrollér filtrene før genkørsel.");
  }

  if (!requireView) return saved;

  const sheet = table.getWorksheet();
  try {
    const view = sheet.getActiveNamedSheetView();
    // En midlertidig arkvisning kan have tomt navn.
    if (!view || typeof view.getName() !== "string") {
      throw new Error("Excel returnerede ingen aktiv arkvisning.");
    }
  } catch (error: unknown) {
    throw new Error("Vælg en personlig arkvisning på Min ressourceplan før indsættelse med aktive filtre. " +
      errorDetail(error));
  }
  return saved;
}

function clearPlanFilter(workbook: ExcelScript.Workbook, name: string): void {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const column = table.getColumnByName(name);
  if (!column) throw new Error("PlanTimer-kolonnen " + name + " mangler.");
  const filter = column.getFilter();
  filter.clear();
}

function restorePlanFilter(workbook: ExcelScript.Workbook, name: string, criterion: string): void {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const column = table.getColumnByName(name);
  if (!column) throw new Error("PlanTimer-kolonnen " + name + " mangler.");
  const filter = column.getFilter();
  filter.applyCustomFilter(criterion);
}

function errorDetail(error: unknown): string {
  return error instanceof Error && error.message ? error.message : String(error);
}
