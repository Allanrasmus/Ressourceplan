type SavedFilter = { column: string; criterion: string };

function readPlanFilters(workbook: ExcelScript.Workbook, headers: string[], requireConfirmation: boolean, confirmed: boolean = false): SavedFilter[] {
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
        criteria.filterOn !== ExcelScript.FilterOn.custom ||
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

  if (requireConfirmation && !confirmed) {
    throw new Error("Vælg din personlige arkvisning på Min ressourceplan. Kør igen, og bekræft valget i dialogen før filtreret indsættelse.");
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
