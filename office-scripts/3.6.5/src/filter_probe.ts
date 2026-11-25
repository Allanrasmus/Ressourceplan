/** Afprøv filterkriterier i PlanTimer på en særskilt Excel-testfil. */
function main(workbook: ExcelScript.Workbook): string {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const headers = table.getHeaderRowRange().getValues()[0].map(value => String(value));
  const saved = readPlanFilters(workbook, headers.length);
  if (!saved.length) throw new Error("Vælg en personlig arkvisning, og sæt mindst ét filter i PlanTimer.");

  const autoFilter = table.getAutoFilter();
  let attempted = 0;
  let failure: string | undefined;
  const restoreErrors: string[] = [];
  try {
    for (const filter of saved) {
      attempted++;
      autoFilter.clearColumnCriteria(filter.column);
    }
    if (autoFilter.getIsDataFiltered()) throw new Error("PlanTimers filtre kunne ikke ryddes.");
  } catch (error: unknown) {
    failure = errorDetail(error);
  } finally {
    restorePlanFilters(workbook, headers, saved, attempted, restoreErrors);
  }
  if (failure !== undefined || restoreErrors.length) {
    throw new Error("Filterprøven fejlede. Kontrollér filtrene manuelt. " +
      (failure ?? "") + " " + restoreErrors.join(", "));
  }

  const restored = readPlanFilters(workbook, headers.length);
  if (JSON.stringify(restored) !== JSON.stringify(saved)) {
    throw new Error("De aflæste kriterier efter gendannelse afviger. Kontrollér filtrene manuelt.");
  }
  return "Ryddede og genskabte " + saved.length + " kolonnefiltre i PlanTimer.";
}
