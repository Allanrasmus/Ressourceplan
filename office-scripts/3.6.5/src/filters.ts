// Kriterier gemmes som lokale værdier. Enum-medlemmer er skalare værdier.
type SavedFilter = {
  column: number;
  filterOn: ExcelScript.FilterOn;
  color?: string;
  criterion1?: string;
  criterion2?: string;
  dynamicCriteria?: ExcelScript.DynamicFilterCriteria;
  icon?: { index: number; set: ExcelScript.IconSet };
  operator?: ExcelScript.FilterOperator;
  subField?: string;
  values?: (string | { date: string; specificity: ExcelScript.FilterDatetimeSpecificity })[];
};

function readPlanFilters(workbook: ExcelScript.Workbook, columnCount: number): SavedFilter[] {
  const table = workbook.getTable("PlanTimer");
  if (!table) throw new Error("Tabellen PlanTimer mangler.");
  const autoFilter = table.getAutoFilter();
  if (!autoFilter.getIsDataFiltered()) return [];

  const criteria = autoFilter.getCriteria();
  const saved: SavedFilter[] = [];
  for (let column = 0; column < criteria.length; column++) {
    if (!criteria[column] || !criteria[column].filterOn) continue;
    const values = criteria[column].values?.map(value => typeof value === "string"
      ? value
      : { date: value.date, specificity: value.specificity });
    const iconIndex = criteria[column].icon?.index;
    const iconSet = criteria[column].icon?.set;
    saved.push({
      column,
      filterOn: criteria[column].filterOn,
      color: criteria[column].color,
      criterion1: criteria[column].criterion1,
      criterion2: criteria[column].criterion2,
      dynamicCriteria: criteria[column].dynamicCriteria,
      operator: criteria[column].operator,
      subField: criteria[column].subField,
      icon: iconIndex === undefined || iconSet === undefined ? undefined : { index: iconIndex, set: iconSet },
      values,
    });
  }
  if (!saved.length || criteria.length > columnCount) {
    throw new Error("PlanTimers aktive filtre kunne ikke aflæses sikkert. Kontrollér filtrene før genkørsel.");
  }
  const sheet = table.getWorksheet();
  try {
    const view = sheet.getActiveNamedSheetView();
    // Et tomt navn er gyldigt for en midlertidig visning.
    if (!view || typeof view.getName() !== "string") {
      throw new Error("Excel returnerede ingen aktiv arkvisning.");
    }
  } catch (error: unknown) {
    throw new Error("Vælg en personlig arkvisning på Min ressourceplan før indsættelse med aktive filtre. " +
      errorDetail(error));
  }
  return saved;
}

function restorePlanFilters(workbook: ExcelScript.Workbook, headers: string[], saved: SavedFilter[],
  attempted: number, failures: string[]): void {
  if (!attempted) return;
  try {
    const table = workbook.getTable("PlanTimer");
    if (!table) throw new Error("Tabellen PlanTimer mangler.");
    const autoFilter = table.getAutoFilter();
    for (let i = 0; i < attempted; i++) {
      const filter = saved[i];
      try {
        const range = table.getRange();
        autoFilter.apply(range, filter.column, {
          filterOn: filter.filterOn,
          color: filter.color,
          criterion1: filter.criterion1,
          criterion2: filter.criterion2,
          dynamicCriteria: filter.dynamicCriteria,
          icon: filter.icon,
          operator: filter.operator,
          subField: filter.subField,
          values: filter.values,
        });
      } catch (error: unknown) {
        failures.push(headers[filter.column] + " (" + errorDetail(error) + ")");
      }
    }
  } catch (error: unknown) {
    failures.push("PlanTimer (" + errorDetail(error) + ")");
  }
}

function errorDetail(error: unknown): string {
  return error instanceof Error && error.message ? error.message : String(error);
}
