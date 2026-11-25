// Subset of the published ExcelScript contract used by these scripts.
// OfficeDev/office-scripts-docs-reference at 97f8f415842e13e004d7c53c1d24fe45653e50c0.
declare namespace ExcelScript {
  interface Workbook {
    getTable(key: string): Table | undefined;
    getWorksheet(key: string): Worksheet | undefined;
    getSelectedRange(): Range;
    getNamedItem(name: string): NamedItem | undefined;
    addNamedItem(name: string, reference: Range | string, comment?: string): NamedItem;
  }
  interface NamedItem {
    getFormula(): string;
  }
  interface ConditionalFormat {
    getId(): string;
    getCustom(): CustomConditionalFormat | undefined;
    getRanges(): RangeAreas;
  }
  interface RangeAreas {
    getAddress(): string;
    getAreaCount(): number;
  }
  interface CustomConditionalFormat {
    getRule(): ConditionalFormatRule;
    getFormat(): ConditionalRangeFormat;
  }
  interface ConditionalFormatRule {
    getFormula(): string;
    setFormula(formula: string): void;
  }
  interface ConditionalRangeFormat {
    getFill(): ConditionalRangeFill;
  }
  interface ConditionalRangeFill {
    setColor(color: string): void;
  }
  interface RangeFormat {
    getFill(): RangeFill;
  }
  interface RangeFill {
    setColor(color: string): void;
  }
  interface Worksheet {
    getRange(address?: string): Range;
    getName(): string;
    getActiveNamedSheetView(): NamedSheetView;
  }
  interface NamedSheetView {
    getName(): string;
  }
  interface Table {
    getAutoFilter(): AutoFilter;
    getRange(): Range;
    getWorksheet(): Worksheet;
    getHeaderRowRange(): Range;
    getRangeBetweenHeaderAndTotal(): Range;
    getRowCount(): number;
    addRows(index?: number, values?: (boolean | string | number)[][]): void;
  }
  interface AutoFilter {
    apply(range: Range | string, columnIndex?: number, criteria?: FilterCriteria): void;
    clearColumnCriteria(columnIndex: number): void;
    getCriteria(): FilterCriteria[];
    getIsDataFiltered(): boolean;
  }
  enum FilterOn {
    bottomItems = "bottomItems",
    bottomPercent = "bottomPercent",
    cellColor = "cellColor",
    custom = "custom",
    dynamic = "dynamic",
    fontColor = "fontColor",
    icon = "icon",
    topItems = "topItems",
    topPercent = "topPercent",
    values = "values",
  }
  enum FilterOperator {
    and = "and",
    or = "or",
  }
  enum DynamicFilterCriteria {
    aboveAverage = "aboveAverage",
    allDatesInPeriodApril = "allDatesInPeriodApril",
    allDatesInPeriodAugust = "allDatesInPeriodAugust",
    allDatesInPeriodDecember = "allDatesInPeriodDecember",
    allDatesInPeriodFebruary = "allDatesInPeriodFebruary",
    allDatesInPeriodJanuary = "allDatesInPeriodJanuary",
    allDatesInPeriodJuly = "allDatesInPeriodJuly",
    allDatesInPeriodJune = "allDatesInPeriodJune",
    allDatesInPeriodMarch = "allDatesInPeriodMarch",
    allDatesInPeriodMay = "allDatesInPeriodMay",
    allDatesInPeriodNovember = "allDatesInPeriodNovember",
    allDatesInPeriodOctober = "allDatesInPeriodOctober",
    allDatesInPeriodQuarter1 = "allDatesInPeriodQuarter1",
    allDatesInPeriodQuarter2 = "allDatesInPeriodQuarter2",
    allDatesInPeriodQuarter3 = "allDatesInPeriodQuarter3",
    allDatesInPeriodQuarter4 = "allDatesInPeriodQuarter4",
    allDatesInPeriodSeptember = "allDatesInPeriodSeptember",
    belowAverage = "belowAverage",
    lastMonth = "lastMonth",
    lastQuarter = "lastQuarter",
    lastWeek = "lastWeek",
    lastYear = "lastYear",
    nextMonth = "nextMonth",
    nextQuarter = "nextQuarter",
    nextWeek = "nextWeek",
    nextYear = "nextYear",
    thisMonth = "thisMonth",
    thisQuarter = "thisQuarter",
    thisWeek = "thisWeek",
    thisYear = "thisYear",
    today = "today",
    tomorrow = "tomorrow",
    unknown = "unknown",
    yearToDate = "yearToDate",
    yesterday = "yesterday",
  }
  enum FilterDatetimeSpecificity {
    day = "day",
    hour = "hour",
    minute = "minute",
    month = "month",
    second = "second",
    year = "year",
  }
  enum IconSet {
    fiveArrows = "fiveArrows",
    fiveArrowsGray = "fiveArrowsGray",
    fiveBoxes = "fiveBoxes",
    fiveQuarters = "fiveQuarters",
    fiveRating = "fiveRating",
    fourArrows = "fourArrows",
    fourArrowsGray = "fourArrowsGray",
    fourRating = "fourRating",
    fourRedToBlack = "fourRedToBlack",
    fourTrafficLights = "fourTrafficLights",
    invalid = "invalid",
    threeArrows = "threeArrows",
    threeArrowsGray = "threeArrowsGray",
    threeFlags = "threeFlags",
    threeSigns = "threeSigns",
    threeStars = "threeStars",
    threeSymbols = "threeSymbols",
    threeSymbols2 = "threeSymbols2",
    threeTrafficLights1 = "threeTrafficLights1",
    threeTrafficLights2 = "threeTrafficLights2",
    threeTriangles = "threeTriangles",
  }
  type FilterCriteria = {
    color?: string;
    criterion1?: string;
    criterion2?: string;
    dynamicCriteria?: DynamicFilterCriteria;
    filterOn: FilterOn;
    icon?: Icon;
    operator?: FilterOperator;
    subField?: string;
    values?: Array<string | FilterDatetime>;
  };
  type FilterDatetime = {
    date: string;
    specificity: FilterDatetimeSpecificity;
  };
  type Icon = {
    index: number;
    set: IconSet;
  };
  interface Range {
    getValues(): (string | number | boolean)[][];
    getFormulas(): string[][];
    getFormulasR1C1(): string[][];
    getConditionalFormats(): ConditionalFormat[];
    getConditionalFormat(id: string): ConditionalFormat | undefined;
    getFormat(): RangeFormat;
    getValue(): string | number | boolean;
    getRowIndex(): number;
    getColumnIndex(): number;
    getRowCount(): number;
    getColumnCount(): number;
    getRowHidden(): boolean;
    getWorksheet(): Worksheet;
    getRow(row: number): Range;
    getCell(row: number, column: number): Range;
    getColumn(column: number): Range;
    getResizedRange(deltaRows: number, deltaColumns: number): Range;
    setValues(values: (string | number | boolean)[][]): void;
    setFormulas(formulas: string[][]): void;
    setFormulaR1C1(formulaR1C1: string): void;
    setNumberFormat(numberFormat: string): void;
  }
}
