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
  }
  interface Table {
    getHeaderRowRange(): Range;
    getRangeBetweenHeaderAndTotal(): Range;
    getRowCount(): number;
    addRows(index?: number, values?: (boolean | string | number)[][]): void;
  }
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
