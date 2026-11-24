// Subset of the published ExcelScript contract used by these scripts.
// OfficeDev/office-scripts-docs-reference at 97f8f415842e13e004d7c53c1d24fe45653e50c0.
declare namespace ExcelScript {
  interface Workbook {
    getTable(key: string): Table | undefined;
    getWorksheet(key: string): Worksheet | undefined;
  }
  interface Worksheet {
    getRange(address?: string): Range;
  }
  interface Table {
    getHeaderRowRange(): Range;
    getRangeBetweenHeaderAndTotal(): Range;
    getRowCount(): number;
    addRows(index?: number, values?: (boolean | string | number)[][]): void;
  }
  interface Range {
    getValues(): (string | number | boolean)[][];
    getValue(): string | number | boolean;
    getRowIndex(): number;
    getColumnIndex(): number;
    getCell(row: number, column: number): Range;
    getColumn(column: number): Range;
    getResizedRange(deltaRows: number, deltaColumns: number): Range;
    setValues(values: (string | number | boolean)[][]): void;
    setFormulas(formulas: string[][]): void;
    setNumberFormat(numberFormat: string): void;
  }
}
