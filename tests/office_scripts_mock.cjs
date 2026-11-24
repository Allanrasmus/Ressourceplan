const assert = require("node:assert/strict");

const textValue = value => typeof value === "string" && value.startsWith("'") ? value.slice(1) : value;
class Range {
  constructor(data, row, column, height, width, base = 0) {
    Object.assign(this, { data, row, column, height, width, base });
  }
  getRowIndex() { return this.base + this.row; }
  getColumnIndex() { return this.column; }
  getCell(y, x) { return new Range(this.data, this.row+y, this.column+x, 1, 1, this.base); }
  getColumn(x) { return new Range(this.data, this.row, this.column+x, this.height, 1, this.base); }
  getResizedRange(y, x) {
    return new Range(this.data, this.row, this.column, this.height+y, this.width+x, this.base);
  }
  getValues() {
    return Array.from({ length: this.height }, (_, y) => Array.from({ length: this.width }, (_, x) =>
      this.data[this.row+y]?.[this.column+x] ?? ""));
  }
  getRow(y) { return new Range(this.data, this.row+y, this.column, 1, this.width, this.base); }
  getValue() { return this.getValues()[0][0]; }
  setValues(values) {
    assert.equal(values.length, this.height);
    values.forEach((row, y) => {
      assert.equal(row.length, this.width);
      this.data[this.row+y] ??= [];
      row.forEach((value, x) => { this.data[this.row+y][this.column+x] = textValue(value); });
    });
  }
  setFormulas(formulas) {
    assert.ok(formulas.every(row => row.every(value => value.startsWith("="))));
    this.setValues(formulas);
  }
  setNumberFormat(format) { assert.equal(format, "0%"); }
}
class Table {
  constructor(data) { Object.assign(this, data); }
  getRowCount() { return this.rows.length; }
  getHeaderRowRange() { return new Range([this.headers], 0, 0, 1, this.headers.length, this.header); }
  getRangeBetweenHeaderAndTotal() {
    return new Range(this.rows, 0, 0, this.rows.length, this.headers.length, this.header+1);
  }
  addRows(index, rows) {
    assert.equal(index, -1);
    if (this.failAdd) throw new Error("Forbigående fejl ved indsættelse");
    rows.forEach(row => {
      assert.equal(row.length, this.headers.length);
      this.rows.push(row.map(textValue));
    });
    if (this.conflict) this.rows.at(-1)[1] = "Interleaved · p02";
  }
}
function testBook(initial) {
  const state = structuredClone(initial);
  const tables = Object.fromEntries(Object.entries(state.tables).map(([k, v]) => [k, new Table(v)]));
  return {
    state, tables, getTable: name => tables[name],
    getSelectedRange() {
      const cell = this.selection;
      if (!cell || cell.multi) throw new Error("Ingen enkelt markering");
      return {
        getRowCount: () => cell.height ?? 1,
        getColumnCount: () => cell.width ?? 1,
        getRowHidden: () => Boolean(cell.hidden),
        getRowIndex: () => cell.row - 1,
        getColumnIndex: () => cell.column - 1,
        getWorksheet: () => ({ getName: () => cell.sheet }),
      };
    },
    getWorksheet: name => {
      assert.equal(name, "Planstart");
      return { getRange(address) {
        const ranges = { P6: [5,15,1,1], "P7:Q13": [6,15,7,2], "F7:F26": [6,5,20,1] };
        assert.ok(address in ranges, address);
        return new Range(state.settings, ...ranges[address]);
      } };
    },
  };
}

module.exports = { testBook };
