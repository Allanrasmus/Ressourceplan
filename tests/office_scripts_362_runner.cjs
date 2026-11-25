const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { testBook } = require("./office_scripts_mock.cjs");

const [compiled, fixture, metricsFile] = process.argv.slice(2);
const initial = JSON.parse(fs.readFileSync(fixture, "utf8"));
const load = name => vm.runInNewContext(
  fs.readFileSync(path.join(compiled, name + ".js"), "utf8") + "\nmain;",
  { console: { log() {} } },
);
const add = load("add_project");
const init = load("initialize_plan_year");

function tracePlan(book) {
  const plan = book.tables.PlanTimer;
  const range = plan.getRangeBetweenHeaderAndTotal.bind(plan);
  const reads = [];
  const writes = [];
  function trace(target) {
    return new Proxy(target, {
      get(object, property) {
        const method = Reflect.get(object, property);
        if (property === "getCell" || property === "getColumn" || property === "getResizedRange") {
          return (...args) => trace(method.apply(object, args));
        }
        if (property === "getValues" || property === "setFormulas") {
          return (...args) => {
            (property === "getValues" ? reads : writes).push({
              row: object.row, column: object.column, height: object.height, width: object.width,
            });
            return method.apply(object, args);
          };
        }
        return typeof method === "function" ? method.bind(object) : method;
      },
    });
  }
  plan.getRangeBetweenHeaderAndTotal = () => trace(range());
  return { plan, reads, writes };
}

const small = testBook(initial);
const traced = tracePlan(small);
assert.match(add(small, "Medarbejder 2", "2026", "Hospital Nord", "Sag", "25"), /tilføjet/);
assert.equal(traced.plan.rows.at(-1)[7], "");
assert.equal(traced.reads.filter(read => read.width === 5).length, 1);
assert.deepEqual(traced.writes.map(write => write.width), [1, 2, 1, 3]);
assert.deepEqual(traced.writes.map(write => write.column), [2, 5, 8, 62]);

const hidden = testBook(initial);
const target = tracePlan(hidden);
hidden.selection = { sheet: "Min ressourceplan", row: 10, column: 5 };
assert.match(add(hidden, "", "", "", "Tilbud", "0"), /p01.*2026/);
assert.equal(target.reads.filter(read => read.height === 1 && read.width === 5).length, 1,
  "Medarbejder, år og projekt skal genbruges fra samme læste række.");
assert.equal(target.plan.rows.at(-1)[7], 0);

const large = testBook(initial);
const expanded = tracePlan(large);
const sample = expanded.plan.rows[0];
while (expanded.plan.rows.length < 20000) {
  const row = Array.from({ length: sample.length }, () => "");
  row[0] = 2027;
  row[1] = "Medarbejder 2 · p02";
  row[2] = "p02";
  row[3] = "Sag";
  // A pasted week 53 value is raw input even though it is excluded from the row total.
  row[61] = 1;
  expanded.plan.rows.push(row);
}
for (const row of expanded.plan.rows) {
  if (row[0] === 2027 && row[1] === "Medarbejder 2 · p02" && row[3] === "Sag" && row[4] === "") {
    row[61] = 1;
    row[9] = -1;
  }
}
expanded.plan.rows[19998][4] = "Hospital Nord · 24015";
expanded.plan.rows[19999][61] = 0;
expanded.plan.rows[19999][9] = 0;
assert.match(add(large, "Medarbejder 2", "2027", "24015", "Sag"), /findes allerede/);
assert.equal(expanded.reads.length, 5);
assert.deepEqual(expanded.reads.map(read => read.height), [4000, 4000, 4000, 4000, 4000]);
assert.ok(expanded.reads.every(read => read.width === 5));
expanded.reads.length = 0;
assert.match(init(large, "Medarbejder 2", "2027"), /0 standardrækker/);
const budgets = {
  smallFormulaWrites: traced.writes.length,
  selectedRowReads: target.reads.filter(read => read.height === 1 && read.width === 5).length,
  largeMetadataReads: expanded.reads.filter(read => read.width === 5).length,
  largeHourReads: expanded.reads.filter(read => read.width === 62).length,
};
assert.equal(budgets.largeMetadataReads, 5);
assert.ok(budgets.largeHourReads <= 41);
assert.ok(expanded.reads.filter(read => read.width === 62).every(read => read.height <= 500));
assert.equal(expanded.plan.rows.length, 20000);
assert.equal(expanded.writes.length, 0);

const changed = testBook(initial);
const conflict = changed.tables.PlanTimer;
const originalCount = conflict.rows.length;
const rowCount = conflict.getRowCount.bind(conflict);
let reads = 0;
conflict.getRowCount = () => {
  if (++reads === 2) conflict.rows.push(conflict.rows[0].slice());
  return rowCount();
};
assert.throws(() => add(changed, "Medarbejder 2", "2026", "24015", "Sag"), /ændrede sig/);
assert.equal(conflict.rows.length, originalCount + 1);
assert.deepEqual(conflict.rows.slice(0, originalCount), initial.tables.PlanTimer.rows);

fs.writeFileSync(metricsFile, JSON.stringify(budgets));
