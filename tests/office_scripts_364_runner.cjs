const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { Range, testBook } = require("./office_scripts_mock.cjs");

const [compiled, fixture, scenario, version = "3.6.4"] = process.argv.slice(2);
const initial = JSON.parse(fs.readFileSync(fixture, "utf8"));
const load = (name, scriptVersion = version) => vm.runInNewContext(
  fs.readFileSync(path.join(compiled, scriptVersion, name + ".js"), "utf8") + "\nmain;",
  { console: { log() {} } },
);

const create = load("create_employee");
const add = load("add_project");
const initialize = load("initialize_plan_year");
const criteria = [
  { filterOn: "values", values: ["2026"] },
  { filterOn: "values", values: ["Medarbejder 1 · p01"] },
  {},
  { filterOn: "custom", criterion1: "=Sag", criterion2: "=Tilbud", operator: "or" },
];
const restored = [
  "clear:0", "clear:1", "clear:3", "insert", "apply:0", "apply:1", "apply:3",
];

function setup(options = {}) {
  const book = testBook(initial);
  const table = book.tables.PlanTimer;
  const events = [];
  const original = options.criteria ?? criteria;
  const active = structuredClone(options.unfiltered ? [] : original);
  let exposed;
  const isFiltered = () => active.some(item => item.filterOn);
  const worksheet = {
    getActiveNamedSheetView() {
      events.push("view");
      return options.defaultView ? undefined : { getName: () => options.viewName ?? "Planlægning" };
    },
  };
  table.getWorksheet = () => worksheet;
  table.getRange = () => new Range(table.rows, -1, 0, table.getRowCount() + 1, table.headers.length, 10);
  table.getAutoFilter = () => ({
    getIsDataFiltered: isFiltered,
    getCriteria() {
      events.push("criteria");
      exposed = structuredClone(options.incomplete ? [{}, {}] : active);
      return exposed;
    },
    clearColumnCriteria(column) {
      events.push("clear:" + column);
      active[column] = {};
      // Model native data that becomes invalid while Excel clears the criteria.
      if (options.invalidateRead) {
        const previous = exposed[column];
        if (previous.values) previous.values.length = 0;
        if (previous.icon) previous.icon.index = 99;
        previous.filterOn = "invalidated";
      }
      if (options.failClear === column) throw new Error("Rydning fejlede efter ændring");
    },
    apply(range, column, filter) {
      events.push("apply:" + column);
      assert.equal(range.getRowIndex(), 9);
      assert.equal(range.height, table.rows.length + 1);
      assert.equal(range.width, table.headers.length);
      if (options.failRestore === column) throw new Error("Excel kunne ikke anvende kriteriet");
      active[column] = JSON.parse(JSON.stringify(filter));
    },
  });
  const originalAdd = table.addRows.bind(table);
  table.addRows = (index, rows) => {
    events.push(index < 0 ? "append" : "insert");
    if (index >= 0 && isFiltered()) {
      throw new Error("InsertDeleteConflict: This won't work because it would move cells in a table on your worksheet.");
    }
    assert.ok(rows.every(row => typeof row[2] === "string" && row[2].startsWith("=")));
    if (options.failInsert) throw new Error("InsertDeleteConflict: indsættelsen blev afvist");
    originalAdd(index, rows);
  };
  return { book, table, active, events };
}

const group = table => table.rows.filter(row => row[0] === 2026 && row[1].endsWith(" · p01"));
const activities = table => group(table).map(row => row[3]);
const addOffer = book => add(book, "p01", "2026", "24051", "Tilbud", "0");
const addCase = book => add(book, "p01", "2026", "24051", "Sag");

switch (scenario) {
  case "project": {
    const previous = setup();
    const oldAdd = load("add_project", "3.6.3");
    assert.match(oldAdd(previous.book, "p01", "2026", "24048", "Sag"), /tilføjet/);
    assert.throws(() => oldAdd(previous.book, "p01", "2026", "24051", "Sag"), /InsertDeleteConflict/);
    assert.deepEqual(previous.events, ["insert"]);
    const { book, table, active, events } = setup({ viewName: "" });
    assert.match(add(book, "p01", "2026", "24048", "Sag"), /tilføjet/);
    const before = structuredClone(table.rows);
    assert.match(addCase(book), /tilføjet/);
    const added = group(table)[4];
    assert.equal(added[4], "Sikringsprojekt · 24051");
    assert.deepEqual(activities(table), ["Sag", "Sag", "Sag", "Sag", "Sag", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"]);
    assert.deepEqual(table.rows.filter(row => row !== added), before);
    assert.equal(table.getRangeBetweenHeaderAndTotal().getCell(4, 7).getNumberFormat(), "0%");
    assert.deepEqual(active, criteria);
    assert.deepEqual(events, ["criteria", "view", ...restored]);
    break;
  }
  case "standard": {
    for (const action of [
      book => initialize(book, "p01", "2026"),
      book => create(book, ...book.tables.PersonRegister.rows[0].slice(1), 2026),
    ]) {
      const { book, table, active, events } = setup();
      const vacation = group(table).find(row => row[3] === "Ferie");
      table.rows.splice(table.rows.indexOf(vacation), 1);
      assert.match(action(book), /1 standardrækker oprettet/);
      assert.deepEqual(activities(table), ["Sag", "Sag", "Sag", "Sag", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"]);
      assert.deepEqual(active, criteria);
      assert.deepEqual(events, ["criteria", "view", ...restored]);
    }
    break;
  }
  case "bypass": {
    const { book, table, active, events } = setup({ defaultView: true });
    const count = table.rows.length;
    assert.match(add(book, "p01", "2026", "24048", "Sag"), /tilføjet/);
    assert.equal(table.rows.length, count);
    assert.deepEqual(events, []);
    assert.match(create(book, "Testperson", "Vest", "EL", "1734.03", 31, 2026), /5 standardrækker/);
    assert.deepEqual(events, ["append"]);
    assert.deepEqual(active, criteria);
    const unfiltered = setup({ unfiltered: true, defaultView: true });
    assert.match(addOffer(unfiltered.book), /tilføjet/);
    assert.deepEqual(unfiltered.events, ["insert"]);
    break;
  }
  case "view": {
    const standard = setup({ defaultView: true });
    assert.throws(() => addOffer(standard.book), /Vælg en personlig arkvisning/);
    assert.deepEqual(standard.events, ["criteria", "view"]);
    assert.deepEqual(standard.active, criteria);
    const unknown = setup({ incomplete: true });
    assert.throws(() => addOffer(unknown.book), /aktive filtre kunne ikke aflæses/);
    assert.deepEqual(unknown.events, ["criteria"]);
    assert.deepEqual(unknown.active, criteria);
    break;
  }
  case "clear": {
    const { book, table, active, events } = setup({ failClear: 1 });
    assert.throws(() => addOffer(book), /Rydning fejlede efter ændring/);
    assert.deepEqual(events, ["criteria", "view", "clear:0", "clear:1", "apply:0", "apply:1"]);
    assert.deepEqual(active, criteria);
    assert.deepEqual(table.rows, initial.tables.PlanTimer.rows);
    break;
  }
  case "insert": {
    const failed = setup({ failInsert: true });
    assert.throws(() => addOffer(failed.book), /Kontrollér de nye rækker.*InsertDeleteConflict/);
    assert.deepEqual(failed.active, criteria);
    assert.deepEqual(failed.events, ["criteria", "view", ...restored]);
    assert.deepEqual(failed.table.rows, initial.tables.PlanTimer.rows);
    const interleaved = setup();
    interleaved.table.conflict = true;
    assert.throws(() => addOffer(interleaved.book), /rækker ændrede sig under oprettelsen/);
    assert.deepEqual(interleaved.active, criteria);
    assert.deepEqual(interleaved.events, ["criteria", "view", ...restored]);
    break;
  }
  case "restore": {
    for (const failInsert of [false, true]) {
      const { book, active, table, events } = setup({ failInsert, failRestore: 1 });
      assert.throws(() => addOffer(book), error => {
        assert.match(error.message, /Genskabelse af PlanTimer-filtre mislykkedes for: Medarbejder/);
        assert.match(error.message, /Genskab disse filtre manuelt/);
        assert.equal(error.message.includes("InsertDeleteConflict"), failInsert);
        return true;
      });
      assert.deepEqual(active, [criteria[0], {}, {}, criteria[3]]);
      assert.deepEqual(events, ["criteria", "view", ...restored]);
      assert.equal(table.rows.length, initial.tables.PlanTimer.rows.length + (failInsert ? 0 : 1));
    }
    break;
  }
  case "snapshot": {
    const allTypes = [
      { filterOn: "values", values: ["2026", { date: "2026-11-01", specificity: "month" }] },
      { filterOn: "icon", icon: { index: 0, set: "threeFlags" } },
      { filterOn: "fontColor", color: "#123456" },
      { filterOn: "custom", criterion1: "=Sag", criterion2: "=Tilbud", operator: "or" },
      { filterOn: "dynamic", dynamicCriteria: "thisYear" },
      { filterOn: "bottomItems", criterion1: "5", subField: "Værdi" },
    ];
    const options = { criteria: allTypes, invalidateRead: true };
    const insert = setup(options);
    assert.match(addOffer(insert.book), /tilføjet/);
    assert.deepEqual(insert.active, allTypes);
    const probe = setup(options);
    assert.match(load("filter_probe")(probe.book), /Ryddede og genskabte 6 kolonnefiltre/);
    assert.deepEqual(probe.active, allTypes);
    assert.deepEqual(probe.table.rows, initial.tables.PlanTimer.rows);
    assert.equal(probe.events.includes("insert"), false);
    break;
  }
  case "restore-setup": {
    const { book, table } = setup({ failInsert: true });
    table.getRange = () => { throw new Error("Tabelområdet kunne ikke hentes"); };
    assert.throws(() => addOffer(book), error => {
      assert.match(error.message, /InsertDeleteConflict/);
      assert.match(error.message, /Tabelområdet kunne ikke hentes/);
      return true;
    });
    break;
  }
  default:
    throw new Error("Ukendt scenarie: " + scenario);
}
