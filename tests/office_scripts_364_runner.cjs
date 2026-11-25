const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { Range, testBook } = require("./office_scripts_mock.cjs");

const [compiled, fixture, scenario, version = "3.6.4"] = process.argv.slice(2);
const initial = JSON.parse(fs.readFileSync(fixture, "utf8"));
const load = (name, scriptVersion = version) => vm.runInNewContext(
  fs.readFileSync(path.join(compiled, scriptVersion, name + ".js"), "utf8") + "\nmain;",
  { console: { log() {} }, ExcelScript: { FilterOn: { custom: "custom" } } },
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
  table.getColumnByName = name => {
    const column = table.headers.indexOf(name);
    if (column < 0) return undefined;
    return { getFilter: () => ({
      getCriteria() {
        events.push("read:" + name);
        return active[column];
      },
      clear() {
        events.push("clear:" + column);
        active[column] = {};
        if (options.failClear === column) throw new Error("Rydning fejlede efter ændring");
      },
      applyCustomFilter(criterion1) {
        events.push("apply:" + column);
        if (options.failRestore === column) throw new Error("Excel kunne ikke anvende kriteriet");
        active[column] = { filterOn: original[column].filterOn, criterion1 };
      },
    }) };
  };
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
  case "custom-columns": {
    const year = { filterOn: "Custom", criterion1: "=2026" };
    const employee = { filterOn: "custom", criterion1: "=Medarbejder 1 · p01" };
    for (const activeFilters of [[], [year], [{}, employee], [year, employee]]) {
      const { book, table, active, events } = setup({ criteria: activeFilters, viewName: "" });
      assert.match(add(book, "p01", "2026", "24048", "Sag"), /tilføjet/);
      assert.deepEqual(events, activeFilters.length ? table.headers.map(name => "read:" + name) : []);
      events.length = 0;
      const before = structuredClone(table.rows);
      assert.match(addCase(book), /tilføjet/);
      const added = group(table)[4];
      assert.equal(added[4], "Sikringsprojekt · 24051");
      assert.equal(added[1], "Medarbejder 1 · p01");
      assert.equal(added[0], 2026);
      assert.equal(group(table)[5][3], "Tilbud");
      assert.deepEqual(table.rows.filter(row => row !== added), before);
      for (const index of [2, 5, 6, 8, 62]) assert.match(added[index], /^=/);
      assert.equal(table.getRangeBetweenHeaderAndTotal().getCell(4, 7).getNumberFormat(), "0%");
      assert.deepEqual(active, activeFilters);
      const columns = activeFilters.flatMap((criteria, column) => criteria.filterOn ? [column] : []);
      const writes = [...columns.map(column => "clear:" + column), "insert",
        ...columns.map(column => "apply:" + column)];
      assert.deepEqual(events, columns.length ? [
        ...table.headers.map(name => "read:" + name), "view", ...writes,
      ] : writes);
      assert.match(addCase(book), /findes allerede/);
      assert.equal(events.filter(event => event === "insert").length, 1);
    }
    break;
  }
  case "custom-refusal": {
    const year = { filterOn: "custom", criterion1: "=2026" };
    for (const activeFilters of [
      [year, { filterOn: "values", values: ["Medarbejder 1 · p01"] }],
      [year, { filterOn: "custom", criterion1: "=Medarbejder*" }],
      [{ filterOn: "custom", criterion1: ">2025" }],
      [{ filterOn: "custom", criterion1: "=2026", criterion2: "=2027" }],
      [{ filterOn: "custom", criterion1: "=2026", subField: "tekst" }],
      [year, {}, {}, { filterOn: "custom", criterion1: "=Sag" }],
      [year, {}, {}, {}, { filterOn: "fontColor", color: "#123456" }],
    ]) {
      const { book, table, active, events } = setup({ criteria: activeFilters });
      assert.throws(() => addOffer(book), /kan ikke genskabes sikkert/);
      assert.deepEqual(active, activeFilters);
      assert.deepEqual(table.rows, initial.tables.PlanTimer.rows);
      assert.ok(events.every(event => event.startsWith("read:")));
    }
    for (const options of [
      { criteria: [year, {}, {}, { filterOn: "custom", criterion1: "=Sag" }] },
      { criteria: [year], defaultView: true },
    ]) {
      const mixed = setup(options);
      const employeeRows = group(mixed.table);
      const otherRows = mixed.table.rows.filter(row => !employeeRows.includes(row));
      const remaining = employeeRows.filter(row => !["Ferie", "Andet"].includes(row[3]));
      mixed.table.rows.splice(0, mixed.table.rows.length, ...otherRows, ...remaining);
      const before = structuredClone(mixed.table.rows);
      assert.throws(() => initialize(mixed.book, "p01", "2026"),
        /kan ikke genskabes sikkert|personlig arkvisning/);
      assert.deepEqual(mixed.table.rows, before);
      assert.ok(mixed.events.every(event => event.startsWith("read:") || event === "view"));
    }
    const standard = setup({ criteria: [year], defaultView: true });
    assert.throws(() => addOffer(standard.book), /personlig arkvisning/);
    assert.deepEqual(standard.active, [year]);
    assert.deepEqual(standard.table.rows, initial.tables.PlanTimer.rows);
    const supportedAppend = setup({ criteria: [year], defaultView: true });
    assert.match(create(supportedAppend.book, "Testperson", "Vest", "EL", "1734.03", 31, 2026), /5 standardrækker/);
    assert.ok(supportedAppend.events.slice(0, -1).every(event => event.startsWith("read:")));
    assert.equal(supportedAppend.events.at(-1), "append");
    assert.deepEqual(supportedAppend.active, [year]);
    const unsupported = setup({ criteria: [{ filterOn: "values", values: ["2026"] }] });
    const beforeWrites = structuredClone(unsupported.book.state);
    assert.throws(() => create(unsupported.book, "Testperson", "Vest", "EL", "1734.03", 31, 2032),
      /kan ikke genskabes sikkert/);
    assert.throws(() => add(unsupported.book, "p01", "2026", "24048", "Sag"), /kan ikke genskabes sikkert/);
    assert.deepEqual(unsupported.book.state, beforeWrites);
    assert.ok(unsupported.events.every(event => event.startsWith("read:")));
    break;
  }
  case "custom-failure": {
    const activeFilters = [
      { filterOn: "Custom", criterion1: "=2026" },
      { filterOn: "Custom", criterion1: "=Medarbejder 1 · p01" },
    ];
    for (const options of [{ failClear: 1 }, { failInsert: true }, { failInsert: true, failRestore: 0 },
      { failRestore: 0 }]) {
      const { book, table, active, events } = setup({ criteria: activeFilters, ...options });
      assert.throws(() => addOffer(book), error => {
        assert.match(error.message, /Ændringer kan være delvist gennemført/);
        if (options.failInsert) assert.match(error.message, /InsertDeleteConflict/);
        if (options.failRestore === 0) assert.match(error.message, /Genskabelse.*Planår/);
        if (options.failClear) assert.match(error.message, /Rydning fejlede/);
        return true;
      });
      assert.deepEqual(events.slice(-2), ["apply:0", "apply:1"]);
      assert.deepEqual(active, options.failRestore === 0 ? [{}, activeFilters[1]] : activeFilters);
      assert.equal(table.rows.length, initial.tables.PlanTimer.rows.length +
        (options.failInsert || options.failClear ? 0 : 1));
    }
    break;
  }
  case "custom-standard": {
    for (const action of [
      book => initialize(book, "p01", "2026"),
      book => create(book, ...book.tables.PersonRegister.rows[0].slice(1), 2026),
    ]) {
      const filter = { filterOn: "custom", criterion1: "=2026" };
      const { book, table, active, events } = setup({ criteria: [filter] });
      const vacation = group(table).find(row => row[3] === "Ferie");
      table.rows.splice(table.rows.indexOf(vacation), 1);
      assert.match(action(book), /1 standardrækker oprettet/);
      assert.ok(group(table).some(row => row[3] === "Ferie"));
      assert.deepEqual(events.slice(-3), ["clear:0", "insert", "apply:0"]);
      assert.deepEqual(active, [filter]);
    }
    break;
  }
  case "confirmed-367": {
    const criteria = [
      { filterOn: "custom", criterion1: "=2026" },
      { filterOn: "custom", criterion1: "=Medarbejder 1 · p01" },
    ];
    const { book, table, active, events } = setup({ criteria, defaultView: true });
    const before = structuredClone(table.rows);
    assert.throws(() => addOffer(book), /bekræft valget/);
    assert.deepEqual(table.rows, before);
    assert.ok(events.every(event => event.startsWith("read:")));
    assert.match(add(book, "p01", "2026", "24048", "Sag", "100"), /^Projektet er tilføjet/);
    assert.equal(group(table)[3][7], "");
    assert.match(add(book, "p01", "2026", "24051", "Sag", undefined, true), /^Projektet er tilføjet/);
    assert.deepEqual(events.slice(-5), ["clear:0", "clear:1", "insert", "apply:0", "apply:1"]);
    assert.deepEqual(active, criteria);
    const after = structuredClone(table.rows);
    assert.throws(() => addCase(book), /findes allerede.*Ingen ændringer er foretaget/);
    assert.throws(() => add(book, "p01", "2026", "Ukendt"), /findes ikke i Projektregister/);
    assert.deepEqual(table.rows, after);
    assert.deepEqual(after.filter(row => !["Belysningsprojekt · 24048", "Sikringsprojekt · 24051"].includes(row[4]) ||
      row[1] !== "Medarbejder 1 · p01" || row[0] !== 2026),
      before.filter(row => row !== before[3]));

    for (const action of [
      (book, confirm) => initialize(book, "p01", "2026", confirm),
      (book, confirm) => create(book, ...book.tables.PersonRegister.rows[0].slice(1), 2026, confirm),
    ]) {
      const item = setup({ criteria });
      const row = group(item.table).find(row => row[3] === "Ferie");
      item.table.rows.splice(item.table.rows.indexOf(row), 1);
      assert.throws(() => action(item.book, false), /bekræft valget/);
      assert.match(action(item.book, true), /1 standardrækker oprettet/);
      assert.deepEqual(item.active, criteria);
    }
    break;
  }
  case "probabilities-367": {
    const { book, table } = setup({ unfiltered: true });
    const cases = [["24022", undefined, 0.5], ["24031", "50", 0.5], ["24048", "0", 0],
      ["24051", "100", 1]];
    for (const [id, input, expected] of cases) {
      assert.match(add(book, "p01", "2026", id, "Tilbud", input), /^Projektet er tilføjet/);
      const index = table.rows.findIndex(row => row[0] === 2026 && row[1].endsWith(" · p01") &&
        row[3] === "Tilbud" && row[4].endsWith(" · " + id));
      assert.equal(table.rows[index][7], expected);
      assert.equal(table.getRangeBetweenHeaderAndTotal().getCell(index, 7).getNumberFormat(), "0%");
    }
    break;
  }
  case "upgrade-367": {
    const workbook = () => {
      let block = structuredClone(initial.upgrade.formulas);
      const banners = { ...initial.upgrade.banners };
      let writes = 0;
      return {
        getWorksheet: name => ({
          getRange: address => ({
            getValue: () => banners[name],
            setValues: ([[value]]) => { banners[name] = value; writes++; },
            getFormulas: () => address === "F8" ? [["=SUM(F590:F689)"]] : structuredClone(block),
            setFormulas: values => { block = structuredClone(values); writes++; },
          }),
        }),
        read: () => ({ block, banners, writes }),
        alter: () => { block[99][52] = "=1"; },
      };
    };
    const run = load("upgrade_workbook");
    const copy = workbook();
    assert.match(run(copy), /3.6.7/);
    assert.deepEqual(copy.read().block, initial.upgrade.expected);
    assert.ok(Object.values(copy.read().banners).every(value => value.endsWith("3.6.7")));
    assert.equal(copy.read().writes, 7);
    run(copy);
    assert.equal(copy.read().writes, 7);
    const custom = workbook();
    custom.alter();
    assert.throws(() => run(custom), /BH689 afviger/);
    assert.equal(custom.read().writes, 0);
    break;
  }
  default:
    throw new Error("Ukendt scenarie: " + scenario);
}
