const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { Range, testBook } = require("./office_scripts_mock.cjs");

const [compiled, fixture, result, packaged] = process.argv.slice(2);
const initial = JSON.parse(fs.readFileSync(fixture, "utf8"));
initial.locale = "da-DK";
const load = name => vm.runInNewContext(
  fs.readFileSync(path.join(compiled, name + ".js"), "utf8") + "\nmain;",
  { console: { log() {} } },
);
const create = load("create_employee");
const add = load("add_project");
const initialize = load("initialize_plan_year");
const upgrade = load("upgrade_workbook");
const originalPlan = initial.tables.PlanTimer.rows;
const projects = ["24015", "24022", "24031", "24048", "24051"];
const group = (book, id, year) => book.tables.PlanTimer.rows.filter(row =>
  row[0] === year && row[1].endsWith(" · " + id));
const activity = rows => rows.map(row => row[3]);

const book = testBook(initial);
const codes = ["7134.01", "1734.03", "7134.10", "01734.03"];
codes.forEach((department, index) => {
  const id = "p0" + (index + 6);
  assert.match(create(book, "Testperson " + (index + 6), "Vest", "EL", department, 31, 2026), /5 standardrækker/);
  const entry = book.tables.PersonRegister.rows.find(row => row[0] === id);
  assert.deepEqual(entry, [id, "Testperson " + (index + 6), "Vest", "EL", department, 31]);
});
const plan = book.tables.PlanTimer;
plan.filters = { employee: "Medarbejder 1 · p01", year: 2026 };
const oldFilters = structuredClone(plan.filters);
const free = group(book, "p06", 2026)[0];
const freeContents = free.slice();
const count = plan.rows.length;
const planRow = plan.rows.indexOf(free) + 10;
book.selection = { sheet: "Min ressourceplan", row: planRow, column: 2 };
assert.match(add(book, "", "", "Hospital Nord", "Sag", "25"), /p06 · 2026/);
assert.equal(plan.rows.length, count);
assert.equal(group(book, "p06", 2026)[0], free);
freeContents[4] = "Hospital Nord · 24015";
assert.deepEqual(free, freeContents);
free[56] = 17;
assert.match(add(book, "Testperson 6", "2026", projects[0], "Sag"), /findes allerede/);
assert.throws(() => add(book, "Testperson 6", "2026", "Ukendt projekt", "Sag"), /først oprettet i Projektregister/);
assert.deepEqual(book.tables.SagRegister.rows, initial.tables.SagRegister.rows);

// Interleave additions so cases must be inserted before existing offers and leave all hours in place.
assert.match(add(book, "Testperson 6", "2026", projects[1], "Tilbud"), /tilføjet/);
assert.match(add(book, "Testperson 6", "2026", projects[2], "Tilbud", "0"), /tilføjet/);
for (const code of projects.slice(1)) {
  assert.match(add(book, "Testperson 6", "2026", code, "Sag"), /tilføjet/);
}
assert.deepEqual(activity(group(book, "p06", 2026)),
  ["Sag", "Sag", "Sag", "Sag", "Sag", "Tilbud", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"]);
const offers = group(book, "p06", 2026).filter(row => row[3] === "Tilbud");
offers[0][56] = 5;
assert.deepEqual(offers.map(row => row[7]), ["", 0]);
assert.equal(group(book, "p06", 2026)[0][56], 17);
assert.equal(initialize(book, "", "2032").includes("5 standardrækker"), true);
assert.match(initialize(book, "", "2032"), /0 standardrækker/);
book.selection = { sheet: "Projektregister", row: 7, column: 1 };
assert.match(add(book, "Testperson 6", "2032", "", "Sag"), /24015.*2032/);
assert.deepEqual(activity(group(book, "p06", 2032)), ["Sag", "Ferie", "Fravær", "Kursus", "Andet"]);
assert.deepEqual(activity(group(book, "p07", 2026)), ["Sag", "Ferie", "Fravær", "Kursus", "Andet"]);
assert.deepEqual(plan.filters, oldFilters);
assert.deepEqual(plan.rows.slice(0, originalPlan.length), originalPlan);

const installed = testBook(JSON.parse(fs.readFileSync(packaged, "utf8")));
const installedCount = installed.tables.PlanTimer.rows.length;
const installedPlan = group(installed, "p01", 2026);
assert.deepEqual(activity(installedPlan), ["Sag", "Sag", "Sag", "Sag", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"]);
assert.equal(installedPlan[3][4], "");
assert.match(add(installed, "Medarbejder 1", "2026", "24048", "Sag"), /24048/);
assert.equal(installed.tables.PlanTimer.rows.length, installedCount);
assert.equal(installedPlan[3][4], "Belysningsprojekt · 24048");
assert.match(add(installed, "Medarbejder 1", "2026", "24051", "Tilbud"), /24051/);
assert.deepEqual(activity(group(installed, "p01", 2026)),
  ["Sag", "Sag", "Sag", "Sag", "Tilbud", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"]);

const existing = testBook(initial);
const vacant = group(existing, "p02", 2026).find(row => row[3] === "Sag" && row[4] === "");
assert.match(add(existing, "Medarbejder 2", "2026", "24015", "Sag"), /p02/);
assert.equal(vacant[4], "Hospital Nord · 24015");
assert.equal(existing.tables.PlanTimer.rows.length, originalPlan.length);
// Keys with leading zeros stay distinct when selected by name or by exact number.
existing.tables.SagRegister.rows[5] = ["0024015", "Projekt med nulprefix"];
assert.match(add(existing, "Medarbejder 2", "2026", "0024015", "Tilbud"), /0024015/);
assert.equal(group(existing, "p02", 2026).find(row => row[3] === "Tilbud")[4], "Projekt med nulprefix · 0024015");
assert.deepEqual(activity(group(existing, "p02", 2026)), ["Sag", "Sag", "Tilbud", "Ferie", "Fravær", "Kursus", "Andet"]);
const legacy = group(existing, "p01", 2026);
assert.match(add(existing, "p01", "2026", "24048", "Sag"), /24048/);
const updated = group(existing, "p01", 2026);
assert.equal(updated[3][4], "Belysningsprojekt · 24048");
assert.deepEqual(updated.filter(row => row !== updated[3]), legacy);

for (const [column, value] of [[7, 0], [9, 0], [61, 5], [9, '=""'], [4, '=""']]) {
  const protectedBook = testBook(initial);
  create(protectedBook, "Testperson 6", "Vest", "EL", "1734.03", 31, 2026);
  const reserved = group(protectedBook, "p06", 2026)[0];
  reserved[column] = value;
  const keep = reserved.slice();
  assert.match(add(protectedBook, "p06", "2026", "24015", "Sag"), /tilføjet/);
  assert.deepEqual(reserved, keep);
  assert.deepEqual(activity(group(protectedBook, "p06", 2026)),
    ["Sag", "Sag", "Ferie", "Fravær", "Kursus", "Andet"]);
}

const race = testBook(initial);
const raceTable = race.tables.PlanTimer;
const originalDataRange = raceTable.getRangeBetweenHeaderAndTotal.bind(raceTable);
let rangeReads = 0;
raceTable.getRangeBetweenHeaderAndTotal = () => {
  if (++rangeReads === 2) raceTable.rows[2][4] = "Ændret af en anden";
  return originalDataRange();
};
assert.throws(() => add(race, "p01", "2026", "24048", "Sag"), /indsættelsesstedet ændrede sig/);
assert.equal(raceTable.rows.length, originalPlan.length);

function upgradeBook() {
  const target = testBook(initial);
  const workbookNames = new Map([["PlanSidsteRaekke", "=IFERROR(ROW(INDEX(PlanTimer[#Data],ROWS(PlanTimer[#Data]),1)),9)"]]);
  const rules = [
    { id: "blue", formula: '=AND($D10="Sag",MOD(ROW(),2)=1)', fill: "#D2EAF7", priority: 14, border: "#DBE5EF" },
    { id: "white", formula: '=AND($D10="Sag",MOD(ROW(),2)=0)', fill: "#FFFFFF", priority: 19, border: "#DBE5EF" },
    { id: "offer", formula: '=$D10="Tilbud"', fill: "#FCE3E3", priority: 9, border: "#DBE5EF" },
  ];
  const wrap = data => ({
    getId: () => data.id,
    getRanges: () => ({
      getAddress: () => "'Min ressourceplan'!A10:BJ20009",
      getAreaCount: () => data.areas ?? 1,
    }),
    getCustom: () => ({
      getRule: () => ({ getFormula: () => data.formula, setFormula: formula => { data.formula = formula; } }),
      getFormat: () => ({ getFill: () => ({ setColor: color => { data.fill = color; } }) }),
    }),
  });
  const settings = target.getWorksheet("Planstart");
  const overview = new Range([["Alle"]], 0, 0, 1, 1);
  const legend = { getValue: () => "Sag", getFormat: () => ({ getFill: () => ({ setColor() {} }) }) };
  const planArea = {
    getConditionalFormats: () => rules.map(wrap),
    getConditionalFormat: id => wrap(rules.find(rule => rule.id === id)),
  };
  target.getWorksheet = name => ({
    "Planstart": settings,
    "Min ressourceplan": { getRange: address => address === "A5" ? legend : planArea },
    "Ressourceoverblik": { getRange: () => overview },
  })[name];
  target.getNamedItem = name => workbookNames.has(name) ? { getFormula: () => workbookNames.get(name) } : undefined;
  target.addNamedItem = (name, formula) => { workbookNames.set(name, formula); };
  return { target, rules, workbookNames, overview };
}
const update = upgradeBook();
const { target, rules, overview } = update;
const originalTables = structuredClone(target.state.tables);
const untouched = structuredClone(rules[2]);
assert.match(upgrade(target), /opdateret/);
assert.deepEqual(target.state.tables, originalTables);
assert.deepEqual(rules[2], untouched);
assert.equal(rules[0].priority, 14);
assert.equal(rules[0].border, "#DBE5EF");
assert.equal(rules[0].fill, "#EAF2FC");
assert.equal(rules[1].fill, "#FFFFFF");
assert.match(rules[0].formula, /INDEX\(PlanSagsstribe,ROW\(\)-9\)=1/);
assert.equal(overview.getNumberFormat(), "@");
assert.equal(target.tables.PersonRegister.getRangeBetweenHeaderAndTotal().getCell(0, 4).getNumberFormat(), "@");
assert.equal(target.tables.PersonRegister.getRangeBetweenHeaderAndTotal().getCell(0, 5).getNumberFormat(), "General");
assert.equal(update.workbookNames.get("PlanSagsstribe"), "='Planstart'!$R$10:$R$20009");
assert.match(upgrade(target), /opdateret/);
assert.deepEqual(target.state.tables, originalTables);
const collision = upgradeBook();
collision.target.state.settings[8][17] = "Brugerdata";
assert.throws(() => upgrade(collision.target), /allerede andre oplysninger/);
assert.equal(collision.workbookNames.has("PlanSagsstribe"), false);
assert.equal(collision.rules[0].fill, "#D2EAF7");

const ranges = upgradeBook();
ranges.rules[0].areas = 2;
assert.throws(() => upgrade(ranges.target), /opdelt/);
assert.equal(ranges.rules[0].fill, "#D2EAF7");

const split = upgradeBook();
split.rules.push({ id: "split", formula: '=AND($D55="Sag",MOD(ROW(),2)=1)', fill: "#D2EAF7" });
assert.throws(() => upgrade(split.target), /opdelt/);
assert.equal(split.workbookNames.has("PlanSagsstribe"), false);
assert.equal(split.rules[0].fill, "#D2EAF7");

fs.writeFileSync(result, JSON.stringify(book.state));
