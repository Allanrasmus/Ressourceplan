const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const [compiled, fixture, result] = process.argv.slice(2);
const initial = JSON.parse(fs.readFileSync(fixture, "utf8"));
const load = name => vm.runInNewContext(
  fs.readFileSync(path.join(compiled, name + ".js"), "utf8") + "\nmain;",
  { console: { log() {} } },
);
const create = load("create_employee");
const add = load("add_project");
const initialize = load("initialize_plan_year");

const { testBook } = require("./office_scripts_mock.cjs");
const book = () => testBook(initial);
function profile(w, year = 2026) { return create(w, "Script Testperson", "Vest", "EL", "QA36", 30, year); }
const before = book();
const recordCount = before.tables.PlanTimer.rows.length;
const beforeRows = JSON.stringify(before.tables.PlanTimer.rows);
assert.throws(() => create(before, " ", "Vest", "EL", "QA36", 30, 2026), /skal udfyldes/);
assert.throws(() => create(before, "Testperson", "Nord", "EL", "QA36", 30, 2026), /gyldigt område/);
assert.throws(() => create(before, "Testperson", "Vest", "EL", "QA36", 81, 2026), /Normtid/);
assert.throws(() => initialize(before, "p01", 2026.5), /helt år/);
assert.throws(() => add(before, "ukendt", 2026, "24015", "Sag"), /entydigt/);
assert.throws(() => add(before, "p01", 2026, "24015", "Ferie"), /Sag eller Tilbud/);
assert.throws(() => add(before, "p01", 2026, "24015", "Sag", 25), /Tilbudsprocent/);
assert.throws(() => add(before, "p01", 2026, "24015", "Tilbud", 101), /Tilbudsprocent/);
assert.equal(JSON.stringify(before.tables.PlanTimer.rows), beforeRows);
assert.deepEqual(before.state, initial);
assert.match(initialize(before, "p05", 2026), /0 standardrækker/);

const retry = book();
retry.tables.PlanTimer.failAdd = true;
assert.throws(() => profile(retry), /Kontrollér de nye rækker/);
assert.equal(retry.tables.PersonRegister.rows.filter(row => row[0] === "p31").length, 1);
assert.equal(retry.tables.PlanTimer.rows.length, recordCount);
retry.tables.PlanTimer.failAdd = false;
assert.match(profile(retry), /p31.*5 standardrækker/);
assert.match(profile(retry), /p31.*0 standardrækker/);

const bad = book();
bad.state.settings[6][16] = "";
assert.throws(() => profile(bad), /formelskabelon/);
assert.equal(bad.tables.PersonRegister.rows.filter(row => row[0] === "p31").length, 0);

const reserved = book();
const staff = reserved.tables.PersonRegister.rows;
for (let n = 30; n < staff.length; n++) {
  staff[n] = ["p" + (n + 1), "Anden person " + n, "Øst", "EL", "QA", 30];
}
staff.reverse();
assert.match(profile(reserved), /p30.*5 standardrækker/);
assert.equal(staff.filter(row => row[0] === "p30").length, 1);

const conflict = book();
conflict.tables.PlanTimer.conflict = true;
assert.throws(() => add(conflict, "p02", 2026, "24015", "Sag"), /ændrede sig/);
assert.equal(JSON.stringify(conflict.tables.PlanTimer.rows.slice(0, recordCount)), beforeRows);

const state = book();
assert.match(profile(state), /p31.*5 standardrækker/);
assert.match(profile(state), /p31.*0 standardrækker/);
const plan = state.tables.PlanTimer;
assert.deepEqual(plan.rows.slice(recordCount).map(row => row[3]), ["Sag", "Ferie", "Fravær", "Kursus", "Andet"]);
assert.ok(plan.rows.slice(recordCount).every(row => row[0] === 2026 && row[1] === "Script Testperson · p31"));
assert.match(initialize(state, "p31", 2032), /5 standardrækker/);
assert.match(initialize(state, "p31", 2032), /0 standardrækker/);
assert.equal(state.state.settings[12][5], 2032);
assert.match(add(state, "p31", 2026, "24015", "Sag"), /tilføjet/);
plan.rows.at(-1)[56] = 8;
plan.rows.at(-1)[61] = 9;
assert.match(add(state, "p31", 2026, "24015", "Sag"), /findes allerede/);
assert.match(add(state, "Script Testperson · p31", 2026, "24015", "Tilbud", 25), /tilføjet/);
plan.rows.at(-1)[56] = 12;
assert.equal(plan.rows.at(-1)[7], 0.25);
assert.match(add(state, "p31", 2027, "24015", "Sag"), /tilføjet/);
plan.rows.at(-1)[61] = 20; // Copied bad input must not count in a 52-week year.
assert.match(add(state, "p31", 2026, "24022", "Tilbud", 0), /tilføjet/);
assert.equal(plan.rows.at(-1)[7], 0);
assert.match(add(state, "p31", 2026, "24031", "Tilbud"), /tilføjet/);
assert.equal(plan.rows.at(-1)[7], "");
assert.equal(JSON.stringify(plan.rows.slice(0, recordCount)), beforeRows);

// A stale caption is resolved using the stable key when checking duplicates.
state.tables.PersonRegister.rows.find(row => row[0] === "p31")[1] = "Omdøbt Testperson";
state.tables.SagRegister.rows[0][1] = "Omdøbt Hospital";
assert.match(add(state, "p31", 2026, "24015", "Sag"), /findes allerede/);
assert.match(initialize(state, "p31", 2026), /0 standardrækker/);

const distinct = book();
profile(distinct);
assert.match(create(distinct, "Script Testperson", "Øst", "EL", "QA36", 30, 2026), /p32/);
assert.throws(() => initialize(distinct, "Script Testperson", 2027), /entydigt/);
distinct.tables.SagRegister.rows[5] = ["00123", "Hospital Nord"];
assert.throws(() => add(distinct, "p31", 2026, "Hospital Nord", "Sag"), /entydigt/);
assert.match(add(distinct, "p31", 2026, "00123", "Sag"), /00123/);
assert.equal(distinct.tables.PlanTimer.rows.at(-1)[4], "Hospital Nord · 00123");

fs.writeFileSync(result, JSON.stringify(state.state));
