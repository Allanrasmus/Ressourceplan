const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { testBook } = require("./office_scripts_mock.cjs");

const [compiled, fixture, result] = process.argv.slice(2);
const initial = JSON.parse(fs.readFileSync(fixture, "utf8"));
const load = name => vm.runInNewContext(
  fs.readFileSync(path.join(compiled, name + ".js"), "utf8") + "\nmain;",
  { console: { log() {} } },
);
const create = load("create_employee");
const add = load("add_project");
const init = load("initialize_plan_year");
const book = () => testBook(initial);
const profile = (w, year = 2026) => create(w, "Testperson 361", "Vest", "EL", "QA361", 30, year);
const select = (sheet, row, column) => ({ sheet, row, column });
const planSelection = (row, column) => select("Min ressourceplan", row, column);
const originalRows = JSON.stringify(initial.tables.PlanTimer.rows);

const history = book();
const oldRow = history.tables.PlanTimer.rows[0].slice();
oldRow[0] = 2028;
oldRow[1] = "Historisk navn · p06";
oldRow[2] = "p06";
history.tables.PlanTimer.rows.push(oldRow);
assert.match(profile(history), /p07.*5 standardrækker/);
assert.deepEqual(history.tables.PersonRegister.rows[5], ["p06", "", "", "", "", ""]);
assert.deepEqual(history.tables.PlanTimer.rows.at(-6), oldRow);

const oldWorkbook = book();
oldWorkbook.state.settings[5][15] = "PlanTimer-formler 3.6";
oldWorkbook.tables.PersonRegister.rows[30] = ["p31", "Testperson 361", "Vest", "EL", "QA361", 30];
assert.match(profile(oldWorkbook), /p31.*5 standardrækker/);
assert.equal(oldWorkbook.tables.PersonRegister.rows[5][1], "");
assert.match(profile(oldWorkbook), /p31.*0 standardrækker/);

const retry = book();
retry.tables.PlanTimer.failAdd = true;
assert.throws(() => profile(retry), /Kontrollér de nye rækker/);
assert.equal(retry.tables.PersonRegister.rows[5][0], "p06");
retry.tables.PlanTimer.failAdd = false;
assert.match(profile(retry), /p06.*5 standardrækker/);
assert.equal(retry.tables.PersonRegister.rows.filter(row => row[0] === "p06").length, 1);

const preflight = book();
const register = preflight.tables.PersonRegister;
let reads = 0;
const previous = register.getRangeBetweenHeaderAndTotal.bind(register);
register.getRangeBetweenHeaderAndTotal = () => {
  if (++reads === 2) register.rows[5][1] = "Samtidig ændring";
  return previous();
};
assert.throws(() => profile(preflight), /blev ændret/);
assert.equal(register.rows[5][1], "Samtidig ændring");
assert.equal(JSON.stringify(preflight.tables.PlanTimer.rows), originalRows);

const state = book();
assert.match(profile(state), /p06.*5 standardrækker/);
const plan = state.tables.PlanTimer;
assert.deepEqual(state.tables.PersonRegister.rows[5], ["p06", "Testperson 361", "Vest", "EL", "QA361", 30]);
initial.tables.PersonRegister.rows.forEach((row, i) => {
  if (i !== 5) assert.deepEqual(state.tables.PersonRegister.rows[i], row);
});
const firstNewRow = initial.tables.PlanTimer.rows.length + 10;
state.selection = planSelection(firstNewRow, 2);
assert.match(init(state), /p06.*2026.*0 standardrækker/);
assert.match(init(state, "", "2032"), /p06.*2032.*5 standardrækker/);
assert.match(init(state, "", "2032"), /p06.*2032.*0 standardrækker/);
assert.equal(state.state.settings[12][5], 2032);
assert.match(add(state, "", "", "Hospital Nord · 24015", "Sag", "uforståelig"), /p06.*2026/);
plan.rows.at(-1)[56] = 8;
assert.equal(plan.rows.at(-1)[7], "");
assert.match(add(state, "Testperson 361", "2026", "Datacenter", "Sag"), /p06/);
assert.equal(plan.rows.at(-1)[7], "");
assert.match(add(state, "p06", "2026", "24015", "Sag", "25"), /findes allerede/);

for (const [code, input, stored] of [
  ["24015", "", ""], ["24022", "0", 0], ["24031", "25", 0.25],
  ["24048", "50", 0.50], ["24051", "100", 1],
]) {
  assert.match(add(state, "p06", "2026", code, "Tilbud", input), /tilføjet/);
  assert.equal(plan.rows.at(-1)[7], stored);
  plan.rows.at(-1)[56] = 8;
}
for (const value of ["-1", "101", "Infinity", "uforståeligt"]) {
  assert.throws(() => add(state, "p06", "2027", "24015", "Tilbud", value), /0–100/);
}
const emptyFields = book();
profile(emptyFields);
for (const [code, value, expected] of [
  ["24015", undefined, ""], ["24022", null, ""], ["24031", " ", ""], ["24048", 0, 0],
]) {
  assert.match(add(emptyFields, "p06", "2026", code, "Tilbud", value), /tilføjet/);
  assert.equal(emptyFields.tables.PlanTimer.rows.at(-1)[7], expected);
}
assert.throws(() => add(state, "p06", "0", "24015", "Tilbud"), /helt år/);
assert.equal(JSON.stringify(plan.rows.slice(0, initial.tables.PlanTimer.rows.length)), originalRows);

const picker = book();
assert.match(profile(picker), /p06/);
picker.selection = select("Projektregister", 7, 2);
assert.match(add(picker, "Testperson 361", "2026", "", "Sag"), /Hospital Nord · 24015/);
picker.selection = planSelection(picker.tables.PlanTimer.rows.length + 9, 5);
assert.match(add(picker, "", "", "", "Tilbud", "25,5"), /Hospital Nord.*p06/);
assert.equal(picker.tables.PlanTimer.rows.at(-1)[7], 0.255);
picker.selection = planSelection(firstNewRow, 2);
assert.throws(() => add(picker, "", "", "", "Sag"), /projektcelle/);
assert.throws(() => init(picker, "Medarbejder 1"), /samme|valgte medarbejder/);
for (const selection of [
  planSelection(9, 2), planSelection(400, 2), planSelection(10, 4),
  { ...planSelection(10, 2), hidden: true }, { ...planSelection(10, 2), width: 2 },
  { ...planSelection(10, 2), multi: true },
]) {
  picker.selection = selection;
  assert.throws(() => init(picker), /Markér|markér/);
}
assert.match(create(picker, "Testperson 361", "Øst", "EL", "QA361", 30, 2026), /p07/);
assert.throws(() => init(picker, "Testperson 361", "2032"), /entydigt/);
picker.selection = planSelection(firstNewRow, 2);
assert.match(init(picker, "", "2032"), /p06/);
picker.tables.SagRegister.rows[5] = ["00123", "Hospital Nord"];
assert.throws(() => add(picker, "p06", "2032", "Hospital Nord", "Sag"), /entydigt/);
assert.match(add(picker, "p06", "2032", "Hospital Nord · 00123", "Sag"), /00123/);
assert.equal(picker.tables.PlanTimer.rows.at(-1)[4], "Hospital Nord · 00123");

fs.writeFileSync(result, JSON.stringify(state.state));
