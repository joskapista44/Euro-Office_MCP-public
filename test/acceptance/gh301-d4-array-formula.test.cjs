'use strict';

/**
 * GH #301 D4: write_array_formula -- the public SetFormulaArray over an EMPTY area, guarded by the LIVE design probe
 * (docs/evidence/gh301-d4-array-formula-probe-20261025.json): SetFormulaArray silently overwrites values and silently
 * builds overlapping arrays, so the write is refused before it runs unless every area cell is empty and no cell of the
 * area or its one-cell ring is an array member; L1 proves the extent (every area cell a member with the written
 * formula, no ring cell a member).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const af = require('../../src/adapters/xlsx/families/array-formula.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const book = () => [{ name: 'A', cells: { A1: { v: 1 }, A2: { v: 2 }, A3: { v: 3 }, B1: { v: 'keep' } } }, { name: 'Other', cells: { A1: { v: 'x' } } }];
const stepOf = (extra = {}) => ({ intent: 'write_array_formula', sheet: 'A', range: 'G1:G3', formula: '=ROW(A1:A3)*2', ...extra });
const task = step => ({ id: 'gh301-d4', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const members = (store, cells) => cells.map(a => (store.durable[0].cells[a] || {}).array || null);

test('D4 schema: sheet + range + formula accepted; an extra field, a non-formula, an external reference, one cell refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(stepOf()).ok, true, JSON.stringify(v(stepOf())));
  assert.equal(v(stepOf({ values: [[1]] })).ok, false, 'an unsupported field');
  assert.equal(v(stepOf({ formula: 'ROW(A1:A3)' })).ok, false, 'no leading =');
  assert.equal(v(stepOf({ formula: '=[book.xlsx]S!A1' })).ok, false, 'an external workbook reference');
  assert.equal(v(stepOf({ range: 'G1' })).ok, false, 'a single cell is not measured');
});

test('D4: an empty area -- APPLIED, every area cell a member of ONE array G1:G3, nothing else changed; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.deepEqual(members(store, ['G1', 'G2', 'G3', 'G4']), ['G1:G3', 'G1:G3', 'G1:G3', null]);
  assert.equal(store.durable[0].cells.B1.v, 'keep');
  const again = await run(store, task(stepOf()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('D4: a 2-D area J1:K2 -- APPLIED, all four cells members', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(stepOf({ range: 'J1:K2', formula: '=A1:B2*1' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(members(store, ['J1', 'K1', 'J2', 'K2']), ['J1:K2', 'J1:K2', 'J1:K2', 'J1:K2']);
});

test('D4: the lowered command is one public SetFormulaArray over the area', () => {
  const plan = af.plan(stepOf(), 0);
  assert.equal(plan.ok, true, plan.detail);
  assert.match(af.lower(plan.operation).body, /ws\.GetRange\("G1:G3"\)\.SetFormulaArray\("=ROW\(A1:A3\)\*2"\)/);
});

const arrayAt = (cells, ref, formula, list) => { for (const a of list) cells[a] = { v: '', f: formula, array: ref, af: formula }; };
for (const [name, mutate, why] of [
  ['a value in the area (SetFormulaArray would overwrite it silently)', c => { c.G2 = { v: 'x' }; }, 'xlsx-array-formula-area-not-empty'],
  ['a plain formula in the area', c => { c.G2 = { v: 4, f: '=2+2' }; }, 'xlsx-array-formula-area-not-empty'],
  ['an array that overlaps the area (the saved file would carry two overlapping refs)', c => arrayAt(c, 'G2:G4', '=ROW(A1:A3)*3', ['G2', 'G3', 'G4']), 'xlsx-array-formula-touches-an-existing-array'],
  ['an array in the one-cell ring (G4)', c => arrayAt(c, 'G4:G5', '=1', ['G4', 'G5']), 'xlsx-array-formula-touches-an-existing-array'],
  ['the same formula but a LARGER array (G1:G4) -- not the asked extent', c => arrayAt(c, 'G1:G4', '=ROW(A1:A3)*2', ['G1', 'G2', 'G3', 'G4']), 'xlsx-array-formula-touches-an-existing-array'],
]) {
  test(`D4 refused before the write, 0 commands: ${name}`, async () => {
    const b = book(); mutate(b[0].cells);
    const store = createStore(b);
    const { result, m } = await run(store, task(stepOf()));
    assert.equal(result.ok, false);
    assert.equal(m.commands, 0);
    assert.match(JSON.stringify(result.execution), new RegExp(why));
  });
}

test('D4: an unobservable GetFormulaArray fails closed before the write, 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(stepOf()), { arrayUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-array-formula-unobservable/);
});

for (const fault of ['array-formula-wrong-formula', 'array-formula-short', 'array-formula-spill-ring', 'side-effect']) {
  test(`D4: a fake write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(stepOf()), { fault });
    assert.equal(m.commands, 1, 'the write was dispatched -- this is the VERIFY failing, not a refusal');
    assert.equal(result.ok, false);
  });
}

test('D4: write_array_formula must be the only step of its task (its ring read could be changed by another step)', () => {
  const p = planOperations({ id: 't', steps: [stepOf(), { intent: 'write_range', sheet: 'A', range: 'C1', values: [[1]] }] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /only step of its task/);
});

test('D4: liveScope reads the area + ring window and asks GetFormulaArray for every one of its cells', () => {
  const s = liveScope(task(stepOf()));
  assert.deepEqual(s.windows, [{ sheet: 'A', range: 'F1:H4' }]);
  assert.equal(s.arrayCells.length, 12);
  assert.ok(s.arrayCells.some(x => x.cell === 'G4') && s.arrayCells.some(x => x.cell === 'G1'));
});
