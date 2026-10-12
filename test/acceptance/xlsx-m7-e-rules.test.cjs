'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M7 E R2 / R3a (tier E probe m7-e-probe1-20261003.json IR / DR / IC): rules before the cut stay; everything else refuses. */

const RULES = () => [
  { kind: 'validation', range: 'A1:A2', formula: '"1,2,3"' },
  { kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', range: 'B1:B2', formula: '15' },
  { kind: 'conditional', range: 'C1:C2', formula: '=$C1>2' },
];
const book = (rules = RULES()) => {
  const cells = { D1: { v: 'd1' }, E1: { v: 'e1' } };
  for (let i = 1; i <= 9; i += 1) { cells[`A${i}`] = { v: `r${i}` }; cells[`B${i}`] = { v: i * 10 }; cells[`C${i}`] = { v: i }; }
  return [{ name: 'P', cells, rules }, { name: 'Q', cells: { A1: { v: 'keep' } } }];
};
const step = (intent, extra) => ({ id: 'm7-e-r2', steps: [{ intent, sheet: 'P', ...extra }] });
const INS = step('insert_rows', { at: 5, count: 2, anchor: { column: 'A', value: 'r5' } });
const DEL = step('delete_rows', { at: 5, count: 2, anchor: { column: 'A', value: 'r7' } });
const COL = step('insert_columns', { at: 'D', count: 1, anchor: { row: 1, value: 'd1' } });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;
const rulesOf = store => JSON.stringify(store.durable[0].rules);

for (const [name, t] of [['insert rows 5:6', INS], ['delete rows 5:6', DEL], ['insert column D', COL]]) {
  test(`R2 / R3a: ${name} with a validation, a cellValue CF and an expression CF before the cut -> APPLIED, rules unchanged, retry 0/0`, async () => {
    const store = createStore(book());
    const before = rulesOf(store);
    const a = await run(store, t);
    assert.equal(a.result.ok, true, JSON.stringify([op(a.result).evidence, a.result.execution.preservationVerification]));
    assert.equal(rulesOf(store), before);
    const again = await run(store, t);
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test('R2 / R3a RED: the editor moves the rules with the step -> L3 FAIL (rulesKept false), not APPLIED', async () => {
  const { result } = await run(createStore(book()), INS, { fault: 'rule-moves' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].evidence.evidence.rulesKept, false);
});

for (const [name, rule, t] of [
  ['a rule ending at the cut', { kind: 'validation', range: 'A4:A5', formula: '"1,2"' }, INS],
  ['a rule after the cut', { kind: 'conditional', range: 'C8:C9', formula: '=$C8>2' }, INS],
  ['a relative reference that reaches the cut over the rule\'s extent', { kind: 'conditional', range: 'C1:C2', formula: '=$C4>2' }, INS],
  ['an absolute reference after the cut', { kind: 'conditional', range: 'C1:C2', formula: '=$C$9>2' }, INS],
  ['a sheet-qualified reference', { kind: 'conditional', range: 'C1:C2', formula: '=Q!$A$1>2' }, INS],
  ['a defined name in the formula', { kind: 'conditional', range: 'C1:C2', formula: '=Limit>2' }, INS],
  ['delete columns (not measured)', { kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', range: 'A1:A2', formula: '15' }, step('delete_columns', { at: 'D', count: 1, anchor: { row: 1, value: 'e1' } })],
]) {
  test(`R2 / R3a refused before the write: ${name}`, async () => {
    const store = createStore(book([rule]));
    if (name.includes('defined name')) store.names = [{ name: 'Limit', ref: 'Q!$A$1' }];
    const r = await run(store, t);
    assert.equal(r.driver.calls.dispatch, 0, JSON.stringify(op(r.result).evidence));
    assert.equal(reason(r.result), 'xlsx-structure-dependent-not-measured');
  });
}

test('R2 / R3a refused: a rule on ANOTHER sheet referencing the shifted sheet', async () => {
  const b = book([]);
  b[1].rules = [{ kind: 'conditional', range: 'A1:A2', formula: '=P!$B$1>2' }];
  const r = await run(createStore(b), INS);
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(reason(r.result), 'xlsx-structure-dependent-not-measured');
});
