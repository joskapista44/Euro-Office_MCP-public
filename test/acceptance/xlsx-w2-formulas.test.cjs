'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const cells = require('../../src/adapters/xlsx/families/cells.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 2: write_range with FORMULAS through the target path. The fake re-renders formulas the way
 * the LIVE probe measured (sets a/b: '= ' prefix, upper-case and space-free outside literals; a
 * case-differing sheet name silently drops the formula), so identity must be canonical, and the exact
 * sheet-name gate must refuse before the editor can lose anything.
 */

const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, B1: { v: 3 }, C2: { v: 6, f: '= B1*2' } } },
  { name: 'Monthly_Plan', cells: { B16: { v: 12020400 } } },
];

const step = (range, formulas, values, sheet = 'Data') => ({ intent: 'write_range', sheet, range, formulas, ...(values ? { values } : {}) });
const task = (...steps) => ({ id: 'm6-w2', steps });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const formulaTask = () => task(step('D1:E1', [['=b1*3', '=Monthly_Plan!B16 + 1']]));
const values = { formulaValues: { D1: 9, E1: 12020401 } };

test('FIRST: formulas written, identity proven on the editor\'s re-rendered text, L3, barrier', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, formulaTask(), values);
  assert.equal(result.ok, true);
  assert.equal(m.writes, 1);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  const d = store.durable.find(s => s.name === 'Data').cells;
  assert.equal(d.D1.f, '= B1*3', 'stored as the editor renders it, not as submitted');
  assert.equal(d.E1.f, '= Monthly_Plan!B16+1', 'the sheet name keeps its live case, spaces gone (measured)');
});

test('SECOND: the same formulas are already there -> SATISFIED, 0 write', async () => {
  const store = createStore(book());
  await run(store, formulaTask(), values);
  const { result, m } = await run(store, formulaTask(), values);
  assert.equal(result.ok, true);
  assert.equal(m.writes, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('values and formulas mixed in one range; values-only cells keep the W1 contract', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(step('D1:E1', [[null, '=D1*2']], [['Total', null]])), { formulaValues: { E1: 0 } });
  assert.equal(result.ok, true);
  const d = store.durable.find(s => s.name === 'Data').cells;
  assert.equal(d.D1.v, 'Total');
  assert.equal(d.E1.f, '= D1*2');
});

test('an EVALUATION error is a written formula AND a failed task, named per cell', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(step('D1', [['=1/0']])), { formulaValues: { D1: '#DIV/0!' } });
  assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS', 'the write itself is proven');
  assert.equal(result.ok, false);
  const failed = result.execution.wholeTaskVerification.evidence.checks.filter(c => !c.ok);
  assert.equal(failed.length, 1);
  assert.match(failed[0].name, /evaluates to an error value/);
  assert.deepEqual(failed[0].detail.evaluationErrors, [{ a: 'D1', value: '#DIV/0!' }]);
});

test('refused BEFORE dispatch: a sheet name that differs only in case (the editor would silently drop it)', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(step('D1', [['=monthly_plan!B16']])));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  const ev = result.execution.appliedOperations[0].evidence.evidence;
  assert.equal(ev.reason, 'xlsx-formula-sheet-reference-not-exact');
  assert.deepEqual(ev.references, ['monthly_plan']);
  assert.deepEqual(store.durable, book());
});

test('NEGATIVE CONTROL: had the gate been bypassed, the lost formula fails L1 (never APPLIED)', async () => {
  // the gate passes (the exact name exists); the transport then loses the write -- L1 must catch it
  const store = createStore(book());
  const { result } = await run(store, task(step('D1', [['=Monthly_Plan!B16']])), { fault: 'ignore' });
  assert.equal(result.ok, false);
  assert.notEqual(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('refused BEFORE dispatch: a DIFFERENT formula over an existing one', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(step('C2', [['=B1*5']])));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-write-over-formula-refused');
  const same = await run(createStore(book()), task(step('C2', [['=b1 * 2']])));
  assert.equal(same.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, 'the same formula, re-rendered, is already there');
});

test('GH79a: replaceFormulas:true -- a DIFFERENT formula over an existing one is accepted, written, verified', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ ...step('C2', [['=B1*5']]), replaceFormulas: true }), { formulaValues: { C2: 15 } });
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(driver.calls.dispatch, 1);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(store.durable.find(s => s.name === 'Data').cells.C2.f, '= B1*5');
});

test('GH79a: replaceFormulas:true does NOT widen `values` -- a plain value over an existing formula stays refused', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, { id: 'm6-w2', steps: [{ intent: 'write_range', sheet: 'Data', range: 'C2', values: [[5]], replaceFormulas: true }] });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-write-over-formula-refused');
  assert.deepEqual(store.durable, book());
});

test('planner: formula shape, value+formula conflict, literals, external refs, formulas-only', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false); assert.match(p.detail, re); };
  refuse(task(step('D1', [['B1*3']])), /must be a string starting with =/);
  refuse(task(step('D1', [['="abc']])), /unterminated/);
  refuse(task(step('D1', [['=[Book2.xlsx]Data!A1']])), /external workbook/);
  refuse(task(step('D1', [['=1']], [['x']])), /both a value and a formula/);
  refuse(task(step('D1:E1', [['=1']])), /1x2 matrix/);
  refuse(task(step('D1', [['=' + 'A'.repeat(8200)]])), /longer than 8192/);
  assert.equal(planOperations(task(step('D1:E1', [['=B1', '="a  B"']]))).ok, true, 'formulas-only, a literal with spaces is fine');
  assert.equal(cells.formulaProblem('=IF(B1>2,"Yes  X","no")'), null);
});

test('the command writes each formula through SetValue, as the probe measured', () => {
  const op = planOperations(task(step('D1:E1', [['=B1*3', null]], [[null, 'x']]))).operations[0];
  const sent = [];
  const ws = { GetName: () => 'Data', GetRange: a => ({ SetValue: v => sent.push([a, v]) }) };
  new Function('Api', cells.lower(op).body)({ GetSheet: () => ws });
  assert.deepEqual(sent, [['D1', '=B1*3'], ['E1', 'x']]);
});

test('the LIVE harness GATE case: a lower-case sheet name on a cell that already holds the right formula is still refused by the sheet gate', async () => {
  const store = createStore(book());
  await run(store, task(step('D1', [['=Monthly_Plan!B16']])), { formulaValues: { D1: 12020400 } });
  const { result, driver } = await run(store, task(step('D1', [['=monthly_plan!B16']])));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-formula-sheet-reference-not-exact');
});
