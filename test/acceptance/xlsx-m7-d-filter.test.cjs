'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const filter = require('../../src/adapters/xlsx/families/filter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M7 R19 (tier D probe, m7-d-probe1-20261003.json FL): a two-value list, an xlAnd pair, several fields on one range. */

const book = () => [{ name: 'FL', cells: { A1: { v: 'id' }, B1: { v: 'tag' }, C1: { v: 'n' }, A2: { v: 1 }, B2: { v: 'x' }, C2: { v: 1 }, A3: { v: 2 }, B3: { v: 'y' }, C3: { v: 2 }, A4: { v: 3 }, B4: { v: 'z' }, C4: { v: 3 } } }];
const step = extra => ({ id: 'm7-d-r19', steps: [{ intent: 'filter_range', sheet: 'FL', range: 'A1:C4', ...extra }] });
const LIST = { field: 2, values: ['x', 'z'] };
const AND = { field: 3, and: ['>1', '<6'] };

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const status = result => result.execution.appliedOperations[0].status;
const filters = store => store.durable[0].autoFilter.filters;

test('R19: a two-value list, then an xlAnd pair on ANOTHER field of the same range -> both APPLIED, both kept, retries 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, step(LIST));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  const b = await run(store, step(AND));
  assert.equal(b.result.ok, true, JSON.stringify([b.result.execution.primitiveVerification, b.result.execution.appliedOperations[0].evidence]));
  assert.equal(b.m.commands, 1);
  assert.deepEqual(filters(store).map(f => [f.field, f.operator, f.criteria1, f.criteria2]), [[2, 'xlOr', '=x', '=z'], [3, 'xlAnd', '>1', '<6']]);
  for (const t of [step(LIST), step(AND)]) {
    const again = await run(store, t);
    assert.equal(again.m.commands, 0);
    assert.equal(status(again.result), OPERATION_STATUS.SATISFIED);
  }
});

test('R19: the command is the measured SetAutoFilter shape for each criteria form', () => {
  const body = extra => filter.lower(planOperations(step(extra)).operations[0]).body;
  assert.match(body(LIST), /SetAutoFilter\(2, \["x","z"\], 'xlFilterValues'\)/);
  assert.match(body(AND), /SetAutoFilter\(3, ">1", 'xlAnd', "<6"\)/);
  assert.match(body({ field: 2, value: 'x' }), /SetAutoFilter\(2, "x", 'xlOr'\)/);
});

test('R19: a one-value filter is NOT satisfied by a two-value list holding the value (criteria2 must be empty)', async () => {
  const store = createStore(book());
  await run(store, step(LIST));
  const one = await run(store, step({ field: 2, value: 'x' }));
  assert.equal(one.result.ok, true, JSON.stringify(one.result.execution.primitiveVerification));
  assert.equal(one.m.commands, 1, 'the list is replaced by the one value, not taken as already there');
});

test('R19 RED: the editor drops the OTHER field -> L1 FAIL, not APPLIED', async () => {
  const store = createStore(book());
  await run(store, step(LIST));
  const { result } = await run(store, step(AND), { fault: 'filter-drops-other' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R19 RED: the second list value reads back wrong -> L1 FAIL', async () => {
  const { result } = await run(createStore(book()), step(LIST), { fault: 'filter-wrong-value' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R19 RED: the xlAnd pair reads back joined by xlOr -> L1 FAIL', async () => {
  const { result } = await run(createStore(book()), step(AND), { fault: 'filter-wrong-operator' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R19 refused at planning: eleven values (GH #301 D2: 3..10 is measured now); one value twice; an unmeasured operator; two criteria forms at once', () => {
  for (const extra of [{ field: 2, values: ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k'] }, { field: 2, values: ['x', 'x'] }, { field: 3, and: ['>=1', '<6'] }, { field: 3, and: ['>1'] }, { field: 2, value: 'x', values: ['x', 'z'] }, { field: 2 }]) {
    assert.equal(planOperations(step(extra)).ok, false, JSON.stringify(extra));
  }
});

/**
 * X-L05 (S26 probe m7-s26-probe1, LIVE MEASURED): applying filter_range with NO
 * active filter yet unhides every data row the engine judges matching -- INCLUDING a manually
 * hidden one. Refused before the write whenever the range has no active filter AND a data row
 * already reads hidden (or that read is unreadable); unaffected by a hidden row OUTSIDE the
 * range, or by a range that ALREADY carries an active filter (today's R19 path, unchanged).
 */
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('X-L05 refused BEFORE the write: a hidden DATA row inside the range, no active filter yet', async () => {
  const b = book();
  b[0].layout = { rows: { 3: { hidden: true } } };
  const { result, m } = await run(createStore(b), step({ field: 2, value: 'x' }));
  assert.equal(result.ok, false);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.equal(reason(result), 'xlsx-filter-would-unhide-manual-rows');
});

test('X-L05: a hidden row OUTSIDE the range does NOT block -- filter_range still APPLIED, retry 0/0', async () => {
  const b = book();
  b[0].cells.A6 = { v: 'outside' };
  b[0].layout = { rows: { 6: { hidden: true } } };
  const store = createStore(b);
  const t = step({ field: 2, value: 'x' });
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
});

test('X-L05: a hidden data row is UNAFFECTED when the range ALREADY carries an active filter -- today\'s R19 multi-field behaviour, unchanged', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:C4', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] };
  b[0].layout = { rows: { 3: { hidden: true } } };
  const { result, m } = await run(createStore(b), step(AND));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
});

// a maintainer review: a CLEARED AutoFilter keeps its range (ShowAllData keeps the controls, B-D13) with no filter on -- the
// hidden row is just as manual there, so the gate keys on "no ACTIVE filter", not on "no AutoFilter range".
test('X-L05 refused BEFORE the write: a hidden data row under a CLEARED AutoFilter (range kept, no filter on)', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:C4', filters: [] };
  b[0].layout = { rows: { 3: { hidden: true } } };
  const { result, m } = await run(createStore(b), step({ field: 2, value: 'x' }));
  assert.equal(result.ok, false);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.equal(reason(result), 'xlsx-filter-would-unhide-manual-rows');
});

test('X-L05 refused BEFORE the write: an unreadable row-hidden state refuses the SAME way, fail-closed', async () => {
  const { result, m } = await run(createStore(book()), step({ field: 2, value: 'x' }), { layoutUnobservable: true });
  assert.equal(result.ok, false);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.equal(reason(result), 'xlsx-filter-rows-unobservable');
});

test('X-L05: clear_filter is NOT touched by this gate -- a hidden data row never refuses a clear', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:C4', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] };
  b[0].layout = { rows: { 3: { hidden: true } } };
  const clearStep = { id: 'm7-d-r19-clear', steps: [{ intent: 'clear_filter', sheet: 'FL', range: 'A1:C4' }] };
  const { result } = await run(createStore(b), clearStep);
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
});
