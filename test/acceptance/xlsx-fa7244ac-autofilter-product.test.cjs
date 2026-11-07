'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const filter = require('../../src/adapters/xlsx/families/filter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * fa7244ac/2 (c19b1065's LIVE measurement, #ref): set_autofilter_range (Q1: criterion-free AutoFilter --
 * filter buttons, no active criteria) / remove_autofilter (Q3: the whole AutoFilter object gone, stronger than
 * clear_filter's ShowAllData) / filter_range's new PRE-MUTATION refusal when `values` names every distinct value
 * already in the column (Q2: the engine's own no-op).
 *
 * Both new intents dispatch the SAME bare SetAutoFilter() call filter.cjs's lower() already documents -- the
 * fake driver models the engine's own measured toggle (xlsx-fa7244ac-autofilter-object-probe.live.cjs; the
 * research harness, not this product unit's own fixture).
 */

const book = () => [{ name: 'F', cells: { A1: { v: 'id' }, B1: { v: 'tag' }, A2: { v: 1 }, B2: { v: 'x' }, A3: { v: 2 }, B3: { v: 'y' }, A4: { v: 3 }, B4: { v: 'z' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const setRange = (range = 'A1:B4') => ({ id: 'fa7244ac2', steps: [{ intent: 'set_autofilter_range', sheet: 'F', range }] });
const removeAf = (range = 'A1:B4') => ({ id: 'fa7244ac2', steps: [{ intent: 'remove_autofilter', sheet: 'F', range }] });
const flt = (extra = {}) => ({ id: 'fa7244ac2', steps: [{ intent: 'filter_range', sheet: 'F', range: 'A1:B4', field: 2, value: 'x', ...extra }] });
const fltValues = values => ({ id: 'fa7244ac2', steps: [{ intent: 'filter_range', sheet: 'F', range: 'A1:B4', field: 2, values }] });

async function run(store, task, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(task, { observation: { live: liveScope(task) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('set_autofilter_range: APPLY on a sheet with no AutoFilter -> buttons only, no filterColumn; retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, setRange());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0] && a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.deepEqual(store.durable[0].autoFilter, { range: 'A1:B4', filters: [], colours: { columns: [] } });
  const b = await run(store, setRange());
  assert.equal(b.m.commands, 0);
  assert.equal(b.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('set_autofilter_range: refused BEFORE dispatch when an AutoFilter already exists -- same range with active criteria, and a different range', async () => {
  const b1 = book();
  b1[0].autoFilter = { range: 'A1:B4', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] };
  const { result: r1, driver: d1 } = await run(createStore(b1), setRange());
  assert.equal(r1.ok, false);
  assert.equal(d1.calls.dispatch, 0);
  assert.equal(reason(r1), 'xlsx-autofilter-already-present');

  const b2 = book();
  b2[0].autoFilter = { range: 'A1:B9', filters: [] };
  const { result: r2, driver: d2 } = await run(createStore(b2), setRange());
  assert.equal(r2.ok, false);
  assert.equal(d2.calls.dispatch, 0);
  assert.equal(reason(r2), 'xlsx-filter-range-conflict');
});

test('set_autofilter_range: refused BEFORE dispatch on a fully empty range (no header, no data) -- fa7244ac/3 LIVE finding (8f23f67b review #ref)', async () => {
  const b = book();
  b.push({ name: 'E', cells: {} });
  const { result: r, driver: d } = await run(createStore(b), { id: 'fa7244ac3', steps: [{ intent: 'set_autofilter_range', sheet: 'E', range: 'A1:B3' }] });
  assert.equal(r.ok, false);
  assert.equal(d.calls.dispatch, 0);
  assert.equal(reason(r), 'xlsx-autofilter-range-empty');
  // NEG. KONTROLL: the SAME shape of range (A1:B3), but on the data-carrying sheet ('F') -- still dispatches.
  // The guard fires on the cells being blank, not on the range shape alone.
  const { result: r2, driver: d2 } = await run(createStore(b), { id: 'fa7244ac3', steps: [{ intent: 'set_autofilter_range', sheet: 'F', range: 'A1:B3' }] });
  assert.equal(r2.ok, true, JSON.stringify(r2.execution.appliedOperations[0] && r2.execution.appliedOperations[0].evidence));
  assert.equal(d2.calls.dispatch, 1);
});

test('set_autofilter_range: a header-only range (one non-blank cell, no data rows) is NOT MEASURED by the empty-range guard -- still dispatches', async () => {
  const b = [{ name: 'H', cells: { A1: { v: 'only-header' } } }];
  const { result, driver } = await run(createStore(b), { id: 'fa7244ac3', steps: [{ intent: 'set_autofilter_range', sheet: 'H', range: 'A1:B3' }] });
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0] && result.execution.appliedOperations[0].evidence));
  assert.equal(driver.calls.dispatch, 1);
});

test('set_autofilter_range: SATISFIED (no dispatch) when the sheet already carries exactly this range, buttons only', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:B4', filters: [] };
  const { result, driver } = await run(createStore(b), setRange());
  assert.equal(result.ok, true);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('remove_autofilter: APPLY removes the whole object (stronger than clear_filter); retry SATISFIED 0/0', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:B4', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] };
  const store = createStore(b);
  const a = await run(store, removeAf());
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0] && a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.m.commands, 1);
  assert.equal(store.durable[0].autoFilter, undefined, 'the AutoFilter object itself is gone, not just inactive');
  const b2 = await run(store, removeAf());
  assert.equal(b2.m.commands, 0);
  assert.equal(b2.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('remove_autofilter: SATISFIED (no dispatch) when the sheet has no AutoFilter at all', async () => {
  const { result, driver } = await run(createStore(book()), removeAf());
  assert.equal(result.ok, true);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('remove_autofilter: RED -- if the engine only leaves an inactive <autoFilter> behind (clear_filter/ShowAllData shape, range intact, every criterion off), verify does NOT accept it as removed', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:B4', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] };
  const { result, driver } = await run(createStore(b), removeAf(), { fault: 'remove-leaves-inactive' });
  assert.equal(driver.calls.dispatch, 1);
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('remove_autofilter: refused BEFORE dispatch when the existing AutoFilter is on a different range', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:B9', filters: [] };
  const { result, driver } = await run(createStore(b), removeAf());
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-filter-range-conflict');
});

test('filter_range: values naming EVERY distinct value already in the column -- refused PRE-MUTATION, named, discriminating from a subset', async () => {
  const all = await run(createStore(book()), fltValues(['x', 'y', 'z']));
  assert.equal(all.result.ok, false);
  assert.equal(all.driver.calls.dispatch, 0);
  assert.equal(reason(all.result), 'xlsx-filter-values-select-everything-is-no-filter');
  assert.deepEqual(new Set(all.result.execution.appliedOperations[0].evidence.evidence.distinct), new Set(['x', 'y', 'z']));

  // the discriminating pair: a SUBSET of the same column's distinct values (2 of the 3) is NOT the no-op shape -- proceeds to dispatch
  const subset = await run(createStore(book()), fltValues(['x', 'y']));
  assert.equal(subset.result.ok, true, JSON.stringify(subset.result.execution.appliedOperations[0] && subset.result.execution.appliedOperations[0].evidence));
  assert.equal(subset.driver.calls.dispatch, 1);
});

test('the command bodies: the bare SetAutoFilter() call is shared by set_autofilter_range and remove_autofilter', () => {
  const log = [];
  const ws = { GetName: () => 'F', GetRange: a => ({ SetAutoFilter: (...x) => log.push([a, 'SetAutoFilter', ...x]) }) };
  assert.deepEqual(new Function('Api', filter.lower(planOperations(setRange()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(new Function('Api', filter.lower(planOperations(removeAf()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['A1:B4', 'SetAutoFilter'], ['A1:B4', 'SetAutoFilter']]);
});

test('planner: set_autofilter_range / remove_autofilter take sheet+range only, no field/criteria', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse({ id: 'x', steps: [{ intent: 'set_autofilter_range', sheet: 'F', range: 'A1:B4', field: 2 }] }, /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'remove_autofilter', sheet: 'F', range: 'A1:B4', value: 'x' }] }, /does not support field/);
  assert.equal(planOperations(setRange()).ok, true);
  assert.equal(planOperations(removeAf()).ok, true);
});
