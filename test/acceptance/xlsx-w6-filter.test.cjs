'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const filter = require('../../src/adapters/xlsx/families/filter.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 6: filter_range / clear_filter on the fake (a proven build); LIVE the field reader is proven on 9.3.4-9475e9f1 (W6 / W6b). */

const book = () => [{ name: 'F', cells: { A1: { v: 'id' }, B1: { v: 'tag' }, A2: { v: 1 }, B2: { v: 'x' }, A3: { v: 2 }, B3: { v: 'y' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const flt = (extra = {}) => ({ id: 'm6-w6f', steps: [{ intent: 'filter_range', sheet: 'F', range: 'A1:B3', field: 2, value: 'x', ...extra }] });
const clr = (range = 'A1:B3') => ({ id: 'm6-w6f', steps: [{ intent: 'clear_filter', sheet: 'F', range }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('filter: APPLIED, retry SATISFIED 0/0; clear: APPLIED (controls kept), retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, flt());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.equal((await run(store, flt())).m.commands, 0);
  const c = await run(store, clr());
  assert.equal(c.result.ok, true, JSON.stringify(c.result.execution.primitiveVerification));
  assert.equal(store.durable[0].autoFilter.range, 'A1:B3', 'ShowAllData keeps the AutoFilter');
  const c2 = await run(store, clr());
  assert.equal(c2.m.commands, 0);
  assert.equal(c2.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const [name, t, why, af, script] of [
  ['an AutoFilter on another range', flt(), 'xlsx-filter-range-conflict', { range: 'A1:B9', filters: [] }],
  // M7 R19: an active filter on another field no longer refuses (test/acceptance/xlsx-m7-d-filter.test.cjs)
  ['an unproven build (not the W6b-proven one)', flt(), 'xlsx-filter-unobservable', null, { filterBuild: false }],
  ['a clear with an unproven reader', clr(), 'xlsx-filter-unobservable', null, { filterBuild: false }],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const b = book();
    if (af) b[0].autoFilter = af;
    const { result, driver } = await run(createStore(b), t, script || {});
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / the wrong field / the wrong value / a second active filter -> L1 FAIL; a cell elsewhere -> L3 FAIL', async () => {
  for (const fault of ['ignore', 'filter-wrong-field', 'filter-wrong-value', 'filter-extra-field']) {
    const { result, driver } = await run(createStore(book()), flt(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  const side = await run(createStore(book()), flt(), { fault: 'side-effect' });
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('RED: a clear that was ignored -> L1 FAIL (the filter is still on)', async () => {
  const b = book();
  b[0].autoFilter = { range: 'A1:B3', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] };
  const { result, driver } = await run(createStore(b), clr(), { fault: 'ignore' });
  assert.equal(driver.calls.dispatch, 1);
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

// B07 cutover (internal tracker point 8; LIVE parity #ref): the field identity comes through the DS Bridge
// sheet.autoFilter.fieldIdentity.read/v1 -- the REAL served block runs on the stand-in (test/fakes/ds-bridge-runtime.cjs).
test('the filter field identity comes only through the DS Bridge; a gap is unobservable, never "no filter"', () => {
  const f = { filter: { ColId: 1 }, GetOn: () => true, GetOperator: () => 'xlOr', GetCriteria1: () => '=x', GetCriteria2: () => null };
  const ws = { GetName: () => 'F', GetIndex: () => 0, GetUsedRange: () => null, GetAutoFilter: () => ({ GetFilters: () => [f], GetRange: () => ({ GetAddress: () => 'A1:B3' }) }) };
  const apiOf = (sheet, bridge = true) => { const a = { GetSheets: () => [sheet], GetSheet: n => (n === 'F' ? sheet : null), WorksheetFunction: { COUNT: () => 0 } }; if (bridge) installBridge(a); return a; };
  const read = api => new Function('Api', readQuery([], { filterSheets: ['F'] }))(api).filterReads.F;
  assert.equal(read(apiOf(ws, false)), 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', 'no bridge -> never a direct ColId read');
  assert.equal(read(withPatch(apiOf(ws))), 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', 'an unaccepted served runtime');
  assert.deepEqual(read(apiOf(ws)), { range: 'A1:B3', filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] });
  const noCol = { ...ws, GetAutoFilter: () => ({ GetFilters: () => [{ ...f, filter: {} }], GetRange: () => ({ GetAddress: () => 'A1:B3' }) }) };
  assert.equal(read(apiOf(noCol)), 'UNOBSERVABLE:BRIDGE:field-unreadable');
  // the measured never-filtered shape: an object without a range and without filters -> no AutoFilter (the bridge is not asked)
  const bare = { ...ws, GetAutoFilter: () => ({ GetFilters: () => [], GetRange: () => null }) };
  assert.deepEqual(read(apiOf(bare, false)), { range: null, filters: [] });
  const odd = { ...ws, GetAutoFilter: () => ({ GetFilters: () => [f], GetRange: () => null }) };
  assert.equal(read(apiOf(odd)), 'UNOBSERVABLE:autofilter-filters-without-range');
  const noAf = { ...ws, GetAutoFilter: () => null };
  assert.deepEqual(read(apiOf(noAf, false)), { range: null, filters: [] });
  assert.doesNotMatch(readQuery([], { filterSheets: ['F'] }), /ColId/, 'the direct field read is gone');
});

test('planner: field inside the range, one plain value', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(flt({ field: 3 }), /field must be a column number inside the range/);
  refuse(flt({ value: '>5' }), /one plain value/);
  refuse(flt({ operator: 'xlAnd' }), /does not support field/);
  refuse(flt({ range: 'A1:B1' }), /header row and data/);
});

test('the command bodies: SetAutoFilter(field, value, xlOr) / GetAutoFilter().ShowAllData()', () => {
  const log = [];
  const ws = { GetName: () => 'F', GetRange: a => ({ SetAutoFilter: (...x) => log.push([a, 'SetAutoFilter', ...x]) }), GetAutoFilter: () => ({ ShowAllData: () => log.push(['ShowAllData']) }) };
  assert.deepEqual(new Function('Api', filter.lower(planOperations(flt()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(new Function('Api', filter.lower(planOperations(clr()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['A1:B3', 'SetAutoFilter', 2, 'x', 'xlOr'], ['ShowAllData']]);
});
