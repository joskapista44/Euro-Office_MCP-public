'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const names = require('../../src/adapters/xlsx/families/names.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M7 R16 (tier D probe, m7-d-probe1-20261003.json NM): sheet-scope defined names; the same name in two scopes refuses. */

const book = () => [{ name: 'NM', cells: { A1: { v: 7 }, A2: { v: 9 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const step = (intent, extra = {}) => ({ id: 'm7-d-r16', steps: [{ intent, name: 'SN', ...extra }] });
const SET = step('set_defined_name', { refersTo: '=NM!$A$1', scope: 'NM' });
const DEL = step('delete_defined_name', { scope: 'NM' });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;
const sheetNames = store => ((store.durable[0].objects || {}).names || []).map(n => n.name);

test('R16: set a sheet-scope name -> on the sheet, not at workbook level; retry SATISFIED 0/0; delete it; retry 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, SET);
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, op(a.result).evidence]));
  assert.deepEqual(sheetNames(store), ['SN']);
  assert.deepEqual(store.names, []);
  const again = await run(store, SET);
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  const d = await run(store, DEL);
  assert.equal(d.result.ok, true, JSON.stringify(d.result.execution.primitiveVerification));
  assert.deepEqual(sheetNames(store), []);
  const d2 = await run(store, DEL);
  assert.equal(d2.m.commands, 0);
  assert.equal(op(d2.result).status, OPERATION_STATUS.SATISFIED);
});

test('R16: the command writes and deletes through the scope sheet (ws.AddDefName bare / ws.GetDefName().Delete)', () => {
  const body = t => names.lower(planOperations(t).operations[0]).body;
  assert.match(body(SET), /const ws = Api\.GetSheet\("NM"\);[\s\S]*ws\.AddDefName\("SN", "NM!\$A\$1"\)/);
  assert.match(body(DEL), /const d = ws\.GetDefName\("SN"\);[\s\S]*d\.Delete\(\)/);
  assert.match(body(step('set_defined_name', { refersTo: '=NM!$A$1' })), /Api\.AddDefName\("SN", "NM!\$A\$1"\)/);
});

test('R16 refused before the write: the same name in the OTHER scope (either direction); an unknown scope sheet', async () => {
  const wb = createStore(book());
  wb.names = [{ name: 'sn', ref: 'NM!$A$2' }];
  const a = await run(wb, SET);
  assert.equal(a.driver.calls.dispatch, 0);
  assert.equal(reason(a.result), 'xlsx-name-exists-in-another-scope');
  const sc = book();
  sc[0].objects = { charts: [], pivots: [], names: [{ name: 'SN', refersTo: '=NM!$A$1' }] };
  for (const t of [step('set_defined_name', { refersTo: '=NM!$A$2' }), step('delete_defined_name')]) {
    const r = await run(createStore(sc), t);
    assert.equal(r.driver.calls.dispatch, 0);
    assert.equal(reason(r.result), 'xlsx-name-exists-in-another-scope');
  }
  const none = await run(createStore(book()), step('set_defined_name', { refersTo: '=NM!$A$1', scope: 'Nope' }));
  assert.equal(none.driver.calls.dispatch, 0);
  assert.equal(reason(none.result), 'xlsx-sheet-not-found');
});

test('R16: a sheet-scope name used on its sheet is not deleted (in use); used from another sheet the use is unresolved -> refused', async () => {
  const scoped = () => { const b = book(); b[0].objects = { charts: [], pivots: [], names: [{ name: 'SN', refersTo: '=NM!$A$1' }] }; return b; };
  const own = scoped();
  own[0].cells.B1 = { v: 14, f: '=SN*2' };
  const r = await run(createStore(own), DEL);
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(reason(r.result), 'xlsx-name-in-use');
  const away = scoped();
  away[1].cells.B1 = { v: 14, f: '=SN*2' };
  const o = await run(createStore(away), DEL);
  assert.equal(o.driver.calls.dispatch, 0);
  assert.equal(reason(o.result), 'xlsx-names-dependency-observation-incomplete');
});

test('R16 RED: the editor puts the name at workbook level -> L1 FAIL, not APPLIED', async () => {
  const { result } = await run(createStore(book()), SET, { fault: 'name-wrong-scope' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R16 refused at planning: an empty or non-string scope', () => {
  for (const scope of ['', 3, null]) assert.equal(planOperations(step('set_defined_name', { refersTo: '=NM!$A$1', scope })).ok, false, JSON.stringify(scope));
});
