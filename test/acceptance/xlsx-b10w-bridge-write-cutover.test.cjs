'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { runViaBridge } = require('../fakes/print-bridge-runtime.cjs');

/**
 * M11.5b.4 (owner: M11.5 continuous mandate, internal tracker): B10-WRITE product cutover to the Bridge WRITE
 * capability sheet.printSettings.write/v1. Part A runs the PRODUCT's command + the driver's injected client against the
 * REAL /7 block (and the REAL /6 block as the wrong-runtime control). Part B runs the canonical executor over the fake
 * driver, which returns what the real driver returns for each body outcome (xlsx-driver b10DispatchResult, pinned in A).
 */

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const V6 = fs.readFileSync(path.join(DIR, 'rev6-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const RT7 = '9.3.4|euro-office-bridge-v1/9.3.4.60/7';
const one = step => ({ id: 'b10c', steps: [{ sheet: 'P', ...step }] });
const opOf = step => planOperations(one(step)).operations[0];
const SCALE = { intent: 'set_print_scale', scale: 75 };

// ---- Part A ------------------------------------------------------------------------------------------------------------

function model({ throws = [] } = {}) {
  const log = [];
  const fn = k => x => { log.push([k, x]); if (throws.includes(k)) throw new Error('boom'); };
  const ps = { asc_setFitToWidth: fn('asc_setFitToWidth'), asc_setFitToHeight: fn('asc_setFitToHeight'), asc_setScale: fn('asc_setScale') };
  return { log, ws: { GetName: () => 'P', worksheet: { PagePrintOptions: { asc_getPageSetup: () => ps }, setFitToPage: fn('setFitToPage') } } };
}

test('A every intent: ONE Bridge WRITE command, no direct internal in the product command, marked for the driver', () => {
  for (const step of [SCALE, { intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 1 }, { intent: 'set_paper_size', paper: 'A4' }, { intent: 'set_first_page_number', value: 2 },
    { intent: 'set_print_centering', horizontal: true }, { intent: 'set_header_footer', slot: 'oddHeader', value: 'x' }, { intent: 'set_header_footer_flag', flag: 'differentFirst', value: true },
    { intent: 'set_page_break', axis: 'row', at: 3 }]) {
    const l = print.lower(opOf(step));
    assert.equal(l.meta.bridgePrintWrite, true, step.intent);
    assert.equal((l.body.match(/__eoB10Write\(/g) || []).length, 1, step.intent);
    assert.match(l.body, /const g10 = __eoB10Gate\(name\); if \(!g10\.ok\) return g10;[\s\S]*__eoB10Write/, `${step.intent}: the gate runs before the write`);
    assert.doesNotMatch(l.body, /PagePrintOptions|asc_set|setFitToPage|changeRowColBreaks|\.headerFooter|AscCommon|\.worksheet/, step.intent);
  }
});

test('A scale through the REAL /7 block: the UI shape, in order; the receipt equals the requested fields', () => {
  const m = model();
  assert.deepEqual(runViaBridge(opOf(SCALE), m.ws, {}), { ok: true });
  assert.deepEqual(m.log, [['asc_setFitToWidth', 0], ['asc_setFitToHeight', 0], ['setFitToPage', false], ['asc_setScale', 75]]);
});

test('A partial write: the command stops, partial + applied + the observed settings are reported; no second write, no rollback', () => {
  const m = model({ throws: ['setFitToPage'] });
  const r = runViaBridge(opOf(SCALE), m.ws, {});
  assert.deepEqual([r.ok, r.stage, r.reason, r.partial, r.applied], [false, 'bridge-write', 'BRIDGE:setFitToPage-threw', true, ['fitW', 'fitH']]);
  assert.ok('observed' in r, 'diagnostics carry the fresh read (or its UNOBSERVABLE reason)');
  assert.deepEqual(m.log.map(c => c[0]), ['asc_setFitToWidth', 'asc_setFitToHeight', 'setFitToPage'], 'nothing after the throw');
});

test('A wrong runtime in the editor (the REAL /6 block): the gate refuses BEFORE any mutation -- no fallback', () => {
  const m = model(); const api = { GetSheet: n => (n === 'P' ? m.ws : null) };
  new Function('Api', `${V6};`)(api);
  const r = new Function('Api', drv.b10BridgeWritePrelude() + print.lower(opOf(SCALE)).body)(api);
  assert.deepEqual(r, { ok: false, stage: 'bridge-write-gate', reason: 'PATCH_VERSION_MISMATCH', mutation: 'NONE' });
  assert.deepEqual(m.log, []);
});

test('A an intent without a Bridge mapping is NEVER lowered to another setting: no gate, no write, not ok', () => {
  const op = { ...opOf(SCALE), intent: 'set_print_future' };
  assert.equal(print.bridgeArgs(op), null);
  const l = print.lower(op);
  assert.doesNotMatch(l.body, /__eoB10/);
  const m = model(); assert.deepEqual(new Function('Api', drv.b10BridgeWritePrelude() + l.body)({ GetSheet: () => m.ws }), { ok: false, stage: 'no-bridge-mapping', intent: 'set_print_future' });
  assert.deepEqual(m.log, []);
});

test('A no injected client: ReferenceError -> caught -> not ok, nothing written', () => {
  const m = model(); const r = new Function('Api', print.lower(opOf(SCALE)).body)({ GetSheet: () => m.ws });
  assert.equal(r.ok, false); assert.equal(r.stage, 'threw'); assert.deepEqual(m.log, []);
});

test('A driver b10DispatchResult: ok ONLY when the command completed AND its body said ok; partial -> partiallyApplied', () => {
  assert.deepEqual(drv.b10DispatchResult({ ok: true, outcome: 'command-completed', result: { ok: true } }), { ok: true, dispatched: true, outcome: 'command-completed', error: null });
  const part = { ok: false, stage: 'bridge-write', partial: true, applied: ['fitW'] };
  assert.deepEqual(drv.b10DispatchResult({ ok: true, result: part }), { ok: true, dispatched: true, partiallyApplied: true, outcome: 'b10-bridge-write-partial', evidence: part, error: null });
  assert.equal(drv.b10DispatchResult({ ok: true, result: { ok: false, stage: 'bridge-write', reason: 'BRIDGE:x', partial: false } }).ok, false);
  assert.equal(drv.b10DispatchResult({ ok: true, result: { ok: false, stage: 'bridge-write-gate' } }).outcome, 'b10-bridge-write-failed');
  assert.equal(drv.b10DispatchResult({ ok: false, outcome: 'timeout' }).outcome, 'timeout');
});

test('A driver second line: a print command is refused WITHOUT an editor on any other runtime, and an UNMARKED print command is refused even on /7', async () => {
  const d = drv.createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const l = print.lower(opOf(SCALE));
  let r = await d.dispatch(null, l);
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b10-bridge-write-runtime-not-proven']);
  r = await d.dispatch(null, { ...l, meta: { ...l.meta, bridgePrintWrite: false } });
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b10-bridge-write-runtime-not-proven'], 'no direct-internal print path exists');
});

// ---- Part B ------------------------------------------------------------------------------------------------------------

const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT7, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const opr = r => r.execution.appliedOperations[0];
const printOf = store => store.durable.find(s => s.name === 'P').print || {};

test('B /7: APPLIED (LIVE_VERIFY on the fresh READ) + FORCESAVE_CONFIRMED with ONE Bridge WRITE; retry SATISFIED 0 command / 0 write', async () => {
  const store = createStore(book());
  const first = await run(store, one(SCALE));
  assert.equal(opr(first.result).status, OPERATION_STATUS.APPLIED);
  assert.equal(first.result.persistenceState, 'FORCESAVE_CONFIRMED');
  assert.equal(first.driver.calls.bridgeWrites, 1);
  assert.deepEqual([printOf(store).fitW, printOf(store).fitFlag, printOf(store).scale], [0, false, 75]);
  const again = await run(store, one(SCALE));
  assert.equal(opr(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.deepEqual([again.m.commands, again.driver.calls.dispatch, again.driver.calls.bridgeWrites], [0, 0, 0]);
});

test('B gate refused in the editor: DISPATCHED_OUTCOME_UNKNOWN, nothing changed, no fallback', async () => {
  const store = createStore(book()); const before = JSON.stringify(store.durable);
  const { result } = await run(store, one(SCALE), { b10Write: 'gate-refused' });
  assert.equal(opr(result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(opr(result).uncertainty.dispatchEvidence.evidence.stage, 'bridge-write-gate');
  assert.equal(JSON.stringify(store.durable), before);
});

test('B partial write: DISPATCHED_OUTCOME_UNKNOWN (ADAPTER_REPORTED_PARTIALLY_APPLIED), no second write; a new invocation APPLIES', async () => {
  const store = createStore(book());
  const first = await run(store, one(SCALE), { b10Write: 'partial' });
  assert.equal(opr(first.result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(opr(first.result).uncertainty.reason, UNCERTAINTY_REASON.ADAPTER_REPORTED_PARTIALLY_APPLIED);
  assert.deepEqual(opr(first.result).uncertainty.dispatchEvidence.evidence.applied, ['fitW']);
  assert.equal(first.driver.calls.bridgeWrites, 1);
  const again = await run(store, one(SCALE));
  assert.equal(opr(again.result).status, OPERATION_STATUS.APPLIED);
});

for (const mode of ['error-applied', 'malformed']) {
  test(`B write error with the TARGET OBSERVED (${mode}): DISPATCHED_OUTCOME_UNKNOWN, NOT APPLIED, no forcesave claimed`, async () => {
    const first = await run(createStore(book()), one(SCALE), { b10Write: mode });
    assert.equal(opr(first.result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
    assert.equal(first.result.ok, false);
    assert.notEqual(first.result.persistenceState, 'FORCESAVE_CONFIRMED');
    assert.equal(first.driver.calls.forceSave, 0);
  });
}

test('B retry after UNKNOWN: a NEW invocation decides from its own fresh prewrite READ (held -> SATISFIED 0 write; not -> APPLY)', async () => {
  const held = createStore(book());
  await run(held, one(SCALE));
  const sat = await run(held, one(SCALE));
  assert.equal(opr(sat.result).status, OPERATION_STATUS.SATISFIED);
  assert.deepEqual([sat.driver.calls.dispatch, sat.driver.calls.bridgeWrites], [0, 0]);
  const store = createStore(book());
  await run(store, one(SCALE), { b10Write: 'error-applied' });
  const again = await run(store, one(SCALE));
  assert.equal(opr(again.result).status, OPERATION_STATUS.APPLIED);
});

test('B setter missing (nothing written): DISPATCHED_OUTCOME_UNKNOWN, store unchanged; retry APPLIED', async () => {
  const store = createStore(book()); const before = JSON.stringify(store.durable);
  const first = await run(store, one(SCALE), { b10Write: 'error-none' });
  assert.equal(opr(first.result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(JSON.stringify(store.durable), before);
  const again = await run(store, one(SCALE));
  assert.equal(opr(again.result).status, OPERATION_STATUS.APPLIED);
});

test('B already holding on a non-/7 runtime: SATISFIED with 0 write (the runtime gate decides only a write)', async () => {
  const b = book(); b[0].print = { fitW: 0, fitH: 0, fitFlag: false, scale: 75 };
  const { result, driver } = await run(createStore(b), one(SCALE), { runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/6' });
  assert.equal(opr(result).status, OPERATION_STATUS.SATISFIED);
  assert.deepEqual([driver.calls.dispatch, driver.calls.bridgeWrites], [0, 0]);
});
