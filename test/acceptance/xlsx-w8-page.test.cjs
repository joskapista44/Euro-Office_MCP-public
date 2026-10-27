'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const page = require('../../src/adapters/xlsx/families/page.cjs');
const { readQuery, EO_PAGE_LAYOUT_PER_SHEET_RUNTIMES } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 8: set_page_layout on the fake (a proven build); LIVE the per-sheet read is proven on 9.3.4-9475e9f1 (W8b L4). */

const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'O', cells: { A1: { v: 'keep' } }, page: { orientation: 'xlPortrait', printGridlines: true, printHeadings: false } }];
const lay = (extra = { orientation: 'landscape', printGridlines: true }) => ({ id: 'm6-w8p', steps: [{ intent: 'set_page_layout', sheet: 'P', ...extra }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('set_page_layout: APPLIED (the requested fields only), retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, lay());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.preservationVerification]));
  assert.equal(a.m.commands, 1);
  assert.deepEqual(store.durable[0].page, { orientation: 'xlLandscape', printGridlines: true, printHeadings: false, topMargin: 19.1, bottomMargin: 19.1, leftMargin: 17.8, rightMargin: 17.8 });
  const r = await run(store, lay());
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('margins: APPLIED in millimetres, the retry SATISFIED 0/0; a margin 0.5 mm off -> L1 FAIL', async () => {
  const store = createStore(book());
  const t = lay({ leftMargin: 25.4, topMargin: 10 });
  const a = await run(store, t);
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.equal(store.durable[0].page.leftMargin, 25.4);
  assert.equal((await run(store, t)).m.commands, 0);
  const d = await run(createStore(book()), t, { fault: 'page-margin-drift' });
  assert.equal(d.result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('the default already holds -> SATISFIED 0/0 (portrait, no gridlines)', async () => {
  const r = await run(createStore(book()), lay({ orientation: 'portrait', printGridlines: false }));
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0);
});

for (const [name, script, why] of [
  ['an unproven build (not the W8b-proven one)', { pageBuild: false }, 'xlsx-page-layout-unobservable'],
  ['a getter that returned nothing (a partial read is not the default)', { pageFieldMissing: true }, 'xlsx-page-layout-unobservable'],
  ['a margin getter that returned nothing', { pageMarginMissing: true }, 'xlsx-page-layout-unobservable'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(book()), lay(), script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / partial -> L1 FAIL; the unrequested fields / another sheet / a cell -> L3 FAIL', async () => {
  for (const fault of ['ignore', 'page-partial']) {
    const { result, driver } = await run(createStore(book()), lay(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  for (const fault of ['page-side', 'page-other-sheet', 'side-effect']) {
    const { result } = await run(createStore(book()), lay(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', fault);
  }
});

test('planner: known fields, booleans, millimetre margins in range, at least one', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(lay({ orientation: 'xlLandscape' }), /'portrait' or 'landscape'/);
  refuse(lay({ printGridlines: 'yes' }), /true or false/);
  refuse(lay({}), /requires orientation/);
  refuse(lay({ leftMargin: -1 }), /millimetres in \[0, 200\]/);
  refuse(lay({ topMargin: '10' }), /millimetres/);
  refuse(lay({ margin: 10 }), /does not support field/);
  assert.equal(planOperations(lay({ leftMargin: 25.4, bottomMargin: 0 })).ok, true);
  refuse(lay({ orientation: 'landscape', paper: 'A4' }), /does not support field/);
});

test('the command body: only the requested setters', () => {
  const log = [];
  const ws = { GetName: () => 'P', SetPageOrientation: x => log.push(['o', x]), SetPrintGridlines: x => log.push(['g', x]), SetPrintHeadings: x => log.push(['h', x]) };
  assert.deepEqual(new Function('Api', page.lower(planOperations(lay()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['o', 'xlLandscape'], ['g', true]]);
  log.length = 0;
  assert.deepEqual(new Function('Api', page.lower(planOperations(lay({ printHeadings: false })).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['h', false]]);
  log.length = 0;
  ws.SetLeftMargin = x => log.push(['left', x]); ws.SetTopMargin = x => log.push(['top', x]);
  assert.deepEqual(new Function('Api', page.lower(planOperations(lay({ topMargin: 10, leftMargin: 25.4 })).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['top', 10], ['left', 25.4]]);
});

test('the per-sheet build list is exactly the W8b-proven build; otherwise the read is unobservable, never "the default"', () => {
  // DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): XLSX_RUNTIME_IDS (aliased here) collapsed to the
  // single accepted 9.3.5.0/16 entry -- the historical 9.3.4.60/5..16 chain this list used to carry (progressively
  // widened by X-O03/6b / X-O08 / X-O05/5-B, as listed in the prior revision of this comment) was retired.
  // GH #326/#327: widened again to /16 + /17 (compatibility-first, sheet.protection.read/v1 client-wiring only).
  assert.deepEqual([...EO_PAGE_LAYOUT_PER_SHEET_RUNTIMES], ['9.3.5|euro-office-bridge-v1/9.3.5.0/16', '9.3.5|euro-office-bridge-v1/9.3.5.0/17']);
  const ws = { GetName: () => 'P', GetIndex: () => 0, GetUsedRange: () => null, GetPageOrientation: () => 'xlLandscape', GetPrintGridlines: () => true, GetPrintHeadings: () => false, GetTopMargin: () => 19.1, GetBottomMargin: () => 19.1, GetLeftMargin: () => 72, GetRightMargin: () => 17.8 };
  const Api = { GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } };
  assert.equal(new Function('Api', readQuery([], { pageSheets: ['*'] }))(Api).pageReads.P, 'UNOBSERVABLE:per-sheet-read-not-verified');
  assert.deepEqual(new Function('Api', readQuery([], { pageSheets: ['*'], pagePerSheet: true }))(Api).pageReads.P, { orientation: 'xlLandscape', printGridlines: true, printHeadings: false, topMargin: 19.1, bottomMargin: 19.1, leftMargin: 72, rightMargin: 17.8 });
});
