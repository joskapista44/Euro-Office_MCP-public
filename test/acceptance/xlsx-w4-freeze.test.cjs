'use strict';

const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');
const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const freeze = require('../../src/adapters/xlsx/families/freeze.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 4: freeze_panes / unfreeze_panes through the target path on the fake transport. The fake stands for a
 * build whose own-pane reader IS proven (script.paneBuild defaults to true); the REAL driver's list is empty
 * until the W4d LIVE probe proves the reader, so LIVE every freeze step refuses before the write (tested below).
 */

const book = () => [
  { name: 'Home', cells: { A1: { v: 'start' } }, pane: { row0: 2, col0: 0 } },
  { name: 'Data', cells: { A1: { v: 'Name' }, B2: { v: 5 } } },
  { name: 'Other', cells: { A1: { v: 'keep me' } }, pane: { row0: 0, col0: 1 } },
];

const task = step => ({ id: 'm6-w4-freeze', steps: [{ sheet: 'Data', ...step }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const sheet = (store, n) => store.durable.find(s => s.name === n);
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

for (const [name, step, pane] of [
  ['rows', { intent: 'freeze_panes', rows: 2 }, { row0: 2, col0: 0 }],
  ['columns', { intent: 'freeze_panes', columns: 3 }, { row0: 0, col0: 3 }],
  ['rows AND columns together', { intent: 'freeze_panes', rows: 4, columns: 3 }, { row0: 4, col0: 3 }],
]) {
  test(`freeze ${name} on a NON-active sheet: APPLIED, the original active sheet restored, retry SATISFIED 0/0 without activation`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(step));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.writes, 1);
    assert.equal(m.commands, 1);
    assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
    assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
    assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    assert.deepEqual(sheet(store, 'Data').pane, pane);
    assert.equal(store.active, 'Home', 'the originally active sheet is active again');
    assert.deepEqual(sheet(store, 'Home').pane, { row0: 2, col0: 0 }, 'the active sheet\'s own freeze untouched');
    const again = await run(store, task(step));
    assert.equal(again.m.writes, 0);
    assert.equal(again.m.commands, 0);
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

test('unfreeze: APPLIED, retry SATISFIED 0/0; an unfrozen sheet is SATISFIED at once', async () => {
  const store = createStore(book());
  const { result } = await run(store, { id: 'u', steps: [{ intent: 'unfreeze_panes', sheet: 'Other' }] });
  assert.equal(result.ok, true);
  assert.equal(sheet(store, 'Other').pane, undefined);
  const again = await run(store, { id: 'u', steps: [{ intent: 'unfreeze_panes', sheet: 'Other' }] });
  assert.equal(again.m.commands, 0);
  const none = await run(createStore(book()), { id: 'u', steps: [{ intent: 'unfreeze_panes', sheet: 'Data' }] });
  assert.equal(none.m.commands, 0);
  assert.equal(none.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('RED: the W4b2 trap (the freeze lands on the ACTIVE sheet) -> L1 FAIL and L3 sees the other sheet; nothing persisted', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ intent: 'freeze_panes', rows: 3 }), { fault: 'wrong-sheet' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
  assert.deepEqual(store.durable, book());
});

test('RED: the active sheet not restored -> L1 FAIL (activeKept false)', async () => {
  const { result } = await run(createStore(book()), task({ intent: 'freeze_panes', rows: 2 }), { fault: 'no-restore' });
  assert.equal(result.ok, false);
  const pv = result.execution.primitiveVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.equal(pv.evidence.evidence.activeKept, false);
});

test('RED: ignored / off by one -> L1 FAIL; another sheet\'s pane changed -> L3 FAIL', async () => {
  for (const fault of ['ignore', 'freeze-off-by-one']) {
    const { result } = await run(createStore(book()), task({ intent: 'freeze_panes', rows: 2 }), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
  const { result } = await run(createStore(book()), task({ intent: 'freeze_panes', rows: 2 }), { fault: 'side-effect' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
  assert.deepEqual(result.execution.preservationVerification[0].evidence.evidence.paneChanged.map(x => x.sheet), ['Home']);
});

for (const [name, t, script, why] of [
  ['an unproven pane reader (never "not frozen")', task({ intent: 'freeze_panes', rows: 2 }), { paneBuild: false }, 'xlsx-freeze-pane-unobservable'],
  ['an unfreeze with an unproven reader', { id: 'u', steps: [{ intent: 'unfreeze_panes', sheet: 'Data' }] }, { paneBuild: false }, 'xlsx-freeze-pane-unobservable'],
  ['a missing sheet', { id: 'x', steps: [{ intent: 'freeze_panes', sheet: 'Nope', rows: 1 }] }, {}, 'xlsx-sheet-not-found'],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, t, script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
    assert.deepEqual(store.durable, book());
  });
}

// B06 cutover (internal tracker point 8; LIVE parity #ref): a sheet's OWN pane is read through the DS Bridge
// sheet.freezeState.read/v1 -- these run the REAL served bridge block on the stand-in (test/fakes/ds-bridge-runtime.cjs).
test('no DS Bridge, or another served runtime -> the own pane is unobservable, never "none"', () => {
  const ws = { GetName: () => 'Data', GetIndex: () => 0, GetUsedRange: () => null, worksheet: { sheetViews: [{ pane: null }] } };
  const api = () => ({ GetSheets: () => [ws], GetSheet: n => (n === 'Data' ? ws : null), GetActiveSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } });
  const run = a => new Function('Api', readQuery([], { panes: true }))(a);
  assert.deepEqual(run(api()).panes, { Data: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' });
  const withBridge = api(); installBridge(withBridge);
  assert.deepEqual(run(withPatch(withBridge)).panes, { Data: 'UNOBSERVABLE:PATCH_VERSION_MISMATCH' });
  assert.deepEqual(run(withBridge).panes, { Data: { state: 'none' } }, 'positive control: the same sheet through the bridge');
  assert.equal(run(withBridge).active, 'Data');
});

test('the driver pane read reads the own pane without activation through the bridge; tri-state (M7 S2a, Talk 30096): any gap or split is unobservable, never none', () => {
  const cellAt = (r, c) => ({ getRow0: () => r, getCol0: () => c });
  let activations = 0;
  const mk = (name, sv) => ({ GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, SetActive: () => { activations += 1; }, worksheet: sv === undefined ? undefined : { sheetViews: sv === 'none' ? [] : [sv] } });
  const sheets = [mk('F', { pane: { state: null, topLeftFrozenCell: cellAt(4, 3) } }), mk('R', { pane: { state: 'frozen', topLeftFrozenCell: cellAt(2, 0) } }), mk('K', { pane: null }),
    mk('U', { pane: { state: null, topLeftFrozenCell: null } }), mk('P', { pane: { state: 'split', xSplit: 2000, ySplit: 1500, topLeftFrozenCell: null } }),
    mk('Q', { pane: { state: 'split', topLeftFrozenCell: cellAt(2, 0) } }), mk('N', { pane: { topLeftFrozenCell: cellAt(2, 0) } }), mk('X', undefined),
    mk('Y', { pane: { state: null, topLeftFrozenCell: { row: -1, col: 0 } } }), mk('Z', { pane: { state: 'frozen', topLeftFrozenCell: { row: 5, col: 4 } } })];
  const Api = { GetSheets: () => sheets, GetSheet: n => sheets.find(x => x.GetName() === n) || null, GetActiveSheet: () => sheets[2], WorksheetFunction: { COUNT: () => 0 } };
  installBridge(Api);
  const out = new Function('Api', readQuery([], { panes: true }))(Api);
  const SPLIT = 'UNOBSERVABLE:BRIDGE:pane-split-or-unknown';
  assert.deepEqual(out.panes, { F: { state: 'frozen', row0: 4, col0: 3 }, R: { state: 'frozen', row0: 2, col0: 0 }, K: { state: 'none' }, U: SPLIT, P: SPLIT, Q: SPLIT, N: SPLIT,
    X: 'UNOBSERVABLE:BRIDGE:worksheet-missing', Y: 'UNOBSERVABLE:BRIDGE:pane-unexpected-shape', Z: 'UNOBSERVABLE:BRIDGE:pane-unexpected-shape' });
  assert.equal(out.active, 'K');
  assert.equal(activations, 0, 'the read never activates a sheet');
  assert.equal(new Function('Api', readQuery())({ GetSheets: () => [], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } }).panes, undefined, 'no pane read unless asked');
});

test('planner: the only step of its task, shapes', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(task({ intent: 'freeze_panes' }), /requires rows and \/ or columns/);
  refuse(task({ intent: 'freeze_panes', rows: 0 }), /rows must be an integer/);
  refuse(task({ intent: 'freeze_panes', rows: 1.5 }), /rows must be an integer/);
  refuse(task({ intent: 'freeze_panes', columns: 101 }), /columns must be an integer/);
  refuse(task({ intent: 'unfreeze_panes', rows: 1 }), /does not support field/);
  refuse(task({ intent: 'freeze_panes', rows: 1, at: 'B2' }), /does not support field/);
  refuse({ id: 'x', steps: [{ intent: 'freeze_panes', sheet: 'Data', rows: 1 }, { intent: 'write_range', sheet: 'Other', range: 'C9', values: [['v']] }] }, /LAST step/); // M7 S5 S23: last, not alone
});

test('the command body: activate (checked), the public freeze call, restore the original -- also after a throw', () => {
  const run1 = (step, opts = {}) => {
    const cmd = freeze.lower(planOperations(task(step)).operations[0]);
    const log = [];
    let active = 'Home';
    const fp = new Proxy({}, { get: (_t, m) => (...args) => { log.push([m, ...args.map(a => (a && a.a) || a)]); if (opts.throwIn === m) throw new Error('boom'); } });
    const ws = n => ({ GetName: () => n, SetActive: () => { if (!(opts.stuck && n === 'Data')) active = n; log.push(['SetActive', n]); }, GetFreezePanes: () => fp, GetRange: a => ({ a }) });
    const Api = { GetSheet: n => ws(n), GetActiveSheet: () => ws(active) };
    return { r: new Function('Api', cmd.body)(Api), log, active: () => active };
  };
  const a = run1({ intent: 'freeze_panes', rows: 4, columns: 3 });
  assert.deepEqual(a.r, { ok: true, origin: 'Home', restored: true });
  assert.deepEqual(a.log, [['SetActive', 'Data'], ['FreezeAt', 'C4'], ['SetActive', 'Home']]);
  assert.deepEqual(run1({ intent: 'freeze_panes', rows: 2 }).log[1], ['FreezeRows', 2]);
  assert.deepEqual(run1({ intent: 'freeze_panes', columns: 1 }).log[1], ['FreezeColumns', 1]);
  assert.deepEqual(run1({ intent: 'unfreeze_panes' }).log[1], ['Unfreeze']);
  const t = run1({ intent: 'freeze_panes', rows: 2 }, { throwIn: 'FreezeRows' });
  assert.equal(t.r.ok, false);
  assert.equal(t.r.stage, 'threw');
  assert.equal(t.r.restored, true, 'restored after the throw');
  assert.equal(t.active(), 'Home');
  const s = run1({ intent: 'freeze_panes', rows: 2 }, { stuck: true });
  assert.equal(s.r.stage, 'activate');
  assert.equal(s.log.some(x => x[0] === 'FreezeRows'), false, 'no freeze when the activation did not take');
});
