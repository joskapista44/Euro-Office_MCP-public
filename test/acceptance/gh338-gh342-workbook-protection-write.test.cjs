'use strict';

/**
 * GH #338-#342 remainder (W3) product side over the Bridge /24 workbook.protection.write/v1: protect_workbook /
 * unprotect_workbook (workbook-protection-write.cjs). Offline: the fake driver models the handler's verdict
 * (62133fca's security fix included -- unprotect on a currently LOCKED workbook is ALWAYS refused
 * credential-unobservable, by design, until a future unit sources a real credential detector); this proves the
 * PRODUCT logic -- the engine side is LIVE-proven by gate G1 on the the production host /24 disposable DS.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/workbook-protection-write.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const RT23 = '9.3.5|euro-office-bridge-v1/9.3.5.0/23';
const RT24 = '9.3.5|euro-office-bridge-v1/9.3.5.0/24';
const RT25 = '9.3.5|euro-office-bridge-v1/9.3.5.0/25'; // W2 regression fix (workbook.protection.read/v1 null branch) -- workbook.protection.write/v1 byte-identical to /24
const RT26 = '9.3.5|euro-office-bridge-v1/9.3.5.0/26'; // review fix (f0451f47, the two guard change) -- workbook.protection.write/v1 untouched, byte-identical to /25
const RT27 = '9.3.5|euro-office-bridge-v1/9.3.5.0/27'; // GH #346 AUTO-REPAIR 1. egyseg (chart.series.style.read/v1 dataPoint only) -- workbook.protection.write/v1 byte-identical to /26
const RT28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28'; // GH #361/#363 pre-install readiness gate (internal tracker) -- byte-identical warrant, ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs
const RT29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker) -- byte-identical warrant, ds935-bridge29-veryhidden-write-handler-identity.test.cjs
const cell = (v, f = '') => ({ v, f });
const book = (ws = {}) => [{ name: 'P', cells: { A1: cell('a') }, objects: { charts: [], pivots: [], names: [] }, ...ws }];
const task = step => ({ id: 'gh338wb', steps: [step] });
const protectWb = (extra = {}) => ({ intent: 'protect_workbook', ...extra });
const unprotectWb = (extra = {}) => ({ intent: 'unprotect_workbook', ...extra });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT24, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver };
}
const evidenceOf = r => { const op = r.execution.appliedOperations && r.execution.appliedOperations[0]; if (op) { const e = op.evidence; return e && e.evidence ? e.evidence : e; } return r.execution; };
const deepReason = o => { if (!o || typeof o !== 'object') return undefined; if (typeof o.reason === 'string') return o.reason; for (const v of Object.values(o)) { const x = deepReason(v); if (x) return x; } return undefined; };

// ---- schema + plan + sole-step ---------------------------------------------------------------------------------

test('schema: protect_workbook / unprotect_workbook take no fields at all; an extra field is refused; mixed with anything else refuses the task', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(protectWb()).ok, true, JSON.stringify(v(protectWb())));
  assert.equal(v(unprotectWb()).ok, true, JSON.stringify(v(unprotectWb())));
  assert.equal(v(protectWb({ sheet: 'P' })).ok, false, 'no sheet field on a workbook-scoped intent');
  assert.equal(W.validateWriteRequest(req([protectWb(), unprotectWb()])).ok, false, 'the only step of its task, within the family');
  assert.equal(W.validateWriteRequest(req([protectWb(), { intent: 'write_range', sheet: 'P', range: 'A1', values: [[1]] }])).ok, false, 'the only step of its task, mixed with an unrelated intent');
});

test('plan: the operation is cell-free, workbook-scoped, want.protect mirrors the intent', () => {
  const p = fam.plan(protectWb(), 0);
  assert.equal(p.ok, true);
  assert.equal(p.operation.want.protect, true);
  assert.equal(p.operation.selector.cellFree, true);
  assert.deepEqual(p.operation.cells, []);
  assert.equal(fam.plan(unprotectWb(), 0).operation.want.protect, false);
  assert.equal(fam.plan({ intent: 'protect_workbook', extra: 1 }, 0).ok, false, 'an unknown field is refused');
});

test('planOperations: the SAME sole-step mechanism as the sheet-level protection family, directly on the adapter -- the literal, named reason', () => {
  const alone = xlsxAdapter.planOperations(task(protectWb()));
  assert.equal(alone.ok, true);
  const other = { intent: 'write_range', sheet: 'P', range: 'A1', values: [[1]] };
  const mixed = xlsxAdapter.planOperations({ id: 't', steps: [protectWb(), other] });
  assert.equal(mixed.ok, false);
  assert.equal(mixed.detail, fam.SOLE_STEP_REASON);
  const reversed = xlsxAdapter.planOperations({ id: 't', steps: [other, unprotectWb()] });
  assert.equal(reversed.ok, false, 'refuses regardless of order');
  assert.equal(reversed.detail, fam.SOLE_STEP_REASON);
});

// ---- runtime -----------------------------------------------------------------------------------------------

test('runtime: /24 + /25 (byte-identical write handler) -- adapter and driver agree; a /23 runtime is FAIL_CLOSED named, with no dispatch', async () => {
  assert.deepEqual([...fam.B28_BRIDGE_WRITE_RUNTIMES], [RT24, RT25, RT26, RT27, RT28, RT29]);
  assert.deepEqual([...drv.B28_BRIDGE_WRITE_RUNTIMES], [RT24, RT25, RT26, RT27, RT28, RT29]);
  const store = createStore(book());
  let n = 0;
  const d = createFakeXlsxDriver(store, { runtime: RT23 });
  const r = await withSession({ driver: d, fileId: 'x', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(task(protectWb()), { observation: { live: liveScope(task(protectWb())) } }));
  assert.equal(deepReason(evidenceOf(r)), 'xlsx-workbook-protection-bridge-write-runtime-not-proven');
  assert.equal(d.calls.bridgeWrites || 0, 0);
  assert.equal(store.workbookLocked, undefined, 'untouched');
});

// ---- GH #366 4th unit: the read-bound refusal, PRE-dispatch, named --------------------------------------------

for (const [name, step] of [['protect_workbook', protectWb()], ['unprotect_workbook', unprotectWb()]]) {
  test(`GH #366 4th unit: ${name} on a truncated observation is refused BEFORE dispatch, named 'xlsx-observation-truncated' (same mechanism as protect_sheet/delete-cells/write_range)`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, task(step), { readTruncated: true });
    assert.equal(result.ok, false);
    assert.equal(deepReason(evidenceOf(result)), 'xlsx-observation-truncated');
    assert.equal(driver.calls.bridgeWrites || 0, 0, 'no dispatch');
    assert.equal(store.workbookLocked, undefined, 'untouched');
  });
}

test("GH #366 4th unit: a COMPLETE observation is unaffected -- protect_workbook still APPLIES exactly as before this unit's change", async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(protectWb()), { readTruncated: false });
  assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
  assert.equal(driver.calls.bridgeWrites, 1);
  assert.equal(store.workbookLocked, true);
});

test('GH #366 4th unit: the truncated refusal fires BEFORE the credential-unobservable check -- a currently-locked workbook over the bound still names the geometric reason, not the security one', async () => {
  const store = createStore(book());
  store.workbookLocked = true;
  const { result, driver } = await run(store, task(unprotectWb()), { readTruncated: true });
  assert.equal(result.ok, false);
  assert.equal(deepReason(evidenceOf(result)), 'xlsx-observation-truncated');
  assert.equal(driver.calls.bridgeWrites || 0, 0);
});

// ---- protect / unprotect -------------------------------------------------------------------------------------

test('protect_workbook: APPLIED on an unlocked workbook, verified by the fresh workbook-protection read; the retry is SATISFIED with 0 write', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(protectWb()));
  assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
  assert.equal(driver.calls.bridgeWrites, 1);
  assert.equal(store.workbookLocked, true);
  const again = await run(store, task(protectWb()));
  assert.equal(again.result.ok, true);
  assert.equal(again.driver.calls.bridgeWrites || 0, 0, 'SATISFIED: no write');
});

test('unprotect_workbook on an UNLOCKED workbook: SATISFIED immediately, 0 writes -- already at the target', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(unprotectWb()));
  assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
  assert.equal(driver.calls.bridgeWrites || 0, 0, 'SATISFIED: no write, nothing to save');
});

test('unprotect_workbook on a LOCKED workbook: ALWAYS refused credential-unobservable, PRE_DISPATCH, 0 write -- by design, until a credential detector is sourced', async () => {
  const store = createStore(book());
  store.workbookLocked = true;
  const { result, driver } = await run(store, task(unprotectWb()));
  assert.equal(result.ok, false);
  assert.equal(deepReason(evidenceOf(result)), 'xlsx-workbook-unprotect-credential-unobservable');
  assert.equal(driver.calls.bridgeWrites || 0, 0, 'refused before any dispatch -- the Bridge is never even called');
  assert.equal(store.workbookLocked, true, 'untouched');
});

test('NEG: a write the engine ignores is never success', async () => {
  const store = createStore(book());
  const r = await run(store, task(protectWb()), { fault: 'ignore' });
  assert.equal(r.result.ok, false);
});

test('the driver second line: the lowered command carries the ONE Bridge write, requesting lockStructure; a direct dispatch against the fake sees the SAME closed door the real Bridge shows on a currently-locked workbook', async () => {
  const l = fam.lower(fam.plan(unprotectWb(), 0).operation);
  assert.equal(l.meta.bridgeWorkbookProtectionWrite, true);
  assert.equal(l.meta.family, 'workbook-protection');
  assert.match(l.body, /__eoWbProtGate\(\)/);
  const args = JSON.parse(/__eoWbProtWrite\((\{.*?\}), \[/.exec(l.body)[1]);
  assert.deepEqual(args, { protect: false });
  const pre = drv.workbookProtectionWriteBridgeWritePrelude();
  assert.match(pre, /euro-office-bridge-v1\/9\.3\.5\.0\/24/);
  assert.doesNotMatch(pre, /9\.3\.5\.0\/23"/);
  // bypassing classify() entirely (as protection-write.cjs's own sibling test does for /22): the SAME
  // fake dispatch function the family's own lower() meta routes to refuses credential-unobservable too.
  const store = createStore(book());
  store.workbookLocked = true;
  const driver = createFakeXlsxDriver(store, { runtime: RT24 });
  await driver.open({ fileId: 'x' });
  const dispatched = await driver.dispatch({}, l);
  assert.equal(dispatched.ok, false);
  assert.match(JSON.stringify(dispatched.evidence), /credential-unobservable/);
});

test('windowsOf / liveExtras: no cell window of its own; opts in to the additive workbook-protection read', () => {
  assert.deepEqual(fam.windowsOf(fam.plan(protectWb(), 0).operation), []);
  assert.deepEqual(fam.liveExtras(), { workbookProtection: true });
  const scope = liveScope(task(protectWb()));
  assert.equal(scope.workbookProtection, true);
  assert.equal(liveScope({ id: 't', steps: [{ intent: 'write_range', sheet: 'P', range: 'A1', values: [[1]] }] }).workbookProtection, undefined, 'not requested by an unrelated task');
});
