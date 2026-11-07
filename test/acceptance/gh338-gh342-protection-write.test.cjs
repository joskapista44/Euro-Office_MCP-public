'use strict';

/**
 * GH #338-#342 product side over the Bridge /22 sheet.protection.write/v1 + sheet.cellProtection.write/v1: protect_sheet / unprotect_sheet /
 * set_cell_protection (protection-write.cjs) and the read_protection mirror (protection-read.cjs). Offline: the fake driver models the
 * handlers' verdicts and the /17 read's shape; this proves the PRODUCT logic -- the engine side is LIVE-proven by gate G1 on the the production host /22 DS.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const rt = require('../../src/server/read-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/protection-write.cjs');
const { projectProtectionRead } = require('../../src/adapters/xlsx/families/protection-read.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const RT22 = '9.3.5|euro-office-bridge-v1/9.3.5.0/22';
const RT23 = '9.3.5|euro-office-bridge-v1/9.3.5.0/23';
const RT24 = '9.3.5|euro-office-bridge-v1/9.3.5.0/24';
const RT25 = '9.3.5|euro-office-bridge-v1/9.3.5.0/25'; // W2 regression fix (workbook.protection.read/v1 null branch) -- sheet.protection.write/v1 byte-identical to /24
const RT26 = '9.3.5|euro-office-bridge-v1/9.3.5.0/26'; // review fix (f0451f47, the two guard change) -- sheet.protection.write/v1 byte-identical to /25
const RT27 = '9.3.5|euro-office-bridge-v1/9.3.5.0/27'; // GH #346 AUTO-REPAIR 1. egyseg (chart.series.style.read/v1 dataPoint only) -- sheet.protection.write/v1 byte-identical to /26
const RT28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28'; // GH #361/#363 pre-install readiness gate (internal tracker) -- byte-identical warrant, ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs
const RT29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker) -- byte-identical warrant, ds935-bridge29-veryhidden-write-handler-identity.test.cjs
const RT21 = '9.3.5|euro-office-bridge-v1/9.3.5.0/21';
const cell = (v, f = '') => ({ v, f });
const book = (ws = {}) => [{ name: 'P', cells: { A1: cell('a'), B2: cell(1), C1: cell(2, '=B2*2') }, objects: { charts: [], pivots: [], names: [] }, ...ws },
  { name: 'Other', cells: { A1: cell('o') }, objects: { charts: [], pivots: [], names: [] } }];
const task = step => ({ id: 'gh338', steps: [step] });
const protect = (extra = {}) => ({ intent: 'protect_sheet', sheet: 'P', forbids: { formatCells: false, insertRows: false }, ...extra });
const unprotect = (extra = {}) => ({ intent: 'unprotect_sheet', sheet: 'P', ...extra });
const cellp = (extra = {}) => ({ intent: 'set_cell_protection', sheet: 'P', range: 'B2:B3', locked: false, ...extra });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT22, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver };
}
const evidenceOf = r => { const op = r.execution.appliedOperations && r.execution.appliedOperations[0]; if (op) { const e = op.evidence; return e && e.evidence ? e.evidence : e; } return r.execution; };
const deepReason = o => { if (!o || typeof o !== 'object') return undefined; if (typeof o.reason === 'string') return o.reason; for (const v of Object.values(o)) { const x = deepReason(v); if (x) return x; } return undefined; };
const sheetOf = store => store.durable[0];

// ---- schema + plan ---------------------------------------------------------------------------------------

test('schema: the three intents take exactly their fields; an unknown flag, a non-boolean, no locked/hidden, a credential-like key are refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  for (const s of [protect(), protect({ forbids: undefined }), unprotect(), cellp(), cellp({ locked: undefined, hidden: true })]) assert.equal(v(s).ok, true, JSON.stringify(v(s)));
  for (const [name, s] of [
    ['unknown flag', protect({ forbids: { sorting: true } })], ['flag not boolean', protect({ forbids: { sort: 'no' } })],
    ['password key', protect({ forbids: { password: 'x' } })], ['password field', protect({ password: 'x' })],
    ['unprotect with forbids', unprotect({ forbids: {} })], ['cell: neither', cellp({ locked: undefined })], ['cell: bad range', cellp({ range: 'B:B' })],
    ['cell: locked string', cellp({ locked: 'no' })], ['no sheet', protect({ sheet: undefined })],
  ]) assert.equal(v(s).ok, false, name);
  assert.equal(W.validateWriteRequest(req([protect(), unprotect()])).ok, false, 'the only step of its task');
});

test('plan: forbids is canonical (every flag, read order, the Excel default for the unnamed); a range is normalised and capped', () => {
  const p = fam.plan(protect(), 0);
  assert.equal(p.ok, true);
  assert.deepEqual(Object.keys(p.operation.want.forbids), fam.FLAGS);
  assert.deepEqual(p.operation.want.forbids, { ...fam.DEFAULT_FORBIDS, formatCells: false, insertRows: false });
  assert.equal(fam.DEFAULT_FORBIDS.selectLockedCells, false); assert.equal(fam.DEFAULT_FORBIDS.sort, true);
  assert.equal(fam.plan(cellp({ range: 'b2:b3' }), 0).operation.selector.range, 'B2:B3');
  assert.equal(fam.plan(cellp({ range: 'A1:Z100' }), 0).ok, false, 'over the 1000-cell ceiling');
});

// ---- runtime -----------------------------------------------------------------------------------------------

test('runtime: /22 + /23 + /24 (the /23 and /24 copies byte-identical) -- adapter and driver agree; a /21 runtime is FAIL_CLOSED named, with no dispatch', async () => {
  assert.deepEqual([...fam.B27_BRIDGE_WRITE_RUNTIMES], [RT22, RT23, RT24, RT25, RT26, RT27, RT28, RT29]);
  assert.deepEqual([...drv.B27_BRIDGE_WRITE_RUNTIMES], [RT22, RT23, RT24, RT25, RT26, RT27, RT28, RT29]);
  const store = createStore(book());
  const { result, driver } = await (async () => { let n = 0; const d = createFakeXlsxDriver(store, { runtime: RT21 }); const r = await withSession({ driver: d, fileId: 'x', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(task(protect()), { observation: { live: liveScope(task(protect())) } })); return { result: r, driver: d }; })();
  assert.equal(deepReason(evidenceOf(result)), 'xlsx-protection-bridge-write-runtime-not-proven');
  assert.equal(driver.calls.bridgeWrites || 0, 0);
  assert.equal(sheetOf(store).protection, undefined);
});

// ---- GH #366 2nd unit: the read-bound refusal, PRE-dispatch, named ---------------------------------------------

for (const [name, step] of [['protect_sheet', protect()], ['unprotect_sheet', unprotect()], ['set_cell_protection', cellp()]]) {
  test(`GH #366 2nd unit: ${name} on a truncated observation is refused BEFORE dispatch, named 'xlsx-observation-truncated' (same mechanism as delete-cells/write_range)`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, task(step), { readTruncated: true });
    assert.equal(result.ok, false);
    assert.equal(deepReason(evidenceOf(result)), 'xlsx-observation-truncated');
    assert.equal(driver.calls.bridgeWrites || 0, 0, 'no dispatch');
    assert.deepEqual(store.durable, book(), 'durable untouched');
  });
}

test("GH #366 2nd unit: a COMPLETE observation is unaffected -- protect_sheet still APPLIES exactly as before this unit's change", async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(protect()), { readTruncated: false });
  assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
  assert.equal(driver.calls.bridgeWrites, 1);
});

// ---- protect / unprotect -------------------------------------------------------------------------------------

test('protect_sheet: APPLIED, verified by the fresh protection read; the retry is SATISFIED with 0 write', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task(protect()));
  assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
  assert.equal(driver.calls.bridgeWrites, 1);
  assert.deepEqual(sheetOf(store).protection.forbids, { ...fam.DEFAULT_FORBIDS, formatCells: false, insertRows: false });
  const again = await run(store, task(protect()));
  assert.equal(again.result.ok, true);
  assert.equal(again.driver.calls.bridgeWrites || 0, 0, 'SATISFIED: no write');
});

test('#342: re-protecting with other forbids applies; the #326 guard then decides per flag on the product-protected sheet', async () => {
  const store = createStore(book());
  await run(store, task(protect({ forbids: { formatCells: false } })));
  const fmt = await run(store, task({ intent: 'format_range', sheet: 'P', range: 'A1', format: { fill: '#FFC000' } }));
  assert.equal(fmt.result.ok, true, 'formatCells allowed -> format_range applies');
  const rows = await run(store, task({ intent: 'insert_rows', sheet: 'P', at: 1, count: 1, anchor: { column: 'A', value: 'a' } }));
  assert.equal(deepReason(evidenceOf(rows.result)), 'xlsx-protected-sheet-forbids:insertRows', 'insertRows left at the Excel default (forbidden)');
  await run(store, task(protect({ forbids: { formatCells: false, insertRows: false } })));
  const rows2 = await run(store, task({ intent: 'insert_rows', sheet: 'P', at: 1, count: 1, anchor: { column: 'A', value: 'a' } }));
  assert.notEqual(deepReason(evidenceOf(rows2.result)), 'xlsx-protected-sheet-forbids:insertRows', 'insertRows now allowed');
});

test('unprotect_sheet: APPLIED on a protected sheet (the #326 guard lets the protection change itself through); on an unprotected one SATISFIED', async () => {
  const store = createStore(book({ protection: { forbids: { formatCells: true } } }));
  const { result, driver } = await run(store, task(unprotect()));
  assert.equal(result.ok, true, JSON.stringify(evidenceOf(result)));
  assert.equal(driver.calls.bridgeWrites, 1);
  assert.equal(sheetOf(store).protection, null);
  const again = await run(store, task(unprotect()));
  assert.equal(again.result.ok, true);
  assert.equal(again.driver.calls.bridgeWrites || 0, 0);
});

test('a credential-protected sheet: the handler refuses (NONE), nothing changes, and the result names it -- protect and unprotect alike', async () => {
  for (const step of [unprotect(), protect()]) {
    const store = createStore(book({ protection: { forbids: { formatCells: true } }, protectionCredential: true }));
    const { result } = await run(store, task(step));
    assert.equal(result.ok, false, step.intent);
    assert.match(JSON.stringify(evidenceOf(result)), /credential-protected/, step.intent);
    assert.deepEqual(sheetOf(store).protection, { forbids: { formatCells: true } }, 'untouched');
  }
});

// ---- set_cell_protection ---------------------------------------------------------------------------------------

test('set_cell_protection: unlock B2:B3 (B3 empty) + hide C1 on an unprotected sheet, verified per cell; the retry is SATISFIED', async () => {
  const store = createStore(book());
  const a = await run(store, task(cellp()));
  assert.equal(a.result.ok, true, JSON.stringify(evidenceOf(a.result)));
  const b = await run(store, task(cellp({ range: 'C1', locked: undefined, hidden: true })));
  assert.equal(b.result.ok, true);
  assert.deepEqual(sheetOf(store).cellProtection, { unlocked: ['B2', 'B3'], hidden: ['C1'] });
  const again = await run(store, task(cellp()));
  assert.equal(again.driver.calls.bridgeWrites || 0, 0, 'SATISFIED');
  assert.equal(sheetOf(store).cells.C1.f, '=B2*2', 'the formula itself is untouched');
});

test('set_cell_protection on a PROTECTED sheet: refused before dispatch by the #326 guard, named, 0 write', async () => {
  const store = createStore(book({ protection: { forbids: {} }, cellProtection: { unlocked: [], hidden: ['C1'] } }));
  const { result, driver } = await run(store, task(cellp({ range: 'C1', locked: undefined, hidden: false })));
  assert.equal(deepReason(evidenceOf(result)), 'xlsx-protected-sheet-forbids:cellProtection');
  assert.equal(driver.calls.bridgeWrites || 0, 0);
  assert.deepEqual(sheetOf(store).cellProtection.hidden, ['C1'], '#327 redaction cannot be turned off by a write');
});

test('NEG: a write the engine ignores is never success; a collateral cell change fails L3', async () => {
  let store = createStore(book());
  let r = await run(store, task(protect()), { fault: 'ignore' });
  assert.equal(r.result.ok, false);
  store = createStore(book());
  r = await run(store, task(cellp()), { fault: 'prot-cell-side-effect' });
  assert.equal(r.result.ok, false, 'Z99 changed: preservation must fail');
});

test('the lowered command: one Bridge write, the capability id in meta, the forbids in read order; the driver prelude gates /22', () => {
  const l = fam.lower(fam.plan(protect(), 0).operation);
  assert.equal(l.meta.bridgeProtectionWrite, 'sheet.protection.write/v1');
  assert.match(l.body, /__eoProtGate\(\)/);
  const args = JSON.parse(/__eoProtWrite\((\{.*?\}), \[/.exec(l.body)[1]);
  assert.deepEqual(Object.keys(args.forbids), fam.FLAGS);
  assert.equal(fam.lower(fam.plan(cellp(), 0).operation).meta.bridgeProtectionWrite, 'sheet.cellProtection.write/v1');
  const pre = drv.protectionWriteBridgeWritePrelude('sheet.protection.write/v1');
  assert.match(pre, /euro-office-bridge-v1\/9\.3\.5\.0\/22/);
  assert.doesNotMatch(pre, /9\.3\.5\.0\/21"/);
});

// ---- read_protection ------------------------------------------------------------------------------------------

test('read_protection: the request names the #326 axes; the projection is a field-select, gaps named, no credential', () => {
  assert.deepEqual(rt.xlsxRequest('read_protection', { sheet: 'P' }), { windows: [], protectionSheets: ['P'], workbookProtection: true });
  assert.deepEqual(rt.xlsxRequest('read_protection', { sheet: 'P', range: 'A1:C3' }), { windows: [], protectionRanges: [{ sheet: 'P', range: 'A1:C3' }], workbookProtection: true });
  const inv = { sheets: [{ name: 'P' }], protectionReads: { P: { protected: true, forbids: { ...fam.DEFAULT_FORBIDS, hashValue: 'X' }, cells: null } } };
  const out = projectProtectionRead(inv, { sheet: 'P' });
  assert.equal(out.protected, true);
  // GH #338-#342 (/24 only, W1): the workbook field is ADDITIVE, independent of the per-sheet read -- not read here, so UNOBSERVABLE
  assert.deepEqual(out.workbook, { lockStructure: null, reason: 'UNOBSERVABLE:not-read' });
  assert.equal(projectProtectionRead({ sheets: [{ name: 'P' }], protectionReads: { P: 'UNOBSERVABLE:CAPABILITY_MISSING' } }, { sheet: 'P' }).reason, 'UNOBSERVABLE:CAPABILITY_MISSING');
  assert.equal(projectProtectionRead({ sheets: [{ name: 'P' }] }, { sheet: 'Nope' }).error, 'SHEET_NOT_FOUND');
  // the workbook field is attached even on SHEET_NOT_FOUND -- it does not depend on the sheet's own existence
  assert.deepEqual(projectProtectionRead({ sheets: [{ name: 'P' }], workbookProtectionRead: { lockStructure: true } }, { sheet: 'Nope' }).workbook, { lockStructure: true });
  const ranged = projectProtectionRead({ sheets: [{ name: 'P' }], protectionRangeReads: { 'P!A1:C3': { protected: false, forbids: null, cells: { range: 'A1:C3', nonEmpty: 3, unlocked: ['B2'], hidden: ['C1'] } } } }, { sheet: 'P', range: 'A1:C3' });
  assert.deepEqual(ranged, { sheet: 'P', range: 'A1:C3', protected: false, forbids: null, cells: { nonEmpty: 3, unlocked: ['B2'], hidden: ['C1'] }, workbook: { lockStructure: null, reason: 'UNOBSERVABLE:not-read' } });
});
