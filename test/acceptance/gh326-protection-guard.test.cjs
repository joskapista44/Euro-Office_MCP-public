'use strict';

/**
 * GH #326 (U4b, owner decision): the sheet-protection WRITE guard --
 * (a) task-level, before the FIRST mutation; (b) fail-closed when the protection itself cannot be read;
 * (c) re-checked, unchanged, right before the mutation (the race condition). docs/DS-9.3.5-BRIDGE17-
 * PROTECTION-READ.md is the capability contract; src/adapters/xlsx/protection-gate.cjs is the decision over
 * it; src/drivers/eurooffice/xlsx-driver.cjs's protectionRacePrelude is (c)'s mechanism.
 *
 * Part A: protection-gate.cjs's own condition table, pure (no editor, no fake driver) -- the exact GH299
 * PR-B02 situations: a locked cell in the write target refuses; every cell unlocked proceeds; a named
 * `forbids` flag refuses format_range/insert_rows; every OTHER intent on a protected sheet refuses, named;
 * an unobservable protection axis (a /16-shaped runtime) fails closed; a missing/case-mismatched sheet is
 * left to the family's own classify() (not this gate's concern).
 * Part B: protectionRacePrelude()'s own generated JS, against the REAL /17 candidate block (the SAME
 * fixture ds935-bridge17-protection-read.test.cjs / gh327-hidden-formula-redaction.test.cjs use) -- unchanged
 * protection between "expected" and the re-read -> the rest of the command runs; a change -> it aborts
 * before a single setter.
 * Part C: the product WRITE tool, end to end, over the fake XLSX driver (test/fakes/fake-xlsx-driver.cjs,
 * extended this round with `sheet.protection` opt-in) -- (a) a locked write is refused with ZERO dispatch,
 * an unlocked one proceeds; (b) a multi-step task whose SECOND step is protected-refused never dispatches
 * its first step either (preflightDecisive); (c) the protection axis itself unreadable fails the whole task
 * closed, named.
 *
 * Part E (GH #338-#342 remainder, W2, docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md "W2 -- the five
 * structural guards"): the workbook-wide structure lock, orthogonal to the per-sheet axis above --
 * `workbookGate`'s own condition table, pure (E1-E4); `protectionRacePrelude`'s new workbook branch against
 * the REAL /24 candidate block (E5, the SAME style as Part B); the product WRITE tool end to end for all
 * five `WORKBOOK_STRUCTURE_INTENTS` (E6); two local NEGATIVE controls proving E1/E6 are not vacuously green
 * (E7).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const protectionGate = require('../../src/adapters/xlsx/protection-gate.cjs');
const { protectionRacePrelude } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { harness, req, createStore } = require('../fakes/write-wiring.cjs');

// ---- Part A: protection-gate.cjs, pure ------------------------------------------------------------------------

// GH #338-#342 remainder (W2): every op now ALSO carries `workbookProtectionRead` (sheets.cjs's liveExtras
// opts every one of its 5 intents into it unconditionally) -- defaulted here to the pre-/24 pass-through
// shape so every PRE-EXISTING test below, none of which is testing the workbook axis, keeps exercising only
// the sheet-level axis it always did. A test that wants the workbook axis itself passes `workbook` explicitly.
const primaryOf = ({ sheets = ['S'], range = {}, sheetLevel = {}, workbook = 'UNOBSERVABLE:CAPABILITY_MISSING' } = {}) => ({
  sheets: sheets.map(name => ({ name })),
  protectionRangeReads: range,
  protectionReads: sheetLevel,
  workbookProtectionRead: workbook,
});
const unprotected = { protected: false, forbids: null };
const protectedWith = (forbids = {}) => ({ protected: true, forbids: Object.assign({ formatCells: false, insertRows: false, insertColumns: false }, forbids) });
const op = (intent, extra) => ({ intent, selector: { sheet: 'S', range: 'A1:B2' }, ...extra });

test('A1: write_range -- every cell of the target in cells.unlocked -> proceeds, carrying a raceCheck', () => {
  const primary = primaryOf({ range: { 'S!A1:B2': { protected: true, forbids: {}, cells: { range: 'A1:B2', nonEmpty: 4, unlocked: ['A1', 'A2', 'B1', 'B2'], hidden: [] } } } });
  const g = protectionGate.gate(primary, op('write_range'));
  assert.equal(g.refusal, null);
  assert.deepEqual(g.raceChecks, [{ sheet: 'S', range: 'A1:B2', expected: primary.protectionRangeReads['S!A1:B2'] }]);
});

test('A2: write_range -- ONE cell missing from cells.unlocked (locked, or simply empty -- the Excel default) -> refused, named, no raceCheck', () => {
  const primary = primaryOf({ range: { 'S!A1:B2': { protected: true, forbids: {}, cells: { range: 'A1:B2', nonEmpty: 3, unlocked: ['A1', 'A2', 'B1'], hidden: [] } } } });
  const g = protectionGate.gate(primary, op('write_range'));
  assert.equal(g.refusal.reason, 'xlsx-protected-locked-cells:B2');
  assert.deepEqual(g.raceChecks, []);
});

test('A3: clear_range -- the SAME cell-level rule as write_range', () => {
  const primary = primaryOf({ range: { 'S!A1:B2': { protected: true, forbids: {}, cells: { range: 'A1:B2', nonEmpty: 0, unlocked: [], hidden: [] } } } });
  // nonEmpty 0: every cell in the range is BLANK, hence never in `unlocked` -- the fail-closed direction
  const g = protectionGate.gate(primary, op('clear_range'));
  assert.equal(g.refusal.reason, 'xlsx-protected-locked-cells:A1,B1,A2,B2');
});

test('A4: copy_range -- ONLY the destination is checked (Copy leaves the source unchanged)', () => {
  const primary = primaryOf({ range: { 'S!C1:D2': { protected: true, forbids: {}, cells: { range: 'C1:D2', nonEmpty: 0, unlocked: [], hidden: [] } } } });
  // the source range ('A1:B2') is NEVER asked -- if the gate looked at it, this would throw (no such key)
  const g = protectionGate.gate(primary, op('copy_range', { dest: { sheet: 'S', range: 'C1:D2' } }));
  assert.equal(g.refusal.reason, 'xlsx-protected-locked-cells:C1,D1,C2,D2');
});

test('A5: move_range -- BOTH the source and the destination are checked (Cut clears the source, writes the destination)', () => {
  const okSrc = { protected: true, forbids: {}, cells: { range: 'A1:B2', nonEmpty: 4, unlocked: ['A1', 'A2', 'B1', 'B2'], hidden: [] } };
  const lockedDst = { protected: true, forbids: {}, cells: { range: 'C1:D2', nonEmpty: 1, unlocked: [], hidden: [] } };
  const primary = primaryOf({ range: { 'S!A1:B2': okSrc, 'S!C1:D2': lockedDst } });
  const g = protectionGate.gate(primary, op('move_range', { dest: { sheet: 'S', range: 'C1:D2' } }));
  assert.equal(g.refusal.reason, 'xlsx-protected-locked-cells:C1,D1,C2,D2', 'the SOURCE passed; the DESTINATION is the one that refuses');
  // and the mirror: a locked SOURCE refuses even with an unlocked destination
  const lockedSrc = { protected: true, forbids: {}, cells: { range: 'A1:B2', nonEmpty: 1, unlocked: [], hidden: [] } };
  const okDst = { protected: true, forbids: {}, cells: { range: 'C1:D2', nonEmpty: 4, unlocked: ['C1', 'C2', 'D1', 'D2'], hidden: [] } };
  const g2 = protectionGate.gate(primaryOf({ range: { 'S!A1:B2': lockedSrc, 'S!C1:D2': okDst } }), op('move_range', { dest: { sheet: 'S', range: 'C1:D2' } }));
  assert.equal(g2.refusal.reason, 'xlsx-protected-locked-cells:A1,B1,A2,B2');
  // and both unlocked -> two raceChecks, one per side
  const g3 = protectionGate.gate(primaryOf({ range: { 'S!A1:B2': okSrc, 'S!C1:D2': okDst } }), op('move_range', { dest: { sheet: 'S', range: 'C1:D2' } }));
  assert.equal(g3.refusal, null);
  assert.deepEqual(g3.raceChecks, [{ sheet: 'S', range: 'A1:B2', expected: okSrc }, { sheet: 'S', range: 'C1:D2', expected: okDst }]);
});

test('A6: format_range -- the sheet\'s own forbids.formatCells decides, no per-cell check', () => {
  const allowed = protectionGate.gate(primaryOf({ sheetLevel: { S: protectedWith({ formatCells: false }) } }), op('format_range'));
  assert.equal(allowed.refusal, null);
  const forbidden = protectionGate.gate(primaryOf({ sheetLevel: { S: protectedWith({ formatCells: true }) } }), op('format_range'));
  assert.equal(forbidden.refusal.reason, 'xlsx-protected-sheet-forbids:formatCells');
});

test('A7: insert_rows -- the SAME shape, forbids.insertRows', () => {
  const allowed = protectionGate.gate(primaryOf({ sheetLevel: { S: protectedWith({ insertRows: false }) } }), op('insert_rows'));
  assert.equal(allowed.refusal, null);
  const forbidden = protectionGate.gate(primaryOf({ sheetLevel: { S: protectedWith({ insertRows: true }) } }), op('insert_rows'));
  assert.equal(forbidden.refusal.reason, 'xlsx-protected-sheet-forbids:insertRows');
});

test('A8: every OTHER intent on a protected sheet is refused, named with the intent -- nothing about it is measured', () => {
  for (const intent of ['merge_range', 'delete_rows', 'sort_range', 'set_validation', 'rename_sheet', 'delete_cells']) {
    const g = protectionGate.gate(primaryOf({ sheetLevel: { S: protectedWith() } }), op(intent));
    assert.equal(g.refusal.reason, `xlsx-protected-sheet-intent-not-measured:${intent}`, intent);
  }
});

test('A9: unprotected -- EVERY intent is unchanged, no restriction at all (but still watched by the race-check)', () => {
  for (const intent of ['write_range', 'clear_range', 'format_range', 'insert_rows', 'merge_range', 'sort_range']) {
    const g = protectionGate.gate(primaryOf({ sheetLevel: { S: unprotected }, range: { 'S!A1:B2': unprotected } }), op(intent));
    assert.equal(g.refusal, null, intent);
    assert.equal(g.raceChecks.length, 1, intent);
  }
});

test('A10: the protection axis itself unreadable (a /16-shaped runtime, or any other gap) -> FAIL_CLOSED, named, no raceCheck', () => {
  const g1 = protectionGate.gate(primaryOf({ range: { 'S!A1:B2': 'UNOBSERVABLE:CAPABILITY_MISSING' } }), op('write_range'));
  assert.equal(g1.refusal.reason, 'xlsx-protection-unobservable:CAPABILITY_MISSING');
  const g2 = protectionGate.gate(primaryOf({ sheetLevel: { S: 'UNOBSERVABLE:CAPABILITY_MISSING' } }), op('format_range'));
  assert.equal(g2.refusal.reason, 'xlsx-protection-unobservable:CAPABILITY_MISSING');
  // not even asked (the key is simply missing) -- the SAME fail-closed degrade, named not-read
  const g3 = protectionGate.gate(primaryOf({}), op('merge_range'));
  assert.equal(g3.refusal.reason, 'xlsx-protection-unobservable:not-read');
  for (const g of [g1, g2, g3]) assert.deepEqual(g.raceChecks, []);
});

test('A11: a missing / case-mismatched sheet is NOT this gate\'s concern -- it passes through for the family\'s own classify() to refuse', () => {
  const primary = { sheets: [{ name: 'Sheet1' }], protectionReads: {}, protectionRangeReads: {} };
  assert.deepEqual(protectionGate.gate(primary, op('write_range', { selector: { sheet: 'sheet1', range: 'A1:B2' } })), { refusal: null, raceChecks: [] }, 'case-mismatch');
  assert.deepEqual(protectionGate.gate(primary, op('merge_range', { selector: { sheet: 'NoSuchSheet', range: 'A1:B2' } })), { refusal: null, raceChecks: [] }, 'missing');
});

test('A12: create_sheet and a `names` family synthetic selector are out of scope -- no sheet exists to be protected', () => {
  // GH #338-#342 remainder (W2): create_sheet IS a WORKBOOK_STRUCTURE_INTENT, so it now runs workbookGate
  // first -- CAPABILITY_MISSING (the pre-/24 pass-through) keeps this test on the per-sheet axis it tests.
  const primary = { sheets: [{ name: 'S' }], protectionReads: { S: protectedWith() }, protectionRangeReads: {}, workbookProtectionRead: 'UNOBSERVABLE:CAPABILITY_MISSING' };
  assert.deepEqual(protectionGate.gate(primary, { intent: 'create_sheet', selector: { name: 'New' } }), { refusal: null, raceChecks: [] });
  assert.deepEqual(protectionGate.gate(primary, { intent: 'set_defined_name', selector: { sheet: '\u0000name:s!x', range: 'A1', name: 'x' } }), { refusal: null, raceChecks: [] });
});

test('A13: targetsOf() -- the SAME pairs gate() will need, pure in `op` alone (liveScope runs before any observation)', () => {
  assert.deepEqual(protectionGate.targetsOf(op('write_range')), [{ sheet: 'S', range: 'A1:B2' }]);
  assert.deepEqual(protectionGate.targetsOf(op('copy_range', { dest: { sheet: 'S', range: 'C1:D2' } })), [{ sheet: 'S', range: 'C1:D2' }]);
  assert.deepEqual(protectionGate.targetsOf(op('move_range', { dest: { sheet: 'T', range: 'C1:D2' } })), [{ sheet: 'S', range: 'A1:B2' }, { sheet: 'T', range: 'C1:D2' }]);
  assert.deepEqual(protectionGate.targetsOf(op('format_range')), [{ sheet: 'S' }]);
  assert.deepEqual(protectionGate.targetsOf(op('merge_range')), [{ sheet: 'S' }]);
  assert.deepEqual(protectionGate.targetsOf({ intent: 'create_sheet', selector: { name: 'New' } }), []);
});

// ---- Part B: protectionRacePrelude(), against the REAL /17 candidate block --------------------------------------

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');
const a1 = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: Number(m[2]) - 1, c: c - 1 }; };
const bbox = range => { const [p, q = p] = range.split(':'); const x = a1(p); const y = a1(q); return { r1: x.r, c1: x.c, r2: y.r, c2: y.c }; };
function sheet(name, { sp = null } = {}) {
  const worksheet = { sheetProtection: sp, getRange3: () => ({ _foreachNoEmpty: () => {} }) };
  return { worksheet, GetName: () => name, GetRange: r => ({ range: { bbox: bbox(r) } }) };
}
const runPrelude = (api, checks) => new Function('Api', `${protectionRacePrelude(checks)}\nreturn { ok: true, ranTheRestOfTheBody: true };`)(api);

test('B1: unchanged protection -- the rest of the command body runs', () => {
  const ws = sheet('P', { sp: { sheet: false } }); // unprotected -- the sheet-level axis strips `cells` to null
  const api = { GetSheet: n => (n === 'P' ? ws : null) };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  // the SAME shape protectionReads (sheet-only) stores: { protected, forbids } -- no `cells` field at all
  const out = runPrelude(api, [{ sheet: 'P', expected: { protected: false, forbids: null } }]);
  assert.deepEqual(out, { ok: true, ranTheRestOfTheBody: true });
});

test('B2: protection CHANGED between planning and the mutation -- aborts BEFORE the rest of the body, named', () => {
  const ws = sheet('P', { sp: { sheet: false } }); // the PLANNING-time read saw: unprotected
  const api = { GetSheet: n => (n === 'P' ? ws : null) };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const expected = { protected: false, forbids: null }; // what classify() saw at planning
  ws.worksheet.sheetProtection = { sheet: true, objects: false, scenarios: false, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: false, insertHyperlinks: false, deleteColumns: false, deleteRows: false, selectLockedCells: false, sort: false, autoFilter: false, pivotTables: false, selectUnlockedCells: false }; // NOW protected -- the race
  const out = runPrelude(api, [{ sheet: 'P', expected }]);
  assert.deepEqual(out, { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: 'P', range: null });
});

test('B3: a RANGE-level check compares cells.unlocked too, and matches the exact args the planning read used', () => {
  const ws = sheet('P', { sp: { sheet: true, objects: false, scenarios: false, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: false, insertHyperlinks: false, deleteColumns: false, deleteRows: false, selectLockedCells: false, sort: false, autoFilter: false, pivotTables: false, selectUnlockedCells: false } });
  ws.worksheet.getRange3 = () => ({ _foreachNoEmpty: fn => fn({ nRow: 0, nCol: 0, xfs: { locked: false, hidden: null } }) }); // A1 unlocked
  const api = { GetSheet: n => (n === 'P' ? ws : null) };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const expected = { protected: true, forbids: Object.fromEntries(['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'].map(f => [f, false])), cells: { range: 'A1', nonEmpty: 1, unlocked: ['A1'], hidden: [] } };
  const pass = runPrelude(api, [{ sheet: 'P', range: 'A1', expected }]);
  assert.deepEqual(pass, { ok: true, ranTheRestOfTheBody: true });
  // now the SAME cell becomes locked between planning and the mutation
  ws.worksheet.getRange3 = () => ({ _foreachNoEmpty: fn => fn({ nRow: 0, nCol: 0, xfs: { locked: true, hidden: null } }) });
  const fail = runPrelude(api, [{ sheet: 'P', range: 'A1', expected }]);
  assert.deepEqual(fail, { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: 'P', range: 'A1' });
});

test('B4: multiple checks (move_range -- source AND destination) -- the FIRST mismatch wins, and it short-circuits', () => {
  const p = sheet('P', { sp: { sheet: false } });
  const q = sheet('Q', { sp: { sheet: true, objects: false, scenarios: false, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: false, insertHyperlinks: false, deleteColumns: false, deleteRows: false, selectLockedCells: false, sort: false, autoFilter: false, pivotTables: false, selectUnlockedCells: false } }); // Q is now protected, planning saw it unprotected
  const api = { GetSheet: n => ({ P: p, Q: q }[n] || null) };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const out = runPrelude(api, [{ sheet: 'P', expected: { protected: false, forbids: null } }, { sheet: 'Q', expected: { protected: false, forbids: null } }]);
  assert.deepEqual(out, { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: 'Q', range: null });
});

// ---- Part C: the product WRITE tool, end to end, over the fake driver ---------------------------------------------

const book = (extra = {}) => [{ name: 'S', cells: { A1: { v: 1 }, B1: { v: 2 } }, ...extra }];
const run = async (steps, store, opts = {}) => { const h = harness({ store, ...opts }); const r = await h.tool.handler(req(steps)); return { r, h }; };

test('C1: write_range into a LOCKED cell on a protected sheet -- refused, ZERO dispatch, the sheet untouched', async () => {
  const store = createStore(book({ protection: { unlocked: [] } })); // A1/B1: neither is named -> both default-locked
  const { r, h } = await run([{ intent: 'write_range', id: 's', sheet: 'S', range: 'A1', values: [['x']] }], store);
  assert.equal(r.ok, false, JSON.stringify(r).slice(0, 300));
  assert.match(JSON.stringify(r), /xlsx-protected-locked-cells:A1/);
  assert.equal(h.calls().dispatch, 0);
  assert.equal(store.durable[0].cells.A1.v, 1, 'the sheet was never touched');
});

test('C2: write_range into an UNLOCKED cell on a protected sheet -- proceeds', async () => {
  const store = createStore(book({ protection: { unlocked: ['A1'] } })); // A1 is explicitly unlocked; B1 stays locked
  const { r } = await run([{ intent: 'write_range', id: 's', sheet: 'S', range: 'A1', values: [['x']] }], store);
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
  assert.equal(store.durable[0].cells.A1.v, 'x');
});

test('C3: format_range -- forbids.formatCells true refuses; false (the default) proceeds', async () => {
  const forbid = createStore(book({ protection: { forbids: { formatCells: true } } }));
  const { r: refused } = await run([{ intent: 'format_range', id: 's', sheet: 'S', range: 'A1', format: { fill: '#FFFF00' } }], forbid);
  assert.equal(refused.ok, false);
  assert.match(JSON.stringify(refused), /xlsx-protected-sheet-forbids:formatCells/);
  const allow = createStore(book({ protection: {} }));
  const { r: ok } = await run([{ intent: 'format_range', id: 's', sheet: 'S', range: 'A1', format: { fill: '#FFFF00' } }], allow);
  assert.equal(ok.ok, true, JSON.stringify(ok).slice(0, 300));
});

test('C4: an unmeasured intent (merge_range) on a protected sheet refuses, named with the intent', async () => {
  const store = createStore(book({ protection: {} }));
  const { r, h } = await run([{ intent: 'merge_range', id: 's', sheet: 'S', range: 'A1:B1' }], store);
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r), /xlsx-protected-sheet-intent-not-measured:merge_range/);
  assert.equal(h.calls().dispatch, 0);
});

test('C5 (MI A KESZ ii): a multi-step task whose SECOND step is protected-refused -- the FIRST step never dispatches either', async () => {
  const store = createStore(book({ protection: { unlocked: [] } })); // A1 locked -- step 2 refuses
  const { r, h } = await run([
    { intent: 'write_range', id: 's1', sheet: 'S', range: 'B1', values: [['first']] },
    { intent: 'write_range', id: 's2', sheet: 'S', range: 'A1', values: [['second']] },
  ], store);
  assert.equal(r.ok, false, JSON.stringify(r).slice(0, 400));
  assert.equal(r.outcome, 'preflight-fail-closed');
  assert.equal(h.calls().dispatch, 0, 'ZERO dispatch -- not even the first, unprotected-looking step');
  assert.equal(store.durable[0].cells.B1.v, 2, 'the first step\'s write ("first") never landed either');
});

test('C6 (MI A KESZ iii): the protection axis itself unreadable -- the whole task fails closed, named, zero dispatch', async () => {
  const store = createStore(book({})); // no `protection` opt-in request here, instead: force the axis to fail
  const { r, h } = await run([{ intent: 'write_range', id: 's', sheet: 'S', range: 'A1', values: [['x']] }], store, { script: { protectionUnread: ['S'] } });
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r), /xlsx-protection-unobservable:/);
  assert.equal(h.calls().dispatch, 0);
});

// ---- Part D: the race prelude composes with EVERY bridge-write prelude dispatch() may put in front of it -----------
// (a maintainer review of the U4b patch: protectionRacePrelude carries its own BRIDGE_CLIENT_SRC, and every bridge-write
// prelude carries one too -- two top-level `const __eoBridgeRead` in ONE callCommand body is a SyntaxError, i.e.
// every b05/b10/b03/B22/B23/B24 write would fail, on an unprotected sheet as well, since a raceCheck rides there too)

test('D1: protectionRacePrelude + each bridge-write prelude still compiles as ONE callCommand body', () => {
  const d = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  const { BRIDGE_WRITES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
  const race = protectionRacePrelude([{ sheet: 'S', expected: { protected: false, forbids: null } }]);
  const anyCap = Object.values(BRIDGE_WRITES)[0];
  const preludes = {
    b05: d.b05BridgeWritePrelude(), b10: d.b10BridgeWritePrelude(), b03: d.b03BridgeWritePrelude(),
    chartStyle: d.chartStyleBridgeWritePrelude(anyCap), chartDelete: d.chartDeleteBridgeWritePrelude(), seriesStyle: d.seriesStyleBridgeWritePrelude(anyCap),
  };
  for (const [name, pre] of Object.entries(preludes)) {
    // the SAME order dispatch() builds: the bridge-write prelude, then (raceChecks ? race : '') + command.body
    assert.doesNotThrow(() => new Function(pre + race + 'return 1;'), `${name} + race prelude`);
  }
});

// ---- Part E: GH #338-#342 remainder (W2) -- the workbook-wide structure lock -----------------------------------

const V24 = fs.readFileSync(path.join(DIR, 'rev24-candidate', 'bridge-block.js'), 'utf8');
const WORKBOOK_OPS = {
  create_sheet: { intent: 'create_sheet', selector: { name: 'New' } },
  rename_sheet: { intent: 'rename_sheet', selector: { sheet: 'S', name: 'Renamed' } },
  delete_sheet: { intent: 'delete_sheet', selector: { sheet: 'S' } },
  move_sheet: { intent: 'move_sheet', selector: { sheet: 'S', reference: 'T', where: 'before' } },
  copy_sheet: { intent: 'copy_sheet', selector: { sheet: 'S', name: 'Copy' } },
};

test('E1: all five WORKBOOK_STRUCTURE_INTENTS, workbook locked -- refused PRE_DISPATCH, named, no raceCheck', () => {
  const primary = primaryOf({ workbook: { lockStructure: true } });
  for (const [intent, op] of Object.entries(WORKBOOK_OPS)) {
    const g = protectionGate.gate(primary, op);
    assert.equal(g.refusal.reason, 'xlsx-protected-workbook-locked-structure', intent);
    assert.deepEqual(g.raceChecks, [], intent);
  }
});

test('E2: all five, workbook UNLOCKED (and the sheet-level axis also unprotected) -- proceeds, carrying the workbook raceCheck', () => {
  const wbRead = { lockStructure: false };
  const primary = primaryOf({ workbook: wbRead, sheetLevel: { S: unprotected } });
  for (const [intent, op] of Object.entries(WORKBOOK_OPS)) {
    const g = protectionGate.gate(primary, op);
    assert.equal(g.refusal, null, intent);
    const wbCheck = g.raceChecks.find(c => c.workbook === true);
    assert.deepEqual(wbCheck, { workbook: true, expected: wbRead }, intent);
  }
  // create_sheet names no EXISTING sheet -- gateSheetLevel's own early exit contributes [], so the
  // workbook's raceCheck is the ONLY one; the defect this guards against (blocker 2.i) is it being lost.
  assert.deepEqual(protectionGate.gate(primary, WORKBOOK_OPS.create_sheet).raceChecks, [{ workbook: true, expected: wbRead }]);
  // the other four ALSO carry the sheet-level axis's own raceCheck (gateSheetLevel's unmeasured(), merged, not replaced).
  for (const intent of ['rename_sheet', 'delete_sheet', 'move_sheet', 'copy_sheet']) {
    const g = protectionGate.gate(primary, WORKBOOK_OPS[intent]);
    assert.equal(g.raceChecks.length, 2, intent);
  }
});

test('E3: CAPABILITY_MISSING (pre-/24 runtime) -- deliberate pass-through, unchanged from today, no raceCheck of its own', () => {
  assert.deepEqual(protectionGate.workbookGate({ workbookProtectionRead: 'UNOBSERVABLE:CAPABILITY_MISSING' }), { refusal: null, raceChecks: [] });
  // and the full gate() composition still falls through to the per-sheet axis unchanged
  const primary = primaryOf({ workbook: 'UNOBSERVABLE:CAPABILITY_MISSING', sheetLevel: { S: protectedWith() } });
  const g = protectionGate.gate(primary, WORKBOOK_OPS.rename_sheet);
  assert.equal(g.refusal.reason, 'xlsx-protected-sheet-intent-not-measured:rename_sheet');
});

test('E4: every OTHER unobservable workbook reason, and a simply-absent read -- fails closed, named, distinctly from CAPABILITY_MISSING', () => {
  assert.equal(protectionGate.workbookGate({ workbookProtectionRead: 'UNOBSERVABLE:protection-flag-not-boolean' }).refusal.reason, 'xlsx-protection-unobservable:workbook-protection-flag-not-boolean');
  assert.equal(protectionGate.workbookGate({ workbookProtectionRead: undefined }).refusal.reason, 'xlsx-protection-unobservable:workbook-not-read');
  assert.equal(protectionGate.workbookGate({}).refusal.reason, 'xlsx-protection-unobservable:workbook-not-read');
  // and fully composed: FAIL_CLOSED regardless of the per-sheet axis's own state
  const primary = primaryOf({ workbook: 'UNOBSERVABLE:model-missing', sheetLevel: { S: unprotected } });
  assert.equal(protectionGate.gate(primary, WORKBOOK_OPS.delete_sheet).refusal.reason, 'xlsx-protection-unobservable:workbook-model-missing');
});

/**
 * E4b (internal tracker review, #ref, blocker 1): `raw === null` and a MALFORMED object read -- fixed here. `null` is
 * the LIVE shape a missing read normalizes to (inventory.cjs:141), not just the theoretical `undefined` this
 * gate originally handled -- `raw.lockStructure` on `null` THREW a TypeError instead of returning a named
 * refusal. And `{}` / `{ lockStructure: null }` / a non-boolean `lockStructure` previously read as "not
 * locked" (fail-OPEN, `!== true`), the exact hazard W3's own gate() (families/workbook-protection-write.cjs)
 * already refuses as `malformed-read` for the SAME capability's read -- W2 must match it, not fall open.
 */
test('E4b: a null read (the LIVE missing-read shape) and a malformed object -- fail closed, named, NO exception, NO silent pass-through', () => {
  assert.equal(protectionGate.workbookGate({ workbookProtectionRead: null }).refusal.reason, 'xlsx-protection-unobservable:workbook-not-read');
  for (const malformed of [{}, { lockStructure: null }, { lockStructure: 'true' }, { lockStructure: 1 }]) {
    const g = protectionGate.workbookGate({ workbookProtectionRead: malformed });
    assert.equal(g.refusal.reason, 'xlsx-protection-unobservable:workbook-malformed-read', JSON.stringify(malformed));
    assert.deepEqual(g.raceChecks, [], JSON.stringify(malformed));
  }
  // composed through gate(): the same five intents, zero of them silently pass a malformed read
  const primary = primaryOf({ workbook: { lockStructure: 'true' } });
  for (const [intent, op] of Object.entries(WORKBOOK_OPS)) {
    assert.equal(protectionGate.gate(primary, op).refusal.reason, 'xlsx-protection-unobservable:workbook-malformed-read', intent);
  }
});

test('E4c NEG. KONTROLL: reverting to the ORIGINAL `raw.lockStructure !== true` check reopens the fail-open hole E4b closes', () => {
  // the exact pre-fix line (`!== true` reads anything-not-true as "proceed"), run against E4b's own fixtures
  const buggyGate = raw => (raw === undefined ? { refusal: { reason: 'xlsx-protection-unobservable:workbook-not-read' } } : (raw.lockStructure !== true ? { refusal: null, raceChecks: [{ workbook: true, expected: raw }] } : { refusal: { reason: 'xlsx-protected-workbook-locked-structure' } }));
  assert.throws(() => buggyGate(null), TypeError, 'the bug: null throws instead of refusing, named');
  assert.equal(buggyGate({}).refusal, null, 'the bug: a malformed {} reads as "proceed", fail-OPEN');
  assert.equal(buggyGate({ lockStructure: 'true' }).refusal, null, 'the bug: a non-boolean truthy-looking string also reads as "proceed"');
  // the real, fixed workbookGate refuses all three, named, no throw
  assert.doesNotThrow(() => protectionGate.workbookGate({ workbookProtectionRead: null }));
  assert.equal(protectionGate.workbookGate({ workbookProtectionRead: null }).refusal.reason, 'xlsx-protection-unobservable:workbook-not-read');
  assert.equal(protectionGate.workbookGate({ workbookProtectionRead: {} }).refusal.reason, 'xlsx-protection-unobservable:workbook-malformed-read');
  assert.equal(protectionGate.workbookGate({ workbookProtectionRead: { lockStructure: 'true' } }).refusal.reason, 'xlsx-protection-unobservable:workbook-malformed-read');
});

// ---- Part E5: protectionRacePrelude's new workbook branch, against the REAL /24 candidate block -----------------

function workbookEngine(lockStructure) {
  const wb = { getWorkbookProtection(t) { if (t !== 0) throw new Error('bad type'); return lockStructure; } };
  const Asc = { editor: { wbModel: wb }, c_oAscWorkbookProtectType: { lockStructure: 0, lockWindows: 1, lockRevisions: 2 }, CWorkbookProtection: function () {} };
  return { Api: { GetSheet: () => null }, Asc, window: { Asc }, AscCommon: { History: { Is_On: () => true } } };
}
const runPrelude24 = (e, checks) => {
  new Function('Api', 'Asc', 'window', 'AscCommon', V24.replace(/\}\s*,\s*$/, '}'))(e.Api, e.Asc, e.window, e.AscCommon);
  return new Function('Api', `${protectionRacePrelude(checks)}\nreturn { ok: true, ranTheRestOfTheBody: true };`)(e.Api);
};

test('E5a: workbook unchanged (still unlocked) -- the rest of the command body runs', () => {
  const out = runPrelude24(workbookEngine(false), [{ workbook: true, expected: { lockStructure: false } }]);
  assert.deepEqual(out, { ok: true, ranTheRestOfTheBody: true });
});

test('E5b: a co-editor locked the workbook between planning and the mutation -- aborts BEFORE the rest of the body, named, sheet: null', () => {
  // the PLANNING-time read saw: unlocked; a second session's protect_workbook ran in between -- NOW locked
  const out = runPrelude24(workbookEngine(true), [{ workbook: true, expected: { lockStructure: false } }]);
  assert.deepEqual(out, { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: null, range: null });
});

test('E5c: a workbook check and a sheet check in the SAME prelude -- either mismatch aborts (the create_sheet-under-race shape)', () => {
  const race = protectionRacePrelude([{ workbook: true, expected: { lockStructure: false } }, { sheet: 'S', expected: { protected: false, forbids: null } }]);
  assert.doesNotThrow(() => new Function(race + 'return 1;'), 'compiles standalone, no bridge-write prelude needed');
});

/**
 * E5d (internal tracker review, #ref, blocker 2): the FRESH re-read itself comes back malformed (bridge ok:true, a
 * non-boolean `lockStructure`) -- race-FAILS, named, same as any other mismatch; never a silent pass-through.
 * The REAL /24 handler's own `typeof v !== "boolean"` guard (handler-wb-protection-read-v24.txt) means this
 * exact bridge-level shape cannot occur through V24 today -- so this is tested against a minimal stub bridge
 * satisfying only __eoBridgeRead's own manifest/capability contract, proving the PRELUDE's OWN defense, not
 * relying on the handler upstream to always keep doing so.
 */
test('E5d: the FRESH re-read itself is malformed (ok:true, non-boolean lockStructure) -- race-fails, named, never silently accepted', () => {
  const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
  const { ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');
  const WB_R = BRIDGE_CAPABILITIES.WORKBOOK_PROTECTION_READ;
  const rt = ACCEPTED_RUNTIMES.xlsx[ACCEPTED_RUNTIMES.xlsx.length - 1]; // the /24 entry
  const manifest = { bridgeVersion: '1.0.0', documentServerVersion: rt.documentServerVersion, documentServerBuild: 'x', patchVersion: rt.patchVersion, capabilities: [{ id: WB_R.id, version: WB_R.version, kind: 'READ', format: 'XLSX', mutation: 'NO' }] };
  const bridge = { manifest: () => manifest, read: id => (id === WB_R.id ? { ok: true, value: { lockStructure: 'not-a-boolean' } } : { ok: false, reason: 'unexpected' }) };
  const Api = { GetEuroOfficeBridge: () => bridge };
  const race = protectionRacePrelude([{ workbook: true, expected: { lockStructure: false } }]);
  const out = new Function('Api', `${race}\nreturn { ok: true, ranTheRestOfTheBody: true };`)(Api);
  assert.deepEqual(out, { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: null, range: null });
});

test('E5e NEG. KONTROLL: without the typeof-boolean check, E5d\'s exact malformed re-read would silently PASS through', () => {
  // the PRE-FIX line this review caught: `{ lockStructure: __r.value.lockStructure }` unconditionally, whatever
  // shape __r.value.lockStructure has -- run the SAME scenario E5d measures through that (buggy) generator.
  const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
  const { ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');
  const { BRIDGE_CLIENT_SRC } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
  const WB_R = BRIDGE_CAPABILITIES.WORKBOOK_PROTECTION_READ;
  const buggyRace = (checks) => `{${BRIDGE_CLIENT_SRC}
  const __eoProtChecks = ${JSON.stringify(checks)};
  let __eoProtFail = null;
  for (const __c of __eoProtChecks) {
    const __r = __eoBridgeRead(${JSON.stringify(WB_R.id)}, ${WB_R.version}, ${JSON.stringify(ACCEPTED_RUNTIMES.xlsx)}, {}, ${JSON.stringify(WB_R.valueType)});
    const __got = !__r.ok ? 'UNOBSERVABLE:' + __r.reason : { lockStructure: __r.value.lockStructure }; // the bug: no boolean check
    if (JSON.stringify(__got) !== JSON.stringify(__c.expected)) { __eoProtFail = { sheet: null, range: null }; break; }
  }
  if (__eoProtFail) return { ok: false, stage: 'protection-race', outcome: 'protection-changed-before-mutation', sheet: __eoProtFail.sheet, range: __eoProtFail.range };
}
`;
  const rt = ACCEPTED_RUNTIMES.xlsx[ACCEPTED_RUNTIMES.xlsx.length - 1];
  const manifest = { bridgeVersion: '1.0.0', documentServerVersion: rt.documentServerVersion, documentServerBuild: 'x', patchVersion: rt.patchVersion, capabilities: [{ id: WB_R.id, version: WB_R.version, kind: 'READ', format: 'XLSX', mutation: 'NO' }] };
  // chosen so the bug's own coincidental shape ({lockStructure: undefined} -> JSON '{}') matches `expected: {}`
  // -- a shape that cannot arise from the REAL (fixed) gate() anymore, but is exactly what the bug let through.
  const bridge = { manifest: () => manifest, read: id => (id === WB_R.id ? { ok: true, value: {} } : { ok: false, reason: 'unexpected' }) };
  const Api = { GetEuroOfficeBridge: () => bridge };
  const out = new Function('Api', `${buggyRace([{ workbook: true, expected: {} }])}\nreturn { ok: true, ranTheRestOfTheBody: true };`)(Api);
  assert.deepEqual(out, { ok: true, ranTheRestOfTheBody: true }, 'the bug: a malformed re-read silently passes, because it coincidentally JSON-matches a malformed expected');
  // the REAL (fixed) prelude, same malformed re-read -- fails closed regardless of what `expected` says
  const realOut = new Function('Api', `${protectionRacePrelude([{ workbook: true, expected: {} }])}\nreturn { ok: true, ranTheRestOfTheBody: true };`)(Api);
  assert.equal(realOut.ok, false, 'the real prelude never accepts a non-boolean lockStructure, even against a degenerate expected');
});

// ---- Part E6: the product WRITE tool, end to end, over the fake driver -- all five intents, both lock states ----

const sheetsBook = () => [{ name: 'S', cells: { A1: { v: 1 } } }, { name: 'T', cells: { A1: { v: 2 } } }];
const wbSteps = {
  create_sheet: { intent: 'create_sheet', id: 's', name: 'New' },
  rename_sheet: { intent: 'rename_sheet', id: 's', sheet: 'S', name: 'Renamed' },
  delete_sheet: { intent: 'delete_sheet', id: 's', sheet: 'S' },
  move_sheet: { intent: 'move_sheet', id: 's', sheet: 'S', after: 'T' },
  copy_sheet: { intent: 'copy_sheet', id: 's', sheet: 'S', name: 'Copy' },
};

test('E6a: all five intents, workbook LOCKED -- refused PRE_DISPATCH, named, ZERO dispatch, the document untouched', async () => {
  for (const [intent, step] of Object.entries(wbSteps)) {
    const store = createStore(sheetsBook());
    store.workbookLocked = true;
    const { r, h } = await run([step], store);
    assert.equal(r.ok, false, `${intent}: ${JSON.stringify(r).slice(0, 300)}`);
    assert.match(JSON.stringify(r), /xlsx-protected-workbook-locked-structure/, intent);
    assert.equal(h.calls().dispatch, 0, intent);
    assert.deepEqual(store.durable.map(s => s.name), ['S', 'T'], `${intent}: the sheet list never changed`);
  }
});

test('E6b: all five intents, workbook UNLOCKED -- proceeds exactly as before W2 (unaffected, APPLY)', async () => {
  for (const [intent, step] of Object.entries(wbSteps)) {
    const store = createStore(sheetsBook());
    store.workbookLocked = false;
    const { r } = await run([step], store);
    assert.equal(r.ok, true, `${intent}: ${JSON.stringify(r).slice(0, 300)}`);
  }
});

test('E6c: pre-/24 runtime shape (CAPABILITY_MISSING) -- unaffected, a locked-looking store still APPLIES', async () => {
  const store = createStore(sheetsBook());
  store.workbookLocked = true; // would refuse on /24 -- but the runtime cannot even observe the axis
  const { r } = await run([wbSteps.delete_sheet], store, { script: { workbookProtectionUnobservable: 'CAPABILITY_MISSING' } });
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
});

test('E6d: a DIFFERENT unobservable workbook reason -- fails closed, named, zero dispatch, even though the sheet axis itself is fine', async () => {
  const store = createStore(sheetsBook());
  const { r, h } = await run([wbSteps.rename_sheet], store, { script: { workbookProtectionUnobservable: 'model-missing' } });
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r), /xlsx-protection-unobservable:workbook-model-missing/);
  assert.equal(h.calls().dispatch, 0);
});

test('E6e (zero\'s review, #ref): a malformed workbook read (null / {} / non-boolean lockStructure) -- all five intents fail closed, named, ZERO dispatch, product tool end to end', async () => {
  for (const malformed of [null, {}, { lockStructure: null }, { lockStructure: 'true' }]) {
    for (const [intent, step] of Object.entries(wbSteps)) {
      const store = createStore(sheetsBook());
      const { r, h } = await run([step], store, { script: { workbookProtectionMalformed: malformed } });
      const label = `${intent} / ${JSON.stringify(malformed)}`;
      assert.equal(r.ok, false, `${label}: ${JSON.stringify(r).slice(0, 300)}`);
      assert.match(JSON.stringify(r), /xlsx-protection-unobservable:workbook-(not-read|malformed-read)/, label);
      assert.equal(h.calls().dispatch, 0, label);
      assert.deepEqual(store.durable.map(s => s.name), ['S', 'T'], `${label}: the sheet list never changed`);
    }
  }
});

// E6e's own mutation control lives at the gate level, where it is simplest and most robust: E4c already
// proves that reverting workbookGate's boolean-typeof check to the original `!== true` rule turns `{}` / `null`
// from a named refusal back into a silent, fail-OPEN "proceed" -- the SAME rule E6e's product-level run depends
// on. E5e proves the SAME mutation for protectionRacePrelude's own re-read. A third, end-to-end-level copy of
// the identical mutation would not catch anything these two do not already catch.

// ---- Part E7: two local NEGATIVE controls -- proving E1/E6a are not vacuously green -----------------------------
// (the same discipline as the fail-closed/race tests elsewhere in this file: a test that cannot go RED on a
// reintroduction of the exact defect it names is not a test of that defect.)

test('E7a NEG. KONTROLL: without workbookGate, E1 (locked refuses) goes RED -- the real gate() does not', () => {
  // the ORIGINAL blocker-1-shaped bug: compose only the per-sheet axis, as if workbookGate were never added
  const buggyGate = (primary, op) => protectionGate.gateSheetLevel(primary, op);
  const primary = primaryOf({ workbook: { lockStructure: true } });
  const buggy = buggyGate(primary, WORKBOOK_OPS.create_sheet);
  assert.equal(buggy.refusal, null, 'the bug: create_sheet sails through a locked workbook');
  const real = protectionGate.gate(primary, WORKBOOK_OPS.create_sheet);
  assert.equal(real.refusal.reason, 'xlsx-protected-workbook-locked-structure', 'the real gate() catches it');
});

test('E7b NEG. KONTROLL: overwriting raceChecks instead of merging loses the workbook check -- the real gate() keeps it', () => {
  // the ORIGINAL blocker-2.i-shaped bug: `{ refusal: null, raceChecks: inner.raceChecks }` (overwrite, not spread)
  const wbRead = { lockStructure: false };
  const primary = primaryOf({ workbook: wbRead });
  const wb = protectionGate.workbookGate(primary);
  const inner = protectionGate.gateSheetLevel(primary, WORKBOOK_OPS.create_sheet); // -> { refusal: null, raceChecks: [] }
  const buggy = { refusal: null, raceChecks: inner.raceChecks }; // overwrite: the workbook check is gone
  assert.deepEqual(buggy.raceChecks, [], 'the bug: create_sheet under an unlocked workbook carries NO race check at all');
  const real = protectionGate.gate(primary, WORKBOOK_OPS.create_sheet);
  assert.deepEqual(real.raceChecks, [{ workbook: true, expected: wbRead }], 'the real gate() keeps it, via spread-merge');
});
