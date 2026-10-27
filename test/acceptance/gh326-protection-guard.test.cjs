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
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const protectionGate = require('../../src/adapters/xlsx/protection-gate.cjs');
const { protectionRacePrelude } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { harness, req, createStore } = require('../fakes/write-wiring.cjs');

// ---- Part A: protection-gate.cjs, pure ------------------------------------------------------------------------

const primaryOf = ({ sheets = ['S'], range = {}, sheetLevel = {} } = {}) => ({
  sheets: sheets.map(name => ({ name })),
  protectionRangeReads: range,
  protectionReads: sheetLevel,
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
  const primary = { sheets: [{ name: 'S' }], protectionReads: { S: protectedWith() }, protectionRangeReads: {} };
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
