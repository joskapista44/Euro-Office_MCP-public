'use strict';

/**
 * GH #327 (U5, owner decision owner channel answer to 41211): read_range on a PROTECTED sheet must not hand back the
 * TEXT of a hidden formula; the cell's VALUE is unaffected on every branch (the owner's own condition -- the secret itself
 * stays hidden, the computed answer does not).
 *
 * Part A (driver): xlsx-driver.cjs's protectionRangeReads axis -- sheet.protection.read/v1 with a `range`, through
 * the SAME bridgeAxis convention cfReads/validationReads already use, against the REAL bridge blocks (not an
 * imitation): the shipped /16 (CAPABILITY_MISSING) and the /17 CANDIDATE (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/rev17-candidate/bridge-block.js), loaded the SAME way test/acceptance/
 * ds935-bridge17-protection-read.test.cjs already does (the range walk needs that file's richer fake worksheet,
 * not the plain one gh326-protection-read-client-wiring.test.cjs uses).
 *
 * Part B (projection): read-tools.cjs's readRange() -- protected+hidden -> REDACTED (value kept); protected,
 * not hidden -> formula kept; unprotected (hidden flag irrelevant, the guard decides what it means) -> formula
 * kept; the protection axis itself unreadable (a /16 runtime, or any other gap) -> the WHOLE window's formula
 * fields fail closed to 'UNOBSERVABLE:protection-unread:<reason>', value kept -- never guessed in the secret's favour.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { readRange } = require('../../src/server/read-tools.cjs');
const { key: cellKey } = require('../../src/adapters/xlsx/inventory.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');

// ---- Part A: the driver axis, against the REAL blocks ---------------------------------------------------------

/** The SAME fake worksheet shape ds935-bridge17-protection-read.test.cjs uses (sheetProtection + getRange3 walk),
 * plus GetName/GetUsedRange (the driver's own per-sheet existence check and sheet-list builder need them). */
const a1 = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return { r: Number(m[2]) - 1, c: c - 1 }; };
const bbox = range => { const [p, q = p] = range.split(':'); const x = a1(p); const y = a1(q); return { r1: x.r, c1: x.c, r2: y.r, c2: y.c }; };
function sheet(name, { sp = null, cells = {} } = {}) {
  const list = Object.entries(cells).map(([ref, xfs]) => { const p = a1(ref); return { nRow: p.r, nCol: p.c, xfs }; });
  const worksheet = { sheetProtection: sp, getRange3: (r1, c1, r2, c2) => ({ _foreachNoEmpty: fn => { for (const cell of list) if (cell.nRow >= r1 && cell.nRow <= r2 && cell.nCol >= c1 && cell.nCol <= c2) fn(cell); } }) };
  return { worksheet, GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, GetRange: r => ({ range: { bbox: bbox(r) } }) };
}
const FLAGS = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
const protectedSp = (over = {}) => Object.assign({ sheet: true }, Object.fromEntries(FLAGS.map(f => [f, false])), over);

const runRanges = (api, protectionRanges) => new Function('Api', readQuery([], { protectionRanges }))(api).protectionRangeReads;

test('A /16 runtime does not advertise the capability at all, even with a range asked -- UNOBSERVABLE:CAPABILITY_MISSING', () => {
  const ws = sheet('P', { sp: protectedSp(), cells: { A1: { locked: false, hidden: true } } });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'P' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  const runtimeId = installBridge(api);
  assert.equal(runtimeId, '9.3.5|euro-office-bridge-v1/9.3.5.0/16', 'calibration: the shipped block really is /16');
  assert.deepEqual(runRanges(api, [{ sheet: 'P', range: 'A1:B2' }]), { 'P!A1:B2': 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('A /17 runtime reads protected + forbids + the range\'s unlocked/hidden cells, keyed sheet!range', () => {
  const ws = sheet('P', { sp: protectedSp({ insertRows: true }), cells: { A1: { locked: false, hidden: true }, B2: { locked: false, hidden: null } } });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'P' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const out = runRanges(api, [{ sheet: 'P', range: 'A1:B2' }]);
  assert.deepEqual(out, { 'P!A1:B2': { protected: true, forbids: { ...Object.fromEntries(FLAGS.map(f => [f, false])), insertRows: true }, cells: { range: 'A1:B2', nonEmpty: 2, unlocked: ['A1', 'B2'], hidden: ['A1'] } } });
});

test('not asked -- protectionRangeReads is generated empty, no bridge call at all (same convention as cfReads/validationReads)', () => {
  assert.doesNotMatch(readQuery(), /protectionRangeAsked/, 'the per-window loop (and its bridgeAxis call) is generated only inside the conditional block');
  const ws = sheet('P', { sp: protectedSp() });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'P' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  assert.deepEqual(new Function('Api', readQuery())(api).protectionRangeReads, {});
});

// ---- Part B: read-tools.cjs's redaction projection -----------------------------------------------------------

const inv = (cells, protectionRangeReads) => ({ sheets: [{ name: 'P' }], cells, protectionRangeReads });
const c = (address, f) => [cellKey('P', address), { sheet: 'P', a: address, v: 'whatever', f, count: 1, merge: null }];
const a = { sheet: 'P', range: 'A1:B1' };

test('B1: protected + the cell IS in cells.hidden -> the formula field is REDACTED, the value is untouched', () => {
  const i = inv(Object.fromEntries([c('A1', '=SECRET()')]), { 'P!A1:B1': { protected: true, forbids: {}, cells: { range: 'A1:B1', nonEmpty: 1, unlocked: [], hidden: ['A1'] } } });
  const out = readRange(i, a);
  assert.deepEqual(out.cells[0], { address: 'A1', v: 'whatever', f: 'REDACTED:hidden-formula-on-protected-sheet', merge: null, type: 'formula' });
  assert.equal(out.cells[1].read, false, 'B1 was never read -- untouched, not a guessed empty cell');
});

test('B2: protected, but the cell is NOT in cells.hidden -> the formula text is unchanged', () => {
  const i = inv(Object.fromEntries([c('A1', '=1+1')]), { 'P!A1:B1': { protected: true, forbids: {}, cells: { range: 'A1:B1', nonEmpty: 1, unlocked: [], hidden: [] } } });
  assert.equal(readRange(i, a).cells[0].f, '=1+1');
});

test('B3: not protected -- the range walk also ran and the cell IS in hidden, but the guard decides what it means: unchanged', () => {
  const i = inv(Object.fromEntries([c('A1', '=1+1')]), { 'P!A1:B1': { protected: false, forbids: null, cells: { range: 'A1:B1', nonEmpty: 1, unlocked: ['A1'], hidden: ['A1'] } } });
  assert.equal(readRange(i, a).cells[0].f, '=1+1');
});

test('B4: the protection axis itself degraded (e.g. a /16 runtime) -- every formula field in the window fails closed, never left readable on a guess', () => {
  const i = inv(Object.fromEntries([c('A1', '=SECRET()')]), { 'P!A1:B1': 'UNOBSERVABLE:CAPABILITY_MISSING' });
  const out = readRange(i, a);
  assert.equal(out.cells[0].f, 'UNOBSERVABLE:protection-unread:CAPABILITY_MISSING');
  assert.equal(out.cells[0].v, 'whatever', 'the VALUE is never touched, on any branch');
});

test('B5: not asked at all (the key is missing) -- the SAME fail-closed degrade, named not-read', () => {
  const i = inv(Object.fromEntries([c('A1', '=SECRET()')]), {});
  assert.equal(readRange(i, a).cells[0].f, 'UNOBSERVABLE:protection-unread:not-read');
});

test('B6: a non-formula cell is never touched by any branch, protected or not, degraded or not', () => {
  const plain = Object.fromEntries([[cellKey('P', 'A1'), { sheet: 'P', a: 'A1', v: '42', f: '', count: 1, merge: null }]]);
  for (const prot of [{ protected: true, forbids: {}, cells: { range: 'A1:B1', nonEmpty: 1, unlocked: [], hidden: ['A1'] } }, 'UNOBSERVABLE:CAPABILITY_MISSING', undefined]) {
    const i = inv(plain, prot === undefined ? {} : { 'P!A1:B1': prot });
    assert.equal(readRange(i, a).cells[0].f, '', JSON.stringify(prot));
  }
});
