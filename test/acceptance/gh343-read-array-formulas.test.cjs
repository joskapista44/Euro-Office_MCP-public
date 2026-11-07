'use strict';

/**
 * GH #343 (F01): read_array_formulas -- anchor / ref / formula / type of the array formulas of one range, from the PUBLIC
 * GetFormulaArray + GetFormula per cell (the LIVE-measured membership rule array-formula.cjs states) over the range plus a
 * one-cell ring. Offline: the observation is built the way the fake driver answers arrayCells (a member -> its array's
 * formula, a plain formula -> null, a value -> the value, an empty cell -> ''), so the membership semantics are the measured ones.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const R = require('../../src/server/read-tools.cjs');
const { projectArrayFormulasRead, readRequest, NOTE } = require('../../src/adapters/xlsx/families/array-formula-read.cjs');
const af = require('../../src/adapters/xlsx/families/array-formula.cjs');
const { key } = require('../../src/adapters/xlsx/inventory.cjs');
const { parseRange, cellsOf } = require('../../src/adapters/xlsx/a1.cjs');

const SHEET = 'Data';
/** an observation over `view` (the cells the request asks): arrays = [{ range, formula }], plain = { A1: '=1+1' }, values = { A1: 5 },
 * unobservable = [addr] -> the cell's GetFormulaArray read is a named gap */
function observe({ view, arrays = [], plain = {}, values = {}, unobservable = [] }) {
  const cells = {};
  const arrayReads = {};
  const memberOfArray = a => arrays.find(x => cellsOf(parseRange(x.range)).includes(a));
  for (const a of cellsOf(parseRange(view))) {
    const arr = memberOfArray(a);
    cells[key(SHEET, a)] = { v: values[a] !== undefined ? values[a] : '', f: arr ? arr.formula : (plain[a] || ''), count: 0, merge: null };
    arrayReads[`${SHEET}!${a}`] = unobservable.includes(a) ? 'UNOBSERVABLE:threw' : arr ? arr.formula : plain[a] ? null : values[a] !== undefined ? values[a] : '';
  }
  return { format: 'xlsx', complete: true, truncated: false, sheets: [{ name: SHEET }], cells, arrayReads };
}
const ask = (range, world) => {
  const a = { sheet: SHEET, range };
  const req = readRequest(a);
  return projectArrayFormulasRead(observe({ view: req.windows[0].range, ...world }), a);
};
const SUM = '=SUM(A1:A3*B1:B3)';

test('catalogue + validation + request: an ACCEPTED xlsx read taking { sheet, range }; the request asks the range PLUS its one-cell ring', () => {
  assert.equal(R.OPERATIONS.xlsx.read_array_formulas.status, R.STATUS.ACCEPTED);
  assert.equal(R.validateReadRequest('xlsx', { fileId: '1', operation: 'read_array_formulas', sheet: SHEET, range: 'b2:c3' }).ok, true);
  assert.equal(R.validateReadRequest('xlsx', { fileId: '1', operation: 'read_array_formulas', sheet: SHEET }).ok, false, 'range required');
  assert.equal(R.validateReadRequest('xlsx', { fileId: '1', operation: 'read_array_formulas', range: 'A1:B2' }).ok, false, 'sheet required');
  assert.equal(R.validateReadRequest('xlsx', { fileId: '1', operation: 'read_array_formulas', sheet: SHEET, range: 'A1:ZZ9999' }).code, 'LIMIT', 'the read ceiling applies');
  const req = R.xlsxRequest('read_array_formulas', { sheet: SHEET, range: 'B2:C3' });
  assert.deepEqual(req.windows, [{ sheet: SHEET, range: 'A1:D4' }]);
  assert.equal(req.arrayCells.length, 16);
  assert.deepEqual(R.xlsxRequest('read_array_formulas', { sheet: SHEET, range: 'A1:B2' }).windows, [{ sheet: SHEET, range: 'A1:C3' }], 'clipped at the sheet edge');
});

test('a column array and a 2x2 array inside the range: anchor / ref / formula / type, extent proven (row-then-column order)', () => {
  const out = ask('A1:F8', { arrays: [{ range: 'E2:F3', formula: '=A1:B2*2' }, { range: 'B4:B6', formula: SUM }] });
  assert.deepEqual(out.arrays, [
    { anchor: 'E2', ref: 'E2:F3', formula: '=A1:B2*2', type: 'array', extent: 'proven' },
    { anchor: 'B4', ref: 'B4:B6', formula: SUM, type: 'array', extent: 'proven' },
  ]);
  assert.equal(out.note, NOTE);
});

test('an array that fills the requested range exactly (it touches the edge but the ring is clean) is proven', () => {
  const out = ask('B2:B4', { arrays: [{ range: 'B2:B4', formula: SUM }] });
  assert.deepEqual(out.arrays, [{ anchor: 'B2', ref: 'B2:B4', formula: SUM, type: 'array', extent: 'proven' }]);
});

test('an array that crosses the requested range boundary: NAMED state, no ref invented, only the part inside is named', () => {
  const out = ask('A1:D5', { arrays: [{ range: 'B2:B9', formula: SUM }] });
  assert.deepEqual(out.arrays, [{ anchor: null, ref: null, formula: SUM, type: 'array', extent: 'crosses-the-requested-range', insideRange: 'B2:B5' }]);
});

test('two adjacent arrays with DIFFERENT formulas are two areas; with the SAME formula they are one block, and the note says so', () => {
  const diff = ask('A1:C8', { arrays: [{ range: 'B2:B3', formula: '=A1:A2' }, { range: 'B4:B5', formula: '=C1:C2' }] });
  assert.deepEqual(diff.arrays.map(x => x.ref), ['B2:B3', 'B4:B5']);
  const same = ask('A1:C8', { arrays: [{ range: 'B2:B3', formula: SUM }, { range: 'B4:B5', formula: SUM }] });
  assert.deepEqual(same.arrays.map(x => x.ref), ['B2:B5'], 'indistinguishable through the public API: one block');
  assert.match(same.note, /adjacent arrays with the same formula cannot be told apart/);
});

test('members of one formula that fill no rectangle (an L) -> not-a-rectangle, no ref', () => {
  const out = ask('A1:E5', { arrays: [{ range: 'B2:B4', formula: SUM }, { range: 'C4:C4', formula: SUM }] });
  assert.equal(out.arrays.length, 1);
  assert.equal(out.arrays[0].extent, 'not-a-rectangle');
  assert.equal(out.arrays[0].ref, null);
  assert.deepEqual(out.arrays[0].cells, ['B2', 'B3', 'B4', 'C4']);
});

test('plain formulas, values and empty cells are NOT arrays (GetFormulaArray null / the value / ""): an empty list, not a gap', () => {
  const out = ask('A1:C3', { plain: { B2: '=1+1', B3: '=B2*2' }, values: { A1: 5, C3: 'x' } });
  assert.deepEqual(out.arrays, []);
});

test('an unreadable cell (range or ring) makes the WHOLE answer a named gap -- no partial list, no guess', () => {
  const inRange = ask('B2:B3', { arrays: [{ range: 'B2:B3', formula: SUM }], unobservable: ['B3'] });
  assert.equal(inRange.arrays, null);
  assert.equal(inRange.reason, 'xlsx-array-formula-unobservable');
  assert.equal(inRange.cell, 'B3');
  const inRing = ask('B2:B3', { arrays: [{ range: 'B2:B3', formula: SUM }], unobservable: ['B4'] });
  assert.equal(inRing.arrays, null, 'the ring is part of the proof');
});

test('a missing sheet is a named error, like the other reads', () => {
  const inv = observe({ view: 'A1:C3' });
  assert.deepEqual(projectArrayFormulasRead(inv, { sheet: 'Nope', range: 'A1:B2' }), { sheet: 'Nope', range: 'A1:B2', error: 'SHEET_NOT_FOUND' });
});

test('guard symmetry: the READ and the write guard share ONE membership definition (a function, not a copy)', () => {
  const src = fs.readFileSync(path.join(__dirname, '..', '..', 'src', 'adapters', 'xlsx', 'families', 'array-formula-read.cjs'), 'utf8');
  assert.ok(/require\('\.\/array-formula\.cjs'\)/.test(src), 'the read imports the write family');
  assert.ok(!/startsWith\('='\)/.test(src), 'the read has no membership rule of its own');
  assert.equal(af.isMember('=A1', '=A1'), true);
  assert.equal(af.isMember('=A1', ''), false, 'a member needs its own GetFormula too');
  assert.equal(af.isMember(null, '=A1'), false);
  assert.equal(af.isMember(5, '=A1'), false);
  const inv = observe({ view: 'A1:B2', arrays: [{ range: 'A1:B2', formula: SUM }], plain: {} });
  for (const a of cellsOf(parseRange('A1:B2'))) {
    const m = af.memberOf(inv, SHEET, a);
    assert.equal(m.member, af.isMember(inv.arrayReads[`${SHEET}!${a}`], inv.cells[key(SHEET, a)].f), a);
  }
});
