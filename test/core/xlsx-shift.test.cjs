'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { shiftFormula, referencesBlock, translateFormula } = require('../../src/adapters/xlsx/shift.cjs');

/**
 * The W5 reference-aware shifter. Every expectation is written by hand from Excel's rules (not derived from the
 * shifter's output); the audit X4 shapes -- LOG10, a string literal that looks like a reference, a reference to
 * ANOTHER sheet -- must stay untouched.
 */

const ins = (at, count, axis = 'rows', sheet = 'P') => ({ sheet, axis, at, count });
const del = (at, k, axis = 'rows', sheet = 'P') => ({ sheet, axis, at, count: -k });
const sh = (f, op, host = 'P') => shiftFormula(f, host, op);

test('insert rows: below the cut moves, above stays, a range across the cut grows, absolute moves too', () => {
  assert.equal(sh('=B5*2', ins(3, 2)).formula, '=B7*2');
  assert.equal(sh('=B1+1', ins(3, 2)).formula, '=B1+1');
  assert.equal(sh('=SUM(B1:B6)', ins(3, 2)).formula, '=SUM(B1:B8)');
  assert.equal(sh('=$B$5', ins(3, 2)).formula, '=$B$7');
  assert.equal(sh('=B3', ins(3, 2)).formula, '=B5', 'the row AT the cut moves');
});

test('the X4 shapes stay: LOG10, a string literal, another sheet, a longer identifier', () => {
  assert.equal(sh('=LOG10(B6)', ins(3, 2)).formula, '=LOG10(B8)');
  assert.equal(sh('="B5"&B2', ins(3, 2)).formula, '="B5"&B2');
  assert.equal(sh('=Q!B5+B5', ins(3, 2)).formula, '=Q!B5+B7');
  assert.equal(sh("='My Q'!B5+B5", ins(3, 2)).formula, "='My Q'!B5+B7");
  assert.equal(sh('=ABC12X+B5', ins(3, 2)).formula, '=ABC12X+B7', 'not a reference: part of an identifier');
});

test('a formula on ANOTHER sheet follows only its references to the shifted sheet', () => {
  assert.equal(sh('=P!B5', ins(3, 2), 'Q').formula, '=P!B7');
  assert.equal(sh('=SUM(P!B2:B5)', ins(3, 2), 'Q').formula, '=SUM(P!B2:B7)');
  assert.equal(sh('=B5', ins(3, 2), 'Q').formula, '=B5', 'a bare ref on Q is Q\'s own cell');
  assert.equal(sh("='P'!B5", ins(3, 2), 'Q').formula, "='P'!B7");
});

test('delete rows: into the block -> #REF!, after it moves up, a range overlapping shrinks, entirely inside -> #REF!', () => {
  assert.deepEqual(sh('=B3*2', del(3, 2)), { formula: '=#REF!*2', broken: true });
  assert.deepEqual(sh('=B6', del(3, 2)), { formula: '=B4', broken: false });
  assert.deepEqual(sh('=SUM(B1:B6)', del(3, 2)), { formula: '=SUM(B1:B4)', broken: false });
  assert.deepEqual(sh('=SUM(B3:B4)', del(3, 2)), { formula: '=SUM(#REF!)', broken: true });
  assert.deepEqual(sh('=SUM(B4:B9)', del(3, 2)), { formula: '=SUM(B3:B7)', broken: false });
  assert.deepEqual(sh('=B2', del(3, 2)), { formula: '=B2', broken: false });
});

test('columns: insert / delete on the column axis; row-only and column-only whole references', () => {
  assert.equal(sh('=C1+A1', ins(2, 1, 'cols')).formula, '=D1+A1');
  assert.equal(sh('=SUM(A1:D1)', ins(2, 1, 'cols')).formula, '=SUM(A1:E1)');
  assert.deepEqual(sh('=B1', del(2, 1, 'cols')), { formula: '=#REF!', broken: true });
  assert.equal(sh('=SUM(3:4)', ins(3, 2)).formula, '=SUM(5:6)');
  assert.equal(sh('=SUM(B:C)', ins(2, 1, 'cols')).formula, '=SUM(C:D)');
  assert.equal(sh('=SUM(B:C)', ins(3, 2)).formula, '=SUM(B:C)', 'a whole-column ref is untouched by a row insert');
  assert.equal(sh('=Z1', ins(26, 1, 'cols')).formula, '=AA1', 'Z -> AA across the letter boundary');
});

test('referencesBlock: a direct ref, a range overlap, another sheet, a literal', () => {
  const block = { sheet: 'P', r1: 3, r2: 4, c1: 1, c2: 16384 };
  assert.equal(referencesBlock('=B3', 'P', block), true);
  assert.equal(referencesBlock('=SUM(B1:B6)', 'P', block), true);
  assert.equal(referencesBlock('=B5', 'P', block), false);
  assert.equal(referencesBlock('=Q!B3', 'P', block), false);
  assert.equal(referencesBlock('=P!B3', 'Q', block), true);
  assert.equal(referencesBlock('="B3"', 'P', block), false);
});

test('translateFormula (copy): relative parts move, absolute stay, out of the sheet -> #REF!', () => {
  assert.deepEqual(translateFormula('=B1*2', 4, 3), { formula: '=E5*2', broken: false });
  assert.deepEqual(translateFormula('=$B$1+A9', 4, 3), { formula: '=$B$1+D13', broken: false });
  assert.deepEqual(translateFormula('=B$1+$A2', 4, 3), { formula: '=E$1+$A6', broken: false });
  assert.deepEqual(translateFormula('=A1', -1, 0), { formula: '=#REF!', broken: true });
  assert.deepEqual(translateFormula('="A1"&LOG10(A1)', 1, 1), { formula: '="A1"&LOG10(B2)', broken: false });
  assert.deepEqual(translateFormula('=SUM(Q!A1:A2)', 1, 0), { formula: '=SUM(Q!A2:A3)', broken: false });
});
