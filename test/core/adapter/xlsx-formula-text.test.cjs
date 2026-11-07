'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const f = require('../../../src/adapters/xlsx/formula.cjs');

/**
 * Formula identity on the editor's MEASURED re-rendering (wave-2 probe set a). The measured pairs are
 * the positive cases; the literal / quoted-name cases pin that the canonical form never changes their
 * case or spaces (owner rule) -- whatever set b measures, a mismatch there must stay a mismatch.
 */

const MEASURED_SET_A = [
  ['=J1*3', '= J1*3'], ['=J1*2', '= J1*2'], ['=j1*4', '= J1*4'], ['=J1 * 5', '= J1*5'],
  ['=1/0', '= 1/0'], ['=NOSUCHFN(1)', '= NOSUCHFN(1)'], ['=Monthly_Plan!B16', '= Monthly_Plan!B16'],
];

const MEASURED_SET_B = [
  ['=J1&" a  B "', '= J1&" a  B "'], ["='Monthly_Plan'!B16", '= Monthly_Plan!B16'], ['=sum(J1, 2)', '= SUM(J1,2)'],
  ['=IF(J1>2,"Yes  X","no")', '= IF(J1>2,"Yes  X","no")'], ['=Monthly_Plan!b16', '= Monthly_Plan!B16'], ['="=Text  q"', '= "=Text  q"'],
];

test('every measured set-b pair is the same formula; the silently lost one is not', () => {
  for (const [sub, read] of MEASURED_SET_B) assert.equal(f.sameFormula(sub, read), true, `${sub} vs ${read}`);
  assert.equal(f.sameFormula('=monthly_plan!B16', ''), false, 'the formula the editor dropped');
});

test('a name that NEEDS quotes stays quoted and verbatim; a cell-shaped name is not bare', () => {
  assert.equal(f.sameFormula("='My Sheet'!B1", '= My Sheet!B1'), false);
  assert.equal(f.bareSheetName('Monthly_Plan'), true);
  assert.equal(f.bareSheetName('A1'), false);
  assert.equal(f.bareSheetName('My Sheet'), false);
});

test('every measured set-a pair is the same formula', () => {
  for (const [sub, read] of MEASURED_SET_A) assert.equal(f.sameFormula(sub, read), true, `${sub} vs ${read}`);
});

test('a different formula is not the same (the comparison can say no)', () => {
  assert.equal(f.sameFormula('=J1*3', '= J1*4'), false);
  assert.equal(f.sameFormula('=J1*3', '= J2*3'), false);
  assert.equal(f.sameFormula('=J1*3', 'J1*3'), false, 'a value text is not a formula');
  assert.equal(f.sameFormula('=J1*3', ''), false);
});

test('string literals and quoted sheet names are compared VERBATIM (case and spaces kept)', () => {
  assert.equal(f.canonical('=J1&" a  B "'), '=J1&" a  B "');
  assert.equal(f.sameFormula('=J1&" a  B "', '= J1&" A B "'), false, 'a literal whose case/spaces changed is NOT the same');
  assert.equal(f.sameFormula("='My Sheet'!B1", "= 'my sheet'!B1"), false);
  assert.equal(f.sameFormula('=IF(J1>2,"Yes  X","no")', '= IF(J1>2,"Yes  X","no")'), true);
  assert.equal(f.canonical('="a""b"'), '="a""b"', 'an escaped quote stays inside the literal');
});

test('an unterminated literal has no canonical form (fail closed)', () => {
  assert.equal(f.canonical('="abc'), null);
  assert.equal(f.sameFormula('="abc', '="abc'), false);
});

test('sheet references: bare, quoted (with escaped quote), literals ignored, external flagged', () => {
  assert.deepEqual(f.sheetRefs('=Monthly_Plan!B16+J1').names, ['Monthly_Plan']);
  assert.deepEqual(f.sheetRefs("='It''s Q1'!A1").names, ["It's Q1"]);
  assert.deepEqual(f.sheetRefs('="Other!A1"&J1').names, [], 'a sheet-like text inside a literal is not a reference');
  assert.equal(f.sheetRefs('=[Book2.xlsx]Data!A1').external, true);
  assert.equal(f.sheetRefs("='[Book2.xlsx]Data'!A1").external, true);
  assert.equal(f.sheetRefs('=SUM(A1:A3)').names.length, 0);
});

// GH #332/#333 (owner decision A): tableRenameRewrite -- a structured (table) reference renamed outside
// string literals, string literals/quoted sheet names kept verbatim (so an INDIRECT(...) text reference
// is detected, never silently rewritten).
test('tableRenameRewrite: a plain structured reference is renamed, case-insensitively, word-boundary on the left', () => {
  assert.deepEqual(f.tableRenameRewrite('=SUM(Table1[Qty])', 'Table1', 'Sales'), { formula: '=SUM(Sales[Qty])', changed: true, sawInLiteral: false });
  assert.deepEqual(f.tableRenameRewrite('=sum(table1[Qty])', 'Table1', 'Sales'), { formula: '=sum(Sales[Qty])', changed: true, sawInLiteral: false });
  assert.deepEqual(f.tableRenameRewrite('=MyTable1[Qty]', 'Table1', 'Sales'), { formula: '=MyTable1[Qty]', changed: false, sawInLiteral: false }, 'word-boundary: MyTable1 is not Table1');
});

test('tableRenameRewrite: two occurrences on different sheets/formulas each rewrite; a reference to another table is untouched', () => {
  assert.deepEqual(f.tableRenameRewrite('=SUM(Table1[Qty])+SUM(Table1[Note]="x")', 'Table1', 'Sales'), { formula: '=SUM(Sales[Qty])+SUM(Sales[Note]="x")', changed: true, sawInLiteral: false });
  assert.deepEqual(f.tableRenameRewrite('=SUM(Other[Qty])', 'Table1', 'Sales'), { formula: '=SUM(Other[Qty])', changed: false, sawInLiteral: false });
});

test('tableRenameRewrite: a relative [@Qty] reference never names the table -- nothing to rewrite', () => {
  assert.deepEqual(f.tableRenameRewrite('=[@Qty]*2', 'Table1', 'Sales'), { formula: '=[@Qty]*2', changed: false, sawInLiteral: false });
});

test('tableRenameRewrite: a string literal or quoted sheet name is NEVER touched -- sawInLiteral flags the unsafe case, not a silent rewrite', () => {
  assert.deepEqual(f.tableRenameRewrite('=INDIRECT("Table1[Qty]")', 'Table1', 'Sales'), { formula: '=INDIRECT("Table1[Qty]")', changed: false, sawInLiteral: true });
  // a literal that merely LOOKS like the pattern, with a real reference ALSO present outside it: both facts hold
  assert.deepEqual(f.tableRenameRewrite('=SUM(Table1[Qty])&"Table1[x]"', 'Table1', 'Sales'), { formula: '=SUM(Sales[Qty])&"Table1[x]"', changed: true, sawInLiteral: true });
});

test('tableRenameRewrite: a non-formula string or an unterminated literal is null (fail closed, same convention as canonical())', () => {
  assert.equal(f.tableRenameRewrite('Table1[Qty]', 'Table1', 'Sales'), null, 'no leading =');
  assert.equal(f.tableRenameRewrite('="abc', 'Table1', 'Sales'), null, 'unterminated literal');
});
