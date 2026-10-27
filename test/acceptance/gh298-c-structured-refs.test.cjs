'use strict';

/**
 * GH #298 C: structured (table) references in write_range formulas. formula.sheetRefs tells a [book] reference (still
 * refused) from a structured one (accepted); the dependency observation keeps a structured reference unresolvable
 * (fail-closed: its table's sheet is not read). The read-back shapes are the LIVE design probe's
 * (docs/evidence/gh298-c-structured-ref-probe-20261026.json).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const formula = require('../../src/adapters/xlsx/formula.cjs');
const W = require('../../src/server/write-tools.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const step = f => ({ intent: 'write_range', sheet: 'S', range: 'F1', formulas: [[f]] });

test('C sheetRefs: a [book] reference is external, a structured reference is structured, a plain one neither', () => {
  for (const f of ['=SUM(Table1[Qty])', '=Table1[[#This Row],[Qty]]*2', '=[@Qty]*2', '=SUM(Table1[[Qty]:[Unit Price]])', '=SUMIF(T[Region],"[x]",T[Qty])']) {
    const r = formula.sheetRefs(f);
    assert.deepEqual([r.external, r.structured], [false, true], f);
  }
  for (const f of ['=[Book2.xlsx]Data!A1', '=[1]Sheet1!A1', '=[book]!Name', "='[b.xlsx]S 1'!A1", '=Table1[Qty]+[b.xlsx]S!A1']) assert.equal(formula.sheetRefs(f).external, true, f);
  assert.deepEqual([formula.sheetRefs('=SUM(Data!A1:A3)').external, formula.sheetRefs('=SUM(Data!A1:A3)').structured], [false, false]);
  assert.equal(formula.sheetRefs('="[x]"&A1').structured, false, 'a bracket inside a string literal is no reference');
});

test('C write_range: a structured reference is accepted; a [book] reference still refused', () => {
  for (const f of ['=SUM(Table1[Qty])', '=Table1[[#This Row],[Qty]]*2', '=SUM(Table1[Unit Price])']) {
    assert.equal(W.validateWriteRequest(req([step(f)])).ok, true, f);
    assert.equal(planOperations({ id: 't', steps: [step(f)] }).ok, true, f);
  }
  const p = planOperations({ id: 't', steps: [step('=[Book2.xlsx]Data!A1')] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /external workbook references are refused/);
});

test('C read-back: the measured re-rendering compares equal; a dropped (unresolvable) formula does not', () => {
  assert.equal(formula.sameFormula('=SUM(Table1[Qty])', '= SUM(Table1[Qty])'), true);
  assert.equal(formula.sameFormula('=sum(table1[qty])', '= SUM(Table1[Qty])'), true, 'case folded to the table\'s own (measured)');
  assert.equal(formula.sameFormula('=SUM(Table1[Unit Price])', '= SUM(Table1[Unit Price])'), true);
  assert.equal(formula.sameFormula('=Table1[[#This Row],[Qty]]*2', '= Table1[[#This Row],[Qty]]*2'), true);
  assert.equal(formula.sameFormula('=SUM(Nope[Qty])', ''), false, 'the engine dropped it: the cell reads empty');
});
