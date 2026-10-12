'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const validation = require('../../src/adapters/xlsx/families/validation.cjs');
const cf = require('../../src/adapters/xlsx/families/cf.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 R17 / R18 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md): the measured validation types /
 * operators / alert styles and the measured conditional-format types / operators. Every expected Add call below is the
 * probe's own call, written out (m7-d-probe1-20261003.json, ACT.v / ACT.c), not derived by the code under test.
 */

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];

// ---- R17 validation --------------------------------------------------------------------------------------------------

const vbook = () => [{ name: 'V', cells: { B2: { v: 1 }, B3: { v: 2 }, Z1: { v: 'a' }, Z2: { v: 'b' }, Z3: { v: 'c' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const vset = extra => ({ id: 'm7-d-r17', steps: [{ intent: 'set_validation', sheet: 'V', range: 'B2:B5', ...extra }] });

for (const [name, extra, addArgs] of [
  ['list inline', { type: 'list', alertStyle: 'stop', items: ['a', 'b', 'c'] }, ['xlValidateList', 'xlValidAlertStop', 'xlBetween', '"a,b,c"']],
  ['list range', { type: 'list', alertStyle: 'stop', listRange: 'Z1:Z3' }, ['xlValidateList', 'xlValidAlertStop', 'xlBetween', '=$Z$1:$Z$3']],
  ['decimal between, warning', { type: 'decimal', alertStyle: 'warning', operator: 'between', min: 0.5, max: 9.5 }, ['xlValidateDecimal', 'xlValidAlertWarning', 'xlBetween', '0.5', '9.5']],
  ['text length <=, information', { type: 'textLength', alertStyle: 'information', operator: 'lessEqual', value: 5 }, ['xlValidateTextLength', 'xlValidAlertInformation', 'xlLessEqual', '5']],
  ['whole notBetween', { type: 'wholeNumber', alertStyle: 'stop', operator: 'notBetween', min: 1, max: 3 }, ['xlValidateWholeNumber', 'xlValidAlertStop', 'xlNotBetween', '1', '3']],
  ['whole greater', { type: 'wholeNumber', alertStyle: 'stop', operator: 'greater', value: 4 }, ['xlValidateWholeNumber', 'xlValidAlertStop', 'xlGreater', '4']],
]) {
  test(`R17 ${name}: the measured Add call; APPLIED, retry SATISFIED 0/0`, async () => {
    const p = planOperations(vset(extra));
    assert.equal(p.ok, true, p.detail);
    const log = [];
    const ws = { GetName: () => 'V', GetRange: () => ({ GetValidation: () => ({ Add: (...x) => log.push(x) }) }) };
    new Function('Api', validation.lower(p.operations[0]).body)({ GetSheet: () => ws });
    assert.deepEqual(log, [addArgs]);
    const store = createStore(vbook());
    const { result, m } = await run(store, vset(extra));
    assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, op(result).evidence]));
    assert.equal(m.commands, 1);
    const again = await run(store, vset(extra));
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test('R17 refused at planning: date / custom / time, unmeasured operators, bad bounds and items', () => {
  for (const [extra, re] of [
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: 1, max: 2 }, /type must be one of/],
    [{ type: 'custom', alertStyle: 'stop' }, /type must be one of/],
    [{ type: 'wholeNumber', alertStyle: 'stop', operator: 'equal', value: 1 }, /operator must be one of/],
    [{ type: 'decimal', alertStyle: 'stop', operator: 'between', min: 2, max: 1 }, /min <= max/],
    [{ type: 'textLength', alertStyle: 'stop', operator: 'greater', value: -1 }, /non-negative/],
    [{ type: 'list', alertStyle: 'stop', items: ['a,b'] }, /without a comma/],
    [{ type: 'list', alertStyle: 'stop', items: ['a'], listRange: 'Z1:Z3' }, /exactly one of/],
    [{ type: 'wholeNumber', alertStyle: 'stop', operator: 'greater', value: 4, items: ['a'] }, /items belongs to type list/],
  ]) {
    const p = planOperations(vset(extra));
    assert.equal(p.ok, false, JSON.stringify(extra));
    assert.match(p.detail, re);
  }
});

test('R17 RED: the second bound lands wrong -> L1 FAIL', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'decimal', alertStyle: 'warning', operator: 'between', min: 0.5, max: 9.5 }), { fault: 'validation-wrong-bounds' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

// ---- R18 conditional formats -----------------------------------------------------------------------------------------

const cbook = () => [{ name: 'C', cells: { A1: { v: 1 }, A2: { v: 6 }, A3: { v: 9 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const cadd = rule => ({ id: 'm7-d-r18', steps: [{ intent: 'add_conditional_format', sheet: 'C', range: 'A1:A3', rule }] });
const cdel = rule => ({ id: 'm7-d-r18', steps: [{ intent: 'delete_conditional_format', sheet: 'C', range: 'A1:A3', rule }] });

for (const [name, rule, addArgs, fill] of [
  ['between + fill', { type: 'cellValue', operator: 'between', formula1: 2, formula2: 7, fill: '#FFC8C8' }, ['xlCellValue', 'xlBetween', '2', '7'], [255, 200, 200]],
  ['less + fill', { type: 'cellValue', operator: 'less', formula1: 3, fill: '#C8FFC8' }, ['xlCellValue', 'xlLess', '3'], [200, 255, 200]],
  ['equal, no fill', { type: 'cellValue', operator: 'equal', formula1: 6 }, ['xlCellValue', 'xlEqual', '6'], null],
  ['notBetween, no fill', { type: 'cellValue', operator: 'notBetween', formula1: 2, formula2: 7 }, ['xlCellValue', 'xlNotBetween', '2', '7'], null],
  ['expression + fill', { type: 'expression', formula: '=$A1>5', fill: '#C8C8FF' }, ['xlExpression', null, '=$A1>5'], [200, 200, 255]],
]) {
  test(`R18 ${name}: the measured Add call; add APPLIED + retry 0/0; delete APPLIED + retry 0/0`, async () => {
    const p = planOperations(cadd(rule));
    assert.equal(p.ok, true, p.detail);
    const log = [];
    const ws = { GetName: () => 'C', GetRange: () => ({ GetFormatConditions: () => ({ Add: (...x) => { log.push(['Add', ...x]); return { SetFillColor: c => log.push(['Fill', c]) }; } }) }) };
    new Function('Api', cf.lower(p.operations[0]).body)({ GetSheet: () => ws, CreateColorFromRGB: (...x) => x });
    assert.deepEqual(log, [['Add', ...addArgs], ...(fill ? [['Fill', fill]] : [])]);
    const store = createStore(cbook());
    const a = await run(store, cadd(rule));
    assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, op(a.result).evidence]));
    assert.equal((await run(store, cadd(rule))).m.commands, 0);
    const d = await run(store, cdel(rule));
    assert.equal(d.result.ok, true, JSON.stringify(d.result.execution.primitiveVerification));
    assert.equal((await run(store, cdel(rule))).m.commands, 0);
  });
}

test('R18 refused: unmeasured types / operators; an expression without a fill; between without formula2', () => {
  for (const [rule, re] of [
    // M7 X-O01 widened the measured set: top10/unique/aboveAverage now have their
    // OWN shape (xlsx-m7-d-o-slice.test.cjs). X-O01 full implementation widened it
    // again: a BARE { type: 'databar' } is now ACCEPTED (default-only create) -- any appearance
    // field on it stays refused, since the criteria/enum shapes are unmeasured.
    [{ type: 'databar', barColor: '#00FF00' }, /a databar rule is/],
    [{ type: 'cellValue', operator: 'notEqual', formula1: 1 }, /operator must be one of/],
    [{ type: 'expression', formula: '=$A1>5' }, /needs a fill/],
    [{ type: 'cellValue', operator: 'between', formula1: 1 }, /needs formula2/],
    [{ type: 'cellValue', operator: 'less', formula1: 1, formula2: 2 }, /formula1 only/],
  ]) {
    const p = planOperations(cadd(rule));
    assert.equal(p.ok, false, JSON.stringify(rule));
    assert.match(p.detail, re);
  }
});

test('R18 RED: formula2 lands wrong / the fill lands wrong -> L1 FAIL', async () => {
  for (const [fault, rule] of [['cf-wrong-formula2', { type: 'cellValue', operator: 'between', formula1: 2, formula2: 7, fill: '#FFC8C8' }], ['cf-wrong-fill', { type: 'cellValue', operator: 'equal', formula1: 6 }]]) {
    const { result } = await run(createStore(cbook()), cadd(rule), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
});
