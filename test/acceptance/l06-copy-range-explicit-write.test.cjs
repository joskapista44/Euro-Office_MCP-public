'use strict';

/**
 * L-06 (owner decision L-06): copy_range on the PRODUCT surface -- the family level (range.cjs) with the fake XLSX editor. The LIVE measurement on the prod DS (DS /12,
 * docs/evidence/l06-range-copy-fail-closed-20261015) found that the editor's Copy shows a formula in session but the SAVED file carries the source's LITERAL text (and
 * its stale value): on the SAME sheet an explicit formula write after the Copy persists correctly; onto ANOTHER sheet the Copy leaves a delayed state that overwrites a formula
 * written after it at the next save, so no Copy is dispatched there -- the typed values and the planner-translated formulas are written (the cell FORMAT then does not travel,
 * named notProven 'cross-sheet-copy-format-not-carried'). The fake models both defects ('copy-saves-literal', 'cross-copy-delayed'); the control tests show the live symptom.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

const status = r => r.execution.appliedOperations[0].status;
const reasonOf = r => { const e = r.execution.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason) || null; };
const sheet = (store, n) => store.durable.find(s => s.name === n).cells;
const runWith = async (store, t, script) => {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const result = await withSession({ driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
};
const book2 = () => [{ name: 'S', cells: { A1: { v: 'h1' }, A2: { v: 1 }, A3: { v: 101, f: '=A2+H1' }, H1: { v: 100 } } }, { name: 'T', cells: {} }];
const copyA = { id: 'l06r5', steps: [{ intent: 'copy_range', sheet: 'S', range: 'A1:A3', target: 'E1' }] };
const bookX = () => [{ name: 'S', cells: { A1: { v: 'h1' }, B1: { v: 'h2', bold: true }, A2: { v: 1 }, B2: { v: 2.5, nf: '0.00' }, A3: { v: 101, f: '=A2+H1' }, H1: { v: 100 } } }, { name: 'T', cells: {} }];
const xcopy = { id: 'l06r8', steps: [{ intent: 'copy_range', sheet: 'S', range: 'A1:B3', target: 'C2', targetSheet: 'T' }] };


test('L-06 r5: the first copy is APPLIED and its SAVED formula is the TRANSLATED one; the reopened retry is SATISFIED with 0 writes / 0 commands', async () => {
  const store = createStore(book2());
  const first = await runWith(store, copyA, { fault: 'copy-saves-literal' });
  assert.equal(first.result.ok, true, JSON.stringify(first.result.execution.appliedOperations[0].evidence));
  assert.equal(status(first.result), OPERATION_STATUS.APPLIED);
  assert.equal(sheet(store, 'S').E3.f, '=E2+L1', 'the SAVED target formula is the translated one (not the source literal)');
  const again = await runWith(store, copyA, { fault: 'copy-saves-literal' });
  assert.equal(status(again.result), OPERATION_STATUS.SATISFIED, `retry: ${status(again.result)} ${reasonOf(again.result)}`);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
});

test('L-06 r5 CONTROL: without the explicit formula write the same fake saves the literal -- the in-session L1 passes, the reopened retry is refused as occupied (the live symptom)', async () => {
  const store = createStore(book2());
  // a hand-built command: the Copy WITHOUT meta.explicitFormulaWrite, as the pre-fix lower() dispatched it
  const driver = createFakeXlsxDriver(store, { fault: 'copy-saves-literal' });
  await driver.open({ fileId: 'disposable' });
  const op = require('../../src/adapters/xlsx/families/range.cjs').plan({ intent: 'copy_range', id: 'c', ...copyA.steps[0] }, 0).operation;
  const cmd = require('../../src/adapters/xlsx/families/range.cjs').lower(op);
  await driver.dispatch(null, { ...cmd, meta: { ...cmd.meta, explicitFormulaWrite: false } });
  await driver.forceSave();
  assert.equal(sheet(store, 'S').E3.f, '=A2+H1', 'the saved file carries the source literal');
  assert.equal(sheet(store, 'S').E3.v, 101);
  const again = await runWith(store, copyA, {});
  assert.equal(again.result.ok, false);
  assert.equal(reasonOf(again.result), 'xlsx-copy-target-occupied');
});

test('L-06 r5: lower() -- a copy_range Copy is followed by an explicit rewrite of the target formulas; a move_range Cut is untouched', () => {
  const range = require('../../src/adapters/xlsx/families/range.cjs');
  const cp = range.lower(range.plan({ intent: 'copy_range', id: 'c', ...copyA.steps[0] }, 0).operation, { formulas: { E3: '=E2+L1' } });
  assert.equal(cp.meta.explicitFormulaWrite, true);
  assert.deepEqual(cp.meta.formulaWrites, { E3: '=E2+L1' });
  assert.match(cp.body, /\.Copy\(/);
  assert.match(cp.body, /GetFormula\(\)/);
  assert.match(cp.body, /SetValue\(/);
  assert.match(cp.body, /"E3":"=E2\+L1"/, 'the PLANNER text is embedded, not read from the editor');
  assert.ok(cp.body.indexOf('SetValue(') > cp.body.indexOf('.Copy('), 'the rewrite comes AFTER the Copy');
  assert.match(cp.body, /xlsx-copy-formula-not-written|copy-formula/, 'a target that lost its formula fails the command, named');
  // L-06/b: move_range no longer accepts `expect` on the step at all (classify() builds it).
  const mv = range.lower(range.plan({ intent: 'move_range', id: 'm', sheet: 'S', range: 'A1:A3', target: 'E1' }, 0).operation);
  assert.notEqual(mv.meta.explicitFormulaWrite, true);
  assert.doesNotMatch(mv.body, /SetValue\(|GetFormula\(/);
});

test('L-06 r6: onto ANOTHER sheet the editor shows (and saves) the source text untranslated -- the planner-translated write repairs it; the reopened retry is SATISFIED', async () => {
  const store = createStore(book2());
  const t = { id: 'l06r6', steps: [{ intent: 'copy_range', sheet: 'S', range: 'A1:A3', target: 'B1', targetSheet: 'T' }] };
  const first = await runWith(store, t, { fault: 'copy-xsheet-shows-literal' });
  assert.equal(first.result.ok, true, JSON.stringify(first.result.execution.appliedOperations[0].evidence));
  assert.equal(sheet(store, 'T').B3.f, '=B2+I1', 'the SAVED target-sheet formula is the planner translation (A2->B2, H1->I1), not the source literal');
  const again = await runWith(store, t, { fault: 'copy-xsheet-shows-literal' });
  assert.equal(status(again.result), OPERATION_STATUS.SATISFIED, `retry: ${status(again.result)} ${reasonOf(again.result)}`);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
});

test('L-06 r8: a cross-sheet copy under the DELAYED-COPY fault is APPLIED with the typed values and the planner formula SAVED, no format carried, a retry SATISFIED 0/0, the format limit NAMED', async () => {
  const store = createStore(bookX());
  const first = await runWith(store, xcopy, { fault: 'cross-copy-delayed' });
  assert.equal(first.result.ok, true, JSON.stringify(first.result.execution.appliedOperations[0].evidence).slice(0, 500));
  assert.equal(status(first.result), OPERATION_STATUS.APPLIED);
  const T = sheet(store, 'T');
  assert.deepEqual([T.C2.v, T.D2.v, T.C3.v, T.D3.v], ['h1', 'h2', 1, 2.5], 'typed values (a number stays a number)');
  assert.equal(T.C4.f, '=C3+J2', 'the planner-translated formula (A2 -> C3, H1 -> J2: dr 1, dc 2) was SAVED');
  assert.equal(T.D2.bold, undefined, 'the cell format is not carried across sheets');
  assert.equal(T.D3.nf, undefined);
  const checks = first.result.execution.wholeTaskVerification;
  assert.ok(JSON.stringify(first.result.execution).includes('cross-sheet-copy-format-not-carried'), 'the receipt names the limit');
  assert.ok(checks, 'whole-task verification ran');
  const again = await runWith(store, xcopy, { fault: 'cross-copy-delayed' });
  assert.equal(status(again.result), OPERATION_STATUS.SATISFIED, `retry: ${status(again.result)} ${reasonOf(again.result)}`);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
});

test('L-06 r8: a cross-sheet copy whose source holds a value the reader cannot type (TRUE) is refused BEFORE the write, named (a text that reads as a formula is already caught earlier by the dependency observation; the resolve guard is defence in depth)', async () => {
  for (const [cells, why] of [[{ A1: { v: 'TRUE' } }, 'xlsx-copy-cross-sheet-value-type-not-observable']]) {
    const store = createStore([{ name: 'S', cells }, { name: 'T', cells: {} }]);
    const r = await runWith(store, { id: 'x', steps: [{ intent: 'copy_range', sheet: 'S', range: 'A1:A1', target: 'C2', targetSheet: 'T' }] }, {});
    assert.equal(r.result.ok, false);
    assert.ok(JSON.stringify(r.result.execution).includes(why), JSON.stringify(r.result.execution).slice(0, 400));
    assert.equal(r.m.commands, 0);
    assert.deepEqual(store.durable.find(x => x.name === 'T').cells, {});
  }
});

test('L-06 r8: lower() of a cross-sheet copy dispatches NO Copy -- the values first, then the planner formulas with a read-back; the same-sheet Copy keeps its Copy', () => {
  const range = require('../../src/adapters/xlsx/families/range.cjs');
  const op = range.plan({ intent: 'copy_range', id: 'c', ...xcopy.steps[0] }, 0).operation;
  const x = range.lower(op, { from: {}, to: op.dest, formulas: { C4: '=C3+J2' }, values: { C2: 'h1', C3: 1 }, crossSheet: true });
  assert.doesNotMatch(x.body, /\.Copy\(|\.Cut\(/);
  assert.ok(x.body.indexOf('"C2":"h1"') > -1 && x.body.indexOf('"C4":"=C3+J2"') > x.body.indexOf('"C2":"h1"'), 'values come before the formulas');
  assert.deepEqual([x.meta.crossSheetWrite, x.meta.explicitFormulaWrite, x.meta.valueWrites, x.meta.formulaWrites], [true, true, { C2: 'h1', C3: 1 }, { C4: '=C3+J2' }]);
  const same = range.lower(range.plan({ intent: 'copy_range', id: 'c', ...copyA.steps[0] }, 0).operation, { formulas: { E3: '=E2+L1' } });
  assert.match(same.body, /\.Copy\(/);
  assert.equal(same.meta.crossSheetWrite, undefined);
});
