'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const formula = require('../../src/adapters/xlsx/formula.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 R26c (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md): sheet names with spaces / accents --
 * quoted in a formula, kept verbatim, followed by a rename (= 'Kiadás x'!A1+1). Copy and the non-cell dependents of
 * such a name stay refused (not measured).
 */

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const task = step => ({ id: 'm7-d-r26c', steps: [step] });
const book = () => [
  { name: 'RF', cells: { A1: { v: 'x' } } },
  { name: 'Q 1', cells: { A1: { v: 22 } } },
];
const sheet = (store, n) => store.durable.find(s => s.name === n);
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('R26c: the renamed reference is the measured quoted text (formula.canonicalRenamed)', () => {
  // the probe's read-back, written out (m7-d-probe1-20261003.json R2.RF.a2: "= 'Kiadás db3617'!A1+1")
  assert.equal(formula.canonicalRenamed("='Q 1 db3617'!A1+1", 'Q 1 db3617', 'Kiadás db3617'), formula.canonical("= 'Kiadás db3617'!A1+1"));
  assert.equal(formula.canonicalRenamed('=Data!B2+1', 'Data', 'Adat lap'), "='Adat lap'!B2+1");
  assert.equal(formula.canonicalRenamed("='Q 1'!A1", 'Q 1', 'Plain'), '=PLAIN!A1');
});

test('R26c: create a quoted-class sheet; write a formula referencing it; rename it -> the formula follows, quoted; retries 0/0', async () => {
  const store = createStore(book());
  const c = await run(store, task({ intent: 'create_sheet', name: 'Bevétel 2026' }));
  assert.equal(c.result.ok, true, JSON.stringify(c.result.execution.primitiveVerification));
  const w = task({ intent: 'write_range', sheet: 'RF', range: 'B1', formulas: [["='Q 1'!A1*2"]] });
  const wr = await run(store, w);
  assert.equal(wr.result.ok, true, JSON.stringify([wr.result.execution.primitiveVerification, wr.result.execution.appliedOperations[0].evidence]));
  assert.equal(sheet(store, 'RF').cells.B1.f, "= 'Q 1'!A1*2");
  const r = task({ intent: 'rename_sheet', sheet: 'Q 1', name: 'Kiadás 1' });
  const rr = await run(store, r);
  assert.equal(rr.result.ok, true, JSON.stringify([rr.result.execution.primitiveVerification, rr.result.execution.preservationVerification]));
  assert.equal(sheet(store, 'RF').cells.B1.f, "= 'Kiadás 1'!A1*2");
  const again = await run(store, r);
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('R26c RED: the quoted reference does not follow the rename -> L3 FAIL', async () => {
  const b = book();
  b[0].cells.B1 = { v: 44, f: "= 'Q 1'!A1*2" };
  const { result } = await run(createStore(b), task({ intent: 'rename_sheet', sheet: 'Q 1', name: 'Kiadás 1' }), { fault: 'rename-no-follow' });
  assert.equal(result.ok, false);
});

test('R26c refused: copy of a quoted-class sheet; (L-21/2: an accented name without a space is now RELEASED); a rename whose sheet a validation references', async () => {
  const copy = await run(createStore(book()), task({ intent: 'copy_sheet', sheet: 'Q 1', name: 'Q 2' }));
  assert.equal(copy.driver.calls.dispatch, 0);
  assert.equal(reason(copy.result), 'xlsx-copy-quoted-name-not-measured');
  assert.equal(planOperations(task({ intent: 'create_sheet', name: 'Ünnep' })).ok, true);
  const b = book();
  b[0].rules = [{ kind: 'validation', type: 1, range: 'C1', formula: "='Q 1'!$A$1", formula2: '', api: { type: 'xlValidateList', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }];
  const rn = await run(createStore(b), task({ intent: 'rename_sheet', sheet: 'Q 1', name: 'Kiadás 1' }));
  assert.equal(rn.driver.calls.dispatch, 0);
  assert.equal(reason(rn.result), 'xlsx-rename-quoted-name-dependent-not-measured');
});
