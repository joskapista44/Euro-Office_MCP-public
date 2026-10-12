'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const validation = require('../../src/adapters/xlsx/families/validation.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 6: set / clear_validation on the fake -- public read for the parameters, internal ranges for the coverage. */

const book = () => [{ name: 'V', cells: { B2: { v: 1 }, B3: { v: 2 }, B4: { v: 3 }, B5: { v: 4 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const set = (range = 'B2:B5', extra = {}) => ({ id: 'm6-w6v', steps: [{ intent: 'set_validation', sheet: 'V', range, type: 'wholeNumber', operator: 'between', alertStyle: 'stop', min: 1, max: 10, ...extra }] });
const clear = (range = 'B2:B5') => ({ id: 'm6-w6v', steps: [{ intent: 'clear_validation', sheet: 'V', range }] });
const rule = (range, f1 = '1', f2 = '10') => ({ kind: 'validation', type: 1, range, formula: f1, formula2: f2, api: { type: 'xlValidateWholeNumber', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const rules = store => store.durable.find(s => s.name === 'V').rules || [];
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('set: APPLIED (exactly one rule on exactly the range), retry SATISFIED 0/0; clear: APPLIED, retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, set());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.deepEqual(rules(store).map(r => r.range), ['B2:B5']);
  const a2 = await run(store, set());
  assert.equal(a2.m.commands, 0);
  assert.equal(a2.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  const c = await run(store, clear());
  assert.equal(c.result.ok, true, JSON.stringify(c.result.execution.primitiveVerification));
  assert.deepEqual(rules(store), []);
  const c2 = await run(store, clear());
  assert.equal(c2.m.commands, 0);
  assert.equal(c2.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const [name, t, why, pre, script] of [
  ['a rule covering the target with another range (partial)', set(), 'xlsx-validation-overlaps-another-range', ['B3:B8']],
  ['an existing rule on exactly the range with other bounds (no silent Modify, B-D23)', set(), 'xlsx-validation-exists-different', ['B2:B5|1|99']],
  ['clear when the rule reaches beyond the target', clear('B2:B3'), 'xlsx-validation-overlaps-another-range', ['B2:B5']],
  ['an unobservable public read', set(), 'xlsx-validation-read-unobservable', [], { validationReadUnobservable: true }],
  ['an unproven internal reader', set(), 'xlsx-validation-rules-unobservable', [], { dependencyBuild: false }],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const b = book();
    b[0].rules = pre.map(x => { const [r, f1, f2] = x.split('|'); return rule(r, f1, f2); });
    const { result, driver } = await run(createStore(b), t, script || {});
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / wrong bounds -> L1 FAIL; a second rule elsewhere / a cell -> L3 FAIL; nothing persisted', async () => {
  for (const fault of ['ignore', 'validation-wrong-bounds']) {
    const store = createStore(book());
    const { result, driver } = await run(store, set(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  const leak = await run(createStore(book()), set(), { fault: 'validation-leaks', leakRange: 'D1:D2' });
  assert.equal(leak.result.ok, false);
  assert.equal(leak.result.execution.preservationVerification[0].verdict, 'FAIL');
  const side = await run(createStore(book()), set(), { fault: 'side-effect' });
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: only the measured shape', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  // M7 R17 widened the measured set (xlsx-m7-d-rules.test.cjs); these shapes stay refused
  refuse(set('B2:B5', { type: 'list' }), /type list does not take operator/);
  refuse(set('B2:B5', { operator: 'greater' }), /takes value, not min/);
  refuse(set('B2:B5', { type: 'date' }), /type must be one of/);
  refuse(set('B2:B5', { alertStyle: 'loud' }), /alertStyle must be one of/);
  refuse(set('B2:B5', { min: 5, max: 1 }), /min <= max/);
  refuse(set('B2:B5', { min: 1.5 }), /integers/);
  refuse(set('B2:B5', { ignoreBlank: true }), /does not support field/);
});

test('the command bodies: GetValidation().Add(...) / Delete() on the exact-name sheet', () => {
  const log = [];
  const ws = { GetName: () => 'V', GetRange: a => ({ GetValidation: () => ({ Add: (...x) => log.push([a, 'Add', ...x]), Delete: () => log.push([a, 'Delete']) }) }) };
  assert.deepEqual(new Function('Api', validation.lower(planOperations(set()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(new Function('Api', validation.lower(planOperations(clear()).operations[0]).body)({ GetSheet: () => ws }), { ok: true });
  assert.deepEqual(log, [['B2:B5', 'Add', 'xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', '1', '10'], ['B2:B5', 'Delete']]);
});
