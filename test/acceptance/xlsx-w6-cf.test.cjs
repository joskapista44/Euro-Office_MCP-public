'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const cf = require('../../src/adapters/xlsx/families/cf.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 6: add / delete a conditional format -- the FULL rule identity (the source matched only named fields, B-D34). */

const RED = 16711680;
const book = () => [{ name: 'C', cells: { A1: { v: 1 }, A2: { v: 6 }, A3: { v: 9 } }, rules: [{ kind: 'conditional', type: 'xlExpression', formula: '=$A1>8', range: 'A1:A3' }] }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const rule = (fill = '#FF0000', formula1 = '5') => ({ type: 'cellValue', operator: 'greater', formula1, fill });
const add = (r = rule()) => ({ id: 'm6-w6c', steps: [{ intent: 'add_conditional_format', sheet: 'C', range: 'A1:A3', rule: r }] });
const del = (r = rule()) => ({ id: 'm6-w6c', steps: [{ intent: 'delete_conditional_format', sheet: 'C', range: 'A1:A3', rule: r }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const rules = store => store.durable.find(s => s.name === 'C').rules;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('add: APPLIED, the other rule kept, retry SATISFIED 0/0; delete: APPLIED, retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, add());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(rules(store).length, 2);
  assert.equal((await run(store, add())).m.commands, 0);
  const d = await run(store, del());
  assert.equal(d.result.ok, true, JSON.stringify(d.result.execution.primitiveVerification));
  assert.deepEqual(rules(store).map(r => r.type), ['xlExpression']);
  const d2 = await run(store, del());
  assert.equal(d2.m.commands, 0);
  assert.equal(d2.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const [name, pre, t, why, script] of [
  ['the same condition with ANOTHER fill (the source called it satisfied, B-D34)', [{ kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', formula: '5', fill: 255, range: 'A1:A3' }], add(), 'xlsx-cf-same-condition-other-fill-or-range'],
  ['two identical rules (identity ambiguous)', [1, 2].map(() => ({ kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', formula: '5', fill: RED, range: 'A1:A3' })), del(), 'xlsx-cf-rule-identity-ambiguous'],
  ['an unobservable range read', [], add(), 'xlsx-cf-read-unobservable', { cfReadUnobservable: true }],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const b = book();
    b[0].rules = [...b[0].rules, ...pre];
    const { result, driver } = await run(createStore(b), t, script || {});
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / wrong fill / duplicated -> L1 FAIL; another rule removed / a cell -> L3 FAIL', async () => {
  for (const fault of ['ignore', 'cf-wrong-fill', 'cf-duplicate']) {
    const store = createStore(book());
    const { result, driver } = await run(store, add(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  const other = await run(createStore(book()), add(), { fault: 'cf-removes-other' });
  assert.equal(other.result.ok, false);
  const altered = await run(createStore(book()), add(), { fault: 'cf-alters-other' });
  assert.equal(altered.result.ok, false);
  assert.equal(altered.result.execution.primitiveVerification[0].verdict, 'PASS', 'the new rule itself is right');
  assert.equal(altered.result.execution.preservationVerification[0].verdict, 'FAIL', 'the OTHER rule changed: L3');
  const side = await run(createStore(book()), add(), { fault: 'side-effect' });
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: only the measured rule shape', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  // M7 R18 widened the measured set (xlsx-m7-d-rules.test.cjs); these shapes stay refused
  refuse(add({ ...rule(), operator: 'between' }), /needs formula2/);
  refuse(add({ ...rule(), type: 'expression' }), /expression rule is/);
  // M7 X-O01 widened the measured set further -- top10/unique/aboveAverage now have
  // their OWN shape (xlsx-m7-d-o-slice.test.cjs). X-O01 full implementation widened
  // it again: a BARE { type: 'databar' } is now ACCEPTED (default-only create) -- cellValue's
  // operator/formula1/fill fields on it stay refused, since databar has no such shape.
  refuse(add({ ...rule(), type: 'databar' }), /a databar rule is/);
  refuse(add({ ...rule(), operator: 'greaterEqual' }), /operator must be one of/);
  refuse(add({ ...rule(), fill: 'red' }), /rule.fill must be/);
  refuse(add({ ...rule(), formula1: '=5' }), /formula1 must be/);
  refuse(add({ ...rule(), priority: 1 }), /rule must be/);
});

test('the command bodies: Add + SetFillColor; delete re-identifies the ONE rule by its getters at dispatch', () => {
  const log = [];
  const item = (type, fill) => ({ GetType: () => type, GetOperator: () => 'xlGreater', GetFormula1: () => '5', GetFillColor: () => ({ GetRGB: () => fill }), GetAppliesTo: () => ({ GetAddress: () => 'A1:A3' }), Delete: () => log.push(['delete', type, fill]) });
  const items = [item('xlExpression', null), item('xlCellValue', RED)];
  const ws = { GetName: () => 'C', GetRange: () => ({ GetFormatConditions: () => ({ Add: (...x) => { log.push(['add', ...x]); return { SetFillColor: c => log.push(['fill', c]) }; }, GetCount: () => items.length, GetItem: i => items[i - 1] }) }) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (...x) => x.join(',') };
  assert.deepEqual(new Function('Api', cf.lower(planOperations(add()).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(new Function('Api', cf.lower(planOperations(del()).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(log, [['add', 'xlCellValue', 'xlGreater', '5'], ['fill', '255,0,0'], ['delete', 'xlCellValue', RED]]);
  items.push(item('xlCellValue', RED));
  assert.deepEqual(new Function('Api', cf.lower(planOperations(del()).operations[0]).body)(Api), { ok: false, stage: 'identity-at-dispatch', hits: 2 });
});
