'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
// GH #293 L7: the set_numbering / set_list_type co-member guards read the word /5 Bridge's whole affected set -- the released paths run on /5
const L7_V5 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5';

/**
 * GH #236b -- set_list_type with the shared-level policy, adapter level on the fake (per numbering and level format slots). The DEFAULT policy ('propagate') is the Wave 2 behaviour and is covered by
 * test/acceptance/docx-w2-list.test.cjs; here only the explicit 'refuse' the product sends: a level other items sit on is stopped BEFORE any dispatch (the DS persists a write whose L3 fails,
 * LIVE g236v1 ROUND6 + guncprobe), a lone level is written, a level that already holds the type is SATISFIED, other levels of the numbering do not count.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const item = (text, numId, level = 0) => ({ text, props: PLAIN, list: { type: 'bullet', level, numId } });
const SOLO = 'GH236b solo';
const A = 'GH236b shared A';
const B = 'GH236b shared B';
const PARENT = 'GH236b parent';
const CHILD = 'GH236b child';
const doc = () => [item(SOLO, 'n1'), item(A, 'n2'), item(B, 'n2'), item(PARENT, 'n3', 0), item(CHILD, 'n3', 1)];
const st = (text, type, extra = {}) => ({ intent: 'set_list_type', selector: { text }, type, sharedDefinition: 'refuse', ...extra });

async function run(store, step, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, { runtime: L7_V5, ...script });
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor({ id: 'g236b', steps: [step] }));
  return { result, driver, x: result.execution, m: result.metrics };
}
const status = x => x.appliedOperations[0].status;
const reason = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };

test('plan: sharedDefinition is "propagate" | "refuse" only; absent keeps the Wave 2 plan (no field on the operation)', () => {
  assert.equal(planOperations({ steps: [{ intent: 'set_list_type', selector: { text: SOLO }, type: 'numbered' }] }).operations[0].sharedDefinition, undefined);
  assert.equal(planOperations({ steps: [st(SOLO, 'numbered')] }).operations[0].sharedDefinition, 'refuse');
  assert.equal(planOperations({ steps: [st(SOLO, 'numbered', { sharedDefinition: 'maybe' })] }).ok, false);
  assert.equal(planOperations({ steps: [{ intent: 'apply_list', selector: { text: SOLO }, type: 'numbered', sharedDefinition: 'refuse' }] }).ok, false, 'only set_list_type takes the policy');
});

test('refuse: an item alone on its numbering level is re-templated (APPLIED, L1/L3 PASS); the retry is SATISFIED with no write', async () => {
  const store = createStore(doc());
  const first = await run(store, st(SOLO, 'numbered'));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED, JSON.stringify(first.x.appliedOperations.map(e => e.evidence)));
  assert.equal(first.x.primitiveVerification[0].verdict, 'PASS');
  assert.equal(first.x.preservationVerification[0].verdict, 'PASS');
  const second = await run(store, st(SOLO, 'numbered'));
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.equal(second.m.writes, 0);
});

test('refuse: a level another item sits on -> FAIL_CLOSED before any dispatch, the numbering, level and count named, the document unchanged', async () => {
  for (const target of [A, B]) {
    const store = createStore(doc());
    const before = JSON.stringify(store.durable.map(p => [p.text, p.list]));
    const { x, driver, m } = await run(store, st(target, 'numbered'));
    assert.equal(status(x), OPERATION_STATUS.FAILED, target);
    assert.equal(reason(x), 'list-shared-definition:n2:level0:1');
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(m.writes, 0);
    assert.equal(JSON.stringify(store.durable.map(p => [p.text, p.list])), before);
  }
});

test('refuse: a shared level that ALREADY holds the asked type is SATISFIED (nothing to change), and the child alone at its level of a numbering the parent uses is written', async () => {
  const store = createStore(doc());
  const same = await run(store, st(A, 'bullet'));
  assert.equal(status(same.x), OPERATION_STATUS.SATISFIED);
  assert.equal(same.m.writes, 0);
  const nested = await run(createStore(doc()), st(CHILD, 'numbered'));
  assert.equal(status(nested.x), OPERATION_STATUS.APPLIED, JSON.stringify(nested.x.appliedOperations.map(e => e.evidence)));
});

test('default policy unchanged: without sharedDefinition a shared level is NOT stopped before dispatch (the Wave 2 path, whose L3 is what catches it)', async () => {
  const { driver } = await run(createStore(doc()), { intent: 'set_list_type', selector: { text: A }, type: 'numbered' });
  assert.ok(driver.calls.dispatch >= 1, 'the default still dispatches');
});
