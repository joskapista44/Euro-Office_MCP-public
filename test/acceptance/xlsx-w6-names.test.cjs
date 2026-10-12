'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const names = require('../../src/adapters/xlsx/families/names.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M6 wave 6: workbook-level defined names through the build-gated names reader (the W3 reader) on the fake. */

const book = () => [{ name: 'Data', cells: { A1: { v: 7 }, B1: { v: 14, f: '=Used*2' } } }, { name: 'Other', cells: { A1: { v: 'keep' } } }];
const store0 = () => { const s = createStore(book()); s.names = [{ name: 'Used', ref: 'Data!$A$1' }, { name: 'Free', ref: 'Other!$A$1' }]; return s; };
const set = (name, refersTo) => ({ id: 'm6-w6n', steps: [{ intent: 'set_defined_name', name, refersTo }] });
const del = name => ({ id: 'm6-w6n', steps: [{ intent: 'delete_defined_name', name }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('set: APPLIED (1 command, L1 + L3, barrier, durable); retry SATISFIED 0/0', async () => {
  const store = store0();
  const { result, m } = await run(store, set('Total', '=Data!$A$1:$A$3'));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.equal(m.commands, 1);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.ok(store.names.some(x => x.name === 'Total' && x.ref === 'Data!$A$1:$A$3'));
  const again = await run(store, set('Total', '=Data!$A$1:$A$3'));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('delete an unused name: APPLIED; retry SATISFIED 0/0 (absent)', async () => {
  const store = store0();
  const { result } = await run(store, del('Free'));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(store.names.some(x => x.name === 'Free'), false);
  const again = await run(store, del('Free'));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

for (const [name, t, why, script] of [
  ['re-pointing an existing name (the source did it silently, B-D28)', set('Used', '=Other!$A$1'), 'xlsx-name-exists-with-another-reference'],
  ['deleting a name a formula uses (it would break)', del('Used'), 'xlsx-name-in-use'],
  ['an unproven names reader (never an empty list)', set('Total', '=Data!$A$1'), 'xlsx-names-unobservable', { dependencyBuild: false }],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const store = store0();
    const { result, driver } = await run(store, t, script || {});
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / the wrong reference landed -> L1 FAIL; a side effect -> L3 FAIL; nothing persisted', async () => {
  for (const fault of ['ignore', 'name-wrong-ref']) {
    const store = store0();
    const { result, driver } = await run(store, set('Total', '=Data!$A$1'), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  const side = await run(store0(), set('Total', '=Data!$A$1'), { fault: 'side-effect' });
  assert.equal(side.result.ok, false);
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planner: identifiers, the absolute-reference rule, two different names may share a task', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(set('A1', '=Data!$A$1'), /valid defined-name identifier/);
  refuse(set('9x', '=Data!$A$1'), /valid defined-name identifier/);
  refuse(set('Ok', '=Data!A1'), /ABSOLUTE reference/);
  refuse(set('Ok', '=5'), /ABSOLUTE reference/);
  refuse(set('Ok', 'Data!$A$1'), /ABSOLUTE reference/);
  refuse({ id: 'x', steps: [{ intent: 'delete_defined_name', name: 'Ok', refersTo: '=Data!$A$1' }] }, /does not support field/);
  assert.equal(planOperations(set('Ok', "='My Sheet'!$A$1:$B$2")).ok, true, 'a quoted sheet name');
  assert.equal(planOperations({ id: 'x', steps: [set('A_', '=Data!$A$1').steps[0], set('B_', '=Data!$A$2').steps[0]] }).ok, true, 'two names');
  refuse({ id: 'x', steps: [set('A_', '=Data!$A$1').steps[0], { intent: 'delete_defined_name', name: 'a_' }] }, /overlapping-writes-refused/);
});

test('the command bodies: the public AddDefName (the reference without its =) / GetDefName(..).Delete()', () => {
  const log = [];
  const Api = { AddDefName: (n, r) => log.push(['add', n, r]), GetDefName: n => ({ Delete: () => log.push(['delete', n]) }) };
  assert.deepEqual(new Function('Api', names.lower(planOperations(set('T', '=Data!$A$1')).operations[0]).body)(Api), { ok: true });
  assert.deepEqual(new Function('Api', names.lower(planOperations(del('T')).operations[0]).body)(Api), { ok: true });
  // the reference goes in WITHOUT its '=' -- a name stored with '=' never resolves (W3e; W5-W9 LIVE run 1)
  assert.deepEqual(log, [['add', 'T', 'Data!$A$1'], ['delete', 'T']]);
  assert.deepEqual(new Function('Api', names.lower(planOperations(del('T')).operations[0]).body)({ GetDefName: () => null }), { ok: false, stage: 'name-missing' });
});
