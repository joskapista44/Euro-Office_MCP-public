'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createFakeSession } = require('../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../fakes/fake-adapter.cjs');
const { createFakeExecutor } = require('../fakes/fake-executor.cjs');

/**
 * M7 save contract on the core (owner GO 2026-10-02, owner channel):
 *   1. canSave false alone is at most EDITOR_CLEAN;
 *   2. the driver's own paired, successful forceSave -> FORCESAVE_CONFIRMED;
 *   3. at close the confirmation is RE-EVALUATED on the observer's state then (a later send / anomaly does not stay
 *      behind a confirmed:true returned earlier);
 *   4. a not-verified close, a superseded confirmation, a bad or missing answer -> no final confirmed save; no rewrite;
 *   5. no general PERSISTED: the authority is never raised; the stored content is never claimed as checked.
 * editOk (the edit) is separate from ok (the whole task): a task that needed a save is not complete on EDITOR_CLEAN;
 * a write-free SATISFIED retry starts no save just to get a confirmation.
 */

function run({ sessionScript = {}, adapterScript = {}, task = { id: 't', operations: [{ id: 'op1' }] } } = {}) {
  let t = 0;
  const driver = createFakeSession(sessionScript);
  const options = { driver, fileId: 'f', now: () => (t += 1), ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 } };
  return withSession(options, createFakeAdapter(adapterScript), createFakeExecutor(task, {})).then(result => ({ result, driver }));
}

const save = r => [r.persistenceOutcome, r.persistenceState, r.saveEvidence.reason, r.editOk, r.ok, r.outcome];

test('C2: the paired confirmation at the barrier AND at a verified close -> FORCESAVE_CONFIRMED, ok, no PERSISTED', async () => {
  const { result } = await run();
  assert.deepEqual(save(result), ['SUCCESS', 'FORCESAVE_CONFIRMED', 'forcesave-confirmed', true, true, 'task-complete']);
  assert.equal(result.saveEvidence.meaning, 'the server reported the save of this acknowledged editor state as successful');
  assert.equal(result.saveEvidence.storedContentVerified, false, 'the product path never checks the stored content');
  assert.notEqual(result.authority, 'PERSISTED');
});

test('C1: canSave false with a driver that has no save observer -> EDITOR_CLEAN; the edit is ok, the task is not', async () => {
  const { result } = await run({ sessionScript: { noSaveConfirmation: true } });
  assert.deepEqual(save(result), ['SUCCESS', 'EDITOR_CLEAN', 'barrier:driver-provides-no-save-confirmation', true, false, 'task-save-not-confirmed']);
  assert.equal(result.saveEvidence.meaning, null);
});

test('C4: a bad or missing answer at the barrier -> no confirmed save, no rewrite (one forceSave only)', async () => {
  for (const reason of ['start-code-4', 'completion-not-successful', 'timeout:awaiting-completion', 'ambiguous:2-requests-in-the-operation']) {
    const { result, driver } = await run({ sessionScript: { forceSaveUnconfirmed: reason } });
    assert.deepEqual(save(result), ['SUCCESS', 'EDITOR_CLEAN', `barrier:${reason}`, true, false, 'task-save-not-confirmed'], reason);
    assert.equal(driver.calls.forceSave, 1, 'never a second save attempt');
  }
});

test('C3: confirmed at the barrier, superseded or anomalous when re-evaluated at close -> no confirmed save', async () => {
  for (const reason of ['superseded-by-a-newer-save', 'anomaly:uninterpreted-ds-frame:text']) {
    const { result } = await run({ sessionScript: { closeConfirmation: { state: 'unconfirmed', reason } } });
    assert.deepEqual(save(result), ['SUCCESS', 'EDITOR_CLEAN', `close:${reason}`, true, false, 'task-save-not-confirmed'], reason);
  }
});

test('C4: a not-verified close -> no confirmed save (even with a confirmation still standing)', async () => {
  const { result } = await run({ sessionScript: { closeFails: true } });
  assert.deepEqual(save(result), ['SUCCESS', 'EDITOR_CLEAN', 'close-not-verified', true, false, 'task-cleanup-failed']);
  const thrown = await run({ sessionScript: { closeThrows: true } });
  assert.equal(thrown.result.persistenceState, 'EDITOR_CLEAN');
});

test('C3: a close that reports no confirmation at all (barrier said confirmed) -> no confirmed save', async () => {
  const driver = createFakeSession({});
  const close = driver.close.bind(driver);
  driver.close = async h => ({ ...(await close(h)), confirmation: null });
  let t = 0;
  const result = await withSession({ driver, fileId: 'f', now: () => (t += 1), ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 } }, createFakeAdapter({}), createFakeExecutor({ id: 't', operations: [{ id: 'op1' }] }, {}));
  assert.deepEqual(save(result), ['SUCCESS', 'EDITOR_CLEAN', 'close:no-confirmation-reported', true, false, 'task-save-not-confirmed']);
});

test('the forceSave precondition failed (no request made) -> FAILED, as any barrier failure', async () => {
  const { result } = await run({ sessionScript: { forceSaveUnavailable: true } });
  assert.deepEqual([result.persistenceOutcome, result.persistenceState, result.ok, result.outcome], ['FAILED', 'FAILED', false, 'task-persistence-barrier-failed']);
  assert.equal(result.authority, 'PRIMITIVE_LIVE_VERIFY_ONLY');
});

test('no-write SATISFIED retry: NOT_REQUIRED, ok, and no save is started to get a confirmation', async () => {
  const { result, driver } = await run({ adapterScript: { classification: 'SATISFIED' } });
  assert.deepEqual(save(result), ['NOT_REQUIRED', 'NOT_REQUIRED', 'no-verified-write', true, true, 'task-complete']);
  assert.equal(driver.calls.forceSave, 0);
  // and the same with a driver that has no save observer: nothing to confirm, nothing missing
  const bare = await run({ adapterScript: { classification: 'SATISFIED' }, sessionScript: { noSaveConfirmation: true } });
  assert.equal(bare.result.ok, true);
});
