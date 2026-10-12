'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { persist } = require('../../../src/core/persistence/barrier.cjs');
const { bindDriver } = require('../../../src/core/session/internal.cjs');
const { SessionHandle } = require('../../../src/core/session/handle.cjs');
const { Metrics } = require('../../../src/core/metrics/collect.cjs');
const { createFakeSession } = require('../../fakes/fake-session.cjs');

function handleWith(script) {
  let t = 0;
  const metrics = new Metrics({ now: () => (t += 1) });
  const handle = new SessionHandle({ id: 'h1', fileId: 'f1', adapterKind: 'fake', metrics });
  const driver = createFakeSession(script);
  bindDriver(handle, driver);
  return { handle, driver, metrics };
}

const opts = { wait: async () => {}, timeoutMs: 500, pollMs: 100 };

test('success: forceSave then canSave false — EDITOR_CLEAN, the driver confirmation carried, never PERSISTED', async () => {
  const { handle, driver } = handleWith({ canSaveSequence: [true, false] });
  const r = await persist(handle, opts);

  assert.equal(r.ok, true);
  assert.equal(r.outcome, 'forceSave-can-save-false');
  assert.equal(r.editorClean, true);
  assert.equal(r.authority, undefined, 'the barrier grants no authority: there is no general PERSISTED');
  assert.deepEqual(r.confirmation, { state: 'confirmed', reason: 'forcesave-confirmed' });
  // a driver without a save observer: nothing is invented for it
  const bare = handleWith({ canSaveSequence: [false], noSaveConfirmation: true });
  assert.deepEqual((await persist(bare.handle, opts)).confirmation, { state: 'unconfirmed', reason: 'driver-provides-no-save-confirmation' });
  assert.equal(driver.calls.forceSave, 1);
  assert.ok(r.reads >= 2, 'the before-probe and at least one poll are counted');
});

test('timeout: canSave never goes false — never reported as persisted', async () => {
  const { handle } = handleWith({ canSaveSequence: [true] });
  const r = await persist(handle, opts);

  assert.equal(r.ok, false);
  assert.equal(r.outcome, 'forceSave-timeout');
  assert.notEqual(r.authority, 'PERSISTED');
  assert.equal(r.authority, 'PRIMITIVE_LIVE_VERIFY_ONLY');
  assert.equal(r.writeAllowed, false);
});

test('unavailable: forceSave refused — fails closed before any polling', async () => {
  const { handle, driver } = handleWith({ forceSaveUnavailable: true });
  const r = await persist(handle, opts);

  assert.equal(r.ok, false);
  assert.equal(r.outcome, 'forceSave-unavailable');
  assert.equal(r.after, null);
  assert.equal(driver.calls.canSave, 1, 'only the before-probe ran');
});

test('a broken canSave probe fails closed rather than being read as "nothing to save"', async () => {
  const { handle } = handleWith({ canSaveBroken: true });
  const r = await persist(handle, opts);
  assert.equal(r.ok, false);
  assert.equal(r.outcome, 'forceSave-timeout');
});

test('every canSave probe is counted as a persistencePoll', async () => {
  const { handle, metrics, driver } = handleWith({ canSaveSequence: [true, true, false] });
  await persist(handle, opts);
  assert.equal(metrics.persistencePolls, driver.calls.canSave);
  assert.ok(metrics.persistencePolls >= 3);
});

test('the timeout is bounded by attempts x pollMs, not by wall-clock sleeping', async () => {
  let waits = 0;
  const { handle } = handleWith({ canSaveSequence: [true] });
  const r = await persist(handle, { wait: async () => { waits += 1; }, timeoutMs: 500, pollMs: 100 });

  assert.equal(r.ok, false);
  assert.equal(waits, 4, 'five attempts, four waits between them — a bounded predicate loop');
});
