'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { validatePlan, PRESERVATION_MODE } = require('../../src/core/execution/plan.cjs');
const { classifyDispatchAttempt } = require('../../src/core/execution/uncertainty.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { createDocument, createDocumentDriver, createDocumentAdapter } = require('../fakes/fake-document.cjs');

/**
 * Preservation placement (owner review R1 of the M4 first checkpoint), on the common executor and a
 * format-free fake. The core never decides the placement: the planner declares it per operation
 * (`preservationMode`), and the core enforces the ORDER.
 *
 *   A  PRIMITIVE_INTEGRAL, L1 PASS + preservation PASS   -> APPLIED + LIVE_VERIFY
 *   B  PRIMITIVE_INTEGRAL, L1 PASS + preservation FAIL   -> NOT APPLIED -> DISPATCHED_OUTCOME_UNKNOWN
 *   C  PRIMITIVE_INTEGRAL, L1 FAIL                       -> uncertain, preservation never invoked
 *   E  NONE                                              -> preservation never invoked
 *
 * D (POST_PRIMITIVE failure -> APPLIED + task FAIL + barrier) is NOT added here: the 28/28 DOCX
 * audit found no intent that uses it. The placement stays the meaning of the M3 form
 * `preservationRequired: true`, proven by m3-executor.test.cjs case G.
 */

const units = () => [{ key: 'u1', text: 'one' }, { key: 'u2', text: 'two' }];
const set = (id, key, text, extra = {}) => ({ id, kind: 'set', selector: { key }, text, ...extra });
const INTEGRAL = { preservationMode: PRESERVATION_MODE.PRIMITIVE_INTEGRAL };

async function runTask(operations, adapterScript = {}) {
  let t = 0;
  const doc = createDocument(units());
  const driver = createDocumentDriver(doc, {});
  const adapter = createDocumentAdapter(adapterScript);
  const options = {
    driver, fileId: 'doc', now: () => (t += 1),
    ready: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
    persistence: { wait: async () => {}, timeoutMs: 500, pollMs: 100 },
  };
  const result = await withSession(options, adapter, createExecutor({ id: 'pp', operations }));
  assert.equal(driver.calls.open, 1);
  assert.equal(driver.calls.close, 1);
  assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, adapter, exec: result.execution };
}

const entry = (r, id) => r.execution.appliedOperations.find(e => e.operationId === id);

test('A. integral: primitive PASS + preservation PASS -> APPLIED + LIVE_VERIFY, then barrier', async () => {
  const r = await runTask([set('a', 'u1', 'A', INTEGRAL)], { preservationFailsFor: [] });
  assert.equal(r.result.ok, true);
  assert.equal(entry(r.result, 'a').status, OPERATION_STATUS.APPLIED);
  assert.equal(entry(r.result, 'a').verification, 'LIVE_VERIFY');
  assert.equal(r.adapter.calls.verifyPreservation, 1);
  assert.equal(r.exec.preservationVerification[0].verdict, 'PASS');
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
});

test('B. integral: primitive PASS + preservation FAIL -> NOT APPLIED, uncertain; uncertain-only -> no barrier', async () => {
  const r = await runTask([set('a', 'u1', 'A', INTEGRAL)], { preservationFailsFor: ['a'] });
  const e = entry(r.result, 'a');
  assert.equal(r.result.ok, false);
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.notEqual(e.verification, 'LIVE_VERIFY', 'never a false APPLIED_VERIFIED');
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  assert.equal(r.result.metrics.writes, 0);
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
  assert.equal(r.driver.calls.forceSave, 0);
});

test('B2. integral failure after a verified prefix -> PERSISTED_UNCERTAIN (persistence is not rollback)', async () => {
  const r = await runTask([set('a', 'u1', 'A', INTEGRAL), set('b', 'u2', 'B', INTEGRAL)], { preservationFailsFor: ['b'] });
  assert.equal(entry(r.result, 'a').status, OPERATION_STATUS.APPLIED);
  assert.equal(entry(r.result, 'b').status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(r.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(r.driver.calls.forceSave, 1);
});

test('C. integral: primitive FAIL -> uncertain, preservation never invoked', async () => {
  const r = await runTask([set('a', 'u1', 'A', { ...INTEGRAL, forceVerifyFail: true })], { preservationFailsFor: [] });
  const e = entry(r.result, 'a');
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
  assert.equal(r.adapter.calls.verifyPreservation, 0);
});

test('E. NONE -> no preservation invocation, even though the adapter has a hook', async () => {
  const r = await runTask([set('a', 'u1', 'A', { preservationMode: PRESERVATION_MODE.NONE })], { preservationFailsFor: ['a'] });
  assert.equal(r.result.ok, true);
  assert.equal(entry(r.result, 'a').status, OPERATION_STATUS.APPLIED);
  assert.equal(r.adapter.calls.verifyPreservation, 0);
  // The absent field means NONE as well.
  const absent = await runTask([set('a', 'u1', 'A')], { preservationFailsFor: ['a'] });
  assert.equal(absent.adapter.calls.verifyPreservation, 0);
});

test('an integral operation with no preservation verifier -> not APPLIED (a required proof that cannot run fails)', async () => {
  const r = await runTask([set('a', 'u1', 'A', INTEGRAL)]); // fake adapter without verifyPreservation
  assert.equal(entry(r.result, 'a').status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(entry(r.result, 'a').uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
});

test('the classifier alone refuses APPLIED for an integral operation whose preservation verdict is missing', () => {
  const base = { generationBefore: 1, generationAfter: 2, dispatchResult: { ok: true }, verification: { verdict: 'PASS' } };
  assert.equal(classifyDispatchAttempt({ ...base, preservationMode: 'PRIMITIVE_INTEGRAL' }).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(classifyDispatchAttempt({ ...base, preservationMode: 'PRIMITIVE_INTEGRAL', integralPreservation: { verdict: 'PASS' } }).status, OPERATION_STATUS.APPLIED);
  // NEGATIVE CONTROL: the same inputs without the integral mode are APPLIED.
  assert.equal(classifyDispatchAttempt(base).status, OPERATION_STATUS.APPLIED);
});

test('validatePlan: an unknown mode and a mode contradicting preservationRequired are refused', () => {
  assert.equal(validatePlan([{ id: 'a', preservationMode: 'SOMETIMES' }]).ok, false);
  assert.equal(validatePlan([{ id: 'a', preservationMode: 'PRIMITIVE_INTEGRAL', preservationRequired: true }]).ok, false);
  assert.equal(validatePlan([{ id: 'a', preservationMode: 'POST_PRIMITIVE', preservationRequired: true }]).ok, true);
  assert.equal(validatePlan([{ id: 'a', preservationMode: 'PRIMITIVE_INTEGRAL' }]).ok, true);
});
