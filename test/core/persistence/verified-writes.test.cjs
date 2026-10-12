'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { AppliedOperations, buildExecutionResult } = require('../../../src/core/execution/result.cjs');
const { hasVerifiedAppliedWrites, verifiedAppliedWrites } = require('../../../src/core/persistence/verified-writes.cjs');

const resultWith = entries => {
  const a = new AppliedOperations();
  for (const e of entries) a.record(e);
  return buildExecutionResult({ executionOk: false, outcome: 'test', appliedOperations: a });
};

test('APPLIED + LIVE_VERIFY makes the predicate true, regardless of executionOk', () => {
  const r = resultWith([{ operationId: 'op1', status: 'APPLIED', verification: 'LIVE_VERIFY' }]);
  assert.equal(r.executionOk, false, 'the fixture is a FAILED execution on purpose');
  assert.equal(hasVerifiedAppliedWrites(r), true, 'M0 8.1: the barrier is mandatory regardless of executionOk');
});

test('APPLIED without LIVE_VERIFY does not fire the barrier', () => {
  assert.equal(hasVerifiedAppliedWrites(resultWith([{ operationId: 'op1', status: 'APPLIED', verification: 'DISPATCH_ONLY' }])), false);
});

test('SATISFIED and FAILED never fire the barrier', () => {
  assert.equal(hasVerifiedAppliedWrites(resultWith([{ operationId: 'op1', status: 'SATISFIED' }])), false);
  assert.equal(hasVerifiedAppliedWrites(resultWith([{ operationId: 'op1', status: 'FAILED' }])), false);
});

test('an empty or malformed execution state fails closed to false', () => {
  assert.equal(hasVerifiedAppliedWrites(resultWith([])), false);
  assert.equal(hasVerifiedAppliedWrites(null), false);
  assert.equal(hasVerifiedAppliedWrites({}), false);
});

test('the predicate reads appliedOperations, not the receipt', () => {
  // A result whose receipt claims a verified write while the authoritative state has none.
  const r = Object.freeze({
    ...resultWith([{ operationId: 'op1', status: 'FAILED' }]),
    receiptEvidence: { operations: [{ operationId: 'ghost', status: 'APPLIED', verification: 'LIVE_VERIFY' }] },
  });
  assert.equal(hasVerifiedAppliedWrites(r), false, 'a receipt must never create a persistence decision');
});

test('verifiedAppliedWrites returns the entries that justified the decision', () => {
  const r = resultWith([
    { operationId: 'op1', status: 'APPLIED', verification: 'LIVE_VERIFY' },
    { operationId: 'op2', status: 'FAILED' },
    { operationId: 'op3', status: 'APPLIED', verification: 'LIVE_VERIFY' },
  ]);
  assert.deepEqual(verifiedAppliedWrites(r).map(e => e.operationId), ['op1', 'op3']);
});
