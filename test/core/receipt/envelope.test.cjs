'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { AppliedOperations, buildExecutionResult } = require('../../../src/core/execution/result.cjs');
const { buildReceipt } = require('../../../src/core/receipt/envelope.cjs');

function execution() {
  const a = new AppliedOperations();
  a.record({ operationId: 'op1', status: 'APPLIED', verification: 'LIVE_VERIFY', evidence: { x: 1 } });
  a.record({ operationId: 'op2', status: 'FAILED' });
  return buildExecutionResult({ executionOk: false, outcome: 'partial', appliedOperations: a });
}

test('the receipt is evidence-only and says where it was derived from', () => {
  const r = buildReceipt({ execution: execution(), sessionId: 's1' });
  assert.equal(r.authority, 'EVIDENCE_ONLY');
  assert.equal(r.derivedFrom, 'ExecutionResult.appliedOperations');
  assert.equal(r.operationCount, 2);
});

test('a serialization failure yields an envelope that reports its own failure', () => {
  const broken = { appliedOperations: { map() { throw new Error('boom'); } } };
  const r = buildReceipt({ execution: broken });
  assert.equal(r.serialized, false);
  assert.equal(r.serializationError, 'boom');
  assert.equal(r.authority, 'EVIDENCE_ONLY', 'even a failed receipt claims no authority');
  assert.equal(r.operationCount, 0);
});

test('the receipt is frozen — downstream cannot edit the evidence', () => {
  const r = buildReceipt({ execution: execution(), sessionId: 's1' });
  assert.throws(() => { 'use strict'; r.operationCount = 99; }, TypeError);
});
