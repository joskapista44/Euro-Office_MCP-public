'use strict';

/**
 * Receipt envelope — M0 section 4 (`buildReceipt`) and invariant 14.
 *
 * The receipt is DERIVED evidence. The direction is one-way:
 *
 *   execution state  ->  persistence decision
 *   execution state  ->  receipt
 *
 * and never receipt -> persistence decision (owner requirement 8).
 *
 * This module therefore imports nothing from core/persistence, and core/persistence imports
 * nothing from here. The static gate in test/static asserts both directions, so the prohibition
 * is measured rather than merely written down.
 *
 * Serialization here may fail. That failure is recorded inside the envelope and must never
 * propagate into a persistence decision (owner requirement 9).
 */

function buildReceipt({ execution, task = null, sessionId = null }) {
  try {
    const entries = (execution.appliedOperations || []).map(entry => ({
      operationId: entry.operationId,
      status: entry.status,
      verification: entry.verification,
      evidence: entry.evidence,
      // Derived copy only. The persistence decision reads the register, never this field.
      uncertainty: entry.uncertainty || null,
    }));

    return Object.freeze({
      kind: 'ReceiptEvidence',
      authority: 'EVIDENCE_ONLY',
      serialized: true,
      sessionId,
      taskId: (task && task.id) || null,
      executionOk: execution.executionOk,
      outcome: execution.outcome,
      operations: Object.freeze(entries),
      operationCount: entries.length,
      // A hash of nothing but the derived view — deliberately not an identity anyone may act on.
      derivedFrom: 'ExecutionResult.appliedOperations',
    });
  } catch (err) {
    // A corrupt or failed receipt is still an envelope. It reports its own failure and stays
    // authority-free, so no decision path can mistake its absence for an absence of writes.
    return Object.freeze({
      kind: 'ReceiptEvidence',
      authority: 'EVIDENCE_ONLY',
      serialized: false,
      serializationError: err.message,
      sessionId,
      operations: Object.freeze([]),
      operationCount: 0,
      derivedFrom: 'ExecutionResult.appliedOperations',
    });
  }
}

module.exports = { buildReceipt };
