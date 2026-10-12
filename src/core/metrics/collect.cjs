'use strict';

const { Clock, BOUNDARIES } = require('./clock.cjs');

/**
 * The single metric schema — M0 section 12.1, with the interaction split of owner requirement R3.
 *
 * PHYSICAL I/O AND SEMANTIC ORCHESTRATION ARE COUNTED SEPARATELY, and this is the correction that
 * matters most in this file:
 *
 *   editorInteractions   physical/common transport interactions ONLY
 *                        freshReads + commands + persistencePolls
 *   verificationInvocations  how many times the core invoked an adapter verifier hook.
 *                        NOT a physical read count -- whatever the hook does inside is the
 *                        adapter's business, and it may well call readLive() itself.
 *   packageInteractions  physical package transport interactions (0 until a common package I/O surface exists)
 *   totalExternalInteractions  editorInteractions + packageInteractions
 *
 * The earlier schema added verificationReads into editorInteractions, which DOUBLE-COUNTED a
 * verifier that read through readLive(): once as freshReads (counted by the handle) and again as
 * a verification read. Measured: three physical reads reported as five interactions. The old name
 * `verificationReads` is gone rather than redefined -- a consumer of the old name now gets
 * `undefined`, which is loud, instead of a plausible wrong number, which is not.
 */

const COUNTERS = Object.freeze([
  'freshReads',
  'packageReads',
  'enrichmentInvocations',
  'commands',
  'writes',
  'verificationInvocations',
  'persistencePolls',
  'operationsPlanned',
  'operationsApplied',
  'operationsSatisfied',
  // M3
  'operationsUncertain',
  'operationsSkipped',
  'batches',
  'batchOperations',
  'dependencyBarriers',
]);

class Metrics {
  constructor({ now } = {}) {
    this.clock = new Clock({ now });
    for (const counter of COUNTERS) this[counter] = 0;
  }

  /** Counter increments go through one door so an unknown counter name fails loudly. */
  count(counter, by = 1) {
    if (!COUNTERS.includes(counter)) {
      throw new Error(`unknown metric counter: ${counter}`);
    }
    this[counter] += by;
    return this;
  }
}

/**
 * Freezes the schema of section 12.1. `operationsFailed` is derived, not counted, exactly as the
 * specification states: planned - applied - satisfied.
 */
function collectMetrics(metrics) {
  const spans = metrics.clock.spans();
  // PHYSICAL transport interactions only. A verifier invocation is NOT one of these: it is
  // semantic orchestration, and any physical read it performs is already counted by the handle.
  const editorInteractions = metrics.freshReads + metrics.commands + metrics.persistencePolls;
  // Physical package transport (M4): counted at core/session/handle.cjs readPackage, the only
  // common package I/O boundary. An enrichment INVOCATION is not a package read.
  const packageInteractions = metrics.packageReads;

  return {
    [BOUNDARIES.SESSION_OPEN]: spans[BOUNDARIES.SESSION_OPEN],
    [BOUNDARIES.READY]: spans[BOUNDARIES.READY],
    [BOUNDARIES.TASK_ELAPSED]: spans[BOUNDARIES.TASK_ELAPSED],
    [BOUNDARIES.TOTAL_ELAPSED]: spans[BOUNDARIES.TOTAL_ELAPSED],

    freshReads: metrics.freshReads,
    packageReads: metrics.packageReads,
    enrichmentInvocations: metrics.enrichmentInvocations,
    commands: metrics.commands,
    writes: metrics.writes,
    persistencePolls: metrics.persistencePolls,

    // Semantic orchestration, deliberately outside every physical total.
    verificationInvocations: metrics.verificationInvocations,

    editorInteractions,
    packageInteractions,
    totalExternalInteractions: editorInteractions + packageInteractions,

    operationsPlanned: metrics.operationsPlanned,
    operationsApplied: metrics.operationsApplied,
    operationsSatisfied: metrics.operationsSatisfied,
    operationsUncertain: metrics.operationsUncertain,
    operationsSkipped: metrics.operationsSkipped,
    // Derived, as in M0 12.1, and now net of the two M3 buckets. The AUTHORITATIVE accounting,
    // with a real identity check, is ExecutionResult.operationAccounting (plan vs register).
    operationsFailed:
      metrics.operationsPlanned - metrics.operationsApplied - metrics.operationsSatisfied -
      metrics.operationsUncertain - metrics.operationsSkipped,

    // M3 batch accounting. SEMANTIC operation counts are never physical command counts:
    // `commands` above is the physical dispatch count; these describe how operations were grouped.
    batches: metrics.batches,
    batchOperations: metrics.batchOperations,
    dependencyBarriers: metrics.dependencyBarriers,
  };
}

module.exports = { Metrics, collectMetrics, COUNTERS };
