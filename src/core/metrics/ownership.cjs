'use strict';

const { COUNTERS } = require('./collect.cjs');

/**
 * Counter ownership — owner requirement 17.
 *
 * Each common counter is incremented by EXACTLY ONE core module. This matters because the two
 * sources meant different things by similar names (M0 section 12): DOCX counted read/write
 * interactions with no writes counter, XLSX counted writes with no interaction counters. A
 * counter that two layers may both increment has no single meaning, and no number derived from
 * it is comparable across formats.
 *
 * An ADAPTER may never increment a common counter. An adapter that counted its own reads would
 * be defining what `freshReads` means for its format, which is exactly the divergence the common
 * schema exists to end. The static gate asserts no adapter calls `metrics.count`.
 */

const OWNER = Object.freeze({
  freshReads: 'core/session/handle.cjs (readLive)',
  packageReads: 'core/session/handle.cjs (readPackage) -- physical package I/O, M4',
  commands: 'core/session/handle.cjs (dispatch)',
  enrichmentInvocations: 'core/observation/enrichment.cjs (runEnrichment)',
  verificationInvocations: 'core/verification/orchestrate.cjs (runHook)',
  persistencePolls: 'core/persistence/barrier.cjs (probeCanSave)',
  writes: 'the task executor, on reaching APPLIED',
  operationsPlanned: 'the task executor, per planned operation',
  operationsApplied: 'the task executor, on reaching APPLIED',
  operationsSatisfied: 'the task executor, on a SATISFIED classification',
  operationsUncertain: 'the task executor, on DISPATCHED_OUTCOME_UNKNOWN',
  operationsSkipped: 'the task executor, on SKIPPED_DUE_TO_PRIOR_FAILURE',
  batches: 'the task executor (core/execution/executor.cjs), per executed group',
  batchOperations: 'the task executor, per operation placed in an executed group',
  dependencyBarriers: 'the task executor, per group boundary forced by a dependency',
});

/**
 * The call-count consistency contract (owner requirement 16), stated BEFORE it is tested.
 *
 * PHYSICAL I/O vs SEMANTIC ORCHESTRATION. These are separate accounts, and mixing them was a
 * measured defect (owner requirement R3):
 *
 *   driver.readLive  ==  freshReads
 *       every readLive goes through core/session/handle.cjs, which counts it -- INCLUDING a
 *       readLive performed by an adapter verifier inside its hook. There is one physical counter
 *       and one place that increments it.
 *
 *   driver.dispatch  ==  commands
 *   driver.canSave   ==  persistencePolls
 *   adapter.enrichObservation invocations  ==  enrichmentInvocations
 *       counted before the call, so a hook that throws is still counted: the package WAS read.
 *
 *   core verifier hook invocations  ==  verificationInvocations
 *       *** NOT A PHYSICAL READ COUNT. *** It counts how many times the core invoked a verifier,
 *       regardless of what the hook does inside. A real adapter verifier may call readLive(),
 *       and that read is already in freshReads.
 *
 * The physical totals therefore EXCLUDE verificationInvocations:
 *
 *   editorInteractions        == freshReads + commands + persistencePolls
 *   packageInteractions       == physical package transport calls (0 until common package I/O exists)
 *   totalExternalInteractions == editorInteractions + packageInteractions
 *
 * The earlier form added verificationInvocations into editorInteractions. A verifier that read
 * through readLive() was then counted twice -- once physically, once semantically -- and three
 * physical reads were reported as five interactions.
 *
 * What makes this worth stating rather than just fixing: the M2 documentation ALREADY said the
 * naive identity `readLive == freshReads + verification reads` was wrong, and explained why. The
 * formula in the code added them together anyway. Writing the lesson down did not protect the
 * line of code two files away.
 */
const CONSISTENCY = Object.freeze({
  'driver.readLive': 'freshReads',
  'driver.readPackage': 'packageReads',
  'driver.dispatch': 'commands',
  'driver.canSave': 'persistencePolls',
  'adapter.enrichObservation': 'enrichmentInvocations',
  'core verifier hook invocations': 'verificationInvocations',
});

/** Every counter in the schema must have exactly one named owner — no unowned counters. */
function unownedCounters() {
  return COUNTERS.filter(c => !OWNER[c]);
}

/** And no owner may be listed for a counter the schema does not have. */
function orphanOwners() {
  return Object.keys(OWNER).filter(c => !COUNTERS.includes(c));
}

module.exports = { OWNER, CONSISTENCY, unownedCounters, orphanOwners };
