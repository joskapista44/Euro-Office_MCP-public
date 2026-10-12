'use strict';

/**
 * The adapter contract as a pure interface — M0 section 5 (M0.3), M1.5.
 *
 * NOTE ON THE TREE: M0 section 3 does not name a file for the contract itself, while M0 section 21
 * (M1.5) requires "the adapter contract as a pure interface" as an M1 deliverable. This module is
 * that interface and nothing else — it holds no execution logic and no format knowledge. The
 * addition is recorded in docs/M1-COMMON-PERSISTENT-SESSION.md rather than being made silently.
 *
 * Nothing here names a paragraph, a list, a table, a cell, a range, a chart, a pivot or a slide.
 * Those live inside an adapter's own implementation.
 */

/** Required on every adapter. An adapter missing one of these cannot drive the lifecycle. */
const REQUIRED = Object.freeze([
  'kind',
  'readinessProbe',
  'normalizeInventory',
  'resolveTarget',
  'planOperations',
  'executePrimitive',
  'verifyPrimitive',
  'verifyWholeTask',
  'classifyIdempotency',
]);

/** Optional. `null` is a legitimate implementation — XLSX returns null for enrichment. */
const OPTIONAL = Object.freeze([
  'enrichObservation',
  'verifyPreservation',
  'terminalContract',
  // M3, additive. Per-operation idempotency from the CURRENT observation; without it the executor
  // falls back to the task-level classifyIdempotency, once, from the planning observation.
  'classifyOperation',
  // M3, additive. Lower several independent operations into ONE physical dispatch. Without it the
  // batch strategy still groups reads and verification, but dispatches per operation.
  'executePrimitiveBatch',
]);

/**
 * Forbidden on an adapter. Lifecycle ownership is the core's: an adapter that exposed any of these
 * would be a second lifecycle implementation, which is the thing the whole architecture exists to
 * prevent (owner requirement 16, M0 5 prohibitions).
 */
const FORBIDDEN = Object.freeze([
  'openSession',
  'closeSession',
  'open',
  'close',
  'persist',
  'forceSave',
  'canSave',
]);

const KINDS = Object.freeze(['docx', 'xlsx', 'pptx', 'fake']);

/**
 * Validates an adapter against the contract. Fail-closed: it returns the reasons rather than
 * throwing on the first one, so a caller sees the whole gap at once.
 */
function validateAdapter(adapter) {
  const problems = [];
  if (!adapter || typeof adapter !== 'object') {
    return { ok: false, problems: ['adapter is not an object'] };
  }

  if (!KINDS.includes(adapter.kind)) {
    problems.push(`adapter.kind must be one of ${KINDS.join('|')}, got ${String(adapter.kind)}`);
  }
  for (const name of REQUIRED) {
    if (name === 'kind') continue;
    if (typeof adapter[name] !== 'function') problems.push(`missing required adapter method: ${name}`);
  }
  for (const name of OPTIONAL) {
    if (name in adapter && adapter[name] !== null && typeof adapter[name] !== 'function') {
      problems.push(`optional adapter member must be a function or null: ${name}`);
    }
  }
  for (const name of FORBIDDEN) {
    if (name in adapter) problems.push(`adapter must not expose lifecycle member: ${name}`);
  }

  return { ok: problems.length === 0, problems };
}

module.exports = { validateAdapter, REQUIRED, OPTIONAL, FORBIDDEN, KINDS };
