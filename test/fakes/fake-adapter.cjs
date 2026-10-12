'use strict';

/**
 * Deterministic fake adapter — owner requirement 19.
 *
 * It fills the M0 section 5 contract with NO format semantics: its "units" are opaque ids, not
 * paragraphs, ranges, charts or slides. That is the point of M1 — the core is proven to drive the
 * whole state machine while knowing nothing about any format.
 *
 * Every behaviour a core test needs is scripted, never timed.
 */

function createFakeAdapter(script = {}) {
  const calls = { readinessProbe: 0, normalizeInventory: 0, executePrimitive: 0, verifyPrimitive: 0, verifyWholeTask: 0, enrichObservation: 0, verifyPreservation: 0 };

  return {
    kind: 'fake',
    calls,

    async readinessProbe() {
      calls.readinessProbe += 1;
      // An adapter boundary can THROW, not just report not-ready. The core must normalize this.
      if (script.readyThrows) throw new Error(script.readyThrowDetail || 'fake readiness probe threw');
      if (script.readyFails) return { ok: false, detail: 'fake never becomes ready' };
      return { ok: true, detail: 'fake ready' };
    },

    normalizeInventory(raw) {
      calls.normalizeInventory += 1;
      if (script.normalizeThrows) throw new Error(script.normalizeThrowDetail || 'fake normalizeInventory threw');
      if (script.normalizeReturnsNothing) return null;
      return { units: (raw && raw.units) || [], source: 'fake' };
    },

    // Enrichment is OPTIONAL. `enrichment: 'absent'` models XLSX, which implements no hook at all;
    // the other modes model a DOCX-shaped adapter that does.
    enrichObservation:
      script.enrichment === 'absent' || script.enrichment === undefined
        ? null
        : async function enrichObservation() {
            calls.enrichObservation += 1;
            if (script.enrichment === 'throws') throw new Error('fake enrichment threw');
            if (script.enrichment === 'fails') return { ok: false, detail: 'fake enrichment refused' };
            if (script.enrichment === 'null') return null;
            return { ok: true, enrichment: { extra: 'fake-package-semantics' } };
          },

    resolveTarget(observation, selector) {
      const matches = (observation.units || []).filter(u => u.id === selector.id);
      if (matches.length === 0) return { ok: false, target: null, selector, ambiguity: 'NOT_FOUND' };
      if (matches.length > 1) return { ok: false, target: null, selector, ambiguity: 'AMBIGUOUS' };
      return { ok: true, target: matches[0], selector, ambiguity: null };
    },

    planOperations(task) {
      return { ok: true, operations: (task.operations || []).slice(), dependencies: [] };
    },

    async executePrimitive(_handle, operation) {
      calls.executePrimitive += 1;
      const outcome = (script.primitives || {})[operation.id];
      if (outcome === 'dispatch-fail') return { ok: false, dispatched: false, authority: 'DISPATCH_ONLY' };
      return { ok: true, dispatched: true, authority: 'DISPATCH_ONLY' };
    },

    async verifyPrimitive(_handle, operation) {
      calls.verifyPrimitive += 1;
      const outcome = (script.primitives || {})[operation.id];
      if (outcome === 'verify-throws') throw new Error(`fake primitive verifier threw on ${operation.id}`);
      if (outcome === 'verify-fail') return { ok: false, evidence: { reason: 'fake verify refused' } };
      return { ok: true, evidence: { unit: operation.id, observed: 'fake-live-state' } };
    },

    async verifyWholeTask() {
      calls.verifyWholeTask += 1;
      if (script.wholeTaskThrows) throw new Error('fake whole-task verifier threw');
      if (script.wholeTaskFails) {
        return { ok: false, checks: [{ name: 'fake-terminal', ok: false }] };
      }
      return { ok: true, checks: [{ name: 'fake-terminal', ok: true }] };
    },

    // Preservation is OPTIONAL: `absent` models an adapter that declares no verifier at all.
    verifyPreservation:
      script.preservation === 'absent' || script.preservation === undefined
        ? null
        : async function verifyPreservation() {
            calls.verifyPreservation += 1;
            if (script.preservation === 'throws') throw new Error('fake preservation verifier threw');
            if (script.preservation === 'fails') return { ok: false, evidence: { changed: 'unrelated-property' } };
            return { ok: true, evidence: { unchanged: 'unrelated-property' } };
          },

    classifyIdempotency(_observation, task) {
      const state = script.classification || 'APPLY';
      return { state, evidence: { taskId: task.id || null, scripted: state } };
    },

    terminalContract: null,
  };
}

module.exports = { createFakeAdapter };
