'use strict';

/**
 * A fake document with REAL state, for the M3 executor.
 *
 * The M1/M2 fakes script outcomes; this one holds an ordered list of opaque units and applies
 * commands to it, so a stale or naive target produces a genuinely WRONG document -- which is what
 * the batch-dependency negative control needs to be able to show. It knows no format: a unit is
 * `{ key, text }`, and the commands are `set`, `append` and `insertAfter` over a PHYSICAL index.
 *
 * The physical index is the point: it is exactly the kind of locator that goes stale when an
 * earlier operation inserts before it. Targets are resolved to an index from an observation; if
 * the observation is older than a mutation, the index can name the wrong unit.
 */

function createDocument(units) {
  return { units: units.map(u => ({ key: u.key, text: u.text })) };
}

function createDocumentDriver(doc, script = {}) {
  const calls = { open: 0, close: 0, forceSave: 0, canSave: 0, readLive: 0, dispatch: 0 };
  let dispatchNo = 0;

  function apply(command) {
    if (command.kind === 'batch') { for (const c of command.commands) apply(c); return; }
    const at = command.index;
    if (!Number.isInteger(at) || at < 0 || at >= doc.units.length) throw new Error(`index ${at} out of range`);
    if (command.kind === 'set') doc.units[at].text = command.text;
    else if (command.kind === 'append') doc.units[at].text += command.suffix;
    else if (command.kind === 'insertAfter') doc.units.splice(at + 1, 0, { key: command.newKey, text: command.text });
    else throw new Error(`unknown command ${command.kind}`);
  }

  return {
    calls,
    doc,
    async open() { calls.open += 1; return { ok: true, id: 'doc-session' }; },
    // M7 save contract: a confirming transport (the generic stand-in for the XLSX driver's paired confirmation)
    async close() { calls.close += 1; return { ok: true, confirmation: calls.forceSave > 0 ? { state: 'confirmed', reason: 'forcesave-confirmed' } : null }; },
    async forceSave() { calls.forceSave += 1; return { ok: true, confirmed: true, forceSave: { state: 'confirmed', reason: 'forcesave-confirmed' } }; },
    async canSave() { calls.canSave += 1; return { ok: true, canSave: false }; },
    async readLive() {
      calls.readLive += 1;
      if (script.readLiveFailsAt === calls.readLive) return { ok: false, detail: 'scripted read failure' };
      return { ok: true, authority: 'LIVE_READ', payload: { units: doc.units.map((u, index) => ({ key: u.key, text: u.text, index })) } };
    },
    async dispatch(_handle, command) {
      calls.dispatch += 1;
      dispatchNo += 1;
      const fault = (script.faults || {})[dispatchNo];
      if (fault === 'throw-before-apply') throw new Error('transport lost before the editor applied it');
      if (fault === 'throw-after-apply') { apply(command); throw new Error('transport lost after the editor applied it'); }
      if (fault === 'refuse') return { ok: false, dispatched: true };
      apply(command);
      return { ok: true, dispatched: true };
    },
  };
}

function find(primary, key) {
  return (primary.units || []).find(u => u.key === key) || null;
}

/**
 * The adapter. Operations:
 *   { id, kind:'set',         selector:{key}|{ref},  text }
 *   { id, kind:'append',      selector:{key}|{ref},  suffix }
 *   { id, kind:'insertAfter', selector:{key}|{ref},  newKey, text }
 * plus optional `dependsOn`, `preservationRequired`, `preservationMode`, `failClosed`.
 */
function createDocumentAdapter(script = {}) {
  const calls = { resolveTarget: 0, executePrimitive: 0, executePrimitiveBatch: 0, verifyPrimitive: 0, verifyWholeTask: 0, classifyOperation: 0, verifyPreservation: 0 };

  const lower = (op, target) => {
    if (op.kind === 'set') return { kind: 'set', index: target.index, text: op.text };
    if (op.kind === 'append') return { kind: 'append', index: target.index, suffix: op.suffix };
    if (op.kind === 'insertAfter') return { kind: 'insertAfter', index: target.index, newKey: op.newKey, text: op.text };
    throw new Error(`cannot lower ${op.kind}`);
  };

  const adapter = {
    kind: 'fake',
    calls,
    async readinessProbe() { return { ok: true }; },
    normalizeInventory(raw) { return { units: (raw && raw.units) || [] }; },
    planOperations(task) {
      if (script.planFailDetail) return { ok: false, detail: script.planFailDetail };
      return { ok: true, operations: (task.operations || []).slice(), dependencies: [] };
    },

    // Per-operation idempotency from the CURRENT observation.
    classifyOperation(primary, op) {
      calls.classifyOperation += 1;
      if (op.failClosed) return { state: 'FAIL_CLOSED', evidence: { reason: 'scripted' } };
      if (op.kind === 'insertAfter') {
        const existing = find(primary, op.newKey);
        return existing && existing.text === op.text ? { state: 'SATISFIED', evidence: { key: op.newKey, text: op.text } } : { state: 'APPLY', evidence: null };
      }
      if (op.kind === 'set' && op.selector && op.selector.key) {
        const u = find(primary, op.selector.key);
        if (u && u.text === op.text) return { state: 'SATISFIED', evidence: { key: u.key, text: u.text } };
      }
      return { state: 'APPLY', evidence: null };
    },
    classifyIdempotency() { return { state: 'APPLY', evidence: null }; },

    // A ref names the earlier operation's SEMANTIC result (its key); it is found again here, in the
    // observation being resolved against. Never an index carried over from another state.
    resolveTarget(primary, selector, ctx = {}) {
      calls.resolveTarget += 1;
      let key = selector.key;
      if (selector.ref) {
        const r = ctx.refs && ctx.refs[selector.ref];
        key = r && r.evidence && r.evidence.key;
        if (!key) return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_RESULT' };
      }
      const matches = (primary.units || []).filter(u => u.key === key);
      if (matches.length !== 1) return { ok: false, target: null, selector, ambiguity: matches.length ? 'AMBIGUOUS' : 'NOT_FOUND' };
      return { ok: true, target: { key, index: matches[0].index }, selector, ambiguity: null };
    },

    async executePrimitive(_handle, op, { dispatch, target }) {
      calls.executePrimitive += 1;
      if (op.lowerFails) return { ok: false, dispatched: false, detail: 'lowering refused before dispatch' };
      return dispatch(lower(op, target));
    },

    async executePrimitiveBatch(_handle, items, { dispatch }) {
      calls.executePrimitiveBatch += 1;
      if (script.noBatchHook) throw new Error('scripted: batch hook must not be used');
      return dispatch({ kind: 'batch', commands: items.map(it => lower(it.operation, it.target)) });
    },

    async verifyPrimitive(_handle, op, _before, after) {
      calls.verifyPrimitive += 1;
      if (op.forceVerifyFail) return { ok: false, evidence: { reason: 'scripted verification failure' } };
      const key = op.kind === 'insertAfter' ? op.newKey : null;
      if (op.kind === 'insertAfter') {
        const u = find(after.primary, key);
        return u && u.text === op.text ? { ok: true, evidence: { key, text: u.text } } : { ok: false, evidence: { key, found: !!u } };
      }
      // set/append: the unit the operation was ABOUT (by key, found freshly) must carry the effect.
      const aboutKey = op.selector.key || op.expectKey;
      const u = aboutKey ? find(after.primary, aboutKey) : null;
      if (!u) return { ok: false, evidence: { reason: 'target unit not found after dispatch', aboutKey } };
      const ok = op.kind === 'set' ? u.text === op.text : u.text.endsWith(op.suffix);
      return ok ? { ok: true, evidence: { key: u.key, text: u.text } } : { ok: false, evidence: { key: u.key, text: u.text } };
    },

    async verifyWholeTask(_handle, task, finalObservation) {
      calls.verifyWholeTask += 1;
      const checks = (task.expect || []).map(e => {
        const u = find(finalObservation.primary, e.key);
        return { name: `unit ${e.key}`, ok: !!u && u.text === e.text };
      });
      return { ok: checks.every(c => c.ok), checks };
    },

    verifyPreservation: script.preservationFailsFor
      ? async (_h, op) => {
        calls.verifyPreservation += 1;
        return script.preservationFailsFor.includes(op.id) ? { ok: false, evidence: { changed: 'scripted' } } : { ok: true, evidence: {} };
      }
      : null,

    terminalContract: null,
  };
  return adapter;
}

function snapshot(doc) {
  return doc.units.map(u => `${u.key}=${u.text}`);
}

module.exports = { createDocument, createDocumentDriver, createDocumentAdapter, snapshot };
