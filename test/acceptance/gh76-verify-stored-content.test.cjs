'use strict';

/**
 * GH #76 MODEL 1 (owner decision 2026-10-18): the optional, task-level `verifyStoredContent` on
 * `xlsx_execute_task` -- docs/GH76-VERIFY-STORED-CONTENT-OPTION-DESIGN.md section 7, offline plan
 * (cases 1-6 + RED CONTROL). Drives the REAL product WRITE path (createWriteTool -> withSession ->
 * the real executor -> mapExecutionResult -> the new post-close verify) with the fake XLSX transport;
 * the SECOND (verify) driver reads back whatever the fake's own store.durable holds, exactly as a
 * fresh product read would after a real save.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');
const { createSessionRegistry, sessionReuseConfig } = require('../../src/server/session-registry.cjs');

/** A configSource whose getEditorConfig returns `keys[0]` on the FIRST call (the write path's own open,
 * and this unit's preKey capture) and `keys[1]` on every call after -- i.e. "the key rotates exactly
 * once, right after the write closes", the SAME shape the measured LIVE runs show. */
function rotatingConfigSource(keys = ['K1', 'K2']) {
  let calls = 0;
  return { kind: 'BROKER', async getEditorConfig() { const k = keys[Math.min(calls, keys.length - 1)]; calls += 1; return { document: { key: k } }; } };
}
/** A configSource whose key never changes -- rotation never completes. */
function stuckConfigSource(key = 'K1') { return { kind: 'BROKER', async getEditorConfig() { return { document: { key } }; } }; }

const FAST = { timeoutMs: 80, pollMs: 5 }; // keyRotation override -- no test waits real 15s/500ms
const wr = (extra = {}) => ({ intent: 'write_range', id: 's', sheet: 'Budget', range: 'A1:A1', values: [['hello']], ...extra });

test('1. verifyStoredContent absent/false -> byte-identical to today: no storedContent field, storedContentVerified stays false', async () => {
  const h = harness({ store: createStore(book()) });
  const r1 = await h.tool.handler(req([wr()]));
  assert.equal(r1.ok, true, JSON.stringify(r1).slice(0, 300));
  assert.equal(r1.persistence.storedContent, undefined);
  assert.equal(r1.persistence.storedContentVerified, false);

  const h2 = harness({ store: createStore(book()) });
  const r2 = await h2.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't', steps: [wr()], verifyStoredContent: false } });
  assert.equal(r2.ok, true);
  assert.equal(r2.persistence.storedContent, undefined);
  assert.equal(r2.persistence.storedContentVerified, false);
});

test('2. verifyStoredContent:true, key rotates within the bound, the fresh read matches the receipt -> VERIFIED', async () => {
  const h = harness({ store: createStore(book()), configSource: rotatingConfigSource(['K1', 'K2']), keyRotation: FAST });
  const r = await h.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't', steps: [wr()], verifyStoredContent: true } });
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 400));
  assert.equal(r.persistence.storedContent.status, 'VERIFIED', JSON.stringify(r.persistence.storedContent));
  assert.equal(r.persistence.storedContentVerified, true);
  assert.equal(r.persistence.storedContent.scope, 'requested-post-state');
  assert.equal(r.persistence.storedContent.failureReason, null);
  assert.ok(r.persistence.storedContent.overheadMs >= 0);
  // the write itself is UNCHANGED by the opt-in -- same ok/code/operations as a non-verifying call
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
});

test('3. the fresh read disagrees with the receipt -> MISMATCH, ok:false, code STORED_CONTENT_MISMATCH, mutation facts unchanged, no re-write', async () => {
  const store = createStore(book());
  // corrupts store.durable (what a fresh open/read sees) the moment the SECOND driver (the verify's own)
  // opens -- the write's own in-session view already closed correctly; only the post-close read catches this,
  // the same shape as the audited 2026-10-15 defect (in-session correct, saved bytes diverged).
  const h = harness({
    store, configSource: rotatingConfigSource(['K1', 'K2']), keyRotation: FAST,
    wrap: (d, n) => new Proxy(d, { get(t, prop) {
      if (prop === 'open') return async args => { if (n.drivers === 2) store.durable.find(x => x.name === 'Budget').cells.A1 = { v: 'CORRUPTED' }; return t.open(args); };
      const v = t[prop]; return typeof v === 'function' ? v.bind(t) : v;
    } }),
  });
  const r = await h.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't', steps: [wr()], verifyStoredContent: true } });
  assert.equal(r.ok, false, JSON.stringify(r).slice(0, 400));
  assert.equal(r.code, 'STORED_CONTENT_MISMATCH');
  assert.equal(r.persistence.storedContent.status, 'MISMATCH');
  assert.ok(r.persistence.storedContent.notProven.some(x => x.startsWith('stored-poststate-mismatch:A1:product:')), JSON.stringify(r.persistence.storedContent.notProven));
  // mutation facts -- the write itself -- are UNCHANGED: the write really did apply
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
  assert.equal(r.mutationCertainty, 'CERTAIN');
  // no re-write: exactly one write dispatch happened (the harness counts dispatch across every driver it created)
  assert.equal(h.calls().dispatch, 1);
});

test('4. the key never rotates within the bound -> UNCERTAIN / key-not-rotated, ok unchanged, bounded (does not hang)', async () => {
  const h = harness({ store: createStore(book()), configSource: stuckConfigSource('K1'), keyRotation: FAST });
  const t0 = Date.now();
  const r = await h.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't', steps: [wr()], verifyStoredContent: true } });
  const elapsed = Date.now() - t0;
  assert.ok(elapsed < 2000, `bounded wait took ${elapsed}ms -- should be well under the FAST timeout+slack`);
  assert.equal(r.ok, true, 'an UNCERTAIN verify is never a write failure');
  assert.equal(r.persistence.storedContent.status, 'UNCERTAIN');
  assert.equal(r.persistence.storedContent.failureReason, 'key-not-rotated');
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
});

test('5. a verifying call never offers or takes a warm #100 session; the NEXT call for the same key opens fresh, not warm', async () => {
  const registry = createSessionRegistry(sessionReuseConfig({ EURO_MCP_SESSION_REUSE: '1', EURO_MCP_SESSION_IDLE_MS: '120000', EURO_MCP_SESSION_MAX_AGE_MS: '600000', EURO_MCP_SESSION_MAX_CALLS: '50' }));
  const configSource = rotatingConfigSource(['K1', 'K2']);
  const reuse = { registry, principal: 'u1', configSignature: 'c1' };
  const h = harness({ store: createStore(book()), configSource, keyRotation: FAST, reuse });
  const r1 = await h.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't1', steps: [wr()], verifyStoredContent: true } });
  assert.equal(r1.ok, true, JSON.stringify(r1).slice(0, 300));
  assert.equal(registry.status().held, false, 'a verifying call must never hold its session in the #100 registry');
  // a second, NON-verifying call for the same file must also open fresh (nothing was ever offered to hold)
  const r2 = await h.tool.handler({ fileId: req([wr({ range: 'A2:A2' })]).fileId, task: { id: 't2', steps: [wr({ range: 'A2:A2' })] } });
  assert.equal(r2.ok, true, JSON.stringify(r2).slice(0, 300));
  assert.equal(h.n.drivers, 3, 'write(verify-path discards) + verify-read + the next write = 3 fresh drivers, never a warm reuse');
});

test('6. RED CONTROL: a deliberately broken targeted-read mapping (expect pointed at a cell the write never touched) surfaces as a MISMATCH, never a false VERIFIED', async () => {
  // same positive setup as case 2, but the write targets A1 while this control expects content at Z9 --
  // simulated by writing to A1 and asserting Z9 directly against expectedCellsFromSteps' own comparator path
  const { cellDiffs, productCells } = require('../../src/adapters/xlsx/stored-content.cjs');
  const results = [{ ok: true, result: { cells: [{ address: 'A1', v: 'hello', f: '', nf: 'General', bold: false, type: 'text' }] } }];
  const pc = productCells(results);
  const diffs = cellDiffs('Z9', { value: 'hello' }, pc.Z9, 'product');
  assert.deepEqual(diffs, ['Z9:product:missing'], 'a targeted read that never covered the expected cell must be named missing, never silently pass');
});

test('7. requested on a write-free retry (SATISFIED, NOT_REQUIRED): NOT_APPLICABLE with a named reason, never VERIFIED, no verify driver opened', async () => {
  const store = createStore(book());
  const h = harness({ store, configSource: rotatingConfigSource(['K1', 'K2']), keyRotation: FAST });
  const first = await h.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't', steps: [wr()] } });
  assert.equal(first.ok, true);
  const opensBefore = h.opens.length;
  const r = await h.tool.handler({ fileId: req([wr()]).fileId, task: { id: 't', steps: [wr()], verifyStoredContent: true } });
  assert.equal(r.ok, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(r.persistence.storedContent.status, 'NOT_APPLICABLE');
  assert.equal(r.persistence.storedContent.failureReason, 'no-confirmed-save:NOT_REQUIRED');
  assert.equal(r.persistence.storedContentVerified, false);
  assert.equal(h.opens.length - opensBefore, 1, 'only the retry\'s own write session opened -- no verify driver');
});
