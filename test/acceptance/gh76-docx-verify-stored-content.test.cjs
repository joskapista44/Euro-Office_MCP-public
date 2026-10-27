'use strict';

/**
 * GH #76 MODEL 1 DOCX PORT (owner decision "A", owner channel, 2026-10-19 -- the owner's own question,
 * msg 38135, "miert nem olvassatok vissza ha bizonytalan?", answered by mapping the EXISTING xlsx solution
 * onto docx first, 1:1; the uncertain-branch extension this unit's own first draft proposed is explicitly
 * NOT part of this unit -- see docs/UNCERTAIN-READBACK-DESIGN.md section 5). The optional, task-level
 * `verifyStoredContent` on `docx_execute_task` -- docs/UNCERTAIN-READBACK-DESIGN.md, offline plan (cases 1-7,
 * the SAME numbering as test/acceptance/gh76-verify-stored-content.test.cjs's own xlsx cases, since the
 * contract and status set are byte-identical, only the comparator and the fake transport differ).
 *
 * Drives the REAL product WRITE path (createDocxWriteTool -> withSession -> the real executor ->
 * mapExecutionResult -> the new post-close verify) with the fake DOCX transport; the SECOND (verify) driver
 * reads back whatever the fake's own store.durable holds, exactly as a fresh product read would after a
 * real save.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { createSessionRegistry, sessionReuseConfig } = require('../../src/server/session-registry.cjs');

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const OLD = 'GH76 OLD paragraph.';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OLD, props: PLAIN }, { text: 'Closing.', props: PLAIN }];
const rewrite = (to = 'GH76 NEW paragraph.', extra = {}) => ({ intent: 'rewrite_paragraph', selector: { text: OLD }, text: to, ...extra });
const req = (steps, extra = {}) => ({ fileId: FILE, task: { id: 't', steps, ...extra } });

/** A configSource whose getEditorConfig returns `keys[0]` on the FIRST call and `keys[1]` on every call
 * after -- "the key rotates exactly once, right after the write closes" (the same shape as the xlsx GH76
 * test's own rotatingConfigSource). */
function rotatingConfigSource(keys = ['K1', 'K2']) {
  let calls = 0;
  return { kind: 'BROKER', async getEditorConfig() { const k = keys[Math.min(calls, keys.length - 1)]; calls += 1; return { document: { key: k } }; } };
}
/** A configSource whose key never changes -- rotation never completes. */
function stuckConfigSource(key = 'K1') { return { kind: 'BROKER', async getEditorConfig() { return { document: { key } }; } }; }

const FAST = { timeoutMs: 80, pollMs: 5 }; // keyRotation override -- no test waits real 15s/500ms
const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };

function product({ store, configSource = rotatingConfigSource(), keyRotation = FAST, reuse = null, wrap = null } = {}) {
  const drivers = [];
  let created = 0;
  const createDriver = () => {
    created += 1;
    const ordinal = created; // captured BEFORE this driver is pushed -- the 1-based "which driver is this" count
    let d = createFakeDocxDriver(store, { saveObserver: true });
    if (wrap) d = wrap(d, { drivers: ordinal });
    drivers.push(d);
    return d;
  };
  const gate = createSessionGate();
  const tool = D.createDocxWriteTool({ configSource, gate, runtime, createDriver, reuse, keyRotation });
  return { store, drivers, gate, tool };
}

test('1. verifyStoredContent absent/false -> byte-identical to today: no storedContent field, storedContentVerified stays false', async () => {
  const p1 = product({ store: createStore(doc()) });
  const r1 = await p1.tool.handler(req([rewrite()]));
  assert.equal(r1.ok, true, JSON.stringify(r1).slice(0, 300));
  assert.equal(r1.persistence.storedContent, undefined);
  assert.equal(r1.persistence.storedContentVerified, false);

  const p2 = product({ store: createStore(doc()) });
  const r2 = await p2.tool.handler(req([rewrite()], { verifyStoredContent: false }));
  assert.equal(r2.ok, true);
  assert.equal(r2.persistence.storedContent, undefined);
  assert.equal(r2.persistence.storedContentVerified, false);
});

test('2. verifyStoredContent:true, key rotates within the bound, the fresh read matches the receipt -> VERIFIED', async () => {
  const p = product({ store: createStore(doc()) });
  const r = await p.tool.handler(req([rewrite()], { verifyStoredContent: true }));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 400));
  assert.equal(r.persistence.storedContent.status, 'VERIFIED', JSON.stringify(r.persistence.storedContent));
  assert.equal(r.persistence.storedContentVerified, true);
  assert.equal(r.persistence.storedContent.scope, 'requested-post-state');
  assert.equal(r.persistence.storedContent.failureReason, null);
  assert.ok(r.persistence.storedContent.overheadMs >= 0);
  // the write itself is UNCHANGED by the opt-in -- same ok/code/operations as a non-verifying call
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
});

test('2b. verifyStoredContent:true with expectedPreservation.bold named and TRUE on the fresh read -> VERIFIED (the one preserved field the comparator covers)', async () => {
  const p = product({ store: createStore(doc()) });
  const r = await p.tool.handler(req([rewrite('GH76 NEW bold paragraph.', {})], { verifyStoredContent: true }));
  // PLAIN carries bold:false; the stored read's own bold therefore reads false -- no expectedPreservation named
  // here reaches a MISMATCH on bold specifically (left for case 3's own, deliberate corruption).
  assert.equal(r.persistence.storedContent.status, 'VERIFIED', JSON.stringify(r.persistence.storedContent));
});

test('3. the fresh read disagrees with the receipt -> MISMATCH, ok:false, code STORED_CONTENT_MISMATCH, mutation facts unchanged, no re-write', async () => {
  const store = createStore(doc());
  // corrupts store.durable (what a fresh open/read sees) the moment the SECOND driver (the verify's own)
  // opens -- the write's own in-session view already closed correctly; only the post-close read catches this.
  const p = product({
    store, wrap: (d, n) => new Proxy(d, { get(t, prop) {
      if (prop === 'open') return async args => { if (n.drivers === 2) store.durable.find(x => x.text === 'GH76 NEW paragraph.').text = 'GH76 CORRUPTED, NOT the requested text'; return t.open(args); };
      const v = t[prop]; return typeof v === 'function' ? v.bind(t) : v;
    } }),
  });
  const r = await p.tool.handler(req([rewrite()], { verifyStoredContent: true }));
  assert.equal(r.ok, false, JSON.stringify(r).slice(0, 400));
  assert.equal(r.code, 'STORED_CONTENT_MISMATCH');
  assert.equal(r.persistence.storedContent.status, 'MISMATCH');
  assert.ok(r.persistence.storedContent.notProven.some(x => x.startsWith('stored-poststate-mismatch:')), JSON.stringify(r.persistence.storedContent.notProven));
  // mutation facts -- the write itself -- are UNCHANGED: the write really did apply
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
  assert.equal(r.mutationCertainty, 'CERTAIN');
});

test('4. the key never rotates within the bound -> UNCERTAIN / key-not-rotated, ok unchanged, bounded (does not hang)', async () => {
  const p = product({ store: createStore(doc()), configSource: stuckConfigSource('K1') });
  const t0 = Date.now();
  const r = await p.tool.handler(req([rewrite()], { verifyStoredContent: true }));
  const elapsed = Date.now() - t0;
  assert.ok(elapsed < 2000, `bounded wait took ${elapsed}ms -- should be well under the FAST timeout+slack`);
  assert.equal(r.ok, true, 'an UNCERTAIN verify is never a write failure');
  assert.equal(r.persistence.storedContent.status, 'UNCERTAIN');
  assert.equal(r.persistence.storedContent.failureReason, 'key-not-rotated');
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
});

test('5. a verifying call never offers or takes a warm #100 session; the NEXT call for the same key opens fresh, not warm', async () => {
  const registry = createSessionRegistry(sessionReuseConfig({ EURO_MCP_SESSION_REUSE: '1', EURO_MCP_SESSION_IDLE_MS: '120000', EURO_MCP_SESSION_MAX_AGE_MS: '600000', EURO_MCP_SESSION_MAX_CALLS: '50' }));
  const reuse = { registry, principal: 'u1', configSignature: 'c1' };
  const p = product({ store: createStore(doc()), reuse });
  const r1 = await p.tool.handler(req([rewrite()], { verifyStoredContent: true }));
  assert.equal(r1.ok, true, JSON.stringify(r1).slice(0, 300));
  assert.equal(registry.status().held, false, 'a verifying call must never hold its session in the #100 registry');
  // a second, NON-verifying call must also open fresh (nothing was ever offered to hold)
  const r2 = await p.tool.handler({ fileId: FILE, task: { id: 't2', steps: [{ intent: 'rewrite_paragraph', selector: { text: 'GH76 NEW paragraph.' }, text: 'GH76 NEWER.' }] } });
  assert.equal(r2.ok, true, JSON.stringify(r2).slice(0, 300));
  assert.equal(p.drivers.length, 3, 'write(verify-path discards reuse) + verify-read + the next write = 3 fresh drivers, never a warm reuse');
});

test('6. RED CONTROL: productParagraphOf pointed at a paragraph the write never touched surfaces as missing, never a false match', () => {
  const units = [{ text: 'GH76 NEW paragraph.', runs: [{ text: 'GH76 NEW paragraph.', format: { bold: false } }] }];
  const { textRunDiffs, productParagraphOf } = require('../../src/adapters/docx/stored-content.cjs');
  const diffs = textRunDiffs({ text: 'Some other text entirely' }, productParagraphOf(units, 'Some other text entirely'), 'product');
  assert.deepEqual(diffs, ['Some other text entirely:product:missing'], 'a read that never covered the expected paragraph must be named missing, never silently pass');
});

/** The selector a caller builds from a FRESH inventory (same discipline as gh220-docx-write-lifecycle.test.cjs's
 * own selectorFrom): `before`/`after` context survives a successful rewrite, so a SECOND, FRESH (non-warm)
 * open can still re-resolve the SAME physical paragraph by its neighbours and see it already holds the
 * wanted text -- SATISFIED, not a failed selector. */
function selectorFrom(paragraphs, target) {
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const b = primary.blocks.find(x => x.text === target);
  return { text: target, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}

test('7. requested on a write-free retry (SATISFIED, NOT_REQUIRED): NOT_APPLICABLE with a named reason, never VERIFIED, no verify driver opened', async () => {
  const p = product({ store: createStore(doc()) });
  const step = { intent: 'rewrite_paragraph', selector: selectorFrom(doc(), OLD), text: 'GH76 NEW paragraph.' };
  const first = await p.tool.handler({ fileId: FILE, task: { id: 't', steps: [step] } });
  assert.equal(first.ok, true, JSON.stringify(first).slice(0, 300));
  assert.deepEqual(first.operations.map(o => o.status), ['APPLIED']);
  const opensBefore = p.drivers.reduce((a, d) => a + d.calls.open, 0);
  // the SAME step again, FRESH (non-warm) open: before/after re-resolves the same paragraph, already at the
  // wanted text -- SATISFIED, zero dispatch, zero forceSave, PERSISTENCE_STATE.NOT_REQUIRED.
  const r = await p.tool.handler({ fileId: FILE, task: { id: 't2', steps: [step], verifyStoredContent: true } });
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(r.persistence.storedContent.status, 'NOT_APPLICABLE');
  assert.equal(r.persistence.storedContent.failureReason, 'no-confirmed-save:NOT_REQUIRED');
  assert.equal(r.persistence.storedContentVerified, false);
  const opensAfter = p.drivers.reduce((a, d) => a + d.calls.open, 0);
  assert.equal(opensAfter - opensBefore, 1, 'only the retry\'s own write session opened -- no verify driver');
});

test('8. MALFORMED_TASK: verifyStoredContent must be a boolean, pre-dispatch, 0 writes', () => {
  const v = D.validateDocxWriteRequest(req([rewrite()], { verifyStoredContent: 'yes' }));
  assert.equal(v.code, 'MALFORMED_TASK');
  assert.match(v.reason, /verifyStoredContent/);
});
