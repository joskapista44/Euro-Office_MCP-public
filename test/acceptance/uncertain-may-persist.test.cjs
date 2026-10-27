'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const V = require('../../src/core/persistence/verified-writes.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * An uncertain-only task ends UNCERTAIN_MAY_PERSIST, not "not persisted" (guncprobe, zero, 2026-10-19, DS 9.3.5: the unsaved edit became durable on its own within ~9 s,
 * and another editor's forcesave saved it too). Offline: the product handler on the fake transport; the lifecycle / receipt wording is the code under test.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const OLD = 'MAY PERSIST OLD';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OLD, props: PLAIN }, { text: 'Closing.', props: PLAIN }];
const req = (text = 'MAY PERSIST NEW') => ({ fileId: FILE, task: { steps: [{ intent: 'rewrite_paragraph', selector: { text: OLD }, text }] } });

function product(script = {}) {
  const store = createStore(doc());
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  return { store, drivers, tool: D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver }) };
}

test('the state set: UNCERTAIN_MAY_PERSIST replaced the old name -- nothing can still compare against a state that is no longer produced', () => {
  assert.equal(V.PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST, 'UNCERTAIN_MAY_PERSIST');
  assert.equal('UNCERTAIN_NOT_PERSISTED' in V.PERSISTENCE_STATE, false);
  assert.deepEqual(Object.keys(V.PERSISTENCE_STATE).sort(), ['EDITOR_CLEAN', 'FAILED', 'FORCESAVE_CONFIRMED', 'NOT_REQUIRED', 'UNCERTAIN_MAY_PERSIST']);
  const u = V.decidePersistence({ appliedOperations: [{ status: 'DISPATCHED_OUTCOME_UNKNOWN' }] });
  assert.deepEqual([u.barrierRequired, u.onSuccess], [false, 'UNCERTAIN_MAY_PERSIST']);
});

test('the receipt of an uncertain-only task: persistence.state UNCERTAIN_MAY_PERSIST, a reason and a meaning that say the edit may become durable and the document must be re-read; no save, no key', async () => {
  const p = product({ fault: 'ignore' });                       // the editor ignores the write: DISPATCHED_OUTCOME_UNKNOWN, L1 cannot prove it
  const r = await p.tool.handler(req());
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(r.mutationCertainty, 'UNCERTAIN');
  assert.equal(r.persistence.state, 'UNCERTAIN_MAY_PERSIST');
  assert.equal(r.persistence.outcome, 'UNCERTAIN_MAY_PERSIST');
  assert.equal(r.persistence.reason, 'only-uncertain-writes-may-persist');
  assert.match(r.persistence.meaning, /may become durable/);
  assert.match(r.persistence.meaning, /re-read the document/);
  assert.equal(r.persistence.storedContentVerified, false);
  assert.equal(p.store.saves, 0, 'the framework still never forceSaves solely for an unproven mutation');
  assert.doesNotMatch(JSON.stringify(r), /NOT_PERSISTED/, 'no field of the receipt claims "not persisted"');
});

test('NEG: a verified write and a certain refusal carry no meaning and are not UNCERTAIN_MAY_PERSIST', async () => {
  const ok = product();
  const r = await ok.tool.handler(req());
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.notEqual(r.persistence.state, 'UNCERTAIN_MAY_PERSIST');
  assert.equal('meaning' in r.persistence, false);
  const none = product();
  const miss = await none.tool.handler({ fileId: FILE, task: { steps: [{ intent: 'rewrite_paragraph', selector: { text: 'NOT THERE' }, text: 'x' }] } });
  assert.equal(miss.ok, false);
  assert.notEqual(miss.persistence && miss.persistence.state, 'UNCERTAIN_MAY_PERSIST');
  assert.equal(miss.mutationCertainty, 'CERTAIN');
});
