'use strict';

const test = require('../helpers/trim-skip.cjs').testFor('add_footnote'); // release 3 trim: skipped while add_footnote is not in the allowlist
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #267 B: add_footnote / add_endnote served by docx_execute_task (LIVE g267v1 PASS R1-R3, docs/GH267-DOCX-NOTES-WRITE-DECISION.md). Scope: ONE target paragraph named by
 * { text } and ONE paragraph of note text. NOT released: { ref } / physical selectors, a line break in the note, multi-paragraph notes, REPLACING a note's text (no routed
 * intent; a different text on the same target appends a second note, g267v1 ROUND4). Offline only: the product handler runs the REAL docx adapter/executor/lifecycle on
 * test/fakes/fake-docx-driver.cjs; the READ side is the product docx_read read_notes on the SAME store.
 */

const FILE = '1303629';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const TARGET = 'GH267 Body.';
const doc = () => [{ text: 'GH267 Heading', level: 1, props: PLAIN }, { text: TARGET, props: PLAIN }];
const req = steps => ({ fileId: FILE, task: { id: 't', steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const step = (intent, extra = {}) => ({ intent, selector: { text: TARGET }, text: 'GH267 note text', ...extra });

test('SCHEMA: a { selector: { text }, text } step is accepted for both intents', () => {
  for (const intent of ['add_footnote', 'add_endnote']) assert.equal(D.validateDocxWriteRequest(req([step(intent)])).ok, true, intent);
});

test('RED (POS/NEG per rule): only the { text } selector, a non-blank target, ONE paragraph of non-blank note text -- each refused pre-dispatch, 0 dispatch', () => {
  const bad = [
    ['no selector', { intent: 'add_footnote', text: 'n' }],
    ['no text', { intent: 'add_footnote', selector: { text: TARGET } }],
    ['empty note text', step('add_footnote', { text: '' })],
    ['blank note text', step('add_footnote', { text: '   ' })],
    ['line break in the note', step('add_footnote', { text: 'a\nb' })],
    ['blank target', step('add_endnote', { selector: { text: '  ' } })],
    ['a ref selector', step('add_footnote', { selector: { ref: 's0' } })],
    ['a fingerprint selector', step('add_footnote', { selector: { text: TARGET, fingerprint: 'x' } })],
    ['an unknown field', step('add_endnote', { marker: 'x' })],
  ];
  for (const [name, s] of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.ok, false, name);
    assert.equal(v.dispatched, false, name);
  }
});

test('RELEASE STATE: both intents are in the accepted allowlist, schema and constraint tables, in the same order', () => {
  const W = require('../../src/server/write-tools.cjs');
  const T = require('../../src/server/tools.cjs');
  for (const i of ['add_footnote', 'add_endnote']) {
    assert.ok(W.DOCX_WRITE_ALLOWLIST.includes(i));
    assert.ok(T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes(i));
    assert.ok(D.DOCX_STEP_SCHEMAS[i] && D.DOCX_CONSTRAINTS[i]);
  }
  assert.deepEqual(Object.keys(D.DOCX_STEP_SCHEMAS), [...W.DOCX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(D.DOCX_CONSTRAINTS), [...W.DOCX_WRITE_ALLOWLIST]);
});

function product() {
  const store = createStore(doc());
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const notes = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_notes' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result; };
  return { store, tool, notes };
}

test('EXECUTION: add_footnote -> APPLIED, the note is read back once; retry -> SATISFIED, no new write; add_endnote lands on the endnote side only', async () => {
  const p = product();
  const first = await p.tool.handler(req([step('add_footnote')]));
  assert.equal(first.ok, true, JSON.stringify(first));
  assert.deepEqual((first.operations || []).map(o => o.status), ['APPLIED']);
  const writes = p.store.saves;
  assert.ok(writes > 0);
  let n = await p.notes();
  assert.deepEqual(n.footnotes.map(x => x.text), ['GH267 note text']);
  assert.deepEqual(n.endnotes.map(x => x.text), []);

  const retry = await p.tool.handler(req([step('add_footnote')]));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.store.saves, writes, 'retry must not write');

  const en = await p.tool.handler(req([step('add_endnote', { text: 'GH267 end text' })]));
  assert.deepEqual((en.operations || []).map(o => o.status), ['APPLIED']);
  n = await p.notes();
  assert.deepEqual(n.footnotes.map(x => x.text), ['GH267 note text']);
  assert.deepEqual(n.endnotes.map(x => x.text), ['GH267 end text']);
});

test('EXECUTION: a DIFFERENT text on the same target APPENDS a second note (no replace intent is released) -- the measured g267v1 ROUND4 behaviour', async () => {
  const p = product();
  await p.tool.handler(req([step('add_footnote')]));
  const second = await p.tool.handler(req([step('add_footnote', { text: 'GH267 other note' })]));
  assert.deepEqual((second.operations || []).map(o => o.status), ['APPLIED']);
  assert.deepEqual((await p.notes()).footnotes.map(x => x.text), ['GH267 note text', 'GH267 other note']);
});

test('EXECUTION: a target that does not exist fails closed -- nothing is written', async () => {
  const p = product();
  const r = await p.tool.handler(req([step('add_footnote', { selector: { text: 'no such paragraph' } })]));
  assert.notEqual((r.operations || []).map(o => o.status)[0], 'APPLIED', JSON.stringify(r));
  assert.deepEqual((await p.notes()).footnotes.map(x => x.text), []);
});
