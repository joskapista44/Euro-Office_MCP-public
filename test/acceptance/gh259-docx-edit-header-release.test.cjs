'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #259 (DOCX P1 WRITE): edit_header served by docx_execute_task, narrowed to the common
 * (default) variant only -- title/even are #261's own scope, and are refused `UNKNOWN_FIELD`
 * because this release's step schema never offers a `variant` key at all. Design and the measured
 * reason for every rule: docs/GH259-DOCX-EDIT-HEADER-COMMON-DECISION.md. Offline only: the product
 * handler runs the REAL docx adapter/executor/lifecycle on test/fakes/fake-docx-driver.cjs; the
 * READ side is the product docx_read read_headers_footers (M4-W5) on the SAME store. The LIVE
 * acceptance is test/live/docx-edit-header-common-write-product.live.cjs (not run by the author).
 */

const FILE = '1303628';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const doc = () => [{ text: 'GH259 Heading', level: 1, props: PLAIN }, { text: 'GH259 Body.', props: PLAIN }];
const req = (steps) => ({ fileId: FILE, task: { id: 't', steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const SECTION = { index: 0, of: 1 };
const TEXT = 'GH259 common header text';

function product(script = {}) {
  const store = createStore(doc());
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readHeadersFooters = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_headers_footers' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.sections; };
  return { store, drivers, tool, readHeadersFooters, writes: () => store.saves };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: a plain { section, text } step is accepted', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION, text: TEXT }]));
  assert.equal(v.ok, true, JSON.stringify(v));
});

// GH #261 1/2 widened the schema to expose `variant: title|even` (additive, see
// gh261-docx-header-footer-variants-release.test.cjs); this card's own narrowing now shows up as
// an enum mismatch, not an unknown key.
test('RED: a `variant` key with a value outside the released enum ("default" -- the omitted-key equivalent -- or anything bogus) is REFUSED_SHAPE, not APPLIED', () => {
  for (const variant of ['default', 'bogus']) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(variant));
    assert.equal(v.dispatched, false);
  }
});

test('RED: a `selector` key is UNKNOWN_FIELD -- the target is section + part, never a text match', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION, text: TEXT, selector: { text: 'GH259 Heading' } }]));
  assert.equal(v.code, 'UNKNOWN_FIELD');
  assert.equal(v.dispatched, false);
});

test('RED: missing section, missing text, or an empty text is REFUSED_SHAPE pre-dispatch, 0 dispatch', () => {
  const noSection = D.validateDocxWriteRequest(req([{ intent: 'edit_header', text: TEXT }]));
  assert.equal(noSection.code, 'REFUSED_SHAPE');
  const noText = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION }]));
  assert.equal(noText.code, 'REFUSED_SHAPE');
  const emptyText = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION, text: '' }]));
  assert.equal(emptyText.code, 'REFUSED_SHAPE');
  for (const r of [noSection, noText, emptyText]) assert.equal(r.dispatched, false);
});

test('RED: section out of bounds (index >= of) is refused at the adapter planner layer', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: { index: 1, of: 1 }, text: TEXT }]));
  assert.equal(v.ok, false, JSON.stringify(v));
});

// ---- execution (the fake transport) --------------------------------------------------------------------------------------------

test('EXECUTION: writing the common header text on a header-less section -> APPLIED/FORCESAVE_CONFIRMED (no default-value trap here -- the fixture has no header part at all, so any text is a real create+write); a fresh read_headers_footers confirms byte-for-byte; retry is SATISFIED with zero new writes', async () => {
  const p = product();
  const receipt = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);
  assert.equal(receipt.persistence && receipt.persistence.state, 'FORCESAVE_CONFIRMED');
  const writesAfterFirst = p.writes();
  assert.ok(writesAfterFirst > 0, 'expected at least one save after an APPLIED step');

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.default.paragraphs, 1, JSON.stringify(sec.header));
  assert.equal(sec.header.default.paras[0].text, TEXT, JSON.stringify(sec.header));
  // the OTHER parts stay untouched (L3): the footer and the title/even header variants, still unset
  assert.equal(sec.footer.default, null, JSON.stringify(sec.footer));
  assert.equal(sec.header.title, null, JSON.stringify(sec.header));
  assert.equal(sec.header.even, null, JSON.stringify(sec.header));

  const retry = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT }]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), writesAfterFirst, 'retry must not produce a new write');
});

test('EXECUTION: a second write with DIFFERENT text replaces the whole part content -- one paragraph, the new text, nothing kept from the old', async () => {
  const p = product();
  await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT }]));
  const receipt2 = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: `${TEXT} v2` }]));
  assert.equal(receipt2.ok, true, JSON.stringify(receipt2));
  assert.deepEqual((receipt2.operations || []).map(o => o.status), ['APPLIED']);
  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.default.paragraphs, 1, JSON.stringify(sec.header));
  assert.equal(sec.header.default.paras[0].text, `${TEXT} v2`, JSON.stringify(sec.header));
});

module.exports = { FILE, SECTION, TEXT };
