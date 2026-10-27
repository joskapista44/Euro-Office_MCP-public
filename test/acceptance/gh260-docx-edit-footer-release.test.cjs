'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #260 (DOCX P1 WRITE): edit_footer served by docx_execute_task -- the SAME shape and the SAME
 * scope as GH #259's edit_header (common/default variant only; title/even are #261's own scope),
 * mirrored onto footer. Design: docs/GH259-DOCX-EDIT-HEADER-COMMON-DECISION.md applies verbatim
 * (header-footer.cjs's own flagOn() keys only on `op.variant`, never on `op.part`). Offline only:
 * the product handler runs the REAL docx adapter/executor/lifecycle on
 * test/fakes/fake-docx-driver.cjs; the READ side is the product docx_read read_headers_footers on
 * the SAME store. The LIVE acceptance is test/live/docx-edit-footer-common-write-product.live.cjs
 * (not run by the author).
 */

const FILE = '1303629';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const doc = () => [{ text: 'GH260 Heading', level: 1, props: PLAIN }, { text: 'GH260 Body.', props: PLAIN }];
const req = (steps) => ({ fileId: FILE, task: { id: 't', steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const SECTION = { index: 0, of: 1 };
const TEXT = 'GH260 common footer text';

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
  const v = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: SECTION, text: TEXT }]));
  assert.equal(v.ok, true, JSON.stringify(v));
});

// GH #261 1/2 widened the schema to expose `variant: title|even` (additive, see
// gh261-docx-header-footer-variants-release.test.cjs); this card's own narrowing now shows up as
// an enum mismatch, not an unknown key.
test('RED: a `variant` key with a value outside the released enum ("default" -- the omitted-key equivalent -- or anything bogus) is REFUSED_SHAPE, not APPLIED', () => {
  for (const variant of ['default', 'bogus']) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: SECTION, text: TEXT, variant }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(variant));
    assert.equal(v.dispatched, false);
  }
});

test('RED: a `selector` key is UNKNOWN_FIELD -- the target is section + part, never a text match', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: SECTION, text: TEXT, selector: { text: 'GH260 Heading' } }]));
  assert.equal(v.code, 'UNKNOWN_FIELD');
  assert.equal(v.dispatched, false);
});

test('RED: missing section, missing text, or an empty text is REFUSED_SHAPE pre-dispatch, 0 dispatch', () => {
  const noSection = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', text: TEXT }]));
  assert.equal(noSection.code, 'REFUSED_SHAPE');
  const noText = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: SECTION }]));
  assert.equal(noText.code, 'REFUSED_SHAPE');
  const emptyText = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: SECTION, text: '' }]));
  assert.equal(emptyText.code, 'REFUSED_SHAPE');
  for (const r of [noSection, noText, emptyText]) assert.equal(r.dispatched, false);
});

test('RED: section out of bounds (index >= of) is refused at the adapter planner layer', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: { index: 1, of: 1 }, text: TEXT }]));
  assert.equal(v.ok, false, JSON.stringify(v));
});

// ---- execution (the fake transport) --------------------------------------------------------------------------------------------

test('EXECUTION: writing the common footer text on a footer-less section -> APPLIED/FORCESAVE_CONFIRMED (no default-value trap here -- no footer part exists at all, so any text is a real create+write); a fresh read_headers_footers confirms byte-for-byte; retry is SATISFIED with zero new writes', async () => {
  const p = product();
  const receipt = await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);
  assert.equal(receipt.persistence && receipt.persistence.state, 'FORCESAVE_CONFIRMED');
  const writesAfterFirst = p.writes();
  assert.ok(writesAfterFirst > 0, 'expected at least one save after an APPLIED step');

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.footer.default.paragraphs, 1, JSON.stringify(sec.footer));
  assert.equal(sec.footer.default.paras[0].text, TEXT, JSON.stringify(sec.footer));
  // the OTHER parts stay untouched (L3): the header and the title/even footer variants, still unset
  assert.equal(sec.header.default, null, JSON.stringify(sec.header));
  assert.equal(sec.footer.title, null, JSON.stringify(sec.footer));
  assert.equal(sec.footer.even, null, JSON.stringify(sec.footer));

  const retry = await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT }]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), writesAfterFirst, 'retry must not produce a new write');
});

test('EXECUTION: a second write with DIFFERENT text replaces the whole part content -- one paragraph, the new text, nothing kept from the old', async () => {
  const p = product();
  await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT }]));
  const receipt2 = await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: `${TEXT} v2` }]));
  assert.equal(receipt2.ok, true, JSON.stringify(receipt2));
  assert.deepEqual((receipt2.operations || []).map(o => o.status), ['APPLIED']);
  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.footer.default.paragraphs, 1, JSON.stringify(sec.footer));
  assert.equal(sec.footer.default.paras[0].text, `${TEXT} v2`, JSON.stringify(sec.footer));
});

test('EXECUTION: a header write and a footer write on the SAME section are independent -- each ends up on its own part, neither overwrites the other', async () => {
  const p = product();
  await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: 'GH260 header text' }]));
  await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT }]));
  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.default.paras[0].text, 'GH260 header text', JSON.stringify(sec.header));
  assert.equal(sec.footer.default.paras[0].text, TEXT, JSON.stringify(sec.footer));
});

module.exports = { FILE, SECTION, TEXT };
