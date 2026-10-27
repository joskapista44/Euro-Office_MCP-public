'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #261 (DOCX P1 WRITE) + GH #261c: edit_header/edit_footer's `variant` field widened to
 * title|even (additive to GH #259/#260's default-only release), WITH activation (flag off -> on).
 * The adapter's own guard (shapeGuard -> activationGuard, header-footer.cjs) already proves an
 * OFF -> ON flag flip is content-safe (every unrequested slot it would also switch is provably
 * unchanged first); the activation write itself (SetTitlePage/SetEvenAndOddHdrFtr) goes through
 * the SAME __eoB21Variant bridge the already-active case already uses. #261 refused the
 * activation transition pending a stored LIVE PASS for that bridge path (withActivationRefused,
 * since removed); #261b supplied it
 * (test/live/docx-hf-activation-bridge-readback-probe.live.cjs, g261bv2 PASS: the bridge reports
 * the new state immediately, persists across save+reopen, and matches the raw internal getter
 * exactly) -- #261c removes the refusal. Offline only: the product handler runs the REAL docx
 * adapter/executor/lifecycle on test/fakes/fake-docx-driver.cjs (which already models title/even
 * + activation, built ahead of this card); the READ side is the product docx_read
 * read_headers_footers. The LIVE acceptance is
 * test/live/docx-header-footer-title-even-variants-write-product.live.cjs (not run by the
 * author -- ROUND4 now exercises activation itself, not a permanent refusal).
 */

const FILE = '1303630';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const doc = () => [{ text: 'GH261 Heading', level: 1, props: PLAIN }, { text: 'GH261 Body.', props: PLAIN }];
const req = (steps) => ({ fileId: FILE, task: { id: 't', steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const SECTION = { index: 0, of: 1 };
const TEXT = 'GH261 variant text';

function product(script = {}, docOpts = {}) {
  const store = createStore(doc(), { doc: docOpts });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readHeadersFooters = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_headers_footers' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.sections; };
  return { store, drivers, tool, readHeadersFooters, writes: () => store.saves };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: variant title|even is now accepted shape (no longer UNKNOWN_FIELD, GH #261 1/2)', () => {
  for (const variant of ['title', 'even']) {
    const h = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant }]));
    assert.equal(h.ok, true, JSON.stringify({ variant, h }));
    const f = D.validateDocxWriteRequest(req([{ intent: 'edit_footer', section: SECTION, text: TEXT, variant }]));
    assert.equal(f.ok, true, JSON.stringify({ variant, f }));
  }
});

test('RED: variant outside the released enum ("default" or bogus) is REFUSED_SHAPE', () => {
  for (const variant of ['default', 'bogus']) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(variant));
    assert.equal(v.dispatched, false);
  }
});

// ---- execution: activation (flag OFF -> ON) is now RELEASED (GH #261c) -------------------------------------------------------

test('GREEN (GH #261c): title variant on a section whose titlePage flag is OFF -> activates it: APPLIED, flag flips true, content written; retry SATISFIED', async () => {
  const p = product();
  const receipt = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant: 'title' }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);
  const writesAfterFirst = p.writes();
  assert.ok(writesAfterFirst > 0);

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.title.paragraphs, 1, JSON.stringify(sec.header));
  assert.equal(sec.header.title.paras[0].text, TEXT, JSON.stringify(sec.header));
  assert.equal((p.store.durableDoc.sections[SECTION.index] || {}).titlePage, true, 'the flag must have flipped on');

  const retry = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant: 'title' }]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), writesAfterFirst, 'retry must not produce a new write');
});

test('GREEN (GH #261c): even variant on a document whose evenAndOdd flag is OFF -> activates it (document-wide): APPLIED, flag flips true, content written', async () => {
  const p = product();
  const receipt = await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT, variant: 'even' }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.footer.even.paragraphs, 1, JSON.stringify(sec.footer));
  assert.equal(sec.footer.even.paras[0].text, TEXT, JSON.stringify(sec.footer));
  assert.equal(p.store.durableDoc.evenAndOdd, true, 'the document-wide flag must have flipped on');
});

// ---- execution: activationGuard refuses an activation that would silently affect existing sibling content (GH #261c) ----------
// Turning title/even ON makes the variant slot shown INSTEAD of the default one, for BOTH kinds
// (title: the same section's first page; even: every section's even pages) -- so an existing
// `default` part on the OTHER kind of the same section would silently go blank on the pages the
// newly-enabled variant now governs, unless the default and the (absent) variant slot are
// provably the same. activationGuard() refuses exactly this, named. (Measured while drafting
// this file: a naive "GREEN, existing sibling survives" test is WRONG for both title and even --
// the adapter's own safety is MORE conservative than that assumption, by design.)

test('RED (GH #261c safety net): title activation is refused when an EXISTING default footer on the SAME section would be silently affected', async () => {
  const p = product();
  await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: 'pre-existing footer' }]));
  const writesAfterSetup = p.writes();

  const receipt = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant: 'title' }]));
  assert.equal(receipt.ok, false, JSON.stringify(receipt));
  assert.equal(receipt.reason, 'operation s0: FAIL_CLOSED_UNSUPPORTED_SHAPE:title-activation-changes-unrequested-section-0-footer', JSON.stringify(receipt));
  assert.equal(p.writes(), writesAfterSetup, 'the refused activation must not reach dispatch');

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.footer.default.paras[0].text, 'pre-existing footer', 'untouched, since the write never dispatched');
  assert.equal(sec.header.title, null, 'the title header part must not have been created');
});

test('RED (GH #261c safety net): even activation is refused when an EXISTING default header on the SAME section would be silently affected', async () => {
  const p = product();
  await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: 'pre-existing header' }]));
  const writesAfterSetup = p.writes();

  const receipt = await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT, variant: 'even' }]));
  assert.equal(receipt.ok, false, JSON.stringify(receipt));
  assert.equal(receipt.reason, 'operation s0: FAIL_CLOSED_UNSUPPORTED_SHAPE:even-activation-changes-unrequested-section-0-header', JSON.stringify(receipt));
  assert.equal(p.writes(), writesAfterSetup, 'the refused activation must not reach dispatch');

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.default.paras[0].text, 'pre-existing header', 'untouched, since the write never dispatched');
  assert.equal(sec.footer.even, null, 'the even footer part must not have been created');
});

// ---- execution: the already-active case is served -----------------------------------------------------------------------------

test('GREEN: title variant on a section whose titlePage flag is ALREADY ON -> APPLIED, content written, flag unchanged; retry SATISFIED', async () => {
  const p = product({}, { sections: [{ titlePage: true }] });
  const receipt = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant: 'title' }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);
  const writesAfterFirst = p.writes();
  assert.ok(writesAfterFirst > 0);

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.title.paragraphs, 1, JSON.stringify(sec.header));
  assert.equal(sec.header.title.paras[0].text, TEXT, JSON.stringify(sec.header));
  // the default variant and the footer stay untouched (L3)
  assert.equal(sec.header.default, null, JSON.stringify(sec.header));
  assert.equal(sec.footer.default, null, JSON.stringify(sec.footer));
  assert.equal(sec.footer.title, null, JSON.stringify(sec.footer));

  const retry = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT, variant: 'title' }]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), writesAfterFirst, 'retry must not produce a new write');
});

test('GREEN: even variant on a document whose evenAndOdd flag is ALREADY ON -> APPLIED, content written', async () => {
  const p = product({}, { evenAndOdd: true });
  const receipt = await p.tool.handler(req([{ intent: 'edit_footer', section: SECTION, text: TEXT, variant: 'even' }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);

  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.footer.even.paragraphs, 1, JSON.stringify(sec.footer));
  assert.equal(sec.footer.even.paras[0].text, TEXT, JSON.stringify(sec.footer));
  assert.equal(sec.footer.default, null, JSON.stringify(sec.footer));
});

// ---- the #260/#259 default-only path stays unaffected --------------------------------------------------------------------------

test('UNCHANGED: an omitted `variant` (default) still works exactly as GH #259/#260 released it, even on a section with title ON', async () => {
  const p = product({}, { sections: [{ titlePage: true }] });
  const receipt = await p.tool.handler(req([{ intent: 'edit_header', section: SECTION, text: TEXT }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);
  const sections = await p.readHeadersFooters();
  const sec = sections.find(s => s.index === SECTION.index);
  assert.equal(sec.header.default.paras[0].text, TEXT, JSON.stringify(sec.header));
  assert.equal(sec.header.title, null, 'the default write must not have touched the title part');
});
