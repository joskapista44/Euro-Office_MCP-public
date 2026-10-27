'use strict';

/**
 * GH #228 (DOCX P1 WRITE, PREPARATION; the release itself is an OWNER DECISION): the PROPOSED state of the DOCX WRITE release --
 * docx_execute_task serving exactly `rewrite_paragraph` behind a PER-FORMAT WRITE release gate. Design:
 * docs/GH228-DOCX-WRITE-RELEASE-GATE-PROPOSAL.md. These tests describe the state AFTER the owner's GO; before it, the gate commit
 * that makes them true is not merged (the proposal's "BEFORE" column).
 *
 * Offline only. The product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs (the same
 * transport docx-rewrite-slice.test.cjs and gh220-docx-write-lifecycle.test.cjs use); the LIVE acceptance is
 * test/live/docx-rewrite-paragraph-product.live.cjs (not run by the author).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const FILE = '1303626';
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const OLD = 'GH228 OLD PARAGRAPH';
const NEW = 'GH228 NEW PARAGRAPH';
const TWIN = 'GH228 TWIN';
const doc = () => [
  { text: 'Heading', level: 1, props: { ...RICH, bold: false, italic: false } },
  { text: OLD, props: RICH },
  { text: TWIN, props: { ...RICH, bold: false } },
  { text: TWIN, props: { ...RICH, bold: false } },
  { text: 'Closing paragraph.', props: { ...RICH, bold: false, italic: false } },
];
const step = (selector, text = NEW, extra = {}) => ({ intent: 'rewrite_paragraph', selector, text, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });

function selectorFrom(paragraphs, target) {
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const b = primary.blocks.find(x => x.text === target);
  return { text: target, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}

function product() {
  const store = createStore(doc());
  const drivers = [];
  // saveObserver: the fake's save-confirmation observer (gh220's setup does the same)
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  const writes = () => store.saves;
  return { store, tool, drivers, writes, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}

// ---- the step schema (pre-dispatch) ------------------------------------------------------------------------------

test('SCHEMA: every stable-selector form the resolver reads is accepted pre-dispatch (text, heading+level, fingerprint, sectionHeading+paragraphOrdinal, qualifiers, precondition)', () => {
  const forms = [
    { text: OLD }, { text: OLD, contains: true, kind: 'paragraph', before: 'x', after: 'y' },
    { text: OLD, headingPath: [{ level: 1, text: 'Heading' }] },
    { heading: 'Heading', level: 1 }, { fingerprint: 'abc123' },
    { sectionHeading: 'Heading', paragraphOrdinal: 2 }, { text: OLD, precondition: { text: OLD } },
  ];
  for (const sel of forms) {
    const v = D.validateDocxWriteRequest(req([step(sel)]));
    assert.equal(v.ok, true, `${JSON.stringify(sel)} -> ${v.code} ${v.reason}`);
  }
  assert.equal(D.validateDocxWriteRequest(req([step({ text: OLD }, NEW, { expectedPreservation: { bold: true, fontSize: 22 }, id: 'r1' })])).ok, true);
});

test('SCHEMA: exact targeting -- physical locators, bookmark / tableCell and ref are NOT in the schema (UNKNOWN_FIELD, 0 dispatch); a selector with no resolvable branch is REFUSED_SHAPE', () => {
  for (const k of ['index', 'position', 'offset', 'paragraphIndex', 'bookmark', 'tableCell', 'ref']) {
    const v = D.validateDocxWriteRequest(req([step({ text: OLD, [k]: 1 })]));
    assert.equal(v.code, 'UNKNOWN_FIELD', k);
    assert.equal(v.dispatched, false, k);
  }
  for (const sel of [{}, { kind: 'heading' }, { before: 'x' }, { precondition: { text: OLD } }, { sectionHeading: 'Heading' }, { paragraphOrdinal: 1 }]) {
    const v = D.validateDocxWriteRequest(req([step(sel)]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(sel));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: malformed values are refused pre-dispatch (empty / missing / non-string text, bad level / ordinal, dependsOn, unknown preserved property)', () => {
  const cases = [
    [step({ text: OLD }, ''), 'REFUSED_SHAPE'], [{ intent: 'rewrite_paragraph', selector: { text: OLD } }, 'REFUSED_SHAPE'],
    [step({ text: OLD }, 5), 'REFUSED_SHAPE'], [step({ heading: 'H', level: 0 }), 'REFUSED_SHAPE'], [step({ heading: 'H', level: 10 }), 'REFUSED_SHAPE'],
    [step({ sectionHeading: 'H', paragraphOrdinal: 0 }), 'REFUSED_SHAPE'], [step({ text: OLD }, NEW, { dependsOn: ['s0'] }), 'UNKNOWN_FIELD'],
    [step({ text: OLD }, NEW, { expectedPreservation: { color: 'red' } }), 'UNKNOWN_FIELD'], [step({ text: OLD }, NEW, { expectedPreservation: {} }), 'REFUSED_SHAPE'],
  ];
  for (const [s, code] of cases) { const v = D.validateDocxWriteRequest(req([s])); assert.equal(v.code, code, `${JSON.stringify(s)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
});

test('ALLOWLIST: every OTHER routed docx intent is REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid rewrite_paragraph (whole task refused, step index named)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i); // GH #229-#231 / #233 / #234 / #235 / #236 / #237 / #249: format_text and the insert, delete, move, list and table-cell intents joined the allowlist
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i); // GH #233 / #234 / #235 / #236: the insert, delete, move and list intents joined the allowlist
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([step({ text: OLD }), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema offers the rewrite_paragraph step first (GH #233 / #234 / #235 / #236 / #237 / #249 / #229 add the insert, delete, move, list, table-cell and format_text steps after it)', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  assert.equal(items.oneOf.length, 38); // release 3 trim + GH #266/#269e/#268d: 38 released intents
  assert.equal(items.oneOf[0].properties.intent.const, 'rewrite_paragraph');
  assert.deepEqual(items.oneOf[0].required, ['intent', 'selector', 'text']);
  assert.deepEqual(items.oneOf.find(x => x.properties.intent.const === 'format_text').required, ['intent', 'selector', 'format']);
  assert.match(D.createDocxWriteTool({}).description, /rewrite_paragraph/);
  assert.match(D.createDocxWriteTool({}).description, /format_text/);
});

// ---- the product handler, end to end on the fake transport ------------------------------------------------------------

test('EXACT TARGETING + PRESERVATION: one unique paragraph is rewritten; the seven run properties survive; the receipt is APPLIED and the rest of the document is untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([step(selectorFrom(doc(), OLD))]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.tool, 'docx_execute_task');
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.noOp, false);
  assert.equal(p.writes() >= 1, true);
  const after = p.store.durable;
  const rewritten = after.find(x => x.text === NEW);
  assert.ok(rewritten, 'the new text is durable');
  for (const k of ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language']) assert.deepEqual(rewritten.props[k], RICH[k], `run property ${k} preserved`);
  assert.equal(after.some(x => x.text === OLD), false, 'the old text is gone');
  assert.deepEqual(after.filter(x => x.text === TWIN).length, 2, 'the other paragraphs are untouched');
  assert.equal(after.length, doc().length);
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes (idempotent, restart-safe)', async () => {
  const p = product();
  const task = req([step(selectorFrom(doc(), OLD))]);
  const first = await p.tool.handler(task);
  assert.equal(first.ok, true, JSON.stringify(first));
  const savesAfterFirst = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), savesAfterFirst, 'no write on the retry');
});

test('AMBIGUOUS TARGET fails closed: two paragraphs with the same text -> the task is refused, ZERO writes, the document is unchanged', async () => {
  const p = product();
  const before = JSON.stringify(p.store.durable);
  const r = await p.tool.handler(req([step({ text: TWIN })]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(JSON.stringify(p.store.durable), before);
  assert.equal(p.writes(), 0);
  assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
});

test('MISSING TARGET fails closed: a text that is nowhere in the document -> refused, ZERO writes', async () => {
  const p = product();
  const before = JSON.stringify(p.store.durable);
  const r = await p.tool.handler(req([step({ text: 'GH228 NOT IN THE DOCUMENT' })]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(JSON.stringify(p.store.durable), before);
  assert.equal(p.writes(), 0);
});

test('PRECONDITION mismatch fails closed: the resolved target text differs from selector.precondition.text -> refused, ZERO writes', async () => {
  const p = product();
  const r = await p.tool.handler(req([step({ ...selectorFrom(doc(), OLD), precondition: { text: 'something else' } })]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(p.writes(), 0);
  assert.equal(p.store.durable.some(x => x.text === NEW), false);
});

test('PRE-DISPATCH refusals never open a session: a refused task opens zero drivers', async () => {
  const p = product();
  // GH #229 (PREPARATION, owner decision pending): format_text joined the allowlist too -- a still-unallowlisted
  // routed intent is needed here instead.
  const notReleased = [...D.docxRoutedIntents()].find(i => !W.DOCX_WRITE_ALLOWLIST.includes(i));
  if (notReleased === undefined) return; // every routed DOCX intent is released in the final union (GH #267 B joined); the release trimming makes this check live again
  const r = await p.tool.handler(req([{ intent: notReleased, selector: { text: OLD } }]));
  assert.equal(r.code, 'REFUSED_NOT_ALLOWLISTED');
  assert.equal(r.dispatched, false);
  assert.equal(p.opens(), 0);
});

// ---- the per-format release gate ----------------------------------------------------------------------------------------

test('GATE: the proposed state passes both halves; the two halves are independent of each other', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  // BEFORE-state docx allowlist ([]): only the docx half fails
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: [], format: 'docx' }).ok, false);
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: [], format: 'xlsx' }).ok, true);
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: [] }).ok, false, 'the all-formats view (product_status) still reports FAIL');
  // a widened XLSX list: only the xlsx half fails
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...xlsx, 'clear_print_titles'], docxAllowlist: [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST], format: 'docx' }).ok, true);
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...xlsx, 'clear_print_titles'], docxAllowlist: [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST], format: 'xlsx' }).ok, false);
});

test('GATE: the BEFORE state (docx allowlist []) leaves xlsx_execute_task served and withholds ONLY docx_execute_task -- the xlsx release is unchanged', () => {
  const before = { xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST], docxAllowlist: [] };
  assert.deepEqual(T.exposedTools({ writeGate: before }), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.deepEqual(T.exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
});

test('GATE: the product_status view carries the all-formats verdict (PASS in the proposed state)', () => {
  const g = T.exposureGate();
  assert.equal(g.writeReleaseGate, 'PASS');
  assert.deepEqual(g.releasedCapabilityTools, ['xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
});
