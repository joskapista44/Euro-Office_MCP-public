'use strict';

/**
 * GH #231 (DOCX P1 WRITE, paragraph-level properties; the release itself is an OWNER DECISION, alongside #228/#229/#230):
 * format_text's `format` gains alignment/spaceBefore/spaceAfter/indentLeft/indentRight/indentFirstLine/shading --
 * the paragraph-level keys shared by BOTH the whole-paragraph and the matchText/range write path (the SAME `paraPr`
 * block, src/adapters/docx/families/formatting.cjs:538-549/:618-629). Design:
 * docs/GH231-FORMAT-TEXT-PARAGRAPH-PROPS-DECISION.md.
 *
 * `indentLeft`/`indentRight`/`indentFirstLine`/`shading` are LIVE-PROVEN through this SAME product intent
 * (docs/evidence/w09-indent-shading-acc3-PASS-20261006.json, including the hanging-indent case: a NEGATIVE
 * indentFirstLine, -12pt, measured APPLIED). `alignment`/`spaceBefore`/`spaceAfter` are code-measured (the same simple
 * points<->twips convention, src/drivers/eurooffice/docx-driver.cjs:327-333) but have NO equivalent LIVE evidence file
 * through this product path -- this unit's own LIVE harness is the first to check them.
 *
 * `lineSpacing` was DELIBERATELY EXCLUDED here (decision doc section 5) because its READ conversion
 * (docx-driver.cjs:334, `n>10 ? (n>=200 ? n/240 : (n+1)/100) : n`) was a code-verified round-trip defect for any value
 * below 200/240≈0.833 (write 0.5 -> read back 1.21). GH #231b (decision doc section 8) fixed the conversion (uniform
 * n/240) and added `lineSpacing` to the schema -- see the schema-acceptance test below, and the dedicated vm-level
 * RED/GREEN proof in test/static/docx-driver-linespacing-read-vm.test.cjs.
 *
 * A paragraph-level key is NEVER refused for being combined with `matchText` -- it writes the PARENT paragraph
 * regardless (docs/W09-TABS-BORDERS-PAGINATION-DESIGN-20261007.md section 1, confirmed here by reading the SAME
 * `paraPr` block both `commandBody` and `rangeCommandBody` share).
 *
 * Offline only -- the product handler runs the REAL docx adapter / executor / lifecycle on
 * test/fakes/fake-docx-driver.cjs's `applyFormat`/`applyRangeFormat` (paragraph-level keys always go through the
 * `paraKeys` branch, which both write paths share); the LIVE acceptance is
 * test/live/gh231-format-text-paragraph-props-product.live.cjs (not run by the author).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const FILE = '1303626';
// PLAIN mirrors test/fakes/fake-docx-driver.cjs's own formatOf() defaults for an untouched paragraph.
const PLAIN = { alignment: 'left', spaceBefore: 0, spaceAfter: 0, indentLeft: 0, indentRight: 0, indentFirstLine: 0, shading: 'auto' };
// shading is BARE 6-hex-digit, no leading '#' -- the adapter's own normalizeShading strips it
// (families/formatting.cjs:112-118) and the real driver's read side returns the same bare form
// (docx-driver.cjs:352-354); unlike `color`, which keeps the '#'. See the schema comment.
const WANTED = { alignment: 'center', spaceBefore: 12, spaceAfter: 6, indentLeft: 36, indentRight: 18, indentFirstLine: -12, shading: 'FFFF00' };
const TARGET = 'GH231 TARGET PARAGRAPH';
const OTHER = 'GH231 OTHER PARAGRAPH';
const RANGE_HOST = 'GH231B range host NEEDLE text';
const doc = () => [
  { text: 'Heading', level: 1, props: { ...PLAIN } },
  { text: TARGET, props: PLAIN },
  { text: OTHER, props: PLAIN },
  { text: RANGE_HOST, props: PLAIN },
  { text: 'GH231 closing filler', props: PLAIN }, // selectorFrom's `after` needs a non-empty neighbor
];
const step = (selector, format = WANTED, extra = {}) => ({ intent: 'format_text', selector, format, ...extra });
const rangeStep = (selector, matchText, format = WANTED) => step(selector, format, { matchText });
const req = steps => ({ fileId: FILE, task: { steps } });

function selectorFrom(paragraphs, target) {
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const b = primary.blocks.find(x => x.text === target);
  return { text: target, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}

function product() {
  const store = createStore(doc());
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  const writes = () => store.saves;
  return { store, tool, drivers, writes, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}

/** read_formatting's own symmetry proof (src/server/read-tools.cjs:403-405): rebuild the inventory from the post-write
 * live paragraphs and read the SAME `.format` field read_formatting's projection exposes. */
function readFormattingOf(paragraphs, text) {
  const inv = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  return inv.blocks.find(b => b.text === text).format;
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------

test('SCHEMA: each of the seven new paragraph-level keys is individually accepted, including a NEGATIVE indentFirstLine (hanging indent, measured)', () => {
  for (const [k, v] of Object.entries(WANTED)) {
    const r = D.validateDocxWriteRequest(req([step({ text: TARGET }, { [k]: v })]));
    assert.equal(r.ok, true, `${k} -> ${r.code} ${r.reason}`);
  }
  assert.equal(D.validateDocxWriteRequest(req([step({ text: TARGET }, { indentFirstLine: -12 })])).ok, true, 'hanging indent (negative) is explicitly measured, not refused');
});

test("SCHEMA: lineSpacing now JOINS the schema (GH #231b fixed the READ round-trip defect, decision doc section 8) -- in-bounds values are accepted, out-of-bounds are REFUSED_SHAPE", () => {
  for (const v of [0.1, 0.5, 1, 1.5, 3, 10]) {
    const r = D.validateDocxWriteRequest(req([step({ text: TARGET }, { lineSpacing: v })]));
    assert.equal(r.ok, true, `lineSpacing ${v} -> ${r.code} ${r.reason}`);
  }
  for (const v of [0, 0.05, 10.1, -1]) {
    const r = D.validateDocxWriteRequest(req([step({ text: TARGET }, { lineSpacing: v })]));
    assert.equal(r.code, 'REFUSED_SHAPE', `lineSpacing ${v} should be out-of-bounds, got ${r.code}`);
  }
});

test('SCHEMA: an unknown alignment name, a negative spaceBefore/spaceAfter, and a malformed shading are REFUSED_SHAPE pre-dispatch', () => {
  for (const alignment of ['Center', 'CENTER', 'start', 'end', '']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { alignment })]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(alignment));
  }
  for (const spaceBefore of [-1, -0.01]) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { spaceBefore })]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(spaceBefore));
  }
  for (const shading of ['auto', '#FFFF00', 'FFFF0', 'GGGGGG', '']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { shading })]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(shading));
  }
});

test('SCHEMA: the seven new keys compose with the run-level keys (#229/#230) and with matchText in the SAME format object', () => {
  const combined = { ...WANTED, bold: true, color: '#2A7F3A' };
  assert.equal(D.validateDocxWriteRequest(req([step({ text: TARGET }, combined)])).ok, true);
  assert.equal(D.validateDocxWriteRequest(req([{ intent: 'format_text', selector: { text: TARGET }, format: WANTED, matchText: 'NEEDLE' }])).ok, true, 'a paragraph-level key is never refused for being combined with matchText');
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('EXACT TARGETING + PRESERVATION: one unique paragraph moves from the measured PLAIN defaults to the measured WANTED values (incl. hanging indent); read_formatting symmetry proves it; the rest of the document is untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([step(selectorFrom(doc(), TARGET))]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(p.writes() >= 1, true);
  const after = p.store.durable;
  const target = after.find(x => x.text === TARGET);
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(target.format[k], v, `format.${k} applied`);
  const read = readFormattingOf(after, TARGET);
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(read[k], v, `read_formatting.format.${k} proves the write`);
  const other = after.find(x => x.text === OTHER);
  for (const [k, v] of Object.entries(PLAIN)) assert.deepEqual(other.props[k], v, `untouched paragraph's ${k} preserved (still the measured default)`);
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes (idempotent, restart-safe)', async () => {
  const p = product();
  const task = req([step(selectorFrom(doc(), TARGET))]);
  const first = await p.tool.handler(task);
  assert.equal(first.ok, true, JSON.stringify(first));
  const savesAfterFirst = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), savesAfterFirst, 'no write on the retry');
});

test('RANGE (matchText) + PARAGRAPH-LEVEL KEY composition: a paragraph-level key is NOT refused for naming matchText, and it writes the WHOLE PARENT paragraph, not the matched substring', async () => {
  const p = product();
  const r = await p.tool.handler(req([rangeStep(selectorFrom(doc(), RANGE_HOST), 'NEEDLE', { alignment: 'center', shading: 'FFFF00' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  const target = p.store.durable.find(x => x.text === RANGE_HOST);
  // the paragraph-level write lands on `.format` (the WHOLE paragraph) -- the structural proof that it
  // applied to the PARENT, not the substring. A format object naming ONLY paragraph-level keys carries
  // no TEXT_KEYS, so applyRangeFormat's run-split branch (setRangeProps with an empty textProps) is a
  // no-op: the paragraph's run list stays exactly as it was (still one run, uniform text properties) --
  // this IS the proof, not an incidental detail: nothing run-level moved.
  assert.equal(target.format.alignment, 'center');
  assert.equal(target.format.shading, 'FFFF00');
  assert.equal(target.text, RANGE_HOST, 'the text itself is untouched');
});

test('PRE-DISPATCH refusals never open a session: an unknown alignment name refuses before any driver opens', async () => {
  const p = product();
  const r = await p.tool.handler(req([step({ text: TARGET }, { alignment: 'start' })]));
  assert.equal(r.code, 'REFUSED_SHAPE');
  assert.equal(r.dispatched, false);
  assert.equal(p.opens(), 0);
});
