'use strict';

/**
 * GH #230 (DOCX P1 WRITE, core run-formatting continued; the release itself is an OWNER DECISION, alongside #228/#229):
 * format_text's `format` gains color/highlight/vertAlign -- EXACTLY the MEASURED W-08 value shapes
 * (src/adapters/docx/families/formatting.cjs's `normalizeW08`/`HEX_RE`/`HIGHLIGHT_VALUES`/`VERT_ALIGN_VALUES`, :166-184).
 * Design: docs/GH230-FORMAT-TEXT-COLOR-HIGHLIGHT-VERTALIGN-DECISION.md. This is productization of an ALREADY-LIVE-PROVEN
 * adapter capability through this SAME product intent (docs/evidence/w08-color-highlight-vertalign-acc3-PASS-20261006.json,
 * predating this unit) -- no new adapter behaviour.
 *
 * `color` accepts ONLY `#RRGGBB` (no 'auto'/reset write value -- never measured); READ can still observe 'auto' on a never-
 * explicitly-coloured run (src/drivers/eurooffice/docx-driver.cjs's `colorHex`, :186-195). `highlight`'s 'none' and
 * `vertAlign`'s 'baseline' ARE legitimate, measured, explicit WRITE values (unlike color's 'auto') -- see decision doc
 * section 2.
 *
 * Offline only -- the product handler runs the REAL docx adapter / executor / lifecycle on
 * test/fakes/fake-docx-driver.cjs's `applyFormat`/`applyRangeFormat` (both already carry color/highlight/vertAlign in
 * TEXT_KEYS, :70); the LIVE acceptance is test/live/gh230-format-text-color-highlight-vertalign-product.live.cjs (not run
 * by the author).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const FILE = '1303626';
// PLAIN mirrors the REAL driver's observed DEFAULT state for a never-explicitly-set run (colorHex's
// null -> 'auto'; highlight/vertAlign's own fixture defaults, docx-driver.cjs :186-195 + the W-08 header
// comment's "the fixture's own default is 'none' / 'baseline'"). WANTED reuses the EXACT values the LIVE
// evidence file already measured (docs/evidence/w08-color-highlight-vertalign-acc3-PASS-20261006.json).
const PLAIN = { color: 'auto', highlight: 'none', vertAlign: 'baseline' };
const WANTED = { color: '#2A7F3A', highlight: 'cyan', vertAlign: 'superscript' };
const TARGET = 'GH230 TARGET PARAGRAPH';
const OTHER = 'GH230 OTHER PARAGRAPH';
const RANGE_HOST = 'GH230B range host NEEDLE text';
const doc = () => [
  { text: 'Heading', level: 1, props: { ...PLAIN } },
  { text: TARGET, props: PLAIN },
  { text: OTHER, props: PLAIN },
  { text: RANGE_HOST, props: PLAIN },
  { text: 'GH230 closing filler', props: PLAIN }, // selectorFrom's `after` needs a non-empty neighbor
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

test('SCHEMA: a valid #RRGGBB color, each of the 17 measured highlight names, and each of the 3 measured vertAlign names are individually accepted', () => {
  for (const color of ['#2A7F3A', '#000000', '#FFFFFF', '#abcdef']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { color })]));
    assert.equal(v.ok, true, `${color} -> ${v.code} ${v.reason}`);
  }
  for (const highlight of ['none', 'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'red', 'white', 'yellow']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { highlight })]));
    assert.equal(v.ok, true, `${highlight} -> ${v.code} ${v.reason}`);
  }
  for (const vertAlign of ['baseline', 'subscript', 'superscript']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { vertAlign })]));
    assert.equal(v.ok, true, `${vertAlign} -> ${v.code} ${v.reason}`);
  }
});

test("SCHEMA: color REFUSES 'auto' and every non-#RRGGBB shape pre-dispatch -- no reset-to-auto write call was ever measured (decision doc section 2)", () => {
  for (const color of ['auto', 'red', '#2A7F3', '#2A7F3AA', '2A7F3A', '#GGGGGG', '', 123]) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { color })]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(color));
    assert.equal(v.dispatched, false, JSON.stringify(color));
  }
});

test('SCHEMA: an unknown highlight/vertAlign name is REFUSED_SHAPE (the schema is an exact enum of the measured set, not a free string)', () => {
  for (const highlight of ['pink', 'Cyan', 'CYAN', '', 1]) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { highlight })]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(highlight));
  }
  for (const vertAlign of ['sub', 'SUPERSCRIPT', '', 1]) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { vertAlign })]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(vertAlign));
  }
});

test('SCHEMA: color/highlight/vertAlign compose with the seven PRESERVED_PROPERTIES keys and with matchText in the SAME format object', () => {
  const combined = { ...WANTED, bold: true, fontSize: 18 };
  assert.equal(D.validateDocxWriteRequest(req([step({ text: TARGET }, combined)])).ok, true);
  assert.equal(D.validateDocxWriteRequest(req([{ intent: 'format_text', selector: { text: TARGET }, format: WANTED, matchText: 'NEEDLE' }])).ok, true);
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('EXACT TARGETING + PRESERVATION: one unique paragraph moves from the measured PLAIN defaults to the measured WANTED values; read_formatting symmetry proves it; the rest of the document is untouched', async () => {
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

test("DEFAULT STATE FIXTURE: the never-explicitly-set state is modeled as literal 'auto'/'none'/'baseline' strings, not null/undefined -- this is the PRODUCT-FACING shape the REAL driver's colorHex() translates null INTO (docx-driver.cjs:186-195, LIVE-measured, not reproduced by the fake transport here: the fake driver passes props through verbatim and never runs colorHex's own null -> 'auto' translation)", () => {
  const target = doc().find(x => x.text === TARGET);
  assert.equal(target.props.color, 'auto');
  assert.equal(target.props.highlight, 'none');
  assert.equal(target.props.vertAlign, 'baseline');
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

test("EXPLICIT RESET via a measured value: highlight:'none' and vertAlign:'baseline' are ordinary, APPLIED writes (unlike color, which has no reset value at all)", async () => {
  const p = product();
  const toRich = await p.tool.handler(req([step(selectorFrom(doc(), TARGET), { highlight: 'yellow', vertAlign: 'subscript' })]));
  assert.equal(toRich.ok, true, JSON.stringify(toRich));
  const reset = await p.tool.handler(req([step(selectorFrom(doc(), TARGET), { highlight: 'none', vertAlign: 'baseline' })]));
  assert.equal(reset.ok, true, JSON.stringify(reset));
  assert.deepEqual(reset.operations.map(o => o.status), ['APPLIED']);
  const target = p.store.durable.find(x => x.text === TARGET);
  assert.equal(target.format.highlight, 'none');
  assert.equal(target.format.vertAlign, 'baseline');
});

test('RANGE (matchText) composition: color/highlight/vertAlign apply correctly inside a substring, the rest of the SAME paragraph keeps the measured PLAIN defaults', async () => {
  const p = product();
  const r = await p.tool.handler(req([rangeStep(selectorFrom(doc(), RANGE_HOST), 'NEEDLE')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  const target = p.store.durable.find(x => x.text === RANGE_HOST);
  const needle = target.runs.find(x => x.text === 'NEEDLE');
  assert.ok(needle, 'the matched substring is its own run');
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(needle.props[k], v, `${k} applied inside the range`);
  const idx = RANGE_HOST.indexOf('NEEDLE');
  const prefix = target.runs.find(x => x.text === RANGE_HOST.slice(0, idx));
  const suffix = target.runs.find(x => x.text === RANGE_HOST.slice(idx + 'NEEDLE'.length));
  assert.ok(prefix && suffix, 'the untouched prefix/suffix survive as their own runs');
  for (const [k, v] of Object.entries(PLAIN)) { assert.deepEqual(prefix.props[k], v, `prefix ${k} unchanged`); assert.deepEqual(suffix.props[k], v, `suffix ${k} unchanged`); }
});

test('PRE-DISPATCH refusals never open a session: a malformed color refuses before any driver opens', async () => {
  const p = product();
  const r = await p.tool.handler(req([step({ text: TARGET }, { color: 'not-a-color' })]));
  assert.equal(r.code, 'REFUSED_SHAPE');
  assert.equal(r.dispatched, false);
  assert.equal(p.opens(), 0);
});
