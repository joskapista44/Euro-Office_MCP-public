'use strict';

/**
 * GH #229 (core run-formatting, PREPARATION; the release itself is an OWNER DECISION, alongside #228): format_text joining
 * the DOCX WRITE allowlist next to rewrite_paragraph, behind the SAME per-format gate. Design: docs/GH229-FORMAT-TEXT-PRODUCT-DECISION.md.
 * Selector reuses REWRITE_SELECTOR verbatim (whole-paragraph targeting for the parent); format accepts the seven
 * PRESERVED_PROPERTIES keys (src/adapters/docx/families/text.cjs:28) PLUS (GH #230, docs/GH230-FORMAT-TEXT-COLOR-HIGHLIGHT-
 * VERTALIGN-DECISION.md) color/highlight/vertAlign PLUS (GH #231, docs/GH231-FORMAT-TEXT-PARAGRAPH-PROPS-DECISION.md)
 * alignment/spaceBefore/spaceAfter/indentLeft/indentRight/indentFirstLine/shading PLUS (GH #232, docs/GH232-FORMAT-TEXT-TABS-BORDERS-
 * PAGINATION-DECISION.md) tabs/border{Top,Bottom,Left,Right}/keepNext/keepLines/pageBreakBefore/widowControl and (GH #231b, decision doc section 8) `lineSpacing`, once its READ round-trip defect (GH231 decision doc section 5) was fixed -- `style` stays
 * out.
 *
 * GH #229 SECOND SLICE (matchText / substring-range targeting, decision doc section 12): the parent's own OPTIONAL literal
 * substring, reusing the adapter's EXISTING ambiguity contract (families/formatting.cjs's `isRange`/`rangeState`/
 * `rangeTerminal`/`parentAddressable`, :260-373) unchanged -- never re-implemented in this schema. A non-empty string is the
 * only pre-dispatch check; substring-within-parent ambiguity and the parent's own document-wide text uniqueness are
 * EXECUTION-time fail-closed classifications, the same way whole-paragraph selector ambiguity already is.
 *
 * These tests describe the state AFTER the owner's GO; before it, the gate commit that makes them true is not merged (the
 * #228 proposal's "BEFORE" column, extended the same way for this intent). Offline only -- the product handler runs the REAL
 * docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs's `applyFormat`/`applyRangeFormat` (dispatch chosen
 * by `meta.matchText`'s presence, test/fakes/fake-docx-driver.cjs:1084); the LIVE acceptance is
 * test/live/gh229-format-text-product.live.cjs (not run by the author).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const FILE = '1303626';
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const WANTED = { bold: false, italic: false, underline: true, strike: true, fontFamily: 'Calibri', fontSize: 14, language: 'en-US' };
const TARGET = 'GH229 TARGET PARAGRAPH';
const OTHER = 'GH229 OTHER PARAGRAPH';
const TWIN = 'GH229 TWIN';
// GH #229 second slice (matchText / range targeting) fixtures:
const RANGE_HOST = 'GH229B prefix NEEDLE suffix'; // a unique parent whose own text occurs once; NEEDLE occurs once inside it
const DOUBLE_HOST = 'GH229B double NEEDLE and NEEDLE again'; // a unique parent, but NEEDLE occurs TWICE inside it
// MIXED_HOST's own text, pre-split into two runs with DIFFERENT starting formats, where the matchText
// substring 'NEE' + 'DLE' spans the run boundary -- the "mixed starting runs" case.
const MIXED_HOST_PREFIX = 'GH229B mixed before NEE';
const MIXED_HOST_SUFFIX = 'DLE after';
const MIXED_HOST = MIXED_HOST_PREFIX + MIXED_HOST_SUFFIX;
const MIXED_RUN_A = { ...RICH, bold: true, fontFamily: 'Arial' };
const MIXED_RUN_B = { ...RICH, bold: false, fontFamily: 'Courier New' };
// RANGE_TWIN's own text repeats TWICE in the document (like TWIN), but with DIFFERENT neighbors each
// time -- the selector (before/after context) can therefore resolve to ONE specific occurrence, unlike
// TWIN. This isolates `parentAddressable`'s document-wide, anchor-BLIND text check (families/
// formatting.cjs:344-353) from ordinary selector ambiguity.
const RANGE_TWIN = 'GH229B range twin NEEDLE host';
const doc = () => [
  { text: 'Heading', level: 1, props: { ...RICH, bold: false, italic: false } },
  { text: TARGET, props: RICH },
  { text: OTHER, props: RICH },
  { text: TWIN, props: { ...RICH, bold: false } },
  { text: TWIN, props: { ...RICH, bold: false } },
  { text: RANGE_HOST, props: RICH },
  { text: RANGE_TWIN, props: RICH },
  { text: DOUBLE_HOST, props: RICH },
  { text: RANGE_TWIN, props: RICH },
  { text: MIXED_HOST, runs: [{ text: MIXED_HOST_PREFIX, props: MIXED_RUN_A }, { text: MIXED_HOST_SUFFIX, props: MIXED_RUN_B }] },
  // trailing filler: selectorFrom's generated `after` context must be non-empty (SEL_TEXT minLength:1) --
  // without a paragraph following it, MIXED_HOST (the last block) would have no valid anchored selector.
  { text: 'GH229B closing filler', props: RICH },
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

/** read_formatting's own symmetry proof (src/server/read-tools.cjs:403-405, `formatOf`/`unitsOf`): rebuild the inventory
 * from the post-write live paragraphs and read the SAME `.format` field read_formatting's projection exposes. */
function readFormattingOf(paragraphs, text) {
  const inv = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  return inv.blocks.find(b => b.text === text).format;
}

/** A range write (`applyRangeFormat`) never touches the paragraph-level `.format` field (only `paraKeys`,
 * always empty for this unit's seven text-only keys) -- it writes straight into `run.props`, with the SAME
 * half-points `fontSize*2` mirror `.props.fontSize` carries everywhere else in this driver (see the
 * SINGLE-PROPERTY write test's own comment). This undoes that ONE conversion so a run's stored `.props` can
 * be compared directly against the product-facing (points) `format` values this unit's schema uses. */
function pointsOf(runProps) {
  return { ...runProps, ...(runProps.fontSize !== undefined ? { fontSize: runProps.fontSize / 2 } : {}) };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------

test('SCHEMA: every stable-selector form the resolver reads is accepted pre-dispatch, with a format step (text, heading+level, fingerprint, sectionHeading+paragraphOrdinal, before/after qualifiers, precondition)', () => {
  const forms = [
    { text: TARGET }, { text: TARGET, kind: 'paragraph', before: 'x', after: 'y' },
    { text: TARGET, headingPath: [{ level: 1, text: 'Heading' }] },
    { heading: 'Heading', level: 1 }, { fingerprint: 'abc123' },
    { sectionHeading: 'Heading', paragraphOrdinal: 2 }, { text: TARGET, precondition: { text: TARGET } },
  ];
  for (const sel of forms) {
    const v = D.validateDocxWriteRequest(req([step(sel)]));
    assert.equal(v.ok, true, `${JSON.stringify(sel)} -> ${v.code} ${v.reason}`);
  }
});

test("ADAPTER LIMITATION, NOT a schema gap: selector.contains:true is schema-valid (it's part of REWRITE_SELECTOR) but refused at plan-validation -- format_text's own plan() has not migrated substring-match-by-contains in M4 Wave 1 (src/adapters/docx/families/formatting.cjs:267); rewrite_paragraph has no such restriction", () => {
  const r = D.validateDocxWriteRequest(req([step({ text: TARGET, contains: true })]));
  assert.equal(r.code, 'REFUSED_SHAPE');
  assert.equal(r.dispatched, false);
  assert.match(r.reason, /not migrated in M4 Wave 1/);
});

test('SCHEMA: each of the seven run properties is individually accepted; physical locators and bookmark/tableCell/ref are NOT in the schema (UNKNOWN_FIELD)', () => {
  for (const [k, v] of Object.entries(WANTED)) {
    const r = D.validateDocxWriteRequest(req([step({ text: TARGET }, { [k]: v })]));
    assert.equal(r.ok, true, `${k} -> ${r.code} ${r.reason}`);
  }
  for (const k of ['index', 'position', 'offset', 'paragraphIndex', 'bookmark', 'tableCell', 'ref']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET, [k]: 1 })]));
    assert.equal(v.code, 'UNKNOWN_FIELD', k);
    assert.equal(v.dispatched, false, k);
  }
});

// ---- GH #229 second slice: matchText (substring/range targeting) --------------------------------------------------------

test('SCHEMA: matchText is now accepted pre-dispatch (a non-empty string); empty/whitespace-only/wrong-type is REFUSED_SHAPE', () => {
  const r = D.validateDocxWriteRequest(req([{ intent: 'format_text', selector: { text: TARGET }, format: WANTED, matchText: 'NEEDLE' }]));
  assert.equal(r.ok, true, `${r.code} ${r.reason}`);
  for (const bad of ['', '   ', 123, null]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'format_text', selector: { text: TARGET }, format: WANTED, matchText: bad }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(bad));
    assert.equal(v.dispatched, false, JSON.stringify(bad));
  }
});

test("SCHEMA: an unknown format field (style/borderBetween; color/highlight/vertAlign/alignment/spaceBefore/indentLeft/shading joined the schema in GH #230/#231, tabs/border*/pagination in GH #232, lineSpacing in GH #231b once its READ defect was fixed) is UNKNOWN_FIELD; an empty format or an unresolvable selector is REFUSED_SHAPE", () => {
  for (const k of ['style', 'borderBetween']) {
    const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, { [k]: 'x' })]));
    assert.equal(v.code, 'UNKNOWN_FIELD', k);
  }
  assert.equal(D.validateDocxWriteRequest(req([step({ text: TARGET }, {})])).code, 'REFUSED_SHAPE', 'empty format names nothing to change');
  for (const sel of [{}, { kind: 'heading' }, { before: 'x' }, { precondition: { text: TARGET } }, { sectionHeading: 'Heading' }, { paragraphOrdinal: 1 }]) {
    const v = D.validateDocxWriteRequest(req([step(sel)]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(sel));
  }
});

test('SCHEMA: malformed property values are refused pre-dispatch (wrong type, out-of-range fontSize, empty fontFamily/language)', () => {
  const cases = [
    [{ bold: 'yes' }, 'REFUSED_SHAPE'], [{ italic: 1 }, 'REFUSED_SHAPE'],
    [{ fontFamily: '' }, 'REFUSED_SHAPE'], [{ fontFamily: 123 }, 'REFUSED_SHAPE'],
    [{ fontSize: 0 }, 'REFUSED_SHAPE'], [{ fontSize: -1 }, 'REFUSED_SHAPE'], [{ fontSize: 1639 }, 'REFUSED_SHAPE'], [{ fontSize: '14' }, 'REFUSED_SHAPE'],
    [{ language: '' }, 'REFUSED_SHAPE'], [{ language: 'x'.repeat(36) }, 'REFUSED_SHAPE'],
  ];
  for (const [format, code] of cases) { const v = D.validateDocxWriteRequest(req([step({ text: TARGET }, format)])); assert.equal(v.code, code, `${JSON.stringify(format)} -> ${v.code} ${v.reason}`); }
  // the boundary values themselves are accepted
  assert.equal(D.validateDocxWriteRequest(req([step({ text: TARGET }, { fontSize: 1 })])).ok, true);
  assert.equal(D.validateDocxWriteRequest(req([step({ text: TARGET }, { fontSize: 1638 })])).ok, true);
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('EXACT TARGETING + PRESERVATION: one unique paragraph gets the new formatting; read_formatting symmetry proves it; the rest of the document is untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([step(selectorFrom(doc(), TARGET))]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.tool, 'docx_execute_task');
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.noOp, false);
  assert.equal(p.writes() >= 1, true);
  const after = p.store.durable;
  const target = after.find(x => x.text === TARGET);
  assert.ok(target, 'the target paragraph survives (format_text never renames text)');
  // `.format` (not `.props`) is the points-convention, read_formatting-symmetric field -- `.props.fontSize` carries the
  // driver's internal half-points mirror (test/fakes/fake-docx-driver.cjs's own SetFontSize convention) and is not part
  // of this product's public contract.
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(target.format[k], v, `format.${k} applied`);
  // read_formatting's own symmetry proof (src/server/read-tools.cjs:403-405, `formatOf`/`unitsOf`): the inventory
  // projection built from the SAME live paragraphs reads the identical field.
  const read = readFormattingOf(after, TARGET);
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(read[k], v, `read_formatting.format.${k} proves the write`);
  const other = after.find(x => x.text === OTHER);
  for (const [k, v] of Object.entries(RICH)) assert.deepEqual(other.props[k], v, `untouched paragraph's ${k} preserved`);
  assert.deepEqual(after.filter(x => x.text === TWIN).length, 2, 'the twin paragraphs are untouched');
  assert.equal(after.length, doc().length, 'no paragraph added or removed');
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

test('AMBIGUOUS TARGET fails closed: two paragraphs with the same text -> the task is refused, ZERO writes, the document is unchanged (the shared selector resolver, not a per-intent behavior)', async () => {
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
  const r = await p.tool.handler(req([step({ text: 'GH229 NOT IN THE DOCUMENT' })]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(JSON.stringify(p.store.durable), before);
  assert.equal(p.writes(), 0);
});

test('PRECONDITION mismatch fails closed: the resolved target text differs from selector.precondition.text -> refused, ZERO writes', async () => {
  const p = product();
  const r = await p.tool.handler(req([step({ ...selectorFrom(doc(), TARGET), precondition: { text: 'something else' } })]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(p.writes(), 0);
  const target = p.store.durable.find(x => x.text === TARGET);
  assert.deepEqual(target.props, RICH, 'the precondition-mismatched paragraph keeps its original formatting');
});

test('PRE-DISPATCH refusals never open a session: a refused task opens zero drivers', async () => {
  const p = product();
  const notReleased = [...D.docxRoutedIntents()].find(i => !W.DOCX_WRITE_ALLOWLIST.includes(i));
  if (notReleased === undefined) return; // every routed DOCX intent is released in the final union (GH #267 B joined); the release trimming makes this check live again
  const r = await p.tool.handler(req([{ intent: notReleased, selector: { text: TARGET } }]));
  assert.equal(r.code, 'REFUSED_NOT_ALLOWLISTED');
  assert.equal(r.dispatched, false);
  assert.equal(p.opens(), 0);
});

test('SINGLE-PROPERTY write: changing only one of the seven leaves the other six exactly as they were', async () => {
  const p = product();
  const r = await p.tool.handler(req([step(selectorFrom(doc(), TARGET), { fontSize: 36 })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const target = p.store.durable.find(x => x.text === TARGET);
  assert.equal(target.format.fontSize, 36);
  for (const k of ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'language']) assert.deepEqual(target.format[k], RICH[k], `${k} untouched by a fontSize-only write`);
});

// ---- GH #229 second slice: matchText (substring/range targeting), end to end on the fake transport ---------------------

test('RANGE APPLIED + PRESERVATION: a unique substring inside a unique parent gets the new formatting; the rest of the SAME paragraph and every other paragraph are untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([rangeStep(selectorFrom(doc(), RANGE_HOST), 'NEEDLE')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(p.writes() >= 1, true);
  const after = p.store.durable;
  const target = after.find(x => x.text === RANGE_HOST);
  assert.ok(Array.isArray(target.runs), 'the range write produces an explicit run split');
  const needle = target.runs.find(x => x.text === 'NEEDLE');
  assert.ok(needle, 'the matched substring is its own run');
  const needlePts = pointsOf(needle.props);
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(needlePts[k], v, `format.${k} applied inside the range`);
  const idx = RANGE_HOST.indexOf('NEEDLE');
  const prefix = target.runs.find(x => x.text === RANGE_HOST.slice(0, idx));
  const suffix = target.runs.find(x => x.text === RANGE_HOST.slice(idx + 'NEEDLE'.length));
  assert.ok(prefix && suffix, 'the untouched prefix/suffix survive as their own runs');
  assert.deepEqual(prefix.props, RICH, 'prefix (outside the range) keeps its original formatting');
  assert.deepEqual(suffix.props, RICH, 'suffix (outside the range) keeps its original formatting');
  for (const text of [TARGET, OTHER]) { const other = after.find(x => x.text === text); for (const [k, v] of Object.entries(RICH)) assert.deepEqual(other.props[k], v, `untouched paragraph ${text}'s ${k} preserved`); }
  assert.equal(after.length, doc().length, 'no paragraph added or removed');
});

test('RANGE on MIXED starting runs: a matchText spanning a run boundary with DIFFERENT starting formats still ends up uniform inside the range, and each side keeps its OWN original (different) formatting outside it', async () => {
  const p = product();
  const r = await p.tool.handler(req([rangeStep(selectorFrom(doc(), MIXED_HOST), 'NEEDLE')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  const target = p.store.durable.find(x => x.text === MIXED_HOST);
  const needle = target.runs.find(x => x.text === 'NEEDLE');
  assert.ok(needle, 'the matched substring (spanning the original run boundary) is its own, now-uniform run');
  const needlePts = pointsOf(needle.props);
  for (const [k, v] of Object.entries(WANTED)) assert.deepEqual(needlePts[k], v, `format.${k} applied uniformly across the former boundary`);
  const idx = MIXED_HOST.indexOf('NEEDLE');
  const before = target.runs.find(x => x.text === MIXED_HOST.slice(0, idx));
  const afterRun = target.runs.find(x => x.text === MIXED_HOST.slice(idx + 'NEEDLE'.length));
  assert.ok(before && afterRun, 'both sides of the range survive as their own runs');
  assert.deepEqual(before.props, MIXED_RUN_A, 'the side that started as run A keeps its OWN original formatting, not run B\'s');
  assert.deepEqual(afterRun.props, MIXED_RUN_B, 'the side that started as run B keeps its OWN original formatting, not run A\'s');
});

test('AMBIGUOUS SUBSTRING WITHIN THE PARAGRAPH fails closed: matchText occurring twice inside an otherwise-unique parent -> refused, ZERO writes, unchanged', async () => {
  const p = product();
  const before = JSON.stringify(p.store.durable);
  const r = await p.tool.handler(req([rangeStep(selectorFrom(doc(), DOUBLE_HOST), 'NEEDLE')]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(JSON.stringify(p.store.durable), before);
  assert.equal(p.writes(), 0);
  assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
});

test("DUPLICATE RAW TEXT fails closed EVEN WHEN THE SELECTOR ITSELF RESOLVES VIA AN ANCHOR -- for BOTH whole-paragraph and range format_text: inventory.cjs's own .format/.props synthesis re-matches by EXACT TEXT (ignoring before/after) and returns null once that text is not globally unique (inventory.cjs:191-194), so classify()'s 'format-state-unreadable' guard fires before the range-specific selector/plan code ever runs", async () => {
  const sel = selectorFrom(doc(), RANGE_TWIN); // anchored to ONE specific occurrence (its own before/after context) --
  // resolveStableTarget (resolver.cjs) DOES pick exactly this one via the anchor; the failure below is NOT a selector-
  // resolution ambiguity, it is the format-family's OWN, anchor-blind re-match on raw text.
  const wholeParagraphProbe = await (async () => { const q = product(); return q.tool.handler(req([step(sel)])); })();
  assert.equal(wholeParagraphProbe.ok, false, JSON.stringify(wholeParagraphProbe));
  assert.match(wholeParagraphProbe.reason || '', /format-state-unreadable/, 'same reason as the range case below -- this is NOT range-specific');
  const p = product();
  const before = JSON.stringify(p.store.durable);
  const r = await p.tool.handler(req([rangeStep(sel, 'NEEDLE')]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.match(r.reason || '', /format-state-unreadable/);
  assert.equal(JSON.stringify(p.store.durable), before);
  assert.equal(p.writes(), 0);
  assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
});

test('RANGE RETRY: the SAME range task again is SATISFIED / noOp with ZERO new writes', async () => {
  const p = product();
  const task = req([rangeStep(selectorFrom(doc(), RANGE_HOST), 'NEEDLE')]);
  const first = await p.tool.handler(task);
  assert.equal(first.ok, true, JSON.stringify(first));
  const savesAfterFirst = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), savesAfterFirst, 'no write on the retry');
});
