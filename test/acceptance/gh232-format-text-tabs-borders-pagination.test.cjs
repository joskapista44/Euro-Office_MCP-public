'use strict';

/**
 * GH #232 (DOCX P1/P2 WRITE, tabs / per-side borders / pagination flags; the release itself is an OWNER DECISION, alongside
 * #228..#231): format_text's `format` gains tabs, borderTop/borderBottom/borderLeft/borderRight and keepNext/keepLines/
 * pageBreakBefore/widowControl -- the MEASURED W-09 axes (docs/W09-TABS-BORDERS-PAGINATION-DESIGN-20261007.md,
 * docs/evidence/w09-3-live-acceptance-PASS-20261007.json), through the SAME paragraph-level `paraPr` block as #231.
 * Design: docs/GH232-FORMAT-TEXT-TABS-BORDERS-PAGINATION-DECISION.md.
 *
 * What the product schema keeps EXPLICIT (the limits the adapter itself enforces, families/formatting.cjs:120-145/:218-226):
 *   tabs        align 'left' | 'right' only; no leader; the write REPLACES the whole list (W09-1e)
 *   border*     style 'single' only; four INDEPENDENT sides (W09-1e); colour '#RRGGBB' on write, BARE on read (asymmetry)
 *   pagination  `true` only -- clearing (false) is not measured and refused
 * plus two grid guards derived from the measured quantisation (twips for pos, eighth-points for size) and an ascending-tab
 * guard, each pre-dispatch (DOCX_CONSTRAINTS.format_text).
 *
 * Offline only -- the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs
 * (whose per-key merge already gives the measured REPLACE / independent-side semantics); the LIVE acceptance is
 * test/live/gh232-format-text-tabs-borders-pagination-product.live.cjs (not run by the author).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const FILE = '1303626';
// PLAIN mirrors test/fakes/fake-docx-driver.cjs's own formatOf() defaults for an untouched paragraph (W-09-3: absence is a DEFINITE state).
const PLAIN = { tabs: [], borderTop: null, borderBottom: null, borderLeft: null, borderRight: null, keepNext: false, keepLines: false, pageBreakBefore: false, widowControl: false };
const WANTED = {
  tabs: [{ pos: 36, align: 'left' }, { pos: 144, align: 'right' }],
  borderTop: { style: 'single', size: 1, color: '#FF0000' },
  borderBottom: { style: 'single', size: 0.5, color: '#0000FF' },
  borderLeft: { style: 'single', size: 2, color: '#00FF00' },
  borderRight: { style: 'single', size: 0.25, color: '#000000' },
  keepNext: true, keepLines: true, pageBreakBefore: true, widowControl: true,
};
// what the READ returns for WANTED: the border colour is BARE upper-case hex (the adapter strips the '#'; the driver read builds it from r/g/b)
const READ_BACK = { ...WANTED, borderTop: { ...WANTED.borderTop, color: 'FF0000' }, borderBottom: { ...WANTED.borderBottom, color: '0000FF' }, borderLeft: { ...WANTED.borderLeft, color: '00FF00' }, borderRight: { ...WANTED.borderRight, color: '000000' } };
const TARGET = 'GH232 TARGET PARAGRAPH';
const OTHER = 'GH232 OTHER PARAGRAPH';
const RANGE_HOST = 'GH232B range host NEEDLE text';
const doc = () => [
  { text: 'Heading', level: 1, props: { ...PLAIN } },
  { text: TARGET, props: PLAIN },
  { text: OTHER, props: PLAIN },
  { text: RANGE_HOST, props: PLAIN },
  { text: 'GH232 closing filler', props: PLAIN }, // selectorFrom's `after` needs a non-empty neighbor
];
const step = (selector, format = WANTED, extra = {}) => ({ intent: 'format_text', selector, format, ...extra });
const rangeStep = (selector, matchText, format = WANTED) => step(selector, format, { matchText });
const req = steps => ({ fileId: FILE, task: { steps } });
const verdict = format => D.validateDocxWriteRequest(req([step({ text: TARGET }, format)]));

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

test('SCHEMA: each of the nine new keys is individually accepted (tabs, the four border sides, the four pagination flags)', () => {
  for (const [k, v] of Object.entries(WANTED)) {
    const r = verdict({ [k]: v });
    assert.equal(r.ok, true, `${k} -> ${r.code} ${r.reason}`);
  }
  assert.equal(verdict(WANTED).ok, true, 'all nine together');
});

test('SCHEMA: tabs -- a decimal align, a leader, an extra field, a missing field, a negative pos are refused pre-dispatch; GH #290: [] (remove the paragraph\'s own tabs) and center are released', () => {
  const cases = [
    // GH #290: 'center' is released (assert below); 'decimal' stays refused (measured: silently stored as left)
    [[{ pos: 36, align: 'decimal' }], 'REFUSED_SHAPE'], [[{ pos: 36, align: 'Left' }], 'REFUSED_SHAPE'],
    [[{ pos: 36, align: 'left', leader: 'dot' }], 'UNKNOWN_FIELD'], [[{ pos: 36 }], 'REFUSED_SHAPE'], [[{ align: 'left' }], 'REFUSED_SHAPE'],
    [[{ pos: -1, align: 'left' }], 'REFUSED_SHAPE'], [[{ pos: 1001, align: 'left' }], 'REFUSED_SHAPE'], [[{ pos: '36', align: 'left' }], 'REFUSED_SHAPE'],
    ['x', 'REFUSED_SHAPE'], [{ pos: 36, align: 'left' }, 'REFUSED_SHAPE'],
  ];
  for (const [tabs, code] of cases) { const v = verdict({ tabs }); assert.equal(v.code, code, `${JSON.stringify(tabs)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
  assert.equal(verdict({ tabs: Array.from({ length: 33 }, (_, i) => ({ pos: i + 1, align: 'left' })) }).code, 'REFUSED_SHAPE', 'more than 32 stops');
  assert.equal(verdict({ tabs: Array.from({ length: 32 }, (_, i) => ({ pos: i + 1, align: 'left' })) }).ok, true, 'exactly 32 stops');
  assert.equal(verdict({ tabs: [{ pos: 36, align: 'center' }] }).ok, true, 'GH #290: a center tab');
  assert.equal(verdict({ tabs: [] }).ok, true, 'GH #290 F-02: [] removes the paragraph\'s own tabs');
});

test('SCHEMA: border sides -- a non-single style, a missing field, a bare / named colour, a zero or out-of-range size, an unknown side are refused pre-dispatch', () => {
  const b = (o = {}) => ({ style: 'single', size: 1, color: '#FF0000', ...o });
  const cases = [
    [{ borderTop: b({ style: 'double' }) }, 'REFUSED_SHAPE'], [{ borderTop: b({ style: 'dashed' }) }, 'REFUSED_SHAPE'],
    [{ borderTop: { size: 1, color: '#FF0000' } }, 'REFUSED_SHAPE'], [{ borderTop: { style: 'single', color: '#FF0000' } }, 'REFUSED_SHAPE'], [{ borderTop: { style: 'single', size: 1 } }, 'REFUSED_SHAPE'],
    [{ borderTop: b({ color: 'FF0000' }) }, 'REFUSED_SHAPE'], [{ borderTop: b({ color: 'red' }) }, 'REFUSED_SHAPE'], [{ borderTop: b({ color: '#FF00' }) }, 'REFUSED_SHAPE'],
    [{ borderTop: b({ size: 0 }) }, 'REFUSED_SHAPE'], [{ borderTop: b({ size: -1 }) }, 'REFUSED_SHAPE'], [{ borderTop: b({ size: 12.5 }) }, 'REFUSED_SHAPE'], [{ borderTop: b({ size: '1' }) }, 'REFUSED_SHAPE'],
    // GH #290: a side is now an anyOf (single | none) -- an extra field inside it is a REFUSED_SHAPE (still pre-dispatch), like format_table's edges
    [{ borderTop: b({ space: 0 }) }, 'REFUSED_SHAPE'], [{ borderTop: null }, 'REFUSED_SHAPE'], [{ borderTop: 'single' }, 'REFUSED_SHAPE'],
    [{ borderBetween: b() }, 'UNKNOWN_FIELD'], [{ borderBar: b() }, 'UNKNOWN_FIELD'], [{ border: b() }, 'UNKNOWN_FIELD'],
  ];
  for (const [format, code] of cases) { const v = verdict(format); assert.equal(v.code, code, `${JSON.stringify(format)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
  // boundary values are accepted: the smallest and the largest size, lower- and upper-case hex
  assert.equal(verdict({ borderTop: b({ size: 0.125 }) }).ok, true);
  assert.equal(verdict({ borderTop: b({ size: 12 }) }).ok, true);
  assert.equal(verdict({ borderTop: b({ color: '#ff00aa' }) }).ok, true);
  // GH #290: { style: 'none' } removes a side -- alone; with a size or a colour it is refused
  assert.equal(verdict({ borderTop: { style: 'none' } }).ok, true);
  for (const bad of [{ style: 'none', size: 1 }, { style: 'none', color: '#FF0000' }]) assert.notEqual(verdict({ borderTop: bad }).ok, true, JSON.stringify(bad));
});

test('SCHEMA: the pagination flags accept a real boolean -- 0, 1, "true", null are REFUSED_SHAPE; GH #290: false is released (widowControl too, measured 2026-10-23)', () => {
  for (const k of ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl']) assert.equal(verdict({ [k]: false }).ok, true, `${k}=false`);
  for (const k of ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl']) {
    for (const bad of [0, 1, 'true', null]) {
      const v = verdict({ [k]: bad });
      assert.equal(v.code, 'REFUSED_SHAPE', `${k}=${JSON.stringify(bad)} -> ${v.code}`);
      assert.equal(v.dispatched, false);
    }
    assert.equal(verdict({ [k]: true }).ok, true, k);
  }
});

test("SCHEMA: `style` stays UNKNOWN_FIELD (lineSpacing joined in GH #231b) -- #232 widened the schema by exactly the nine keys, not by 'the rest of the adapter'", () => {
  for (const k of ['style']) {
    const v = verdict({ [k]: 'x' });
    assert.equal(v.code, 'UNKNOWN_FIELD', k);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: the nine new keys compose with the run-level keys (#229/#230), the #231 paragraph keys and matchText in the SAME format object', () => {
  const combined = { ...WANTED, bold: true, color: '#2A7F3A', alignment: 'center', shading: 'FFFF00' };
  assert.equal(verdict(combined).ok, true);
  assert.equal(D.validateDocxWriteRequest(req([{ intent: 'format_text', selector: { text: TARGET }, format: WANTED, matchText: 'NEEDLE' }])).ok, true, 'a paragraph-level key is never refused for being combined with matchText');
});

// ---- the grid / order guards (DOCX_CONSTRAINTS.format_text) ---------------------------------------------------------------

test('GRID: a tab position must land on the 1/20-point (twip) grid within the adapter\'s 0.01pt tolerance; a border size on the 1/8-point grid -- else REFUSED_SHAPE pre-dispatch', () => {
  for (const pos of [36.123, 0.07, 36.08]) {
    const v = verdict({ tabs: [{ pos, align: 'left' }] });
    assert.equal(v.code, 'REFUSED_SHAPE', `pos ${pos}`);
    assert.match(v.reason, /\(twip\) grid/);
    assert.equal(v.dispatched, false);
  }
  for (const pos of [0, 36, 36.05, 36.5, 72.25, 0.05, 36.004, 36.009]) assert.equal(verdict({ tabs: [{ pos, align: 'left' }] }).ok, true, `pos ${pos}`);
  for (const size of [0.3, 0.226, 1.07, 0.2]) {
    const v = verdict({ borderTop: { style: 'single', size, color: '#FF0000' } });
    assert.equal(v.code, 'REFUSED_SHAPE', `size ${size}`);
    assert.match(v.reason, /1\/8-point grid/);
  }
  for (const size of [0.125, 0.25, 0.5, 1, 1.375, 6, 12, 1.001]) assert.equal(verdict({ borderTop: { style: 'single', size, color: '#FF0000' } }).ok, true, `size ${size}`);
});

test('ORDER: tab stops must be strictly ascending (sorted + unique) -- a descending or duplicated list is refused pre-dispatch (whether the editor re-orders the list is NOT measured)', () => {
  for (const tabs of [[{ pos: 72, align: 'left' }, { pos: 36, align: 'left' }], [{ pos: 36, align: 'left' }, { pos: 36, align: 'right' }], [{ pos: 10, align: 'left' }, { pos: 20, align: 'left' }, { pos: 15, align: 'right' }]]) {
    const v = verdict({ tabs });
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(tabs));
    assert.match(v.reason, /strictly greater/);
  }
  assert.equal(verdict({ tabs: [{ pos: 36, align: 'left' }, { pos: 72, align: 'right' }, { pos: 108, align: 'left' }] }).ok, true);
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('EXACT TARGETING + PRESERVATION: one unique paragraph moves from the measured PLAIN (absent = none/off) to WANTED; read_formatting symmetry proves it (border colour BARE on read); the rest of the document is untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([step(selectorFrom(doc(), TARGET))]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(p.writes() >= 1, true);
  const after = p.store.durable;
  const read = readFormattingOf(after, TARGET);
  for (const [k, v] of Object.entries(READ_BACK)) assert.deepEqual(read[k], v, `read_formatting.format.${k} proves the write`);
  const other = after.find(x => x.text === OTHER);
  for (const [k, v] of Object.entries(PLAIN)) assert.deepEqual(other.props[k], v, `untouched paragraph's ${k} preserved (still none/off)`);
});

test('READ ASYMMETRY (named, not hidden): a border colour written WITH the "#" and in lower case comes back BARE and upper-case; a tab pos / border size come back in the SAME points', async () => {
  const p = product();
  const r = await p.tool.handler(req([step(selectorFrom(doc(), TARGET), { borderLeft: { style: 'single', size: 1.375, color: '#ff00aa' }, tabs: [{ pos: 36.05, align: 'right' }] })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const read = readFormattingOf(p.store.durable, TARGET);
  assert.deepEqual(read.borderLeft, { style: 'single', size: 1.375, color: 'FF00AA' });
  assert.deepEqual(read.tabs, [{ pos: 36.05, align: 'right' }]);
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

test('REPLACE (W09-1e): a second tabs request REPLACES the whole list; the other keys of the first request are untouched', async () => {
  const p = product();
  const sel = selectorFrom(doc(), TARGET);
  assert.equal((await p.tool.handler(req([step(sel, { tabs: [{ pos: 36, align: 'left' }, { pos: 144, align: 'right' }], keepNext: true })]))).ok, true);
  const second = await p.tool.handler(req([step(sel, { tabs: [{ pos: 72, align: 'left' }] })]));
  assert.equal(second.ok, true, JSON.stringify(second));
  const read = readFormattingOf(p.store.durable, TARGET);
  assert.deepEqual(read.tabs, [{ pos: 72, align: 'left' }], 'REPLACED, not appended');
  assert.equal(read.keepNext, true, 'an unrelated key from the first request survives');
});

test('INDEPENDENT SIDES (W09-1e): a second border request on ONE side adds/changes that side only; the other sides keep their state', async () => {
  const p = product();
  const sel = selectorFrom(doc(), TARGET);
  assert.equal((await p.tool.handler(req([step(sel, { borderTop: WANTED.borderTop, borderLeft: WANTED.borderLeft })]))).ok, true);
  const second = await p.tool.handler(req([step(sel, { borderBottom: WANTED.borderBottom })]));
  assert.equal(second.ok, true, JSON.stringify(second));
  const read = readFormattingOf(p.store.durable, TARGET);
  assert.deepEqual(read.borderTop, READ_BACK.borderTop);
  assert.deepEqual(read.borderLeft, READ_BACK.borderLeft);
  assert.deepEqual(read.borderBottom, READ_BACK.borderBottom);
  assert.equal(read.borderRight, null, 'a side nobody wrote stays none');
});

test('RANGE (matchText) + the new keys: not refused for naming matchText -- they write the WHOLE PARENT paragraph, the text itself is untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([rangeStep(selectorFrom(doc(), RANGE_HOST), 'NEEDLE', { tabs: [{ pos: 36, align: 'left' }], borderTop: WANTED.borderTop, keepLines: true })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  const target = p.store.durable.find(x => x.text === RANGE_HOST);
  assert.deepEqual(target.format.tabs, [{ pos: 36, align: 'left' }]);
  assert.equal(target.format.keepLines, true);
  assert.equal(target.text, RANGE_HOST);
});

test('PRE-DISPATCH refusals never open a session: an off-grid tab position, a descending tab list and a decimal tab refuse before any driver opens', async () => {
  for (const format of [{ tabs: [{ pos: 36.123, align: 'left' }] }, { tabs: [{ pos: 72, align: 'left' }, { pos: 36, align: 'left' }] }, { tabs: [{ pos: 36, align: 'decimal' }] }]) {
    const p = product();
    const r = await p.tool.handler(req([step({ text: TARGET }, format)]));
    assert.equal(r.code, 'REFUSED_SHAPE', JSON.stringify(format));
    assert.equal(r.dispatched, false);
    assert.equal(p.opens(), 0);
  }
});
