'use strict';

/**
 * GH #346 AUTO-REPAIR, megvalositas 2 -- chart.series.dataPoint.style.write/v1, the normalising WRITE
 * capability in the SAME /27 candidate as the READ fix (a maintainer's decision (a):
 * docs/GH346-DPT-AUTO-REPAIR-DESIGN-20261104.md's "nem nyitunk /28-at").
 *
 * Offline: the REAL candidate block is evaluated against a fake Api/AscFormat, reusing the same spPr/solid
 * fixture shapes ds935-bridge27-datapoint-effective.test.cjs (and ds935-bridge20-read.test.cjs before it)
 * already proved against the factory 9.3.5 sdk-all.js model paths for the FILL/OUTLINE DECODING half. This
 * file additionally fakes the ENCODING half (AscFormat.CreateSolidFillRGB / CSpPr / CLn / CDPt, s.removeDPt /
 * s.addDPt / s.getNumPts) -- these are NOT live-verified anywhere in this repo (see the module docstring's
 * own "megvalositas 2" section for the exact citation/uncertainty boundary); this proves the HANDLER LOGIC
 * against a fake that matches the design doc's and a maintainer's own cited mechanism (sdk-all.js :7:4151589 for the
 * gap-fill shape, :7:2220807 for CLn.prototype.createDuplicate), not that the real engine exposes these
 * methods under these exact names.
 *
 * v2 (a maintainer's review #ref) IMPLEMENTS the hits.length===0 + explicit-formatting gap-fill path (v1 refused
 * it NAMED, "dpt-gap-write-unimplemented"), adds the point-out-of-range boundary guard on both branches,
 * moves every constructor/method existence check to a single PREFLIGHT block before removeDPt (v1 could
 * leave a series partially normalized on a missing-constructor refusal), and preserves an existing line's
 * OTHER properties (e.g. prstDash) via CLn.prototype.createDuplicate instead of always building fresh.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V27 = fs.readFileSync(path.join(DIR, 'rev27-candidate', 'bridge-block.js'), 'utf8');
const WR = 'chart.series.dataPoint.style.write/v1';

// `missing`: an optional array of AscFormat key names to delete from the mock after building it -- lets a
// test simulate a single missing constructor/method without hand-rolling a whole separate mock.
function withAscFormat(fn, missing) {
  const prev = global.AscFormat;
  global.AscFormat = {
    CreateSolidFillRGB: (r, g, b) => solid(r, g, b),
    CSpPr: function () {
      this.Fill = null; this.ln = null;
      this.setFill = function (v) { this.Fill = v; };
      this.setLn = function (v) { this.ln = v; };
    },
    CLn: function () {
      this.Fill = null; this.w = null;
      this.setFill = function (v) { this.Fill = v; };
      this.setW = function (v) { this.w = v; };
      this.createDuplicate = function () {
        const dup = new global.AscFormat.CLn();
        dup.Fill = this.Fill; dup.w = this.w;
        return dup;
      };
    },
    CDPt: function () {
      this.idx = null; this.spPr = null;
      this.setIdx = function (v) { this.idx = v; };
      this.setSpPr = function (v) { this.spPr = v; };
    },
  };
  if (missing) for (const k of missing) delete global.AscFormat[k];
  try {
    return fn();
  } finally {
    if (prev === undefined) delete global.AscFormat; else global.AscFormat = prev;
  }
}

function bridgeOf(Api) {
  new Function('Api', 'Asc', V27.replace(/\}\s*,\s*$/, '}'))(Api, {});
  return Api.GetEuroOfficeBridge();
}

// ---- the SAME fixture shapes as ds935-bridge27-datapoint-effective.test.cjs / ds935-bridge20-read.test.cjs

const solid = (r, g, b) => ({ fill: { color: { color: { h: null, s: null, l: null, RGBA: { R: r, G: g, B: b, A: 255 } }, Mods: null } } });

// unlike the READ-side fixture (plain objects), a WRITE-side spPr/ln needs setFill/setLn/setW -- these are
// the methods the handler calls, named by direct symmetry with the ALREADY-established spPr.setLn pair
// (build-rev15.py/9.3.4.60's chart.series.lineStyle.write/v1). `lnExtra`: extra properties the real engine's
// CLn carries (e.g. prstDash) that a preservation test needs to survive a createDuplicate() round-trip.
function makeLn(fill, w, lnExtra) {
  const ln = Object.assign({ Fill: fill === undefined ? null : fill, w: w === undefined ? null : w }, lnExtra || {});
  ln.setFill = function (v) { this.Fill = v; };
  ln.setW = function (v) { this.w = v; };
  ln.createDuplicate = function () {
    const dup = {};
    for (const k of Object.keys(this)) { if (typeof this[k] !== 'function') dup[k] = this[k]; }
    dup.setFill = ln.setFill; dup.setW = ln.setW; dup.createDuplicate = ln.createDuplicate;
    return dup;
  };
  return ln;
}

function sppr(fill, lnFill, w, lnExtra) {
  const wantsLn = (lnFill !== undefined && lnFill !== null) || (w !== undefined && w !== null) || !!lnExtra;
  const o = {
    Fill: fill === undefined ? null : fill,
    ln: wantsLn ? makeLn(lnFill === undefined ? null : lnFill, w === undefined ? null : w, lnExtra) : null,
  };
  o.setFill = function (v) { this.Fill = v; };
  o.setLn = function (v) { this.ln = v; };
  return o;
}

// `numPts`: the series' own point COUNT (s.getNumPts().length) -- independent of the dPt array length, since
// most points carry no dPt entry at all. Defaults large enough that no existing test's idx trips the new
// point-out-of-range guard.
function makeSeries(dPtEntries, opts = {}) {
  const arr = dPtEntries.slice();
  const removedPositions = [];
  const addDPtCalls = [];
  const numPts = opts.numPts === undefined ? 1000 : opts.numPts;
  const series = {
    spPr: sppr(solid(100, 110, 120), null),
    dPt: arr,
    removeDPt(pos) { removedPositions.push(pos); arr.splice(pos, 1); },
    addDPt(p) { addDPtCalls.push(p); arr.push(p); },
    getNumPts() { return Array.from({ length: numPts }); },
  };
  if (opts.noRemoveDPt) delete series.removeDPt;
  if (opts.noAddDPt) delete series.addDPt;
  if (opts.noGetNumPts) delete series.getNumPts;
  series._removedPositions = removedPositions;
  series._addDPtCalls = addDPtCalls;
  return series;
}

function chartApi(series, name = 'Chart1', plots = 1) {
  const plotCharts = Array.from({ length: plots }, () => ({ series: [series] }));
  const model = { chart: { plotArea: { charts: plotCharts } } };
  return { GetSheet: n => (n === 'S1' ? { GetAllCharts: () => [{ GetName: () => name, Chart: model }] } : null) };
}

const hexOf = (r, g, b) => [r, g, b].map(n => n.toString(16).toUpperCase().padStart(2, '0')).join('');
const dptFillRgba = entry => entry.spPr.Fill.fill.color.color.RGBA;
const writeOn = (series, args, missing) => withAscFormat(() => bridgeOf(chartApi(series)).write(WR, Object.assign({ sheet: 'S1', chart: 'Chart1' }, args)), missing);
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

test('/27 capability list carries the new WRITE capability, kind WRITE, mutation TARGETED', () => {
  const b = withAscFormat(() => bridgeOf(chartApi(makeSeries([]))));
  const cap = b.manifest().capabilities.find(c => c.id === WR);
  assert.ok(cap, 'chart.series.dataPoint.style.write/v1 advertised');
  assert.equal(cap.kind, 'WRITE');
  assert.equal(cap.mutation, 'TARGETED');
  assert.equal(cap.version, 1);
});

test('S1 (no duplicate, count 1), no explicit formatting -> SATISFIED, 0 mutation, effective unchanged', () => {
  const series = makeSeries([{ idx: 0, spPr: sppr(solid(16, 17, 18), solid(19, 20, 21)) }]);
  const r = writeOn(series, { series: 0, point: 0 });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'SATISFIED');
  assert.equal(r.value.removed, 0);
  assert.deepEqual(r.value.before, r.value.after, 'nothing changed');
  assert.deepEqual(r.value.after.effective, { brush: { form: 'rgb', hex: hexOf(16, 17, 18) }, pen: { form: 'rgb', hex: hexOf(19, 20, 21) } });
  assert.equal(series.dPt.length, 1, 'no entry removed');
});

test('MUTATION CONTROL: a 0 entries + no explicit call must not call removeDPt/addDPt at all', () => {
  const series = makeSeries([], { noRemoveDPt: true });
  const r = writeOn(series, { series: 0, point: 3 });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'SATISFIED');
  assert.deepEqual(r.value.before, { count: 0, effective: null });
  assert.deepEqual(r.value.after, { count: 0, effective: null });
});

test('S2: identical duplicate -- normalized to count 1, the FIRST (document order) entry survives, no explicit', () => {
  const series = makeSeries([
    { idx: 1, spPr: sppr(solid(10, 20, 30), solid(11, 21, 31)) },
    { idx: 1, spPr: sppr(solid(10, 20, 30), solid(11, 21, 31)) },
  ]);
  const r = writeOn(series, { series: 0, point: 1 });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'NORMALIZED');
  assert.equal(r.value.removed, 1);
  assert.equal(series.dPt.length, 1, 'duplicate removed');
  assert.deepEqual(dptFillRgba(series.dPt[0]), { R: 10, G: 20, B: 30, A: 255 });
});

test('MUTATION CONTROL: removeDPt is called with the DISCARDED positions in DECREASING order, never the kept position', () => {
  const series = makeSeries([
    { idx: 2, spPr: sppr(solid(1, 1, 1), null) },
    { idx: 9, spPr: sppr(solid(9, 9, 9), null) }, // unrelated idx, interleaved -- positions must be absolute
    { idx: 2, spPr: sppr(solid(2, 2, 2), null) },
    { idx: 2, spPr: sppr(solid(3, 3, 3), null) },
  ]);
  const r = writeOn(series, { series: 0, point: 2 });
  assert.equal(r.ok, true);
  assert.equal(r.value.removed, 2);
  assert.deepEqual(series._removedPositions, [3, 2], 'decreasing absolute position order -- position 0 (the kept, first entry) never removed');
  assert.equal(series.dPt.length, 2, 'the idx-9 entry and the surviving idx-2 entry remain');
  assert.deepEqual(dptFillRgba(series.dPt[0]), { R: 1, G: 1, B: 1, A: 255 }, 'the FIRST idx-2 entry (document order, position 0) survives');
  assert.deepEqual(dptFillRgba(series.dPt[1]), { R: 9, G: 9, B: 9, A: 255 }, 'the unrelated idx-9 entry (position 1) is untouched');
});

test('contradictory duplicate, order A,B -- A (first) survives with NO explicit formatting', () => {
  const series = makeSeries([{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }]);
  const r = writeOn(series, { series: 0, point: 1 });
  assert.equal(r.value.outcome, 'NORMALIZED');
  assert.deepEqual(r.value.after.effective.brush, { form: 'rgb', hex: hexOf(1, 1, 1) }, 'A, the first, wins');
});

test('contradictory duplicate, order B,A -- B (now first) survives with NO explicit formatting (order flips the winner, proves first-match not a fixed value)', () => {
  const series = makeSeries([{ idx: 1, spPr: sppr(solid(2, 2, 2), null) }, { idx: 1, spPr: sppr(solid(1, 1, 1), null) }]);
  const r = writeOn(series, { series: 0, point: 1 });
  assert.equal(r.value.outcome, 'NORMALIZED');
  assert.deepEqual(r.value.after.effective.brush, { form: 'rgb', hex: hexOf(2, 2, 2) }, 'B, now the first, wins');
});

test('contradictory duplicate WITH explicit fill -- the NEW value replaces the first-entry inherited value, not just a no-op normalize', () => {
  const series = makeSeries([{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }]);
  const r = writeOn(series, { series: 0, point: 1, fill: { form: 'rgb', hex: 'ABCDEF' } });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'WRITTEN');
  assert.equal(r.value.removed, 1);
  assert.deepEqual(r.value.before.effective.brush, { form: 'rgb', hex: hexOf(1, 1, 1) }, 'before reflects the pre-write first entry');
  assert.deepEqual(r.value.after.effective.brush, { form: 'rgb', hex: 'ABCDEF' }, 'after is the CALLER value, not the inherited first-entry value');
  assert.equal(series.dPt.length, 1);
  assert.deepEqual(dptFillRgba(series.dPt[0]), { R: 0xAB, G: 0xCD, B: 0xEF, A: 255 });
});

test('complementary duplicate (fill-only A + line-only B) -- A survives, effective is UNCHANGED before/after (no cross-entry merge introduced by normalizing)', () => {
  const series = makeSeries([
    { idx: 1, spPr: sppr(solid(5, 5, 5), undefined) }, // first: fill only, no ln
    { idx: 1, spPr: sppr(undefined, solid(9, 9, 9)) }, // second: line only
  ]);
  const r = writeOn(series, { series: 0, point: 1 });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'NORMALIZED');
  assert.deepEqual(r.value.before.effective, { brush: { form: 'rgb', hex: hexOf(5, 5, 5) }, pen: null });
  assert.deepEqual(r.value.after.effective, r.value.before.effective, 'the effective value does not change -- B\'s line is NOT backfilled onto the survivor');
  assert.equal(series.dPt[0].spPr.ln, null, 'the surviving entry still has no line -- not merged from the discarded entry');
});

test('N=3 duplicate -- 1 survives, 2 removed in decreasing position order, generalises past the live-measured 2-case', () => {
  const series = makeSeries([
    { idx: 4, spPr: sppr(solid(1, 1, 1), null) },
    { idx: 4, spPr: sppr(solid(2, 2, 2), null) },
    { idx: 4, spPr: sppr(solid(3, 3, 3), null) },
  ]);
  const r = writeOn(series, { series: 0, point: 4 });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'NORMALIZED');
  assert.equal(r.value.removed, 2);
  assert.deepEqual(series._removedPositions, [2, 1]);
  assert.equal(series.dPt.length, 1);
  assert.deepEqual(dptFillRgba(series.dPt[0]), { R: 1, G: 1, B: 1, A: 255 });
});

test('preservation: a non-targeted point (different idx) and the series-level spPr are byte-identical before/after', () => {
  const untouchedEntry = { idx: 7, spPr: sppr(solid(77, 77, 77), null) };
  const series = makeSeries([{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }, untouchedEntry]);
  const seriesSpPrBefore = JSON.stringify(series.spPr);
  const untouchedBefore = JSON.stringify(untouchedEntry);
  writeOn(series, { series: 0, point: 1, fill: { form: 'rgb', hex: 'FFEEDD' } });
  assert.equal(JSON.stringify(untouchedEntry), untouchedBefore, 'idx 7 (not targeted) is untouched');
  assert.equal(JSON.stringify(series.spPr), seriesSpPrBefore, 'the series-level spPr is untouched');
  assert.equal(series.dPt.some(e => e.idx === 7), true, 'the untouched point is still present');
});

test('unresolvable: a discarded entry carrying a non-spPr property (marker) refuses NAMED, 0 mutation -- UNCONDITIONALLY, even with explicit formatting on the call', () => {
  const series = makeSeries([
    { idx: 1, spPr: sppr(solid(1, 1, 1), null) },
    { idx: 1, spPr: sppr(solid(2, 2, 2), null), marker: { symbol: 3 } },
  ]);
  const r1 = writeOn(series, { series: 0, point: 1 });
  assert.equal(reason(r1), 'xlsx-datapoint-duplicate-unresolvable:marker');
  assert.equal(series.dPt.length, 2, "no mutation -- a maintainer's decision (b): explicit formatting does not bypass this");
  const r2 = writeOn(series, { series: 0, point: 1, fill: { form: 'rgb', hex: 'ABCDEF' } });
  assert.equal(reason(r2), 'xlsx-datapoint-duplicate-unresolvable:marker', 'explicit fill on the call does NOT override the refusal');
  assert.equal(series.dPt.length, 2, 'still 0 mutation');
});

test('unresolvable: each of the six named dPt-only properties refuses, named by its own property', () => {
  const props = ['marker', 'dLbl', 'explosion', 'invertIfNegative', 'bubble3D', 'pictureOptions'];
  for (const p of props) {
    const extra = {};
    extra[p] = {};
    const series = makeSeries([{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, Object.assign({ idx: 1, spPr: sppr(solid(2, 2, 2), null) }, extra)]);
    const r = writeOn(series, { series: 0, point: 1 });
    assert.equal(reason(r), 'xlsx-datapoint-duplicate-unresolvable:' + p, p);
  }
});

test('idempotency: a second, identical call after an explicit write is SATISFIED with 0 mutation (no new History point, no removeDPt/addDPt call)', () => {
  const series = makeSeries([{ idx: 0, spPr: sppr(solid(1, 1, 1), null) }]);
  const first = writeOn(series, { series: 0, point: 0, fill: { form: 'rgb', hex: 'FF0000' } });
  assert.equal(first.value.outcome, 'WRITTEN');
  const beforeSecondCount = series.dPt.length;
  const second = writeOn(series, { series: 0, point: 0, fill: { form: 'rgb', hex: 'FF0000' } });
  assert.equal(second.ok, true);
  assert.equal(second.value.outcome, 'SATISFIED');
  assert.equal(second.value.removed, 0);
  assert.equal(series.dPt.length, beforeSecondCount, 'the stored entry count does not grow on retry');
});

test('idempotency: retry after a pure dedup (no explicit) is itself SATISFIED, not NORMALIZED again, with no further removeDPt call', () => {
  const series = makeSeries([{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }]);
  const first = writeOn(series, { series: 0, point: 1 });
  assert.equal(first.value.outcome, 'NORMALIZED');
  assert.equal(series._removedPositions.length, 1);
  const second = writeOn(series, { series: 0, point: 1 });
  assert.equal(second.value.outcome, 'SATISFIED');
  assert.equal(second.value.removed, 0);
  assert.equal(series._removedPositions.length, 1, 'no further removeDPt call on the second run');
});

// ---- v2 (a maintainer's review #ref, point 1): the gap-fill path IS implemented -----------------------------

test('GAP-FILL: hezagos lista (idx0+idx2), explicit write at idx1 appends a NEW entry via addDPt (idx-based, never the position-based public setter) -- idx0/idx2 stay byte-identical', () => {
  const idx0 = { idx: 0, spPr: sppr(solid(10, 10, 10), null) };
  const idx2 = { idx: 2, spPr: sppr(solid(20, 20, 20), null) };
  const series = makeSeries([idx0, idx2]);
  const idx0Before = JSON.stringify(idx0);
  const idx2Before = JSON.stringify(idx2);
  const r = writeOn(series, { series: 0, point: 1, fill: { form: 'rgb', hex: '112233' } });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'WRITTEN', 'a gap-fill is always WRITTEN -- there is nothing to have already been the same against');
  assert.equal(r.value.removed, 0);
  assert.deepEqual(r.value.before, { count: 0, effective: null });
  // a POSITIONAL overwrite (the engine's own public setter's bug, GH #346's root cause) would have hit
  // dPt[1] -- which, in document order, is the idx2 entry. Byte-identity here proves idx-based resolution.
  assert.equal(JSON.stringify(idx0), idx0Before, 'idx0 (position 0) untouched');
  assert.equal(JSON.stringify(idx2), idx2Before, 'idx2 (position 1) untouched -- a positional write would have hit THIS entry');
  assert.equal(series.dPt.length, 3, 'a new entry was APPENDED via addDPt, nothing overwritten in place');
  assert.equal(series._addDPtCalls.length, 1);
  const newEntry = series._addDPtCalls[0];
  assert.equal(newEntry.idx, 1, 'setIdx(args.point) was called on the new CDPt before addDPt');
  assert.deepEqual(dptFillRgba(newEntry), { R: 0x11, G: 0x22, B: 0x33, A: 255 });
  assert.notEqual(newEntry, idx0);
  assert.notEqual(newEntry, idx2);
});

test('GAP-FILL: a second, identical call after the gap-fill write is SATISFIED, not WRITTEN again -- idempotent retry on the newly-created entry', () => {
  const series = makeSeries([{ idx: 0, spPr: sppr(solid(1, 1, 1), null) }]);
  const first = writeOn(series, { series: 0, point: 4, fill: { form: 'rgb', hex: 'AABBCC' } });
  assert.equal(first.value.outcome, 'WRITTEN');
  assert.equal(series.dPt.length, 2);
  assert.equal(series._addDPtCalls.length, 1);
  const second = writeOn(series, { series: 0, point: 4, fill: { form: 'rgb', hex: 'AABBCC' } });
  assert.equal(second.ok, true);
  assert.equal(second.value.outcome, 'SATISFIED');
  assert.equal(second.value.removed, 0);
  assert.equal(series.dPt.length, 2, 'no further entry added');
  assert.equal(series._addDPtCalls.length, 1, 'addDPt not called again');
});

// ---- v2 (a maintainer's review, point 1's boundary half): point-out-of-range, on BOTH branches -------------------

test('point-out-of-range: point >= s.getNumPts().length refuses BEFORE collecting hits, on the GAP branch -- addDPt is never called', () => {
  const series = makeSeries([], { numPts: 3, noAddDPt: true });
  const r = writeOn(series, { series: 0, point: 10, fill: { form: 'rgb', hex: '445566' } });
  assert.equal(r.ok, false);
  assert.equal(r.reason, 'point-out-of-range');
  assert.equal(series.dPt.length, 0);
});

test('point-out-of-range: the SAME guard fires on the existing-entry branch too -- the check runs before hits are even collected', () => {
  const series = makeSeries([{ idx: 5, spPr: sppr(solid(9, 9, 9), null) }], { numPts: 2 });
  const r = writeOn(series, { series: 0, point: 5, fill: { form: 'rgb', hex: '112233' } });
  assert.equal(r.ok, false);
  assert.equal(r.reason, 'point-out-of-range');
  assert.equal(series.dPt.length, 1, 'no mutation');
});

// ---- v2 (a maintainer's review, point 2): ELOSZOR ELLENORIZ, AZTAN MUTAL -- no partial normalize on a missing ctor

test('ELOSZOR ELLENORIZ, AZTAN MUTAL: duplicate + explicit fill + missing CreateSolidFillRGB refuses NAMED BEFORE removeDPt is called -- no partially-normalized series', () => {
  const series = makeSeries([{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }]);
  const r = writeOn(series, { series: 0, point: 1, fill: { form: 'rgb', hex: 'ABCDEF' } }, ['CreateSolidFillRGB']);
  assert.equal(r.ok, false);
  assert.equal(r.reason, 'createsolidfillrgb-missing');
  assert.equal(series.dPt.length, 2, 'removeDPt was NOT called -- the duplicate is still there');
  assert.equal(series._removedPositions.length, 0);
});

// ---- v2 (a maintainer's review, point 3): outline rewrite preserves the existing lines OTHER properties ----------

test('outline rewrite PRESERVES the existing lines other properties (prstDash) via CLn.prototype.createDuplicate -- only Fill/w change', () => {
  const entry = { idx: 0, spPr: sppr(solid(1, 1, 1), solid(9, 9, 9), 12700, { prstDash: 'dash' }) };
  const originalLn = entry.spPr.ln;
  const series = makeSeries([entry]);
  const r = writeOn(series, { series: 0, point: 0, outline: { form: 'rgb', hex: 'FF00FF' } });
  assert.equal(r.ok, true);
  assert.equal(r.value.outcome, 'WRITTEN');
  const ln = series.dPt[0].spPr.ln;
  assert.equal(ln.prstDash, 'dash', 'prstDash preserved via createDuplicate -- a fresh CLn would have dropped it');
  assert.deepEqual(ln.Fill.fill.color.color.RGBA, { R: 0xFF, G: 0x00, B: 0xFF, A: 255 }, 'the explicitly given outline colour applies');
  assert.equal(ln.w, 12700, 'width preserved (not explicitly given, carried over from the duplicate)');
  assert.notEqual(ln, originalLn, 'the SURVIVING ln is a clone, the original ln object is never mutated in place');
  assert.deepEqual(originalLn.Fill.fill.color.color.RGBA, { R: 9, G: 9, B: 9, A: 255 }, 'the original ln object itself is untouched');
});

test('MUTATION CONTROL: when there is NO existing ln, a fresh CLn is built (createDuplicate has nothing to clone)', () => {
  const series = makeSeries([{ idx: 0, spPr: sppr(solid(1, 1, 1), null) }]);
  const r = writeOn(series, { series: 0, point: 0, outline: { form: 'rgb', hex: '00FF00' } }, ['CLn']);
  assert.equal(r.ok, false, 'no existing ln to duplicate from -- a fresh CLn is required and we removed it');
  assert.equal(r.reason, 'cln-missing');
});

// ---- regression / bad-args -----------------------------------------------------------------------------

test('bad-args / not-found reasons, unaffected regression', () => {
  const series = makeSeries([{ idx: 0, spPr: sppr(solid(1, 1, 1), null) }]);
  assert.equal(reason(writeOn(series, {})), 'bad-args');
  assert.equal(reason(writeOn(series, { series: 0 })), 'bad-args');
  const bNotFound = withAscFormat(() => bridgeOf(chartApi(series)).write(WR, { sheet: 'S1', chart: 'Nope', series: 0, point: 0 }));
  assert.equal(reason(bNotFound), 'chart-not-found');
  const bAmbig = withAscFormat(() => bridgeOf(chartApi(series, 'Chart1', 2)));
  assert.equal(reason(bAmbig.write(WR, { sheet: 'S1', chart: 'Chart1', series: 0, point: 0 })), 'chart-type-not-single', 'two plots in one chart is chart-type-not-single, not chart-name-ambiguous -- the ambiguity axis is the NAME match, not the plot count');
  const twoCharts = { GetSheet: n => (n === 'S1' ? { GetAllCharts: () => [{ GetName: () => 'Chart1', Chart: { chart: { plotArea: { charts: [{ series: [series] }] } } } }, { GetName: () => 'Chart1', Chart: { chart: { plotArea: { charts: [{ series: [series] }] } } } }] } : null) };
  const bDup = withAscFormat(() => bridgeOf(twoCharts));
  assert.equal(reason(bDup.write(WR, { sheet: 'S1', chart: 'Chart1', series: 0, point: 0 })), 'chart-name-ambiguous', 'strict exactly-one-match write convention (chart.drawing.delete.write/v1 /13 style), unlike the READ side\'s first-match');
  assert.equal(reason(writeOn(series, { series: 0, point: 0, fill: 'not-an-object' })), 'invalid-fill');
  assert.equal(reason(writeOn(series, { series: 0, point: 0, outlineWidthEmu: -1 })), 'invalid-outlineWidthEmu');
  const badShape = makeSeries([]); badShape.dPt = 5;
  assert.equal(reason(writeOn(badShape, { series: 0, point: 0 })), 'dPt-unexpected-shape');
});
