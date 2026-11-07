'use strict';

/**
 * GH #346 AUTO-REPAIR, megvalositas 1 -- the euro-office-bridge-v1 /27 READ candidate for DS 9.3.5
 * (documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev27-candidate/): chart.series.style.read/v1's
 * dataPoint branch no longer refuses datapoint-ambiguous on more than one stored <c:dPt> entry at the
 * requested idx -- it answers the renderer's own first-match-wins value (fill/outline from the FIRST
 * document-order entry, brush and pen read INDEPENDENTLY, mirroring sdk-all.js :7:4099122 / :7:4100364),
 * plus effective:{brush,pen} and, only when more than one entry is stored, duplicated:{count,order}.
 *
 * Offline: the REAL candidate block is evaluated against a fake Api, reusing the exact fixture shapes
 * (solid/scheme/sppr/chartApi) ds935-bridge20-read.test.cjs already proved against the factory 9.3.5
 * sdk-all.js model paths. This proves the HANDLER LOGIC, not that the model paths hold on a live session
 * (that is docs/GH346-DPT-DUPLICATE-PRECEDENCE-RESEARCH-20261104.md's LIVE, calibrated gate).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V26 = fs.readFileSync(path.join(DIR, 'rev26-candidate', 'bridge-block.js'), 'utf8');
const V27 = fs.readFileSync(path.join(DIR, 'rev27-candidate', 'bridge-block.js'), 'utf8');
const ST = 'chart.series.style.read/v1';
const LEAK = 'LEAK-SENTINEL';

function bridgeOf(Api, block = V27) {
  new Function('Api', 'Asc', block.replace(/\}\s*,\s*$/, '}'))(Api, {});
  return Api.GetEuroOfficeBridge();
}
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);
const noLeak = r => assert.ok(!JSON.stringify(r).includes(LEAK), 'no internal object leaks into the result');
const apiWith = sheets => ({ GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) });

// ---- the SAME chart.series.style.read/v1 fixture shapes as ds935-bridge20-read.test.cjs ---------------

const solid = (r, g, b) => ({ fill: { color: { color: { h: null, s: null, l: null, RGBA: { R: r, G: g, B: b, A: 255 }, _sentinel: LEAK }, Mods: null } } });
const scheme = id => ({ fill: { color: { color: { id, RGBA: { R: 0, G: 0, B: 0, A: 255 } }, Mods: null } } });
const sppr = (fill, ln, w) => ({ Fill: fill, ln: ln === undefined ? null : { Fill: ln, w: w === undefined ? null : w } });
function chartApi({ chartSpPr, plotSpPr, title, legend, series, plots = 1, name = 'Chart1' } = {}) {
  const ser = series || [{ spPr: sppr(solid(1, 2, 3), solid(4, 5, 6), 12700), dPt: [{ idx: 0, spPr: sppr(solid(16, 17, 18), solid(19, 20, 21)) }] }];
  const plotCharts = Array.from({ length: plots }, () => ({ series: ser }));
  const model = { spPr: chartSpPr === undefined ? sppr(solid(100, 110, 120), solid(101, 111, 121), 9525) : chartSpPr, chart: { plotArea: { spPr: plotSpPr === undefined ? sppr(solid(10, 20, 30), null) : plotSpPr, charts: plotCharts }, title: title === undefined ? { spPr: sppr(solid(40, 50, 60), solid(41, 51, 61)) } : title, legend: legend === undefined ? { spPr: sppr(null, null) } : legend }, _sentinel: LEAK };
  return apiWith({ S1: { GetAllCharts: () => [{ GetName: () => name, Chart: model }] } });
}
const readST = (o, args, block) => bridgeOf(chartApi(o), block).read(ST, Object.assign({ sheet: 'S1', chart: 'Chart1' }, args));
const hexOf = (r, g, b) => [r, g, b].map(n => n.toString(16).toUpperCase().padStart(2, '0')).join('');

test('/27 patchVersion and capability list unchanged apart from the version bump', () => {
  const b = bridgeOf(chartApi());
  assert.equal(b.manifest().patchVersion, 'euro-office-bridge-v1/9.3.5.0/27');
  const cap = b.manifest().capabilities.find(c => c.id === ST);
  assert.ok(cap, 'chart.series.style.read/v1 still advertised');
  assert.equal(cap.version, 1, 'same capability version -- the /27 unit extends the response shape in place');
});

test('POZ: count === 1 (no duplicate) -- effective mirrors the stored value, no duplicated field, byte-identical to /26 for every non-dataPoint target', () => {
  const r = readST({}, { target: 'dataPoint', series: 0, point: 0 });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value, { element: 'present', fill: { form: 'rgb', hex: hexOf(16, 17, 18) }, outline: { form: 'rgb', hex: hexOf(19, 20, 21) }, outlineWidthEmu: null, effective: { brush: { form: 'rgb', hex: hexOf(16, 17, 18) }, pen: { form: 'rgb', hex: hexOf(19, 20, 21) } } });
  assert.equal('duplicated' in r.value, false, 'count 1 carries no duplicated flag');
  for (const t of ['chart', 'plotArea', 'title', 'legend', 'series']) {
    const args = t === 'series' ? { target: t, series: 0 } : { target: t };
    assert.deepEqual(readST({}, args, V26).value, readST({}, args, V27).value, `${t}: /26 and /27 agree (only dataPoint changed)`);
    assert.equal('effective' in readST({}, args, V27).value, false, `${t}: effective is a dataPoint-only field`);
  }
  noLeak(r);
});

test('POZ: count === 1, absent / no-dPt -- unaffected regression (element absent, no effective)', () => {
  const r = readST({}, { target: 'dataPoint', series: 0, point: 9 });
  assert.deepEqual(r.value, { element: 'absent', fill: null, outline: null, outlineWidthEmu: null });
  assert.equal('effective' in r.value, false);
});

test('count === 2, identical duplicate -- effective = the FIRST (document-order) entry, duplicated.count 2, order document-order', () => {
  const dup = [{ spPr: null, dPt: [
    { idx: 1, spPr: sppr(solid(10, 20, 30), solid(11, 21, 31)) },
    { idx: 1, spPr: sppr(solid(90, 90, 90), solid(91, 91, 91)) },
  ] }];
  const r = readST({ series: dup }, { target: 'dataPoint', series: 0, point: 1 });
  assert.equal(r.ok, true, 'no longer datapoint-ambiguous');
  assert.deepEqual(r.value.fill, { form: 'rgb', hex: hexOf(10, 20, 30) }, 'fill = the FIRST entry');
  assert.deepEqual(r.value.outline, { form: 'rgb', hex: hexOf(11, 21, 31) }, 'outline = the FIRST entry');
  assert.deepEqual(r.value.effective, { brush: { form: 'rgb', hex: hexOf(10, 20, 30) }, pen: { form: 'rgb', hex: hexOf(11, 21, 31) } });
  assert.deepEqual(r.value.duplicated, {
    count: 2,
    order: [
      { fill: { form: 'rgb', hex: hexOf(10, 20, 30) }, line: { form: 'rgb', hex: hexOf(11, 21, 31) } },
      { fill: { form: 'rgb', hex: hexOf(90, 90, 90) }, line: { form: 'rgb', hex: hexOf(91, 91, 91) } },
    ],
  });
  noLeak(r);
});

test('MUTATION CONTROL: swapping document order flips which entry is "effective" -- a second-wins implementation would fail this', () => {
  const abFirst = [{ spPr: null, dPt: [{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }] }];
  const baFirst = [{ spPr: null, dPt: [{ idx: 1, spPr: sppr(solid(2, 2, 2), null) }, { idx: 1, spPr: sppr(solid(1, 1, 1), null) }] }];
  const rAB = readST({ series: abFirst }, { target: 'dataPoint', series: 0, point: 1 });
  const rBA = readST({ series: baFirst }, { target: 'dataPoint', series: 0, point: 1 });
  assert.deepEqual(rAB.value.effective.brush, { form: 'rgb', hex: hexOf(1, 1, 1) }, 'A,B order -> A (the first) wins');
  assert.deepEqual(rBA.value.effective.brush, { form: 'rgb', hex: hexOf(2, 2, 2) }, 'B,A order -> B (now the first) wins');
  assert.notDeepEqual(rAB.value.effective.brush, rBA.value.effective.brush, 'the two orders disagree -- proves first-match, not a fixed value');
});

test('brush and pen are read INDEPENDENTLY off the FIRST entry -- the first entry\'s own missing line is never backfilled from the second entry', () => {
  const dup = [{ spPr: null, dPt: [
    { idx: 1, spPr: sppr(solid(5, 5, 5), undefined) }, // first: fill only, no ln at all
    { idx: 1, spPr: sppr(null, solid(9, 9, 9)) }, // second: line only
  ] }];
  const r = readST({ series: dup }, { target: 'dataPoint', series: 0, point: 1 });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.effective.brush, { form: 'rgb', hex: hexOf(5, 5, 5) }, 'brush from the first entry');
  assert.equal(r.value.effective.pen, null, 'pen is null -- NOT the second entry\'s line colour (no cross-entry fallback)');
  assert.deepEqual(r.value.duplicated.order[0], { fill: { form: 'rgb', hex: hexOf(5, 5, 5) }, line: null });
  assert.deepEqual(r.value.duplicated.order[1], { fill: null, line: { form: 'rgb', hex: hexOf(9, 9, 9) } });
});

test('RED (the /26 behaviour): the SAME duplicate shape still refuses datapoint-ambiguous on /26 -- the fix is isolated to /27', () => {
  const dup = [{ spPr: null, dPt: [{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }] }];
  assert.equal(reason(readST({ series: dup }, { target: 'dataPoint', series: 0, point: 1 }, V26)), 'datapoint-ambiguous');
  assert.equal(reason(readST({ series: dup }, { target: 'dataPoint', series: 0, point: 1 }, V27)) !== 'datapoint-ambiguous', true);
});

test('datapoint-ambiguous cannot be produced by /27 and is dropped from its own reasons list', () => {
  assert.equal(V27.includes('datapoint-ambiguous'), false, 'the string no longer appears anywhere in the /27 block');
  const dupA = [{ spPr: null, dPt: [{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(solid(2, 2, 2), null) }, { idx: 1, spPr: sppr(solid(3, 3, 3), null) }] }];
  const r3 = readST({ series: dupA }, { target: 'dataPoint', series: 0, point: 1 });
  assert.equal(r3.ok, true, 'three stored entries (N>2, not live-measured) still resolve via the same first-match rule, not a refusal');
  assert.equal(r3.value.duplicated.count, 3);
});

test('unsupported-target / bad-args / not-found reasons unaffected (regression)', () => {
  assert.equal(reason(readST({}, { target: 'axis' })), 'unsupported-target');
  assert.equal(reason(readST({}, {})), 'bad-args');
  assert.equal(reason(readST({ series: [{ spPr: null, dPt: 5 }] }, { target: 'dataPoint', series: 0, point: 1 })), 'dPt-unexpected-shape');
  assert.equal(reason(bridgeOf(chartApi()).read(ST, { sheet: 'S1', chart: 'Nope', target: 'chart' })), 'chart-not-found');
});
