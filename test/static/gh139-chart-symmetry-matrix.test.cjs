'use strict';

/**
 * GH #139 U1 (internal tracker): a
 * STATIC, code-grounded check of docs/GH139-CHART-DRAWING-SYMMETRY-MATRIX.md's own SYMMETRIC /
 * WRITE_ONLY / MISSING claims -- repo-internal only (no network, no live editor, no git history).
 * Every row the matrix calls SYMMETRIC or WRITE_ONLY is checked HERE against the actual code:
 *   - the claimed WRITE intent is (or is not) in write-tools.cjs's XLSX_WRITE_ALLOWLIST;
 *   - the claimed READ operation exists in read-tools.cjs's OPERATIONS, ACCEPTED, with the named
 *     field mentioned in its own evidence text;
 *   - a WRITE_ONLY style target actually carries `semanticsNotProven` (via charts.cjs's own
 *     wholeTaskChecks), and a verifiable one (marker/gridlines/gapOverlap/seriesOrder) does NOT.
 * A stale row (the matrix claiming something the code no longer does) turns ONE of these checks
 * red -- this file never re-derives the matrix, only re-measures what it asserts.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { XLSX_WRITE_ALLOWLIST } = require('../../src/server/write-tools.cjs');
const { OPERATIONS, STATUS } = require('../../src/server/read-tools.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const drawingsInsert = require('../../src/adapters/xlsx/families/drawings-insert.cjs');

// ---- 1. WRITE-intent release claims (SYMMETRIC / WRITE_ONLY rows need the intent released;
// MISSING rows need it ABSENT -- a drift in either direction is the whole point of this gate) ----

// GH #139 U2 (same #141 roadmap GO as GH #138): move_chart released -- the matrix's C16 row moved
// from MISSING to SYMMETRIC (read_drawings's own pos fields on the chart kind verify the position).
const RELEASED_CHART_INTENTS = Object.freeze(['set_chart', 'rename_chart', 'set_chart_title', 'set_chart_legend', 'set_chart_axis_title', 'set_chart_style', 'delete_chart', 'copy_chart', 'move_chart']);
// GH #139 U3 (same #141 roadmap GO): insert_shape/insert_image released -- the matrix's D1/D2/D5
// rows moved from MISSING to SYMMETRIC (D3 stays SYMMETRIC with a NAMED ±1 EMU caveat, never an
// exact-match comparison; D7 stays READ_ONLY -- no text write exists, before or after this
// release). Nothing in the D-table is MISSING any more (the drawing family has no further
// unreleased intent).
// GH #139 D4/D6 (2026-10-17): fill/stroke colour moved from unconditional WRITE_ONLY to SYMMETRIC
// on an accepted /14+ runtime (drawing.fillLineColor.read/v1 closes the PUBLIC-getter gap) --
// conditionally WRITE_ONLY again only on a pre-/14 accepted runtime, same as before this
// capability existed. See test/acceptance/xlsx-x-o08-drawings-insert.test.cjs for both branches.
const RELEASED_DRAWING_INTENTS = Object.freeze(['insert_shape', 'insert_image']);

test('every chart intent the matrix calls released (SYMMETRIC/WRITE_ONLY rows) IS in XLSX_WRITE_ALLOWLIST', () => {
  for (const intent of RELEASED_CHART_INTENTS) assert.ok(XLSX_WRITE_ALLOWLIST.includes(intent), `${intent} should be released`);
});

test('every drawing intent the matrix calls released (SYMMETRIC/WRITE_ONLY D-rows) IS in XLSX_WRITE_ALLOWLIST -- the matrix no longer has a MISSING row', () => {
  for (const intent of RELEASED_DRAWING_INTENTS) assert.ok(XLSX_WRITE_ALLOWLIST.includes(intent), `${intent} should be released`);
});

test('both released drawing intents exist at the ADAPTER level too (insert_shape/insert_image in drawings-insert.cjs); move_chart is likewise adapter-backed (charts.cjs)', () => {
  assert.ok(charts.INTENTS.includes('move_chart'));
  assert.ok(drawingsInsert.INTENTS.includes('insert_shape'));
  assert.ok(drawingsInsert.INTENTS.includes('insert_image'));
});

// ---- 2. READ-side claims (read_charts / read_drawings, ACCEPTED, with the named field present in
// the operation's own evidence text -- the SAME text the matrix's "reopen-bizonyitek" column cites) ----

test('read_charts is ACCEPTED and its own evidence names every field the matrix marks SYMMETRIC/READ_ONLY for the chart rows', () => {
  const op = OPERATIONS.xlsx.read_charts;
  assert.equal(op.status, STATUS.ACCEPTED);
  for (const word of ['name/type/title/series/legend/axis-titles', 'dataLabels', 'markerStyle', 'majorGridlinesColor', 'series\\[\\].order', 'gapOverlap', 'markerSymbol', 'lineStyle']) {
    assert.match(op.evidence, new RegExp(word), `read_charts evidence should mention ${word}`);
  }
});

test('read_drawings is ACCEPTED and its own evidence names identity/position/stroke-width/text -- the drawing rows\' READ column', () => {
  const op = OPERATIONS.xlsx.read_drawings;
  assert.equal(op.status, STATUS.ACCEPTED);
  for (const word of ['parent/kind/name/pos/width/height', 'stroke width', 'text']) {
    assert.match(op.evidence, new RegExp(word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')), `read_drawings evidence should mention "${word}"`);
  }
});

// ---- 3. WRITE_ONLY vs SYMMETRIC style targets: `wholeTaskChecks`'s own `notProven` field, called
// with a real `plan()`-built operation and NO observation (verifyPrimitive's `gate()` tolerates a
// missing one and returns ok:false cleanly -- `notProven` does not depend on it either way, so this
// isolates exactly the claim the matrix makes, with nothing else in the result to go stale around it). ----

const styleOp = extra => {
  const step = { intent: 'set_chart_style', id: 's', sheet: 'S', name: 'Chart1', ...extra };
  const p = charts.plan(step, 0);
  assert.ok(p.ok, `plan() refused a well-formed set_chart_style/${extra.target} step: ${p.detail}`);
  return p.operation;
};
const notProvenOf = op => charts.wholeTaskChecks(op, null)[0].notProven || [];

test('the six WRITE-ONLY style targets (no public getter) each name their own written field in notProven', () => {
  const cases = [
    [{ target: 'plotArea', fill: '#FF0000' }, 'chart-style-plotarea-fill'],
    [{ target: 'title', outline: { width: 1, color: '#000000' } }, 'chart-style-title-outline'],
    [{ target: 'legend', fill: '#00FF00' }, 'chart-style-legend-fill'],
    [{ target: 'chart', fill: '#0000FF' }, 'chart-style-chart-fill'],
    [{ target: 'series', seriesIndex: 0, fill: '#AA00AA' }, 'chart-style-series-fill'],
    [{ target: 'dataPoint', seriesIndex: 0, dataPointIndex: 1, fill: '#123456' }, 'chart-style-datapoint-fill'],
  ];
  for (const [extra, expect] of cases) {
    const notProven = notProvenOf(styleOp(extra));
    assert.ok(notProven.includes(expect), `target=${extra.target}: expected ${expect} in ${JSON.stringify(notProven)}`);
  }
});

test('markerSymbol/lineStyle still name themselves in notProven on the WRITE side (the matrix calls C12/C13 READ+WRITE, not SYMMETRIC: the /16 READ exists, the write L1 does not consume it yet)', () => {
  assert.deepEqual(notProvenOf(styleOp({ target: 'markerSymbol', seriesIndex: 0, symbol: 'circle' })), ['chart-style-markersymbol-symbol']);
  assert.deepEqual(notProvenOf(styleOp({ target: 'lineStyle', seriesIndex: 0, dash: 'dot' })), ['chart-style-linestyle-dash']);
});

test('the four bridge-VERIFIABLE style targets (marker, gridlines, gapOverlap, seriesOrder) carry NO notProven -- the matrix calls these SYMMETRIC, not WRITE_ONLY, and this is the code fact that backs it', () => {
  const cases = [
    { target: 'marker', seriesIndex: 0, fill: '#FF0000' },
    { target: 'gridlines', axis: 'vertical', outline: { width: 1, color: '#000000' } },
    { target: 'gapOverlap', gapWidth: 150 },
    { target: 'seriesOrder', order: [1, 0] },
  ];
  for (const extra of cases) assert.deepEqual(notProvenOf(styleOp(extra)), [], `target=${extra.target} should have no notProven`);
});

test('copy_chart v1 unconditionally omits every style/marker/gridline field (never even attempted; data labels are copied since GH #302 G6) -- the matrix\'s "copy" row names this explicitly, not as a silent gap', () => {
  const p = charts.plan({ intent: 'copy_chart', id: 's', sheet: 'S', name: 'Chart2', sourceSheet: 'S', sourceName: 'Chart1', at: { col: 0, row: 10 } }, 0);
  assert.ok(p.ok, p.detail);
  const notProven = notProvenOf(p.operation);
  for (const field of ['chart-style-plotarea', 'chart-style-marker', 'chart-style-seriesorder']) assert.ok(notProven.includes(field), field);
  // GH #302 G6: data labels left the unconditional list (the copy carries them; what it cannot is named per copy)
  assert.equal(notProven.includes('chart-data-labels'), false);
});

// ---- 4. A deliberately broken row MUST go red (the card's own "MI A KESZ") ----

test('RED CONTROL: a non-existent intent is correctly reported as NOT released -- proves check #1 actually discriminates, not just passes by construction', () => {
  assert.ok(!XLSX_WRITE_ALLOWLIST.includes('set_chart_marquee_definitely_not_a_real_intent'));
});
