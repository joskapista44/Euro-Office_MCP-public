'use strict';

/**
 * GH #323 (offline): the `insert_chart` family (src/adapters/docx/families/chart-insert.cjs), the driver's page-side chart read (docx-charts-read.cjs) and the inventory's
 * `document.charts` rows. plan/classify/resolve/verify against fabricated `normalizeInventory` observations; the read and the callCommand body run against SCRIPTED ApiDocument stand-ins
 * -- no editor, no VM. What the real editor does is covered by the LIVE probes (docx-gh323-native-chart-probe / -chart-readback-and-types-probe / -chart-tojson-series-probe) and the LIVE
 * product harness (docx-gh323-insert-chart-product.live.cjs).
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ci = require('../../src/adapters/docx/families/chart-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { readChartsInto, CHARTS_READ_CALL } = require('../../src/drivers/eurooffice/docx-charts-read.cjs');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

const ANCHOR = 'GH323 anchor';
const SIBLING = 'GH323 sibling';
const OTHER = 'GH323 other';
const MM = 36000;
const W = 120 * MM;
const H = 70 * MM;
const CATS = ['1', '2', '3'];
const SERIES = [{ name: 'Toke', values: [1000, 920.5, 840] }, { name: 'Kamat', values: [50, 46, 42] }];
const CHART = { type: 'bar', categories: CATS, series: SERIES, widthMm: 120, heightMm: 70 };

// a driver-chart row (docx-charts-read.cjs row shape: series / categories as JSON strings) through the REAL normalizeInventory
const row = (previous, over = {}) => ({ index: 0, type: 'bar', cx: W, cy: H, paragraph: '', previous, seriesJson: JSON.stringify(SERIES), categoriesJson: JSON.stringify(CATS), ...over });
function obsOf(html, charts = [], extra = {}) {
  return { primary: normalizeInventory({ html, document: { charts, images: [], shapes: [], ...extra } }) };
}
const step = (over = {}) => ({ intent: 'insert_chart', selector: { afterParagraph: ANCHOR }, chart: { ...CHART }, ...over });
const op = (over = {}) => ci.plan(step(over), 0).operation;

// ---- plan -----------------------------------------------------------------------------------------

test('plan: accepts { afterParagraph } + a full chart; mm -> EMU (x36000); title / legend are optional and carried only when given', () => {
  const p = ci.plan(step(), 0);
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
  assert.equal(p.operation.chartType, 'bar');
  assert.equal(p.operation.widthEmu, W);
  assert.equal(p.operation.heightEmu, H);
  assert.deepEqual(p.operation.categories, CATS);
  assert.deepEqual(p.operation.series, SERIES);
  assert.equal('title' in p.operation, false);
  const full = ci.plan(step({ chart: { ...CHART, title: 'T', legend: 'bottom' } }), 0);
  assert.equal(full.ok, true, full.detail);
  assert.equal(full.operation.title, 'T');
  assert.equal(full.operation.legend, 'bottom');
  for (const t of ci.TYPES) assert.equal(ci.plan(step({ chart: { ...CHART, type: t, series: ci.SINGLE_SERIES_TYPES.includes(t) ? [SERIES[0]] : SERIES } }), 0).ok, true, t);
});

test('plan: refuses a wrong selector, an UNMEASURED type (line returned null LIVE), bad sizes, bad categories/series, a values/category length mismatch, duplicate names, >1 series on pie/doughnut, a bad title / legend', () => {
  const bad = (over, why) => assert.equal(ci.plan(step(over), 0).ok, false, why);
  const chart = over => ({ chart: { ...CHART, ...over } });
  bad({ selector: {} }, 'empty selector');
  bad({ selector: { afterParagraph: '  ' } }, 'blank anchor');
  bad({ chart: undefined }, 'no chart');
  bad({ chart: [CHART] }, 'chart is an array');
  for (const k of ['type', 'categories', 'series', 'widthMm', 'heightMm']) { const c = { ...CHART }; delete c[k]; bad({ chart: c }, `missing ${k}`); }
  bad(chart({ colour: 'red' }), 'extra chart key');
  for (const t of ['line', 'lineStacked', 'scatter', 'stock', 'bar3D', 'pie3D', 'BAR', '', undefined]) bad(chart({ type: t }), `type ${t}`);
  for (const v of [9, 401, 50.5, '120', undefined]) { bad(chart({ widthMm: v }), `widthMm ${v}`); bad(chart({ heightMm: v }), `heightMm ${v}`); }
  for (const c of [[], 'a', [''], ['a\nb'], [1, 2, 3], Array.from({ length: 101 }, (_, i) => String(i))]) bad(chart({ categories: c }), `categories ${JSON.stringify(c).slice(0, 20)}`);
  bad(chart({ series: [] }), 'no series');
  bad(chart({ series: Array.from({ length: 11 }, (_, i) => ({ name: `s${i}`, values: [1, 2, 3] })) }), '11 series');
  bad(chart({ series: [{ name: 'a', values: [1, 2] }] }), 'values shorter than categories');
  bad(chart({ series: [{ name: 'a', values: [1, 2, NaN] }] }), 'NaN value');
  bad(chart({ series: [{ name: 'a', values: [1, 2, '3'] }] }), 'string value');
  bad(chart({ series: [{ name: '', values: [1, 2, 3] }] }), 'blank name');
  bad(chart({ series: [{ name: 'a', values: [1, 2, 3], colour: 'x' }] }), 'extra series key');
  bad(chart({ series: [{ name: 'a', values: [1, 2, 3] }, { name: 'a', values: [4, 5, 6] }] }), 'duplicate series names');
  bad(chart({ type: 'pie', series: SERIES }), 'pie with two series');
  bad(chart({ type: 'doughnut', series: SERIES }), 'doughnut with two series');
  for (const t of ['', ' ', 'a\nb', 5, 'x'.repeat(201)]) bad(chart({ title: t }), `title ${JSON.stringify(t).slice(0, 12)}`);
  for (const l of ['top', 'none', 'b', 1]) bad(chart({ legend: l }), `legend ${l}`);
});

test('resolve / preDispatchCheck: round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [step()] });
  assert.equal(planned.ok, true, planned.detail);
  const o = planned.operations[0];
  const r = ci.resolve(obsOf(`<p>${ANCHOR}</p>`).primary, o.selector);
  assert.equal(r.ok, true);
  assert.equal(ci.preDispatchCheck(o, r.target).ok, true);
  assert.equal(ci.preDispatchCheck(o, { kind: 'insert_chart_anchor', anchorText: 'other' }).ok, false);
});

// ---- classify ----------------------------------------------------------------------------------------

test('classify: APPLY when nothing follows the anchor, the next block is not empty, or no matching chart row exists', () => {
  assert.equal(ci.classify(obsOf(`<p>${ANCHOR}</p>`).primary, op()).state, 'APPLY');
  assert.equal(ci.classify(obsOf(`<p>${ANCHOR}</p><p>${SIBLING}</p>`, [row(ANCHOR)]).primary, op()).state, 'APPLY', 'next block not empty');
  assert.equal(ci.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [row(SIBLING)]).primary, op()).state, 'APPLY', 'chart after another paragraph');
});

test('classify: SATISFIED only when the empty paragraph after the anchor holds a chart with the SAME type, categories, series names AND values, and size (+-1 EMU)', () => {
  const html = `<p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR)]).primary, op()).state, 'SATISFIED');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { cy: H - 0.12 })]).primary, op()).state, 'SATISFIED', 'the measured height rounding drift');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { cx: W - 2 })]).primary, op()).state, 'APPLY', '+-2 EMU is not');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { type: 'barStacked' })]).primary, op()).state, 'APPLY', 'other type');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { categoriesJson: JSON.stringify(['1', '2', 'X']) })]).primary, op()).state, 'APPLY', 'other category');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { seriesJson: JSON.stringify([SERIES[0], { name: 'Kamat', values: [50, 46, 43] }]) })]).primary, op()).state, 'APPLY', 'one value differs');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { seriesJson: JSON.stringify([SERIES[1], SERIES[0]]) })]).primary, op()).state, 'APPLY', 'series order differs');
  assert.equal(ci.classify(obsOf(html, [row(ANCHOR, { seriesJson: JSON.stringify([SERIES[0]]) })]).primary, op()).state, 'APPLY', 'a series is missing');
});

test('classify: FAIL_CLOSED on a missing / ambiguous anchor and on an UNOBSERVABLE chart list (a retry could duplicate -- never guessed as "none")', () => {
  assert.equal(ci.classify(obsOf(`<p>${OTHER}</p>`).primary, op()).state, 'FAIL_CLOSED');
  assert.equal(ci.classify(obsOf(`<p>${ANCHOR}</p><p>${ANCHOR}</p>`).primary, op()).state, 'FAIL_CLOSED');
  for (const unread of ['UNOBSERVABLE:no-GetAllCharts', undefined]) {
    const primary = normalizeInventory({ html: `<p>${ANCHOR}</p>`, document: { charts: unread } });
    const c = ci.classify(primary, op());
    assert.equal(c.state, 'FAIL_CLOSED');
    assert.equal(c.evidence.reason, 'insert-chart-charts-unobservable');
  }
});

// ---- inventory --------------------------------------------------------------------------------------

test('normalizeInventory: chart rows inflate series / categories; a failed read stays its named UNOBSERVABLE string; unparsable JSON is named; an unread list is a named string', () => {
  const p = normalizeInventory({ html: '<p>x</p>', document: { charts: [row(ANCHOR), row(null, { seriesJson: 'UNOBSERVABLE:boom', categoriesJson: 'UNOBSERVABLE:boom' }), row(null, { seriesJson: '{oops' })] } });
  const [a, b, c] = p.document.charts;
  assert.deepEqual(a.series, SERIES);
  assert.deepEqual(a.categories, CATS);
  assert.equal(a.widthEmu, W);
  assert.equal(b.series, 'UNOBSERVABLE:boom');
  assert.equal(c.series, 'UNOBSERVABLE:chart-json-unparsable');
  assert.equal(normalizeInventory({ html: '<p>x</p>', document: {} }).document.charts, 'UNOBSERVABLE:unread');
  assert.equal(ci.chartRows(normalizeInventory({ html: '<p>x</p>', document: { charts: 'UNOBSERVABLE:no-GetAllCharts' } })), null);
});

test('NO OTHER INTENT MOVES: the charts list never reaches a block -- blocks are byte-identical with and without charts', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  assert.deepEqual(normalizeInventory({ html, document: { charts: [row(ANCHOR)] } }).blocks, normalizeInventory({ html, document: { charts: [] } }).blocks);
});

// ---- verify -----------------------------------------------------------------------------------------

test('verifyPrimitive: ok only with the empty paragraph right after the anchor AND the matching chart row; an empty paragraph alone fails; an unread list fails', () => {
  const withNew = `<p>${ANCHOR}</p><p></p>`;
  assert.equal(ci.verifyPrimitive(op(), obsOf(`<p>${ANCHOR}</p>`), obsOf(withNew, [row(ANCHOR)])).ok, true);
  assert.equal(ci.verifyPrimitive(op(), obsOf(`<p>${ANCHOR}</p>`), obsOf(withNew, [])).ok, false);
  assert.equal(ci.verifyPrimitive(op(), obsOf(`<p>${ANCHOR}</p>`), { primary: normalizeInventory({ html: withNew, document: { charts: 'UNOBSERVABLE:x' } }) }).ok, false);
});

test('verifyPreservation: OK on exactly one new chart + exactly the one new empty paragraph; FAILS on a lost / mutated chart, no or two new charts, a changed image or shape, a shifted block', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
  const withNew = `<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  const existing = row(SIBLING, { index: 0, type: 'pie', seriesJson: JSON.stringify([SERIES[0]]) });
  const before = obsOf(html, [existing]);
  assert.equal(ci.verifyPreservation(op(), before, obsOf(withNew, [existing, row(ANCHOR)])).ok, true);
  assert.equal(ci.verifyPreservation(op(), before, obsOf(withNew, [row(ANCHOR)])).ok, false, 'existing chart lost');
  assert.equal(ci.verifyPreservation(op(), before, obsOf(withNew, [{ ...existing, cx: 1 }, row(ANCHOR)])).ok, false, 'existing chart mutated');
  assert.equal(ci.verifyPreservation(op(), before, obsOf(withNew, [existing])).ok, false, 'no new chart');
  assert.equal(ci.verifyPreservation(op(), before, obsOf(withNew, [existing, row(ANCHOR), row(ANCHOR)])).ok, false, 'two new charts');
  const img = { index: 0, part: 'body', drawingType: 'inline', cx: 1, cy: 1, paragraph: OTHER };
  assert.equal(ci.verifyPreservation(op(), obsOf(html, [existing], { images: [img] }), obsOf(withNew, [existing, row(ANCHOR)], { images: [] })).ok, false, 'image lost');
  const shape = { index: 0, text: 'S', cx: 1, cy: 1, paragraph: '', previous: OTHER };
  assert.equal(ci.verifyPreservation(op(), obsOf(html, [existing], { shapes: [shape] }), obsOf(withNew, [existing, row(ANCHOR)], { shapes: [] })).ok, false, 'shape lost');
  assert.equal(ci.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING} changed</p>`, [existing, row(ANCHOR)])).ok, false, 'a later block changed');
  assert.equal(ci.verifyPreservation(op(), before, obsOf(html, [existing, row(ANCHOR)])).ok, false, 'no new block');
});

test('wholeTaskChecks: one named check, true only on the matching final chart; producedText: null', () => {
  const [c] = ci.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR)]));
  assert.equal(c.ok, true);
  assert.match(c.name, /bar chart \(2 series x 3 categories/);
  assert.equal(ci.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p><p></p>`, []))[0].ok, false);
  assert.equal(ci.producedText(), null);
});

// ---- the callCommand body ---------------------------------------------------------------------------

function scripted(texts, over = {}) {
  const calls = [];
  const mkPara = text => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, AddDrawing(c) { calls.push('AddDrawing'); this.drawing = c; }, text });
  const elements = texts.map(mkPara);
  const doc = { GetElementsCount: () => elements.length, GetElement: i => elements[i], AddElement(i, p) { calls.push(`AddElement@${i}`); elements.splice(i, 0, p); } };
  const Api = {
    GetDocument: () => doc,
    CreateParagraph: () => mkPara(''),
    CreateChart: (type, values, names, cats, w, h, style) => {
      calls.push(`CreateChart(${type},${JSON.stringify(values)},${JSON.stringify(names)},${JSON.stringify(cats)},${w},${h},${style})`);
      return { GetClassType: () => 'chart', SetTitle: (t, sz, b) => calls.push(`title(${t},${sz},${b})`), SetLegendPos: p => calls.push(`legend(${p})`) };
    },
    ...over,
  };
  return { Api, elements, calls };
}
const run = (body, Api) => new Function('Api', body)(Api);
const lowered = o => ci.lower(o, { kind: 'insert_chart_anchor', anchorText: ANCHOR });

test('commandBody: builds the chart (+ title + legend) FIRST, then a fresh paragraph carries it and lands at anchorIndex + 1 -- the insert is the LAST step', () => {
  const s = scripted([OTHER, ANCHOR, SIBLING]);
  run(lowered(op({ chart: { ...CHART, title: 'T', legend: 'bottom' } })).body, s.Api);
  assert.deepEqual(s.calls, [`CreateChart(bar,[[1000,920.5,840],[50,46,42]],["Toke","Kamat"],["1","2","3"],${W},${H},24)`, 'title(T,13,true)', 'legend(bottom)', 'AddDrawing', 'AddElement@2']);
  assert.equal(s.elements.length, 4);
  assert.equal(s.elements[1].text, ANCHOR, 'the anchor itself is untouched');
  const plain = scripted([OTHER, ANCHOR, SIBLING]);
  run(lowered(op()).body, plain.Api);
  assert.ok(!plain.calls.some(c => c.startsWith('title') || c.startsWith('legend')), 'no title / legend call when none was asked');
});

test('commandBody: any guard miss leaves the document UNCHANGED -- missing / ambiguous anchor, no CreateChart, CreateChart returning null (the measured "line" result), a non-chart, a throw', () => {
  const body = lowered(op({ chart: { ...CHART, title: 'T' } })).body;
  const untouched = (s, why) => { run(body, s.Api); assert.equal(s.elements.length, 3, why); assert.ok(!s.calls.some(c => c.startsWith('AddElement')), why); };
  untouched(scripted([OTHER, 'not the anchor', SIBLING]), 'no anchor');
  untouched(scripted([ANCHOR, SIBLING, ANCHOR]), 'ambiguous anchor');
  untouched(scripted([OTHER, ANCHOR, SIBLING], { CreateChart: undefined }), 'no CreateChart');
  untouched(scripted([OTHER, ANCHOR, SIBLING], { CreateChart: () => null }), 'CreateChart returned null');
  untouched(scripted([OTHER, ANCHOR, SIBLING], { CreateChart: () => ({ GetClassType: () => 'shape' }) }), 'not a chart');
  untouched(scripted([OTHER, ANCHOR, SIBLING], { CreateChart: () => ({ GetClassType: () => 'chart' }) }), 'title asked but no SetTitle');
  untouched(scripted([OTHER, ANCHOR, SIBLING], { CreateChart: () => { throw new Error('boom'); } }), 'a throw');
});

// ---- the page-side read --------------------------------------------------------------------------

const pt = (v, idx) => ({ v, idx });
const chartJson = (series, cats) => JSON.stringify({ graphic: { chart: { plotArea: { charts: [{ ser: series.map(s => ({
  tx: { strRef: { strCache: { ptCount: 1, pt: [pt(s.name, 0)] } } },
  cat: { strRef: { strCache: { ptCount: cats.length, pt: cats.map(pt) } } },
  val: { numRef: { numCache: { ptCount: s.values.length, pt: s.values.map((v, i) => pt(String(v), i)) } } },
})) }] } } } });
const chartObj = (type, w, h, json) => ({ GetClassType: () => 'chart', GetChartType: () => type, GetWidth: () => w, GetHeight: () => h, ToJSON: () => json });
const para = (text, objs = []) => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetAllDrawingObjects: () => objs });
const docOf = elements => ({ GetAllCharts: () => [], GetElementsCount: () => elements.length, GetElement: i => elements[i] });
const readOf = doc => { const info = {}; readChartsInto(doc, info); return info.charts; };

test('readChartsInto: one primitive row per CHART (class "chart" only) -- type, EMU size, its own (empty) paragraph, the element before it, series / categories as JSON strings (values as numbers)', () => {
  const rows = readOf(docOf([para(OTHER), para(ANCHOR), para('', [{ GetClassType: () => 'image' }, chartObj('bar', W, H, chartJson(SERIES, CATS))])]));
  assert.equal(rows.length, 1);
  assert.deepEqual({ ...rows[0], seriesJson: undefined, categoriesJson: undefined }, { index: 0, type: 'bar', cx: W, cy: H, paragraph: '', previous: ANCHOR, seriesJson: undefined, categoriesJson: undefined });
  assert.deepEqual(JSON.parse(rows[0].seriesJson), SERIES);
  assert.deepEqual(JSON.parse(rows[0].categoriesJson), CATS);
  assert.ok(Object.values(rows[0]).every(v => v === null || ['string', 'number'].includes(typeof v)), 'every field is a primitive');
});

test('readChartsInto: points land by their own idx (a missing idx stays null, never shifted); series whose categories differ, or more than one plot-area chart, read as a NAMED gap', () => {
  const holes = JSON.stringify({ graphic: { chart: { plotArea: { charts: [{ ser: [{ tx: { strRef: { strCache: { ptCount: 1, pt: [pt('S', 0)] } } }, cat: { strRef: { strCache: { ptCount: 3, pt: [pt('c', 2), pt('a', 0)] } } }, val: { numRef: { numCache: { ptCount: 3, pt: [pt('7', 2), pt('5', 0)] } } } }] }] } } } });
  const [r] = readOf(docOf([para(ANCHOR), para('', [chartObj('bar', W, H, holes)])]));
  assert.deepEqual(JSON.parse(r.categoriesJson), ['a', null, 'c']);
  assert.deepEqual(JSON.parse(r.seriesJson), [{ name: 'S', values: [5, null, 7] }]);
  const mixed = JSON.parse(chartJson(SERIES, CATS));
  mixed.graphic.chart.plotArea.charts[0].ser[1].cat.strRef.strCache.pt[0].v = 'X';
  assert.match(readOf(docOf([para(''), para('', [chartObj('bar', W, H, JSON.stringify(mixed))])]))[0].seriesJson, /^UNOBSERVABLE:series-categories-differ/);
  const two = JSON.parse(chartJson(SERIES, CATS));
  two.graphic.chart.plotArea.charts.push(two.graphic.chart.plotArea.charts[0]);
  assert.match(readOf(docOf([para('', [chartObj('bar', W, H, JSON.stringify(two))])]))[0].seriesJson, /^UNOBSERVABLE:plot-area-charts-not-one/);
});

test('readChartsInto: a failed read is NAMED -- throwing ToJSON / size / type, missing GetAllCharts; `previous` null when the chart paragraph is first; no charts -> []', () => {
  const [r] = readOf(docOf([para('', [{ GetClassType: () => 'chart', GetChartType: () => { throw new Error('t'); }, GetWidth: () => { throw new Error('w'); }, GetHeight: () => 1, ToJSON: () => { throw new Error('j'); } }])]));
  assert.match(r.type, /^UNOBSERVABLE:/);
  assert.match(r.cx, /^UNOBSERVABLE:/);
  assert.match(r.seriesJson, /^UNOBSERVABLE:/);
  assert.equal(r.previous, null);
  assert.deepEqual(readOf(docOf([para(ANCHOR)])), []);
  const info = {}; readChartsInto({}, info); assert.equal(info.charts, 'UNOBSERVABLE:no-GetAllCharts');
});

test('the page body: CHARTS_READ_CALL is self-contained, holds no backtick / dollar-brace / backslash, and sits in the driver\'s readLive body; the body still compiles', () => {
  const info = {};
  new Function('doc', 'docInfo', CHARTS_READ_CALL)(docOf([para(ANCHOR), para('', [chartObj('pie', W, H, chartJson([SERIES[0]], CATS))])]), info);
  assert.equal(info.charts[0].type, 'pie');
  assert.ok(!/`|\$\{|\\/.test(readChartsInto.toString()), 'the injected source must hold no backtick / dollar-brace / backslash (it lives inside a template literal)');
  const body = propertyQuery({});
  assert.ok(body.includes('GetAllCharts') && body.includes('docInfo.charts'), 'the driver body carries the projection');
  assert.doesNotThrow(() => new Function(body), 'the whole readLive body must still be valid JavaScript');
});
