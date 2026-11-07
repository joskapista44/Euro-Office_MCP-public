'use strict';

const { norm } = require('../inventory.cjs');
const { imageRows } = require('./image.cjs');
const { canonicalBlock } = require('./table-insert.cjs');
const { everyBeforeImagePreserved } = require('./image-insert.cjs');
const { resolveAnchor, shapeRows, everyBeforeShapePreserved } = require('./shape-insert.cjs');

/**
 * DOCX `chart-insert` family (GH #323, NEW_CAPABILITY): insert_chart -- a NEW inline NATIVE chart, carrying its own data (type, categories, named series of numbers, an optional title and a
 * bottom legend), in a NEW paragraph right after a named anchor paragraph. Not a picture of a chart (#306) and not a shape (#309).
 *
 * LIVE-measured (prod NC/DS, disposable copies; test/live/docx-gh323-native-chart-probe.live.cjs, -chart-readback-and-types-probe, -chart-tojson-series-probe, all PASS):
 *   - `Api.CreateChart(type, seriesValues, seriesNames, categories, wEmu, hEmu, 24)` + `paragraph.AddDrawing(chart)` lands a native chart that survives save + a fresh reopen: the stored
 *     word/charts/chartN.xml carries the chart element of the type, every series name, every category, every value, the title (`SetTitle(text, 13, true)`) and `c:legendPos val="b"`
 *     (`SetLegendPos('bottom')`); an embedded workbook is written alongside; doc.GetAllImages() does NOT list the chart.
 *   - type -> stored element (probe 2): bar = barChart col clustered | barStacked = col stacked | barStackedPercent = col percentStacked | horizontalBar = barChart bar clustered |
 *     horizontalBarStacked = bar stacked | area = areaChart standard | areaStacked = areaChart stacked | pie = pieChart | doughnut = doughnutChart. ONLY these are offered; 'line'
 *     returned NULL from CreateChart (refused by name, never tried), and the rest of the measured set (lineStacked, scatter, stock, 3D) is not offered in v1.
 *   - READ: the Word API's chart series objects have NO GetName / GetValues; the chart's own ToJSON plotArea series caches are the read-back (docx-charts-read.cjs -> document.charts rows
 *     { type, widthEmu, heightEmu, paragraph, previous, series: [{ name, values }], categories }) -- the L1 below is built on THAT, in the after-state's own fresh observation.
 *   NOT measured, and therefore NOT claimed by L1: the title text and the legend position round-trip through the product read (the LIVE product harness reads them from the stored XML);
 *   a chart in a table cell / header / footer; floating placement; number formats; colours; a style other than 24.
 *   `doc.AddElement(anchorIndex + 1, paragraph)` placing the chart paragraph is the SAME mechanism insert_shape / insert_image / insert_table already measured; the chart probes used Push --
 *   the LIVE product harness is the measurement for AddElement with a chart.
 *
 * ADDRESSING: `{ afterParagraph }` -- the SAME unique body-level paragraph / heading contract insert_shape uses (its own resolveAnchor, REUSED).
 *
 * RETRY IDENTITY: SATISFIED when the block right after the anchor is an EMPTY paragraph AND a chart row sits in an empty paragraph whose `previous` is the anchor, with the requested type,
 * categories, series (names AND values) and size (+-1 EMU). Stronger than insert_shape's (the chart's own data is part of the key), still a NAMED limit: two identical charts at the SAME anchor
 * cannot be told apart from one call that already landed. An UNREADABLE chart list FAILS CLOSED.
 *
 * Preservation placement: PRIMITIVE_INTEGRAL. Every OTHER chart preserved (multiset), exactly ONE new chart; every image and every shape unchanged; the body block sequence unchanged except the
 * ONE new empty paragraph immediately after the anchor (the anchor block alone gets the border-alt equivalence -- the SAME discipline insert_shape uses, canonicalBlock REUSED).
 */

const TYPES = Object.freeze(['bar', 'barStacked', 'barStackedPercent', 'horizontalBar', 'horizontalBarStacked', 'area', 'areaStacked', 'pie', 'doughnut']);
const SINGLE_SERIES_TYPES = Object.freeze(['pie', 'doughnut']);
const LEGEND_POSITIONS = Object.freeze(['bottom']); // the one measured position
const STYLE_INDEX = 24; // the one measured style
const TITLE_FONT_SIZE = 13; // the measured SetTitle arguments
const EMU_PER_MM = 36000;
const MIN_MM = 10;
const MAX_MM = 400;
const MAX_CATEGORIES = 100;
const MAX_SERIES = 10;
const MAX_LABEL = 100;
const MAX_TITLE = 200;
const CHART_KEYS = Object.freeze(['categories', 'heightMm', 'legend', 'series', 'title', 'type', 'widthMm']);
const REQUIRED_CHART_KEYS = Object.freeze(['categories', 'heightMm', 'series', 'type', 'widthMm']);
const PLAIN_TEXT_RE = /^[^\u0000-\u001f\u007f]+$/;

const isMm = v => Number.isInteger(v) && v >= MIN_MM && v <= MAX_MM;
const isLabel = v => typeof v === 'string' && !!norm(v) && v.length <= MAX_LABEL && PLAIN_TEXT_RE.test(v);

function planChart(c, index) {
  const bad = detail => ({ ok: false, detail: `step ${index}: insert_chart ${detail}` });
  if (!c || typeof c !== 'object' || Array.isArray(c)) return bad('requires a chart object');
  const keys = Object.keys(c);
  if (keys.some(k => !CHART_KEYS.includes(k)) || REQUIRED_CHART_KEYS.some(k => !keys.includes(k))) return bad(`chart must have ${REQUIRED_CHART_KEYS.join(', ')} and may have title, legend -- nothing else`);
  if (!TYPES.includes(c.type)) return bad(`chart.type must be one of ${TYPES.join(', ')} (only the measured types)`);
  if (!isMm(c.widthMm) || !isMm(c.heightMm)) return bad(`chart.widthMm / heightMm must be integers ${MIN_MM}..${MAX_MM}`);
  if (!Array.isArray(c.categories) || c.categories.length < 1 || c.categories.length > MAX_CATEGORIES || !c.categories.every(isLabel)) {
    return bad(`chart.categories must be 1..${MAX_CATEGORIES} non-empty single-line strings of at most ${MAX_LABEL} characters`);
  }
  const maxSeries = SINGLE_SERIES_TYPES.includes(c.type) ? 1 : MAX_SERIES;
  if (!Array.isArray(c.series) || c.series.length < 1 || c.series.length > maxSeries) return bad(`chart.series must hold 1..${maxSeries} series for type ${c.type}`);
  for (const s of c.series) {
    if (!s || typeof s !== 'object' || Array.isArray(s) || Object.keys(s).sort().join(',') !== 'name,values') return bad('each chart.series item must be exactly { name, values }');
    if (!isLabel(s.name)) return bad(`each series name must be a non-empty single-line string of at most ${MAX_LABEL} characters`);
    if (!Array.isArray(s.values) || s.values.length !== c.categories.length || !s.values.every(v => typeof v === 'number' && Number.isFinite(v))) {
      return bad('each series values must be finite numbers, one per category');
    }
  }
  if (new Set(c.series.map(s => s.name)).size !== c.series.length) return bad('series names must be unique');
  if (c.title !== undefined && !(typeof c.title === 'string' && !!norm(c.title) && c.title.length <= MAX_TITLE && PLAIN_TEXT_RE.test(c.title))) {
    return bad(`chart.title must be a non-empty single-line string of at most ${MAX_TITLE} characters`);
  }
  if (c.legend !== undefined && !LEGEND_POSITIONS.includes(c.legend)) return bad(`chart.legend must be ${LEGEND_POSITIONS.join(' | ')} (the measured position)`);
  return { ok: true };
}

function plan(step, index) {
  if (step.intent !== 'insert_chart') return { ok: false, detail: `step ${index}: ${step.intent} is not a chart-insert intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object' || Object.keys(sel).sort().join(',') !== 'afterParagraph') {
    return { ok: false, detail: `step ${index}: insert_chart requires a selector of exactly { afterParagraph }` };
  }
  if (typeof sel.afterParagraph !== 'string' || !norm(sel.afterParagraph)) {
    return { ok: false, detail: `step ${index}: insert_chart selector.afterParagraph must be a non-empty string` };
  }
  const p = planChart(step.chart, index);
  if (!p.ok) return p;
  const c = step.chart;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, selector: Object.freeze({ afterParagraph: sel.afterParagraph }),
      preservationMode: 'PRIMITIVE_INTEGRAL', chartType: c.type, widthEmu: c.widthMm * EMU_PER_MM, heightEmu: c.heightMm * EMU_PER_MM,
      categories: Object.freeze([...c.categories]), series: Object.freeze(c.series.map(s => Object.freeze({ name: s.name, values: Object.freeze([...s.values]) }))),
      ...(c.title !== undefined ? { title: c.title } : {}), ...(c.legend !== undefined ? { legend: c.legend } : {}),
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- the observation ------------------------------------------------------------------------------

const blocksOf = primary => (primary && primary.blocks) || [];
const fieldReadable = v => !(typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0);

/** document.charts as an array, or null when the driver could not read it (a string / absent). */
function chartRows(primary) {
  const d = primary && primary.document;
  const rows = d ? d.charts : undefined;
  return Array.isArray(rows) ? rows : null;
}

/** +-1 EMU: the measured height read-back was 2999999.88 for 3000000 -- this absorbs that rounding drift, never a wrong size. */
const sizeMatches = (row, op) => fieldReadable(row.widthEmu) && fieldReadable(row.heightEmu) && typeof row.widthEmu === 'number' && typeof row.heightEmu === 'number'
  && Math.abs(row.widthEmu - op.widthEmu) <= 1 && Math.abs(row.heightEmu - op.heightEmu) <= 1;

/** The chart's own data, exactly as requested: type, categories, series names and values (numbers compared by value). */
function dataMatches(row, op) {
  if (row.type !== op.chartType) return false;
  if (!Array.isArray(row.categories) || JSON.stringify(row.categories) !== JSON.stringify(op.categories)) return false;
  if (!Array.isArray(row.series) || row.series.length !== op.series.length) return false;
  return op.series.every((s, i) => {
    const r = row.series[i];
    return r && r.name === s.name && Array.isArray(r.values) && r.values.length === s.values.length && r.values.every((v, k) => v === s.values[k]);
  });
}

/** The chart this op would create: in an EMPTY paragraph whose preceding body element is the anchor, with the op's own data and size. */
const rowMatches = (row, anchorText, op) => fieldReadable(row.paragraph) && norm(row.paragraph) === '' && fieldReadable(row.previous) && row.previous !== null
  && norm(row.previous) === norm(anchorText) && sizeMatches(row, op) && dataMatches(row, op);

/** See the header's RETRY IDENTITY note. `rows` null (unreadable) is handled by the callers (fail closed), never treated as 'none'. */
function matchingChartAfter(primary, anchorIndex, op, anchorText) {
  const blocks = blocksOf(primary);
  const next = blocks[anchorIndex + 1];
  if (!next || next.kind !== 'paragraph' || norm(next.text) !== '') return false;
  const rows = chartRows(primary);
  return Array.isArray(rows) && rows.some(r => rowMatches(r, anchorText, op));
}

// ---- classify / resolve ----------------------------------------------------------------------------

function classify(primary, op) {
  const anchor = resolveAnchor(primary, op.selector.afterParagraph);
  if (!anchor.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `insert-chart-anchor-${String(anchor.ambiguity).toLowerCase().replace(/_/g, '-')}` } };
  }
  if (chartRows(primary) === null) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'insert-chart-charts-unobservable', detail: 'the chart list could not be read, so a retry could not tell whether the chart already exists' } };
  }
  const anchorIndex = blocksOf(primary).indexOf(anchor.target);
  if (matchingChartAfter(primary, anchorIndex, op, anchor.target.text)) {
    return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, chartType: op.chartType, widthEmu: op.widthEmu, heightEmu: op.heightEmu, proof: 'fresh-live-chart-with-own-data-in-empty-paragraph-immediately-after-anchor' } };
  }
  return { state: 'APPLY', evidence: { anchor: anchor.target.text } };
}

function resolve(primary, selector) {
  const sel = selector || {};
  const anchor = resolveAnchor(primary, sel.afterParagraph);
  if (!anchor.ok) return { ok: false, target: null, selector, ambiguity: anchor.ambiguity };
  return { ok: true, target: { kind: 'insert_chart_anchor', anchorText: norm(anchor.target.text) }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'insert_chart_anchor' && target.anchorText === norm(op.selector.afterParagraph)
    ? { ok: true }
    : { ok: false, dispatched: false, detail: 'insert-chart-anchor-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({
    family: 'chart-insert', action: 'insert', anchorText: target.anchorText, chartType: op.chartType, widthEmu: op.widthEmu, heightEmu: op.heightEmu, styleIndex: STYLE_INDEX,
    seriesValues: op.series.map(s => [...s.values]), seriesNames: op.series.map(s => s.name), categories: [...op.categories],
    title: op.title === undefined ? null : op.title, titleFontSize: TITLE_FONT_SIZE, legend: op.legend === undefined ? null : op.legend,
  });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand. Re-locates the anchor by the SAME text-matching walk every other DOCX write family uses, re-checked for uniqueness inside the sandbox. The chart, its title and legend are
 * built BEFORE anything touches the document: any guard miss or throw returns with the document unchanged (the insert is the LAST step). Then a fresh paragraph carries `AddDrawing(chart)` and
 * lands at `doc.AddElement(anchorIndex + 1, ...)`.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  const textOpts = { Numbering: false };
  const lookupTextOf = p => String(p.GetText(textOpts) || '').replace(/[\\r\\n]+$/, '');
  const n = doc.GetElementsCount();
  let anchorIndex = -1;
  for (let i = 0; i < n; i += 1) {
    const e = doc.GetElement(i);
    if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph' && lookupTextOf(e) === m.anchorText) {
      if (anchorIndex >= 0) return; // ambiguous at dispatch time: refuse rather than guess
      anchorIndex = i;
    }
  }
  if (anchorIndex < 0) return;
  if (typeof Api.CreateChart !== 'function' || typeof doc.AddElement !== 'function') return;
  const chart = Api.CreateChart(m.chartType, m.seriesValues, m.seriesNames, m.categories, m.widthEmu, m.heightEmu, m.styleIndex);
  if (!chart || typeof chart.GetClassType !== 'function' || chart.GetClassType() !== 'chart') return;
  if (m.title !== null) { if (typeof chart.SetTitle !== 'function') return; chart.SetTitle(m.title, m.titleFontSize, true); }
  if (m.legend !== null) { if (typeof chart.SetLegendPos !== 'function') return; chart.SetLegendPos(m.legend); }
  const np = Api.CreateParagraph();
  if (typeof np.AddDrawing !== 'function') return;
  np.AddDrawing(chart);
  doc.AddElement(anchorIndex + 1, np);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

function verifyPrimitive(op, before, after) {
  const anchorAfter = resolveAnchor(after && after.primary, op.selector.afterParagraph);
  if (!anchorAfter.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the after-state' } };
  if (chartRows(after.primary) === null) return { ok: false, evidence: { reason: 'the chart list is unobservable in the after-state' } };
  const anchorIndexAfter = blocksOf(after.primary).indexOf(anchorAfter.target);
  const ok = matchingChartAfter(after.primary, anchorIndexAfter, op, anchorAfter.target.text);
  return { ok, evidence: { anchor: anchorAfter.target.text, anchorIndexAfter, chartType: op.chartType } };
}

/** A chart's own key (type / size / own paragraph / data) -- NOT `previous`: inserting after the anchor moves the preceding-element text of an existing chart that followed the anchor. */
const chartKey = r => JSON.stringify([r.type, r.widthEmu, r.heightEmu, r.paragraph, r.series, r.categories]);

/** Every BEFORE chart still exists AFTER, by key (a multiset subtraction) -- insert never loses or mutates an existing chart. */
function everyBeforeChartPreserved(beforeRows, afterRows) {
  const pool = afterRows.map(chartKey);
  for (const r of beforeRows) {
    const i = pool.indexOf(chartKey(r));
    if (i < 0) return false;
    pool.splice(i, 1);
  }
  return true;
}

/**
 * PRIMITIVE_INTEGRAL: every OTHER chart preserved (multiset) AND exactly one new chart; every image and every shape unchanged (same count, same members); the body block sequence unchanged except
 * the ONE empty paragraph this write creates, immediately after the anchor. Only the ANCHOR block gets the border-alt equivalence (insert_shape's own discipline).
 */
function verifyPreservation(op, before, after) {
  const anchorBefore = resolveAnchor(before && before.primary, op.selector.afterParagraph);
  if (!anchorBefore.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const beforeCharts = chartRows(before.primary);
  const afterCharts = chartRows(after.primary);
  if (beforeCharts === null || afterCharts === null) return { ok: false, evidence: { reason: 'the chart list is unobservable (before or after)' } };
  const chartsPreserved = everyBeforeChartPreserved(beforeCharts, afterCharts);
  const countOk = afterCharts.length === beforeCharts.length + 1;
  const beforeImages = imageRows(before.primary) || [];
  const afterImages = imageRows(after.primary) || [];
  const imagesOk = beforeImages.length === afterImages.length && everyBeforeImagePreserved(beforeImages, afterImages);
  const beforeShapes = shapeRows(before.primary);
  const afterShapes = shapeRows(after.primary);
  const shapesOk = beforeShapes !== null && afterShapes !== null && beforeShapes.length === afterShapes.length && everyBeforeShapePreserved(beforeShapes, afterShapes);

  const beforeBlocks = blocksOf(before.primary);
  const afterBlocks = blocksOf(after.primary);
  const anchorIndex = beforeBlocks.indexOf(anchorBefore.target);
  const newBlock = afterBlocks[anchorIndex + 1];
  const newBlockOk = !!newBlock && newBlock.kind === 'paragraph' && norm(newBlock.text) === '';
  const prefixKey = blocks => blocks.map((b, i) => canonicalBlock(b, i === blocks.length - 1));
  const prefixSame = JSON.stringify(prefixKey(beforeBlocks.slice(0, anchorIndex + 1))) === JSON.stringify(prefixKey(afterBlocks.slice(0, anchorIndex + 1)));
  const restSame = JSON.stringify(beforeBlocks.slice(anchorIndex + 1).map(b => canonicalBlock(b))) === JSON.stringify(afterBlocks.slice(anchorIndex + 2).map(b => canonicalBlock(b)));
  const lengthOk = afterBlocks.length === beforeBlocks.length + 1;
  const ok = chartsPreserved && countOk && imagesOk && shapesOk && newBlockOk && prefixSame && restSame && lengthOk;
  return { ok, evidence: { chartsPreserved, countOk, imagesOk, shapesOk, newBlockOk, prefixSame, restSame, lengthOk } };
}

function wholeTaskChecks(op, finalObservation) {
  const anchor = resolveAnchor(finalObservation.primary, op.selector.afterParagraph);
  let ok = false;
  if (anchor.ok) {
    const idx = blocksOf(finalObservation.primary).indexOf(anchor.target);
    ok = matchingChartAfter(finalObservation.primary, idx, op, anchor.target.text);
  }
  return [{ name: `${op.id}: a ${op.chartType} chart (${op.series.length} series x ${op.categories.length} categories, ${op.widthEmu}x${op.heightEmu} EMU) sits immediately after "${op.selector.afterParagraph}"`, ok }];
}

// refs: none -- a chart is not a body text unit a later step can name
function producedText() {
  return null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  chartRows, rowMatches, dataMatches, matchingChartAfter, sizeMatches, everyBeforeChartPreserved, chartKey,
  TYPES, SINGLE_SERIES_TYPES, LEGEND_POSITIONS, STYLE_INDEX, EMU_PER_MM, MIN_MM, MAX_MM, MAX_CATEGORIES, MAX_SERIES, MAX_LABEL, MAX_TITLE,
};
