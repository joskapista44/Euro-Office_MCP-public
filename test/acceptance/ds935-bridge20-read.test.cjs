'use strict';

/**
 * GH #344 / #346 / #354 / #355 -- the euro-office-bridge-v1 /20 READ candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev20-candidate/): THREE new READ capabilities -- sheet.printSettings.read/v2, sheet.autoFilter.colorFilter.read/v1, chart.series.style.read/v1.
 * Offline: the REAL candidate block is evaluated against a fake Api in the SHAPES read from the factory 9.3.5 sdk-all.js (asc_CPageMargins,
 * asc_CPageSetup.pageOrder, ColorFilter{CellColor,dxf}, CSpPr{Fill,ln}, CDPt{idx,spPr}). This proves the HANDLER LOGIC -- NOT that those model paths
 * hold on the real engine in a live session; that is the design's R1-R2 gate on the disposable DS (NOT MEASURED).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V19 = fs.readFileSync(path.join(DIR, 'rev19-candidate', 'bridge-block.js'), 'utf8');
const V20 = fs.readFileSync(path.join(DIR, 'rev20-candidate', 'bridge-block.js'), 'utf8');
const HANDLER = fs.readFileSync(path.join(DIR, 'rev20-candidate', 'handler-read-v20.txt'), 'utf8');
const PS2 = 'sheet.printSettings.read/v2';
const PS1 = 'sheet.printSettings.read/v1';
const CF = 'sheet.autoFilter.colorFilter.read/v1';
const ST = 'chart.series.style.read/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const LEAK = 'LEAK-SENTINEL';

function bridgeOf(Api, block = V20) {
  new Function('Api', 'Asc', block.replace(/\}\s*,\s*$/, '}'))(Api, {});
  return Api.GetEuroOfficeBridge();
}
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);
const noLeak = r => assert.ok(!JSON.stringify(r).includes(LEAK), 'no internal object leaks into the result');

// ---- the fake worksheet / chart models ---------------------------------------------------------------

/** a worksheet model in the shape sheet.printSettings.read/v1 needs, plus the /20 fields */
function printModel({ header = 0.3, footer = 0.4, order = 1, noGetHeader = false, noGetFooter = false, noMargins = false } = {}) {
  const hf = {};
  for (const k of ['OddHeader', 'OddFooter', 'FirstHeader', 'FirstFooter', 'EvenHeader', 'EvenFooter']) hf[`get${k}`] = () => ({ str: `t-${k}` });
  hf.getDifferentFirst = () => true; hf.getDifferentOddEven = () => 0;
  const pm = { asc_getLeft: () => 0.7, _sentinel: LEAK };
  if (!noGetHeader) pm.asc_getHeader = () => header;
  if (!noGetFooter) pm.asc_getFooter = () => footer;
  const ps = { asc_getFitToWidth: () => 1, asc_getFitToHeight: () => 0, asc_getScale: () => 100, asc_getWidth: () => 210, asc_getHeight: () => 297, useFirstPageNumber: false, firstPageNumber: 1, pageOrder: order, _sentinel: LEAK };
  const ppo = { asc_getPageSetup: () => ps, asc_getHorizontalCentered: () => true, asc_getVerticalCentered: () => false };
  if (!noMargins) ppo.asc_getPageMargins = () => pm;
  return { PagePrintOptions: ppo, headerFooter: hf, rowBreaks: null, sheetPr: { FitToPage: true } };
}
const apiWith = sheets => ({ GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) });

// ---- the build ---------------------------------------------------------------------------------------

test('the /20 block = the /19 block + exactly the builder\'s substitutions, and the builder is deterministic', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);print(json.dumps({"entries":m.ENTRIES,"impl":m.handler_text(),"sha":m.V19_BLOCK_SHA}))',
    path.join(DIR, 'rev20-candidate', 'build-rev20.py')], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V19), b.sha, 'the builder pins the /19 candidate block');
  for (const frag of [b.entries, b.impl]) assert.equal(V20.split(frag).length - 1, 1, `fragment occurs once: ${frag.slice(0, 40)}`);
  const back = V20.split(b.entries).join('').split(b.impl).join('').replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/20"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/19"');
  assert.equal(back, V19, 'nothing else moved: every /19 capability (manifest entry and handler) is byte-for-byte unchanged');
  assert.equal(fs.readFileSync(path.join(DIR, 'rev20-candidate', 'old.txt'), 'utf8'), V19);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev20-candidate', 'new.txt'), 'utf8'), V20);
  // the builder also (re)writes manifest-mechanism-proof.json next to bridge-block.js; snapshot and restore it so
  // this re-run leaves the tree exactly as committed (the published copy is redacted by the derivation -- a rebuild from the
  // published, redaction-touched build-rev20.py would otherwise write back a byte-different proof file).
  const proofPath = path.join(DIR, 'rev20-candidate', 'manifest-mechanism-proof.json');
  const proofBefore = fs.readFileSync(proofPath);
  const re = spawnSync('python3', ['-I', path.join(DIR, 'rev20-candidate', 'build-rev20.py')], { encoding: 'utf8', cwd: '/' });
  assert.equal(re.status, 0, re.stderr);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev20-candidate', 'bridge-block.js'), 'utf8'), V20, 'a re-run reproduces the block exactly');
  assert.equal(JSON.parse(re.stdout).v20_block_sha, sha(V20));
  fs.writeFileSync(proofPath, proofBefore);
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V19, V20]) assert.doesNotThrow(() => new Function('Api', 'Asc', t.replace(/\}\s*,\s*$/, '}')));
});

test('the handler text enumerates no object and carries no line comment', () => {
  for (const w of ['Object.keys', 'Object.values', 'Object.entries', 'getOwnPropertyNames', 'for (var k in', 'for (var p in', ' in args']) assert.ok(!HANDLER.includes(w), w);
  assert.ok(!HANDLER.includes('//'));
});

test('the manifest advertises exactly the three new READ capabilities, once each, mutation NO, and the patch version is /20', () => {
  const m = bridgeOf(apiWith({})).manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/20');
  const old = bridgeOf(apiWith({}), V19).manifest();
  const oldIds = old.capabilities.map(c => c.id);
  const ids = m.capabilities.map(c => c.id);
  assert.deepEqual(ids.slice(0, oldIds.length), oldIds, 'the /19 capabilities keep their order and identity');
  assert.deepEqual(ids.slice(oldIds.length), [PS2, CF, ST]);
  for (const c of m.capabilities.slice(oldIds.length)) { assert.equal(c.kind, 'READ'); assert.equal(c.mutation, 'NO'); assert.equal(c.format, 'XLSX'); }
  assert.equal(m.capabilities.find(c => c.id === PS2).version, 2);
  assert.equal(new Set(ids).size, ids.length);
  assert.ok(ids.includes(PS1), 'v1 stays served');
});

// ---- sheet.printSettings.read/v2 --------------------------------------------------------------------

test('printSettings v2: the v1 value UNCHANGED + headerMargin, footerMargin, pageOrder (POZ), and v1 itself is identical in /19 and /20', () => {
  const sheets = { S1: { worksheet: printModel() } };
  const v1 = bridgeOf(apiWith(sheets), V19).read(PS1, { sheet: 'S1' });
  const v1b = bridgeOf(apiWith(sheets)).read(PS1, { sheet: 'S1' });
  assert.deepEqual(v1b, v1, 'the v1 handler is byte-identical: same value in /19 and /20');
  const v2 = bridgeOf(apiWith(sheets)).read(PS2, { sheet: 'S1' });
  assert.equal(v2.ok, true);
  const { headerMargin, footerMargin, pageOrder, ...rest } = v2.value;
  assert.deepEqual(rest, v1.value, 'every v1 field unchanged');
  assert.deepEqual([headerMargin, footerMargin, pageOrder], [0.3, 0.4, 'overThenDown']);
  noLeak(v2);
  const v2d = bridgeOf(apiWith({ S1: { worksheet: printModel({ order: 0, header: 0, footer: 1.25 }) } })).read(PS2, { sheet: 'S1' });
  assert.deepEqual([v2d.value.headerMargin, v2d.value.footerMargin, v2d.value.pageOrder], [0, 1.25, 'downThenOver']);
});

test('printSettings v2: null margins / null pageOrder are the model\'s own unset state; anything else refuses with a NAMED reason (NEG)', () => {
  const run = (opts, args = { sheet: 'S1' }) => bridgeOf(apiWith({ S1: { worksheet: printModel(opts) } })).read(PS2, args);
  const unset = run({ header: null, footer: null, order: null }); assert.equal(unset.ok, true);
  assert.deepEqual([unset.value.headerMargin, unset.value.footerMargin, unset.value.pageOrder], [null, null, null]);
  assert.equal(reason(run({ order: 2 })), 'pageOrder-unexpected');
  assert.equal(reason(run({ order: 'downThenOver' })), 'pageOrder-unexpected');
  assert.equal(reason(run({ header: '0.3' })), 'headerMargin-not-number');
  assert.equal(reason(run({ footer: NaN })), 'footerMargin-not-number');
  assert.equal(reason(run({ header: -1 })), 'headerMargin-not-number');
  assert.equal(reason(run({ noGetHeader: true })), 'margin-getters-missing');
  assert.equal(reason(run({ noGetFooter: true })), 'margin-getters-missing');
  assert.equal(reason(run({ noMargins: true })), 'model-missing');
  assert.equal(reason(run({}, {})), 'bad-args');
  assert.equal(reason(run({}, { sheet: 'Nope' })), 'sheet-not-found');
  const broken = printModel(); broken.rowBreaks = 5;
  assert.equal(reason(bridgeOf(apiWith({ S1: { worksheet: broken } })).read(PS2, { sheet: 'S1' })), 'rowBreaks-unexpected-shape', 'a v1 refusal is passed through, never masked');
});

// ---- sheet.autoFilter.colorFilter.read/v1 ----------------------------------------------------------

const rgbC = n => ({ rgb: n, theme: null, tint: null, _sentinel: LEAK });
const themeC = (theme, tint) => ({ rgb: null, theme, tint });
const dxfFill = (fg, bg) => ({ fill: { patternFill: { fgColor: fg, bgColor: bg } }, font: null });
const col = (id, cf) => ({ filter: { ColId: id, ColorFilter: cf, _sentinel: LEAK } });
const cfOf = (cellColor, dxf) => ({ CellColor: cellColor, dxf });
function filterApi(filters, { af = true, throwAf = false, throwFilters = false } = {}) {
  const sheet = {
    GetAutoFilter: () => { if (throwAf) throw new Error('x'); return af ? { GetFilters: () => { if (throwFilters) throw new Error('x'); return filters; } } : null; },
  };
  return apiWith({ S1: sheet });
}
const readCF = (filters, opts, args = { sheet: 'S1' }) => bridgeOf(filterApi(filters, opts)).read(CF, args);

test('colorFilter: cell-colour and font-colour filters resolve to rgb / theme with the kind and the slot the colour was found in (POZ)', () => {
  const r = readCF([
    col(1, cfOf(null, dxfFill(rgbC(0xFF0000), rgbC(0xFF0000)))),
    col(2, cfOf(false, dxfFill(rgbC(0x0000FF), rgbC(0x0000FF)))),
    col(3, cfOf(false, { fill: null, font: { c: rgbC(0x123456) } })),
    col(4, cfOf(true, dxfFill(themeC(3, 0.4), themeC(3, 0.4)))),
    col(5, null),
    col(6, cfOf(null, dxfFill(null, null))),
  ]);
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.columns, [
    { field: 2, kind: 'cellColor', colour: { form: 'rgb', hex: 'FF0000' }, slot: 'fgColor' },
    { field: 3, kind: 'fontColor', colour: { form: 'rgb', hex: '0000FF' }, slot: 'fgColor' },
    { field: 4, kind: 'fontColor', colour: { form: 'rgb', hex: '123456' }, slot: 'fontColor' },
    { field: 5, kind: 'cellColor', colour: { form: 'theme', theme: 3, tint: 0.4 }, slot: 'fgColor' },
    { field: 7, kind: 'cellColor', colour: null, slot: null },
  ].map(x => x), 'the field is the 1-based ColId; a column with no ColorFilter (ColId 5) is omitted');
  noLeak(r);
});

test('colorFilter: no AutoFilter / no colour filter -> an empty list; the alpha byte is dropped; an indexed colour object is reported raw', () => {
  assert.deepEqual(readCF([], { af: false }).value, { columns: [] });
  assert.deepEqual(readCF([col(0, null)]).value, { columns: [] });
  assert.equal(readCF([col(0, cfOf(null, dxfFill(rgbC(0xFF112233), rgbC(0xFF112233))))]).value.columns[0].colour.hex, '112233');
  assert.deepEqual(readCF([col(0, cfOf(null, dxfFill({ indexed: 2 }, { indexed: 2 })))]).value.columns[0].colour, { form: 'indexed', indexed: 2 });
});

test('colorFilter: unreadable shapes refuse the WHOLE read with a named reason (NEG)', () => {
  const ok1 = col(1, cfOf(null, dxfFill(rgbC(1), rgbC(1))));
  assert.equal(reason(readCF([ok1, col(2, cfOf(null, dxfFill(rgbC(1), rgbC(2))))])), 'colour-slots-disagree');
  assert.equal(reason(readCF([ok1, col(2, cfOf(null, dxfFill({ nothing: 1 }, null)))])), 'colour-form-unknown');
  assert.equal(reason(readCF([ok1, col(2, cfOf('x', dxfFill(rgbC(1), rgbC(1))))])), 'cellcolor-unexpected');
  assert.equal(reason(readCF([ok1, col(2, cfOf(null, null))])), 'colorfilter-dxf-missing');
  assert.equal(reason(readCF([ok1, col(2, 5)])), 'colorfilter-unexpected-shape');
  assert.equal(reason(readCF([ok1, { filter: { ColorFilter: cfOf(null, dxfFill(rgbC(1), rgbC(1))) } }])), 'field-unreadable');
  assert.equal(reason(readCF([ok1, { filter: { ColId: -1, ColorFilter: null } }])), 'field-unreadable');
  assert.equal(reason(readCF([ok1, null])), 'field-unreadable');
  assert.equal(reason(readCF({})), 'filters-unexpected-shape');
  assert.equal(reason(readCF([], { throwAf: true })), 'autofilter-threw');
  assert.equal(reason(readCF([], { throwFilters: true })), 'autofilter-threw');
  assert.equal(reason(readCF([], {}, {})), 'bad-args');
  assert.equal(reason(readCF([], {}, { sheet: 'Nope' })), 'sheet-not-found');
});

// ---- chart.series.style.read/v1 --------------------------------------------------------------------

const solid = (r, g, b) => ({ fill: { color: { color: { h: null, s: null, l: null, RGBA: { R: r, G: g, B: b, A: 255 }, _sentinel: LEAK }, Mods: null } } });
const scheme = id => ({ fill: { color: { color: { id, RGBA: { R: 0, G: 0, B: 0, A: 255 } }, Mods: null } } });
const noFill = { fill: {} };
const sppr = (fill, ln, w) => ({ Fill: fill, ln: ln === undefined ? null : { Fill: ln, w: w === undefined ? null : w } });
function chartApi({ chartSpPr, plotSpPr, title, legend, series, plots = 1, name = 'Chart1' } = {}) {
  const ser = series || [{ spPr: sppr(solid(1, 2, 3), solid(4, 5, 6), 12700), dPt: [{ idx: 0, spPr: sppr(solid(16, 17, 18), solid(19, 20, 21)) }, { idx: 2, spPr: sppr(scheme(7), null) }] }, { spPr: null, dPt: [] }];
  const plotCharts = Array.from({ length: plots }, () => ({ series: ser }));
  const model = { spPr: chartSpPr === undefined ? sppr(solid(100, 110, 120), solid(101, 111, 121), 9525) : chartSpPr, chart: { plotArea: { spPr: plotSpPr === undefined ? sppr(solid(10, 20, 30), null) : plotSpPr, charts: plotCharts }, title: title === undefined ? { spPr: sppr(solid(40, 50, 60), solid(41, 51, 61)) } : title, legend: legend === undefined ? { spPr: sppr(null, null) } : legend }, _sentinel: LEAK };
  return apiWith({ S1: { GetAllCharts: () => [{ GetName: () => name, Chart: model }] } });
}
const readST = (o, args) => bridgeOf(chartApi(o)).read(ST, Object.assign({ sheet: 'S1', chart: 'Chart1' }, args));
const hexOf = (r, g, b) => [r, g, b].map(n => n.toString(16).toUpperCase().padStart(2, '0')).join('');

test('chart style: all six targets read their own spPr, fill and outline independently (POZ)', () => {
  assert.deepEqual(readST({}, { target: 'chart' }).value, { element: 'present', fill: { form: 'rgb', hex: hexOf(100, 110, 120) }, outline: { form: 'rgb', hex: hexOf(101, 111, 121) }, outlineWidthEmu: 9525 });
  assert.deepEqual(readST({}, { target: 'plotArea' }).value, { element: 'present', fill: { form: 'rgb', hex: hexOf(10, 20, 30) }, outline: null, outlineWidthEmu: null });
  assert.deepEqual(readST({}, { target: 'title' }).value, { element: 'present', fill: { form: 'rgb', hex: hexOf(40, 50, 60) }, outline: { form: 'rgb', hex: hexOf(41, 51, 61) }, outlineWidthEmu: null });
  assert.deepEqual(readST({}, { target: 'legend' }).value, { element: 'present', fill: null, outline: null, outlineWidthEmu: null });
  assert.deepEqual(readST({}, { target: 'series', series: 0 }).value, { element: 'present', fill: { form: 'rgb', hex: hexOf(1, 2, 3) }, outline: { form: 'rgb', hex: hexOf(4, 5, 6) }, outlineWidthEmu: 12700 });
  assert.deepEqual(readST({}, { target: 'dataPoint', series: 0, point: 0 }).value, { element: 'present', fill: { form: 'rgb', hex: hexOf(16, 17, 18) }, outline: { form: 'rgb', hex: hexOf(19, 20, 21) }, outlineWidthEmu: null });
  for (const t of ['chart', 'plotArea', 'title', 'legend']) noLeak(readST({}, { target: t }));
});

test('chart style: a scheme colour is named, never guessed as rgb; a non-solid fill is "other"; an absent element / empty spPr is not an error', () => {
  assert.deepEqual(readST({}, { target: 'dataPoint', series: 0, point: 2 }).value, { element: 'present', fill: { form: 'scheme', id: 7 }, outline: null, outlineWidthEmu: null });
  assert.deepEqual(readST({ chartSpPr: sppr(noFill, noFill) }, { target: 'chart' }).value, { element: 'present', fill: { form: 'other' }, outline: { form: 'other' }, outlineWidthEmu: null });
  assert.deepEqual(readST({}, { target: 'series', series: 1 }).value, { element: 'present', fill: null, outline: null, outlineWidthEmu: null }, 'spPr === null is the model\'s own no-style state');
  assert.deepEqual(readST({}, { target: 'dataPoint', series: 0, point: 5 }).value, { element: 'absent', fill: null, outline: null, outlineWidthEmu: null }, 'a point with no own dPt');
  assert.deepEqual(readST({ title: null }, { target: 'title' }).value, { element: 'absent', fill: null, outline: null, outlineWidthEmu: null });
});

test('chart style: refusals are NAMED and never silently coerce (NEG)', () => {
  assert.equal(reason(readST({}, { target: 'axis' })), 'unsupported-target');
  assert.equal(reason(readST({}, {})), 'bad-args');
  assert.equal(reason(readST({}, { target: 'series' })), 'bad-args');
  assert.equal(reason(readST({}, { target: 'series', series: '0' })), 'bad-args');
  assert.equal(reason(readST({}, { target: 'series', series: -1 })), 'bad-args');
  assert.equal(reason(readST({}, { target: 'dataPoint', series: 0 })), 'bad-args');
  assert.equal(reason(readST({}, { target: 'chart', series: 0 })), 'bad-args');
  assert.equal(reason(readST({}, { target: 'series', series: 0, point: 0 })), 'bad-args');
  assert.equal(reason(readST({}, { target: 'series', series: 9 })), 'series-index-out-of-range');
  assert.equal(reason(readST({ plots: 2 }, { target: 'series', series: 0 })), 'chart-type-not-single');
  assert.equal(readST({ plots: 2 }, { target: 'chart' }).ok, true, 'the chart-level targets do not need a single plot');
  assert.equal(reason(bridgeOf(chartApi()).read(ST, { sheet: 'S1', chart: 'Nope', target: 'chart' })), 'chart-not-found');
  assert.equal(reason(bridgeOf(chartApi()).read(ST, { sheet: 'Nope', chart: 'Chart1', target: 'chart' })), 'sheet-not-found');
  const dup = [{ spPr: null, dPt: [{ idx: 1, spPr: sppr(solid(1, 1, 1), null) }, { idx: 1, spPr: sppr(null, solid(2, 2, 2)) }] }];
  assert.equal(reason(readST({ series: dup }, { target: 'dataPoint', series: 0, point: 1 })), 'datapoint-ambiguous');
  assert.equal(reason(readST({ series: [{ spPr: null, dPt: 5 }] }, { target: 'dataPoint', series: 0, point: 1 })), 'dPt-unexpected-shape');
  const mod = solid(1, 2, 3); mod.fill.color.Mods = { Mods: [{ name: 'lumMod' }] };
  assert.equal(reason(readST({ chartSpPr: sppr(mod, null) }, { target: 'chart' })), 'sppr-chain-broken', 'a colour with modifiers is not reported as its base rgb');
  const broken = solid(1, 2, 3); broken.fill.color.color.RGBA = { R: 'x', G: 1, B: 2 };
  assert.equal(reason(readST({ chartSpPr: sppr(broken, null) }, { target: 'chart' })), 'sppr-chain-broken');
  assert.equal(reason(readST({ chartSpPr: sppr(solid(1, 2, 3), broken) }, { target: 'chart' })), 'sppr-chain-broken', 'fill and outline are read independently: a broken outline refuses');
});

test('the three new ids are served by read() only: write() and hydrate() refuse them, an unknown id is refused by name', () => {
  const sheets = { S1: { worksheet: printModel() } };
  assert.equal(bridgeOf(apiWith(sheets)).read(PS2, { sheet: 'S1' }).ok, true);
  assert.equal(readCF([col(0, cfOf(null, dxfFill(rgbC(1), rgbC(1))))]).ok, true);
  assert.equal(readST({}, { target: 'chart' }).ok, true);
  for (const id of [PS2, CF, ST]) { assert.equal(bridgeOf(apiWith(sheets)).write(id, { sheet: 'S1' }).reason, 'capability-not-advertised', id); assert.equal(bridgeOf(apiWith(sheets)).hydrate(id, { sheet: 'S1' }).reason, 'capability-not-advertised', id); }
  assert.equal(bridgeOf(apiWith(sheets)).read('sheet.nonexistent/v1', {}).reason, 'capability-not-advertised');
});
