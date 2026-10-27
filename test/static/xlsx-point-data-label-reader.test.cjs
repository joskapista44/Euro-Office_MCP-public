'use strict';

// Offline unit tests of point-data-label-reader.cjs (X-O05b U2), on synthetic OOXML.
// Every PASS case and every named fail-closed case per the design doc / dispatch.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const r = require('../../src/adapters/xlsx/ooxml/point-data-label-reader.cjs');

// every mkdtemp dir is removed when the file's tests end (a LIVE mkdtemp-leak finding)
const made = [];
const mkdtempTracked = prefix => { const d = fs.mkdtempSync(path.join(os.tmpdir(), prefix)); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });

const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const SML = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** The chart part: a barChart with `seriesCount` series, the FIRST series carrying `dLblsInner`
 * as its own `c:dLbls` body (raw XML, caller-supplied) when given. */
function chartPartXml({ seriesCount = 1, dLblsInner } = {}) {
  const ser = i => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${i === 0 && dLblsInner !== undefined ? `<c:dLbls>${dLblsInner}</c:dLbls>` : ''}</c:ser>`;
  const sers = Array.from({ length: seriesCount }, (_, i) => ser(i)).join('');
  return `<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${sers}</c:barChart></c:plotArea></c:chart></c:chartSpace>`;
}

/** Full multi-part package: one workbook/sheet/drawing, one named chart part. */
function buildPackage({ chartName = 'Chart 1', chartXml = chartPartXml() } = {}) {
  const dir = mkdtempTracked('pdl-pkg-');
  const write = (rel, content) => { const p = path.join(dir, rel); fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, content); };

  write('xl/workbook.xml', `<workbook xmlns="${SML}" xmlns:r="${R}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  write('xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>`);
  write('xl/worksheets/sheet1.xml', `<worksheet xmlns="${SML}" xmlns:r="${R}"><drawing r:id="rId1"/></worksheet>`);
  write('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="${PR}"><Relationship Id="rId1" Type="t" Target="../drawings/drawing1.xml"/></Relationships>`);
  write('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}"><xdr:twoCellAnchor><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${chartName}"/></xdr:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>`);
  write('xl/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="${PR}"><Relationship Id="rIdChart" Type="t" Target="../charts/chart1.xml"/></Relationships>`);
  write('xl/charts/chart1.xml', chartXml);

  const zipPath = path.join(dir, '_.zip');
  const rel = p => path.relative(dir, p);
  const allFiles = [];
  (function walk(d) { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else allFiles.push(rel(p)); } })(dir);
  execFileSync('zip', ['-q', zipPath, ...allFiles.filter(f => f !== '_.zip')], { cwd: dir });
  return fs.readFileSync(zipPath);
}

const dLbl = (idx, flags, { deleted } = {}) => deleted
  ? `<c:dLbl><c:idx val="${idx}"/><c:delete val="1"/></c:dLbl>`
  : `<c:dLbl><c:idx val="${idx}"/><c:showSerName val="${flags.serName ? 1 : 0}"/><c:showCatName val="${flags.catName ? 1 : 0}"/><c:showVal val="${flags.val ? 1 : 0}"/><c:showPercent val="${flags.percent ? 1 : 0}"/></c:dLbl>`;

test('readPointDataLabel PASS: the point\'s own c:dLbl flags, matched by c:idx, among several points', () => {
  const inner = dLbl(0, { serName: false, catName: true, val: true, percent: false }) + dLbl(1, { serName: true, catName: false, val: false, percent: true });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 0), { ok: true, value: { serName: false, catName: true, val: true, percent: false } });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 1), { ok: true, value: { serName: true, catName: false, val: false, percent: true } });
});

test('readPointDataLabel fail-closed: point-label-absent -- the series has c:dLbls but no c:dLbl for THIS point', () => {
  const inner = dLbl(5, { serName: true, catName: true, val: true, percent: true });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'point-label-absent' });
});

test('readPointDataLabel fail-closed: series-dlbls-not-present -- the series has no c:dLbls at all', () => {
  const buf = buildPackage({ chartXml: chartPartXml({ seriesCount: 1 }) }); // no dLblsInner given
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'series-dlbls-not-present' });
});

test('readPointDataLabel fail-closed: series-index-out-of-range / series-index-malformed', () => {
  const inner = dLbl(0, { serName: true, catName: true, val: true, percent: true });
  const buf = buildPackage({ chartXml: chartPartXml({ seriesCount: 2, dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 2, 0), { ok: false, reason: 'series-index-out-of-range' });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', -1, 0), { ok: false, reason: 'series-index-malformed' });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 1.5, 0), { ok: false, reason: 'series-index-malformed' });
});

test('readPointDataLabel fail-closed: point-index-malformed', () => {
  const inner = dLbl(0, { serName: true, catName: true, val: true, percent: true });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, -1), { ok: false, reason: 'point-index-malformed' });
});

test('readPointDataLabel fail-closed: point-label-ambiguous -- two c:dLbl entries carry the SAME c:idx', () => {
  const inner = dLbl(0, { serName: true, catName: true, val: true, percent: true }) + dLbl(0, { serName: false, catName: false, val: false, percent: false });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'point-label-ambiguous' });
});

test('readPointDataLabel fail-closed: point-label-deleted -- c:delete val="1" on the matched entry, no flags to read', () => {
  const inner = dLbl(0, {}, { deleted: true });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'point-label-deleted' });
});

test('readPointDataLabel fail-closed: point-label-field-missing / point-label-field-malformed', () => {
  const missing = `<c:dLbl><c:idx val="0"/><c:showSerName val="1"/><c:showCatName val="1"/><c:showVal val="1"/></c:dLbl>`; // showPercent absent
  const bufMissing = buildPackage({ chartXml: chartPartXml({ dLblsInner: missing }) });
  assert.deepEqual(r.readPointDataLabel(bufMissing, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'point-label-field-missing' });

  const malformed = `<c:dLbl><c:idx val="0"/><c:showSerName val="yes"/><c:showCatName val="1"/><c:showVal val="1"/><c:showPercent val="1"/></c:dLbl>`;
  const bufMalformed = buildPackage({ chartXml: chartPartXml({ dLblsInner: malformed }) });
  assert.deepEqual(r.readPointDataLabel(bufMalformed, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'point-label-field-malformed' });
});

test('readPointDataLabel fail-closed: chart/sheet resolution reuses resolveChartPart\'s own reasons', () => {
  const inner = dLbl(0, { serName: true, catName: true, val: true, percent: true });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.readPointDataLabel(buf, 'NoSuchSheet', 'Chart 1', 0, 0), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'No Such Chart', 0, 0), { ok: false, reason: 'chart-frame-not-found' });
});

test('readPointDataLabel fail-closed: chart-type-not-found (no barChart/pieChart/... element under plotArea)', () => {
  const noType = `<c:chartSpace xmlns:c="${C}"><c:chart><c:plotArea/></c:chart></c:chartSpace>`;
  const buf = buildPackage({ chartXml: noType });
  assert.deepEqual(r.readPointDataLabel(buf, 'Sheet1', 'Chart 1', 0, 0), { ok: false, reason: 'chart-type-not-found' });
});

test('resolveChartPart is re-exported and resolves the same part this reader uses internally', () => {
  const inner = dLbl(0, { serName: true, catName: true, val: true, percent: true });
  const buf = buildPackage({ chartXml: chartPartXml({ dLblsInner: inner }) });
  assert.deepEqual(r.resolveChartPart(buf, 'Sheet1', 'Chart 1'), { ok: true, value: { chartPartPath: 'xl/charts/chart1.xml' } });
});
