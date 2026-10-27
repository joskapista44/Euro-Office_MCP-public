'use strict';

/**
 * X-O05b, unit U2 (design doc `docs/XO05B-FILE-VERIFIED-DESIGN-20261006.md`). PRODUCT reader
 * (not test tooling): the post-close file-verification phase
 * (U1, engine) needs to read the SAVED package's own bytes to prove a point data-label write
 * persisted, since the public API has no step-level readback (`GetDataLabels` is chart-level
 * only -- design doc "Measured facts").
 *
 * Ported from the FILE_VERIFIER test/evidence module (`docs/FILE-VERIFIER-DESIGN-20261006.md`,
 * LIVE-proven chart-part-by-name resolution chain), NOT required from it -- that module is
 * explicitly test/evidence-only (its own header: "NOT product runtime... binding it into the
 * product adapters/drivers is a separate, later owner question"); this IS that later question,
 * scoped to exactly the one shape this unit needs. Only the zip reader, the XML tokenizer, the
 * tree helpers, and the sheet-name -> drawing -> chart-by-name resolution chain are ported;
 * everything else in that module (cell style, row/col geometry, page setup, drawing anchors,
 * media, chart legend/axis-titles/chart-level-labels) is out of scope for this reader.
 *
 * Every function returns `{ ok: true, value }` or `{ ok: false, reason }`. NEVER a guessed
 * default -- `reason` is always one of the named states below, never a raw exception message
 * (which could leak document content into a log/evidence record).
 */

const zlib = require('node:zlib');

// ---- namespace URIs, by name (same discipline as the source module: written out, never matched
// by prefix) -- only the five this chain actually touches. ------------------------------------

const NS = Object.freeze({
  SPREADSHEET: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  RELATIONSHIPS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  PACKAGE_RELATIONSHIPS: 'http://schemas.openxmlformats.org/package/2006/relationships',
  DRAWING_SPREADSHEET: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  CHART: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
});

// ================================================================================================
// ZIP reader (ported verbatim from the FILE_VERIFIER module's validated EOCD/central-directory/
// local-header pattern).
// ================================================================================================

const EOCD_SIG = 0x06054b50;
const CENTRAL_SIG = 0x02014b50;
const LOCAL_SIG = 0x04034b50;

function findEOCD(buf) {
  for (let i = buf.length - 22; i >= 0; i -= 1) if (buf.readUInt32LE(i) === EOCD_SIG) return i;
  return -1;
}

function parseZipCentralDirectory(buf) {
  const eocdPos = findEOCD(buf);
  if (eocdPos < 0) return { ok: false, reason: 'zip-eocd-not-found' };
  const totalEntries = buf.readUInt16LE(eocdPos + 10);
  let pos = buf.readUInt32LE(eocdPos + 16);
  const entries = [];
  for (let i = 0; i < totalEntries; i += 1) {
    if (pos + 46 > buf.length || buf.readUInt32LE(pos) !== CENTRAL_SIG) return { ok: false, reason: 'zip-central-signature-mismatch' };
    const compressionMethod = buf.readUInt16LE(pos + 10);
    const crc32 = buf.readUInt32LE(pos + 16);
    const compressedSize = buf.readUInt32LE(pos + 20);
    const filenameLen = buf.readUInt16LE(pos + 28);
    const extraLen = buf.readUInt16LE(pos + 30);
    const commentLen = buf.readUInt16LE(pos + 32);
    const localHeaderOffset = buf.readUInt32LE(pos + 42);
    const filename = buf.toString('utf8', pos + 46, pos + 46 + filenameLen);
    entries.push({ filename, compressionMethod, crc32, compressedSize, localHeaderOffset });
    pos += 46 + filenameLen + extraLen + commentLen;
  }
  return { ok: true, entries };
}

function readZipEntryData(buf, entry) {
  const lp = entry.localHeaderOffset;
  if (lp + 30 > buf.length || buf.readUInt32LE(lp) !== LOCAL_SIG) return { ok: false, reason: 'zip-local-signature-mismatch' };
  const localFilenameLen = buf.readUInt16LE(lp + 26);
  const localExtraLen = buf.readUInt16LE(lp + 28);
  const dataStart = lp + 30 + localFilenameLen + localExtraLen;
  const raw = buf.subarray(dataStart, dataStart + entry.compressedSize);
  if (entry.compressionMethod === 0) return { ok: true, data: raw };
  if (entry.compressionMethod === 8) {
    try { return { ok: true, data: zlib.inflateRawSync(raw) }; } catch { return { ok: false, reason: 'zip-inflate-failed' }; }
  }
  return { ok: false, reason: 'zip-unsupported-compression' };
}

/** Reads and decompresses one named entry's text (UTF-8 only). A duplicate-named entry ->
 * UNOBSERVABLE, never "take the first". CRC-32 checked when `zlib.crc32` is available on this
 * Node runtime; `crcChecked: false` on success names the gap explicitly otherwise. */
function readZipPartText(buf, partName) {
  const dir = parseZipCentralDirectory(buf);
  if (!dir.ok) return dir;
  const matches = dir.entries.filter(e => e.filename === partName);
  if (matches.length === 0) return { ok: false, reason: 'zip-part-not-found' };
  if (matches.length > 1) return { ok: false, reason: 'zip-duplicate-entry' };
  const entry = matches[0];
  const got = readZipEntryData(buf, entry);
  if (!got.ok) return got;
  const crcChecked = typeof zlib.crc32 === 'function';
  if (crcChecked && zlib.crc32(got.data) !== entry.crc32) return { ok: false, reason: 'zip-crc-mismatch' };
  let text;
  try { text = new TextDecoder('utf-8', { fatal: true }).decode(got.data); } catch { return { ok: false, reason: 'encoding-error' }; }
  return { ok: true, text, crcChecked };
}

// ================================================================================================
// The minimal XML tokenizer (ported verbatim). Element/attribute/namespace-scope aware.
// DOCTYPE/ENTITY -> hard reject (closes the classic XXE vector). A non-UTF-8 encoding
// declaration -> hard reject.
// ================================================================================================

function splitQName(name) {
  const i = name.indexOf(':');
  return i < 0 ? { prefix: '', local: name } : { prefix: name.slice(0, i), local: name.slice(i + 1) };
}

function resolveUri(scope, prefix) {
  return prefix in scope ? scope[prefix] : null;
}

const PREDEFINED_ENTITIES = Object.freeze({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" });

function decodeEntities(s) {
  let out = '';
  let i = 0;
  while (i < s.length) {
    const amp = s.indexOf('&', i);
    if (amp < 0) return out + s.slice(i);
    out += s.slice(i, amp);
    const semi = s.indexOf(';', amp + 1);
    if (semi < 0 || semi - amp > 10) return null;
    const body = s.slice(amp + 1, semi);
    if (body[0] === '#') {
      const hex = body[1] === 'x' || body[1] === 'X';
      const digits = hex ? body.slice(2) : body.slice(1);
      if (!digits || !new RegExp(`^[0-9${hex ? 'a-fA-F' : ''}]+$`).test(digits)) return null;
      const code = parseInt(digits, hex ? 16 : 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return null;
      out += String.fromCodePoint(code);
    } else if (Object.prototype.hasOwnProperty.call(PREDEFINED_ENTITIES, body)) {
      out += PREDEFINED_ENTITIES[body];
    } else {
      return null;
    }
    i = semi + 1;
  }
  return out;
}

const ATTR_RE = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;

function parseAttrText(raw) {
  const out = [];
  let i = 0;
  while (i < raw.length) {
    while (i < raw.length && /\s/.test(raw[i])) i += 1;
    if (i >= raw.length) break;
    ATTR_RE.lastIndex = i;
    const m = ATTR_RE.exec(raw);
    if (!m || m.index !== i) return null;
    const decoded = decodeEntities(m[3] !== undefined ? m[3] : m[4]);
    if (decoded === null) return null;
    out.push({ name: m[1], value: decoded });
    i = ATTR_RE.lastIndex;
  }
  return out;
}

function parseStartTagBody(body, parentScope) {
  const nameMatch = body.match(/^([^\s/>]+)/);
  if (!nameMatch) return { ok: false, reason: 'malformed-start-tag' };
  const { prefix: tagPrefix, local: tagLocal } = splitQName(nameMatch[1]);
  const rawAttrs = parseAttrText(body.slice(nameMatch[1].length));
  if (rawAttrs === null) return { ok: false, reason: 'malformed-attributes' };
  const scope = Object.create(parentScope);
  for (const a of rawAttrs) {
    if (a.name === 'xmlns') scope[''] = a.value;
    else if (a.name.startsWith('xmlns:')) scope[a.name.slice(6)] = a.value;
  }
  const tagUri = tagPrefix ? resolveUri(scope, tagPrefix) : resolveUri(scope, '');
  if (tagPrefix && tagUri === null) return { ok: false, reason: 'unresolved-namespace-prefix' };
  const attrs = [];
  for (const a of rawAttrs) {
    if (a.name === 'xmlns' || a.name.startsWith('xmlns:')) continue;
    const { prefix: ap, local: al } = splitQName(a.name);
    const auri = ap ? resolveUri(scope, ap) : null;
    if (ap && auri === null) return { ok: false, reason: 'unresolved-namespace-prefix' };
    attrs.push({ uri: auri, local: al, prefix: ap, value: a.value });
  }
  return { ok: true, scope, node: { type: 'element', uri: tagUri, local: tagLocal, prefix: tagPrefix, attrs, children: [] } };
}

function matchEndTag(name, stack) {
  if (!stack.length) return { ok: false, reason: 'unmatched-end-tag' };
  const top = stack[stack.length - 1];
  const { prefix, local } = splitQName(name);
  if (top.node.local !== local || top.node.prefix !== prefix) return { ok: false, reason: 'mismatched-end-tag' };
  stack.pop();
  return { ok: true };
}

function checkDeclarationAndDoctype(xml) {
  const xmlDeclMatch = xml.match(/^\s*<\?xml\s+[^?]*\?>/);
  if (xmlDeclMatch) {
    const enc = xmlDeclMatch[0].match(/encoding\s*=\s*["']([^"']+)["']/i);
    if (enc && enc[1].toLowerCase() !== 'utf-8') return { ok: false, reason: 'encoding-error' };
  }
  if (/<!DOCTYPE/i.test(xml) || /<!ENTITY/i.test(xml)) return { ok: false, reason: 'doctype-or-entity-not-supported' };
  return { ok: true, bodyStart: xmlDeclMatch ? xmlDeclMatch[0].length : 0 };
}

function parseXml(xml) {
  const pre = checkDeclarationAndDoctype(xml);
  if (!pre.ok) return pre;

  let i = pre.bodyStart;
  const len = xml.length;
  const rootScope = Object.create(null);
  const stack = [];
  let root = null;

  const pushText = (parent, value, cdata) => { if (value !== '' || cdata) parent.children.push({ type: 'text', value, cdata }); };

  while (i < len) {
    if (xml[i] !== '<') {
      const next = xml.indexOf('<', i);
      const text = next < 0 ? xml.slice(i) : xml.slice(i, next);
      if (stack.length) {
        const decoded = decodeEntities(text);
        if (decoded === null) return { ok: false, reason: 'malformed-entity-reference' };
        pushText(stack[stack.length - 1].node, decoded, false);
      } else if (text.trim() !== '') return { ok: false, reason: 'text-outside-root' };
      i = next < 0 ? len : next;
      continue;
    }
    if (xml.startsWith('<!--', i)) {
      const end = xml.indexOf('-->', i + 4);
      if (end < 0) return { ok: false, reason: 'unterminated-comment' };
      i = end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', i)) {
      const end = xml.indexOf(']]>', i + 9);
      if (end < 0) return { ok: false, reason: 'unterminated-cdata' };
      if (!stack.length) return { ok: false, reason: 'text-outside-root' };
      pushText(stack[stack.length - 1].node, xml.slice(i + 9, end), true);
      i = end + 3;
      continue;
    }
    if (xml.startsWith('<?', i)) {
      const end = xml.indexOf('?>', i + 2);
      if (end < 0) return { ok: false, reason: 'unterminated-pi' };
      i = end + 2;
      continue;
    }
    if (xml.startsWith('</', i)) {
      const end = xml.indexOf('>', i + 2);
      if (end < 0) return { ok: false, reason: 'unterminated-end-tag' };
      const matched = matchEndTag(xml.slice(i + 2, end).trim(), stack);
      if (!matched.ok) return matched;
      i = end + 1;
      continue;
    }
    const end = xml.indexOf('>', i + 1);
    if (end < 0) return { ok: false, reason: 'unterminated-start-tag' };
    let body = xml.slice(i + 1, end);
    const selfClosing = body.endsWith('/');
    if (selfClosing) body = body.slice(0, -1);
    const parentScope = stack.length ? stack[stack.length - 1].scope : rootScope;
    const tag = parseStartTagBody(body, parentScope);
    if (!tag.ok) return tag;
    if (stack.length) stack[stack.length - 1].node.children.push(tag.node); else { if (root) return { ok: false, reason: 'multiple-root-elements' }; root = tag.node; }
    if (!selfClosing) stack.push({ node: tag.node, scope: tag.scope });
    i = end + 1;
  }
  if (stack.length) return { ok: false, reason: 'unclosed-element' };
  if (!root) return { ok: false, reason: 'no-root-element' };
  return { ok: true, root };
}

// ---- tree helpers (ported verbatim) -------------------------------------------------------------

const children = (el, uri, local) => el.children.filter(c => c.type === 'element' && c.uri === uri && c.local === local);
const child = (el, uri, local) => { const c = children(el, uri, local); return c.length === 1 ? c[0] : null; };
const attr = (el, uri, local) => { const a = el.attrs.find(x => x.uri === uri && x.local === local); return a ? a.value : undefined; };
function findAll(el, uri, local, out = []) {
  for (const c of el.children) { if (c.type === 'element') { if (c.uri === uri && c.local === local) out.push(c); findAll(c, uri, local, out); } }
  return out;
}

function readPartXml(zipBuf, partName) {
  const got = readZipPartText(zipBuf, partName);
  if (!got.ok) return got;
  return parseXml(got.text);
}

/** `.rels` Targets are relative to the directory the .rels file's OWN part lives in (ported
 * verbatim), not to the .rels file's own path. Resolves `..` segments. */
function normalizeZipPath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const segments = (baseDir ? baseDir.split('/') : []).concat(target.split('/'));
  const out = [];
  for (const seg of segments) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') out.pop();
    else out.push(seg);
  }
  return out.join('/');
}

const dirOf = partPath => { const i = partPath.lastIndexOf('/'); return i < 0 ? '' : partPath.slice(0, i); };
const relsPathFor = partPath => { const i = partPath.lastIndexOf('/'); const dir = i < 0 ? '' : partPath.slice(0, i); const name = i < 0 ? partPath : partPath.slice(i + 1); return `${dir ? `${dir}/` : ''}_rels/${name}.rels`; };

function resolveRelTarget(relsRoot, rId, basePartDir) {
  const matches = findAll(relsRoot, NS.PACKAGE_RELATIONSHIPS, 'Relationship').filter(r => attr(r, null, 'Id') === rId);
  if (matches.length !== 1) return { ok: false, reason: matches.length === 0 ? 'relationship-id-not-found' : 'relationship-id-ambiguous' };
  const target = attr(matches[0], null, 'Target');
  if (!target) return { ok: false, reason: 'relationship-target-missing' };
  return { ok: true, path: normalizeZipPath(basePartDir, target) };
}

function resolveSheetRid(zipBuf, sheetName) {
  const wb = readPartXml(zipBuf, 'xl/workbook.xml');
  if (!wb.ok) return wb;
  const sheets = findAll(wb.root, NS.SPREADSHEET, 'sheet').filter(s => attr(s, null, 'name') === sheetName);
  if (sheets.length !== 1) return { ok: false, reason: sheets.length === 0 ? 'sheet-not-found' : 'sheet-name-ambiguous' };
  const rid = attr(sheets[0], NS.RELATIONSHIPS, 'id');
  if (!rid) return { ok: false, reason: 'sheet-relationship-id-missing' };
  return { ok: true, rid };
}

function resolveSheetDrawingRid(sheetRoot) {
  const drawings = findAll(sheetRoot, NS.SPREADSHEET, 'drawing');
  if (drawings.length !== 1) return { ok: false, reason: drawings.length === 0 ? 'sheet-drawing-not-found' : 'sheet-drawing-ambiguous' };
  const rid = attr(drawings[0], NS.RELATIONSHIPS, 'id');
  if (!rid) return { ok: false, reason: 'drawing-relationship-id-missing' };
  return { ok: true, rid };
}

function resolveChartRidInDrawing(drawingRoot, chartName) {
  const frames = findAll(drawingRoot, NS.DRAWING_SPREADSHEET, 'graphicFrame');
  const named = frames.filter(f => { const nv = findAll(f, NS.DRAWING_SPREADSHEET, 'cNvPr')[0]; return nv && attr(nv, null, 'name') === chartName; });
  if (named.length !== 1) return { ok: false, reason: named.length === 0 ? 'chart-frame-not-found' : 'chart-frame-name-ambiguous' };
  const chartRefs = findAll(named[0], NS.CHART, 'chart');
  if (chartRefs.length !== 1) return { ok: false, reason: chartRefs.length === 0 ? 'chart-ref-not-found' : 'chart-ref-ambiguous' };
  const rid = attr(chartRefs[0], NS.RELATIONSHIPS, 'id');
  if (!rid) return { ok: false, reason: 'chart-relationship-id-missing' };
  return { ok: true, rid };
}

/** Sheet NAME + chart NAME -> the chart part's own zip path. Ported verbatim from
 * `resolveChartPart` (FILE_VERIFIER, LIVE-proven chain): workbook sheet -> rels -> sheet part ->
 * its one `<drawing>` -> rels -> drawing part -> the one `graphicFrame` whose `cNvPr/@name`
 * matches -> rels -> chart part. */
function resolveChartPart(zipBuf, sheetName, chartName) {
  const sheetRid = resolveSheetRid(zipBuf, sheetName);
  if (!sheetRid.ok) return sheetRid;
  const wbRels = readPartXml(zipBuf, 'xl/_rels/workbook.xml.rels');
  if (!wbRels.ok) return wbRels;
  const sheetPathR = resolveRelTarget(wbRels.root, sheetRid.rid, 'xl');
  if (!sheetPathR.ok) return sheetPathR;

  const sheetXml = readPartXml(zipBuf, sheetPathR.path);
  if (!sheetXml.ok) return sheetXml;
  const drawingRid = resolveSheetDrawingRid(sheetXml.root);
  if (!drawingRid.ok) return drawingRid;
  const sheetRels = readPartXml(zipBuf, relsPathFor(sheetPathR.path));
  if (!sheetRels.ok) return sheetRels;
  const drawingPathR = resolveRelTarget(sheetRels.root, drawingRid.rid, dirOf(sheetPathR.path));
  if (!drawingPathR.ok) return drawingPathR;

  const drawingXml = readPartXml(zipBuf, drawingPathR.path);
  if (!drawingXml.ok) return drawingXml;
  const chartRid = resolveChartRidInDrawing(drawingXml.root, chartName);
  if (!chartRid.ok) return chartRid;
  const drawingRels = readPartXml(zipBuf, relsPathFor(drawingPathR.path));
  if (!drawingRels.ok) return drawingRels;
  const chartPathR = resolveRelTarget(drawingRels.root, chartRid.rid, dirOf(drawingPathR.path));
  if (!chartPathR.ok) return chartPathR;

  return { ok: true, value: { chartPartPath: chartPathR.path } };
}

// ================================================================================================
// readPointDataLabel -- the one shape this unit needs: c:ser/c:dLbls/c:dLbl[c:idx=point]'s four
// show-flags (design doc "Measured facts": `SetShowPointDataLabel` persists exactly this element,
// X-O05c probe2). Deliberately narrower than the chart-level `readChartDataLabels` in the source
// module -- this never reads/needs the chart-level or series-level (non-per-point) c:dLbls.
// ================================================================================================

const CHART_TYPE_TAGS = Object.freeze(['barChart', 'pieChart', 'lineChart', 'areaChart', 'doughnutChart', 'scatterChart', 'stockChart']);
const CHART_ELEMENT_RE = /Chart$/;

/** The chart part's ONE chart-type element (barChart/pieChart/...) under c:plotArea -- absence
 * or more than one is UNOBSERVABLE, never a guess at which one is "the" chart. Narrower than the
 * source module's `readChartType`: this never needs the bar-chart subtype fields (barDir/
 * grouping), which are irrelevant to reading a point's data-label flags. */
function resolveSingleChartTypeElement(chartRoot) {
  const plotArea = findAll(chartRoot, NS.CHART, 'plotArea')[0];
  if (!plotArea) return { ok: false, reason: 'plot-area-not-found' };
  const present = plotArea.children.filter(c => c.type === 'element' && c.uri === NS.CHART && CHART_ELEMENT_RE.test(c.local));
  if (present.length !== 1) return { ok: false, reason: present.length === 0 ? 'chart-type-not-found' : 'chart-type-combo-ambiguous' };
  if (!CHART_TYPE_TAGS.includes(present[0].local)) return { ok: false, reason: 'chart-type-unsupported' };
  return { ok: true, value: present[0] };
}

const POINT_LABEL_FLAGS = Object.freeze({ serName: 'showSerName', catName: 'showCatName', val: 'showVal', percent: 'showPercent' });
const BOOL_VAL = Object.freeze({ '1': true, true: true, '0': false, false: false });

function isNonNegativeInt(n) {
  return Number.isInteger(n) && n >= 0;
}

/**
 * `bytes` (the saved .xlsx package) + sheet NAME + chart NAME + zero-based series/point index ->
 * `{ ok: true, value: { serName, catName, val, percent } }` (the point's own four show-flags, as
 * booleans, keyed the SAME way as the `set_chart_point_data_label` intent's `labels` field -- an
 * engine comparison needs no key translation) or `{ ok: false, reason }`.
 *
 * Fail-closed states beyond `resolveChartPart`'s own (sheet/drawing/chart-frame resolution) and
 * the XML/zip parse errors above: `series-index-malformed`/`-out-of-range`, `series-dlbls-not-
 * present` (no per-point override is even possible without a series-level `c:dLbls`),
 * `point-index-malformed`, `point-label-absent` (no `c:dLbl` carries this `c:idx` -- the DISTINCT
 * "no override written for this point" state, never conflated with an error), `point-label-
 * ambiguous` (more than one `c:dLbl` with the same `c:idx` -- a malformed file, not "last wins"),
 * `point-label-deleted` (`c:delete val="1"` on the matched entry -- no show-flags to read),
 * `point-label-field-missing`/`-malformed` (one of the four flags absent or not a recognized
 * boolean literal).
 */
function readPointDataLabel(bytes, sheetName, chartName, seriesIdx, pointIdx) {
  if (!isNonNegativeInt(seriesIdx)) return { ok: false, reason: 'series-index-malformed' };
  if (!isNonNegativeInt(pointIdx)) return { ok: false, reason: 'point-index-malformed' };

  const partR = resolveChartPart(bytes, sheetName, chartName);
  if (!partR.ok) return partR;
  const chartXml = readPartXml(bytes, partR.value.chartPartPath);
  if (!chartXml.ok) return chartXml;

  const typeEl = resolveSingleChartTypeElement(chartXml.root);
  if (!typeEl.ok) return typeEl;

  const series = children(typeEl.value, NS.CHART, 'ser');
  if (seriesIdx >= series.length) return { ok: false, reason: 'series-index-out-of-range' };
  const ser = series[seriesIdx];

  const dLbls = child(ser, NS.CHART, 'dLbls');
  if (!dLbls) return { ok: false, reason: 'series-dlbls-not-present' };

  const matches = children(dLbls, NS.CHART, 'dLbl').filter(d => {
    const idxEl = child(d, NS.CHART, 'idx');
    return idxEl && attr(idxEl, null, 'val') === String(pointIdx);
  });
  if (matches.length === 0) return { ok: false, reason: 'point-label-absent' };
  if (matches.length > 1) return { ok: false, reason: 'point-label-ambiguous' };
  const dLbl = matches[0];

  const deleteEl = child(dLbl, NS.CHART, 'delete');
  if (deleteEl && attr(deleteEl, null, 'val') in BOOL_VAL && BOOL_VAL[attr(deleteEl, null, 'val')]) {
    return { ok: false, reason: 'point-label-deleted' };
  }

  const value = {};
  for (const [outKey, tag] of Object.entries(POINT_LABEL_FLAGS)) {
    const el = child(dLbl, NS.CHART, tag);
    if (!el) return { ok: false, reason: 'point-label-field-missing' };
    const val = attr(el, null, 'val');
    if (!(val in BOOL_VAL)) return { ok: false, reason: 'point-label-field-malformed' };
    value[outKey] = BOOL_VAL[val];
  }
  return { ok: true, value };
}

module.exports = { readPointDataLabel, resolveChartPart };
