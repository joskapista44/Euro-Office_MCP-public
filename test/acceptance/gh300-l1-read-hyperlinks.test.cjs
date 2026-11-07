'use strict';

/**
 * GH #300 L1 (owner decision, path A, order A: after read_tables): read_hyperlinks -- the XLSX READ of every
 * hyperlink of one sheet, through the DS Bridge sheet.hyperlinks.read/v1, served only on patchVersion /18. The
 * driver's hyperlinkSheets read (xlsx-driver.cjs, the SAME per-sheet bridgeAxis convention tableReads already
 * uses, run here against the REAL bridge-block code -- not an imitation), the projection
 * (families/hyperlinks-read.cjs), the request it makes and the catalogue entry. A /16 (the shipped block,
 * test/fakes/ds-bridge-runtime.cjs) or /17 (the protection-only candidate) runtime does not advertise the
 * capability at all -- CAPABILITY_MISSING, never a guessed value.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { OPERATIONS, STATUS, validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { projectHyperlinksRead } = require('../../src/adapters/xlsx/families/hyperlinks-read.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');
const V18 = fs.readFileSync(path.join(DIR, 'rev18-candidate', 'bridge-block.js'), 'utf8');

function sheet(name, worksheet = {}) { return { GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, worksheet }; }
/** A fake hyperlink record, in the LIVE-measured shape (bbox.getAbsName / data.{Hyperlink,Location,Tooltip,
 * getHyperlinkType}); `extra` plants an UNNAMED sentinel field on `data`, to prove it never leaks. */
function link(ref, { address = null, location = null, tooltip = null, kind = 1, extra } = {}) {
  const data = { Hyperlink: address, Location: location, Tooltip: tooltip, getHyperlinkType: () => kind };
  if (extra !== undefined) data._sentinel = extra;
  return { bbox: { getAbsName: () => ref }, data };
}
const run = (api, hyperlinkSheets) => new Function('Api', readQuery([], { hyperlinkSheets }))(api).hyperlinkReads;

test('GH #300 L1: read_hyperlinks is an ACCEPTED xlsx read operation taking { sheet }, asking the driver for that sheet\'s hyperlinks only', () => {
  assert.strictEqual(OPERATIONS.xlsx.read_hyperlinks.status, STATUS.ACCEPTED);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_hyperlinks', sheet: 'Data' }).ok, true);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_hyperlinks' }).ok, false);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_hyperlinks', sheet: 'Data', range: 'A1' }).ok, false);
  assert.deepEqual(xlsxRequest('read_hyperlinks', { sheet: 'Data' }), { windows: [], hyperlinkSheets: ['Data'] });
});

test('a /16 runtime (the shipped block) does not advertise the capability at all -- UNOBSERVABLE:CAPABILITY_MISSING, never a guessed value', () => {
  const ws = sheet('Sheet1', { hyperlinkManager: { getAll: () => [link('$A$1', { address: 'x' })] } });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'Sheet1' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  const runtimeId = installBridge(api);
  assert.equal(runtimeId, '9.3.5|euro-office-bridge-v1/9.3.5.0/16', 'calibration: the shipped block really is /16');
  assert.deepEqual(run(api, ['Sheet1']), { Sheet1: 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('a /17 runtime (the protection-only candidate) does not advertise sheet.hyperlinks.read/v1 either -- CAPABILITY_MISSING', () => {
  const ws = sheet('Sheet1', { hyperlinkManager: { getAll: () => [link('$A$1', { address: 'x' })] } });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'Sheet1' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const m = api.GetEuroOfficeBridge().manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/17', 'calibration: the candidate block really is /17');
  assert.deepEqual(run(api, ['Sheet1']), { Sheet1: 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('a /18 runtime (the real, F1/F2-fixed candidate) reads every link, row-then-column order, through the SAME bridgeAxis convention', () => {
  const ws = sheet('T', { hyperlinkManager: { getAll: () => [
    link('$A$3', { address: 'a3-url', tooltip: 'tip', kind: 1 }),
    link('$C$1', { address: 'c1-url', kind: 2, extra: 'SENTINEL-HLINK' }),
    link('$A$1', { address: 'a1-url', location: 'Other!B2', kind: 7 }),
  ] } });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'T' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(api);
  const m = api.GetEuroOfficeBridge().manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/18', 'calibration: the candidate block really is /18');
  const out = run(api, ['T']);
  assert.deepEqual(out.T.links.map(l => l.ref), ['A1', 'C1', 'A3']);
  assert.deepEqual(out.T.links[0], { ref: 'A1', address: 'a1-url', location: 'Other!B2', tooltip: null, kind: 'unknown:7' });
  assert.deepEqual(out.T.links[1], { ref: 'C1', address: 'c1-url', location: null, tooltip: null, kind: 'internal' });
  assert.deepEqual(out.T.links[2], { ref: 'A3', address: 'a3-url', location: null, tooltip: 'tip', kind: 'external' });
  // the owner condition (never an unnamed field in the output), re-checked at THIS client's own call site
  assert.ok(!JSON.stringify(out).includes('SENTINEL-HLINK'), 'an unnamed hyperlink payload field leaked');
});

test('not asked -- hyperlinkReads is left off the return entirely, no bridge call at all (same convention as tableReads/commentReads)', () => {
  assert.doesNotMatch(readQuery(), /hyperlinkAsked/, 'the per-sheet loop (and its bridgeAxis call) is generated only inside the conditional block');
  const ws = sheet('T', { hyperlinkManager: { getAll: () => [link('$A$1', { address: 'x' })] } });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'T' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(api);
  assert.strictEqual(new Function('Api', readQuery())(api).hyperlinkReads, undefined);
});

test('"*" asks every sheet, the SAME all-sheets convention tableSheets/protectionSheets already use', () => {
  const a = sheet('A', { hyperlinkManager: { getAll: () => [] } });
  const b = sheet('B', { hyperlinkManager: { getAll: () => [link('$A$1', { address: 'x' })] } });
  const api = { GetSheets: () => [a, b], GetSheet: n => ({ A: a, B: b }[n] || null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(api);
  const out = run(api, ['*']);
  assert.deepEqual(Object.keys(out).sort(), ['A', 'B']);
  assert.deepEqual(out.A, { links: [] });
  assert.equal(out.B.links.length, 1);
});

test('a missing sheet names the gap, never a guessed empty object', () => {
  const api = { GetSheets: () => [], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  assert.deepEqual(run(api, ['NoSuchSheet']), { NoSuchSheet: 'UNOBSERVABLE:sheet-missing' });
});

const inv = (reads, sheets = ['T', 'Plain']) => ({ sheets: sheets.map(name => ({ name })), hyperlinkReads: reads });

test('the projection: a field-select over the already-shaped bridge row, nothing added, nothing dropped', () => {
  const reads = { T: { links: [{ ref: 'A1', address: 'x', location: null, tooltip: 'tip', kind: 'external' }] } };
  assert.deepEqual(projectHyperlinksRead(inv(reads), { sheet: 'T' }), { sheet: 'T', links: [{ ref: 'A1', address: 'x', location: null, tooltip: 'tip', kind: 'external' }] });
});

test('the projection: no link -> []; a named axis gap -> links null + the reason; not read -> NAMED; no such sheet -> SHEET_NOT_FOUND', () => {
  assert.deepEqual(projectHyperlinksRead(inv({ Plain: { links: [] } }), { sheet: 'Plain' }), { sheet: 'Plain', links: [] });
  assert.deepEqual(projectHyperlinksRead(inv({ T: 'UNOBSERVABLE:CAPABILITY_MISSING' }), { sheet: 'T' }), { sheet: 'T', links: null, reason: 'UNOBSERVABLE:CAPABILITY_MISSING' });
  assert.deepEqual(projectHyperlinksRead(inv(null), { sheet: 'T' }), { sheet: 'T', links: null, reason: 'UNOBSERVABLE:not-read' });
  assert.deepEqual(projectHyperlinksRead(inv({ T: { rows: [] } }), { sheet: 'T' }), { sheet: 'T', links: null, reason: 'UNOBSERVABLE:malformed-read' });
  assert.deepEqual(projectHyperlinksRead(inv({}), { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
});
