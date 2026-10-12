'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { BRIDGE_CAPABILITIES, BRIDGE_CLIENT_SRC } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
const { RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { withPatch, OTHER_PATCH } = require('../fakes/ds-bridge-runtime.cjs');

/**
 * Wave B READ (/4) -- range.alignmentExt.read/v1 (B05 READ), sheet.printSettings.read/v1 (B09), pivot.state.read/v1
 * (B13), OFFLINE (design docs/DS-BRIDGE-WAVE-B-READ-DESIGN-20261008.md). The ACTUAL shipped bridge-block.js text,
 * evaluated with `new Function('Api', source)` against a fake Api exposing only the measured shapes (the B11 test
 * pattern): positives, every fail-closed reason, whole-read refusal on one malformed part, no setter touched, the
 * manifest, and the /3-runtime refusal through the MCP-side client.
 */

const BLOCK_PATH = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'bridge-block.js');
const SOURCE = fs.readFileSync(BLOCK_PATH, 'utf8').replace(/,\s*$/, '');
const bridgeOf = api => new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(api);

// ---- fakes: every setter is a spy that must never be called by a READ ----
const writes = [];
const spy = name => () => { writes.push(name); };

function alignApi(al, { sheet = 'S' } = {}) {
  const ws = { GetRange: a1 => (a1 === 'ZZ0' ? null : { range: { getAlign: () => al, setIndent: spy('setIndent'), setShrinkToFit: spy('setShrinkToFit') } }) };
  return { GetSheet: n => (n === sheet ? ws : undefined) };
}
const al = over => ({ getReadingOrder: () => 2, getIndent: () => 3, getShrinkToFit: () => true, ...over });

function printModel(over = {}) {
  const ps = { asc_getFitToWidth: () => 1, asc_getFitToHeight: () => 1, asc_getScale: () => 90, asc_getWidth: () => 210, asc_getHeight: () => 297,
    useFirstPageNumber: true, firstPageNumber: 3, asc_setScale: spy('asc_setScale') };
  const txt = s => (s === null ? null : { str: s });
  const hf = { getOddHeader: () => txt('OH'), getOddFooter: () => txt('OF'), getFirstHeader: () => txt('FH'), getFirstFooter: () => null,
    getEvenHeader: () => txt('EH'), getEvenFooter: () => txt('EF'), getDifferentFirst: () => 1, getDifferentOddEven: () => false, setOddHeader: spy('setOddHeader') };
  return { PagePrintOptions: { asc_getPageSetup: () => ps, asc_getHorizontalCentered: () => true, asc_getVerticalCentered: () => null, asc_setHorizontalCentered: spy('hc') },
    headerFooter: hf, sheetPr: { FitToPage: 1 }, rowBreaks: { getBreaks: () => [{ id: 4, man: true, min: null, max: 16383, pt: null }] }, ...over };
}
const sheetApi = m => ({ GetSheet: n => (n === 'S' ? { worksheet: m } : undefined) });

function pivotApi(pivots, { throws = false } = {}) {
  const ws = { GetAllPivotTables: () => { if (throws) throw new Error('boom'); return pivots; } };
  return { GetSheet: n => (n === 'P' ? ws : undefined) };
}
const pivot = (name, ref) => ({ GetName: () => name, pivot: ref === undefined ? undefined : { asc_getDataRef: () => ref, asc_setDataRef: spy('asc_setDataRef') } });

// ---- range.alignmentExt.read/v1 ----
test('alignmentExt: the three keys exactly; a never-set readingOrder is the measured default 0', () => {
  const r = id => bridgeOf(alignApi(id)).read('range.alignmentExt.read/v1', { sheet: 'S', cell: 'A1' });
  assert.deepEqual(r(al()), { ok: true, value: { readingOrder: 2, indent: 3, shrink: true } });
  assert.deepEqual(r(al({ getReadingOrder: () => undefined })).value.readingOrder, 0);
  assert.deepEqual(r(al({ getReadingOrder: undefined })).value.readingOrder, 0, 'a missing getter reads undefined -> 0, as the driver did');
});

test('alignmentExt: every gap refuses the WHOLE read, never a default', () => {
  const r = (a, args = { sheet: 'S', cell: 'A1' }) => bridgeOf(alignApi(a)).read('range.alignmentExt.read/v1', args);
  assert.deepEqual(r(null), { ok: false, reason: 'getAlign-missing' });
  assert.deepEqual(r(al({ getReadingOrder: () => '1' })), { ok: false, reason: 'readingOrder-not-number' });
  assert.deepEqual(r(al({ getIndent: () => null })), { ok: false, reason: 'indent-missing' });
  assert.deepEqual(r(al({ getIndent: undefined })), { ok: false, reason: 'indent-missing' });
  assert.deepEqual(r(al({ getShrinkToFit: () => 0 })), { ok: false, reason: 'shrink-missing' });
  assert.deepEqual(r(al(), { sheet: 'S' }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(r(al(), { sheet: 'X', cell: 'A1' }), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(r(al(), { sheet: 'S', cell: 'ZZ0' }), { ok: false, reason: 'cell-not-found' });
  assert.deepEqual(r(al({ getIndent: () => { throw new Error('x'); } })), { ok: false, reason: 'internal-threw' });
});

// ---- sheet.printSettings.read/v1 ----
test('printSettings: the full field set with the driver normalisations', () => {
  const r = bridgeOf(sheetApi(printModel())).read('sheet.printSettings.read/v1', { sheet: 'S' });
  assert.deepEqual(r, { ok: true, value: { fitW: 1, fitH: 1, fitFlag: true, oddHeader: 'OH', rowBreaks: [{ id: 4, man: true, min: null, max: 16383 }],
    useFirstPage: true, firstPage: 3, hCenter: true, scale: 90, width: 210, height: 297, oddFooter: 'OF', vCenter: false,
    firstHeader: 'FH', firstFooter: null, evenHeader: 'EH', evenFooter: 'EF', diffFirst: true, diffOddEven: false } });
  assert.deepEqual(Object.keys(r.value), ['fitW', 'fitH', 'fitFlag', 'oddHeader', 'rowBreaks', 'useFirstPage', 'firstPage', 'hCenter', 'scale', 'width', 'height',
    'oddFooter', 'vCenter', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter', 'diffFirst', 'diffOddEven'], 'the driver projects by this order');
});

test('printSettings: the measured defaults (W8b) -- no sheetPr, null row breaks, null centering, no first page', () => {
  const m = printModel({ sheetPr: null, rowBreaks: null });
  m.PagePrintOptions.asc_getHorizontalCentered = () => null;
  const ps = m.PagePrintOptions.asc_getPageSetup(); ps.useFirstPageNumber = false; ps.firstPageNumber = null;
  const v = bridgeOf(sheetApi(m)).read('sheet.printSettings.read/v1', { sheet: 'S' }).value;
  assert.deepEqual([v.fitFlag, v.rowBreaks, v.hCenter, v.useFirstPage, v.firstPage], [false, [], false, false, null]);
});

test('printSettings: row-break and model gaps refuse the WHOLE read', () => {
  const r = m => bridgeOf(sheetApi(m)).read('sheet.printSettings.read/v1', { sheet: 'S' });
  const absent = printModel(); delete absent.rowBreaks;
  assert.deepEqual(r(absent), { ok: false, reason: 'rowBreaks-absent' });
  assert.deepEqual(r(printModel({ rowBreaks: {} })), { ok: false, reason: 'rowBreaks-unexpected-shape' });
  assert.deepEqual(r(printModel({ rowBreaks: { getBreaks: () => { throw new Error('x'); } } })), { ok: false, reason: 'rowBreaks-unexpected-shape' });
  assert.deepEqual(r(printModel({ rowBreaks: { getBreaks: () => [{ id: 1048576, man: true }] } })), { ok: false, reason: 'rowBreaks-element-invalid' }, 'past the last row');
  assert.deepEqual(r(printModel({ rowBreaks: { getBreaks: () => [{ id: 2, man: true, max: 16384 }] } })), { ok: false, reason: 'rowBreaks-element-invalid' }, 'a span past XFD');
  assert.deepEqual(r(printModel({ rowBreaks: { getBreaks: () => [{ id: 2, man: 1 }] } })), { ok: false, reason: 'rowBreaks-element-invalid' });
  assert.deepEqual(r(printModel({ headerFooter: { getOddHeader: () => null } })), { ok: false, reason: 'model-missing' }, 'a header/footer getter absent');
  assert.deepEqual(r(printModel({ PagePrintOptions: null })), { ok: false, reason: 'model-missing' });
  assert.deepEqual(r(null), { ok: false, reason: 'model-missing' });
  assert.deepEqual(bridgeOf(sheetApi(printModel())).read('sheet.printSettings.read/v1', { sheet: 'Q' }), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(bridgeOf(sheetApi(printModel())).read('sheet.printSettings.read/v1', {}), { ok: false, reason: 'bad-args' });
});

// ---- pivot.state.read/v1 ----
test('pivotState: the stored data reference by the pivot name; ambiguity and gaps refuse', () => {
  const r = (pivots, args = { sheet: 'P', pivot: 'PT' }, o) => bridgeOf(pivotApi(pivots, o)).read('pivot.state.read/v1', args);
  assert.deepEqual(r([pivot('X', 'Data!$A$1:$C$4'), pivot('PT', "'My Sheet'!$A$1:$B$9")]), { ok: true, value: { dataRef: "'My Sheet'!$A$1:$B$9" } });
  assert.deepEqual(r([pivot('X', 'a')]), { ok: false, reason: 'pivot-not-found' });
  assert.deepEqual(r([pivot('PT', 'a'), pivot('PT', 'b')]), { ok: false, reason: 'pivot-name-ambiguous' }, 'never picks one');
  assert.deepEqual(r([pivot('PT')]), { ok: false, reason: 'pivot-model-missing' });
  assert.deepEqual(r([pivot('PT', 7)]), { ok: false, reason: 'dataRef-not-string' });
  assert.deepEqual(r([], undefined, { throws: true }), { ok: false, reason: 'pivots-threw' });
  assert.deepEqual(r([{ GetName: () => { throw new Error('x'); } }]), { ok: false, reason: 'pivots-threw' });
  assert.deepEqual(r([], { sheet: 'P' }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(r([], { sheet: 'P', pivot: 3 }), { ok: false, reason: 'bad-args' }, 'the pivot is a NAME string');
  assert.deepEqual(r([], { sheet: 'Q', pivot: 'PT' }), { ok: false, reason: 'sheet-not-found' });
});

test('no Wave B READ calls a setter (READ only, mutation NO)', () => {
  writes.length = 0;
  bridgeOf(alignApi(al())).read('range.alignmentExt.read/v1', { sheet: 'S', cell: 'A1' });
  bridgeOf(sheetApi(printModel())).read('sheet.printSettings.read/v1', { sheet: 'S' });
  bridgeOf(pivotApi([pivot('PT', 'D!$A$1:$B$2')])).read('pivot.state.read/v1', { sheet: 'P', pivot: 'PT' });
  assert.deepEqual(writes, []);
});

test('manifest /5: the three Wave B ids stay READ / XLSX / mutation NO (/5 adds only the B15 hydration entry)', () => {
  const m = bridgeOf({ GetSheet: () => undefined }).manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.4.60/5');
  assert.equal(m.capabilities.length, 16);
  for (const id of ['range.alignmentExt.read/v1', 'sheet.printSettings.read/v1', 'pivot.state.read/v1']) {
    const c = m.capabilities.filter(x => x.id === id);
    assert.equal(c.length, 1, id);
    assert.deepEqual([c[0].version, c[0].kind, c[0].format, c[0].mutation], [1, 'READ', 'XLSX', 'NO'], id);
  }
  assert.deepEqual([BRIDGE_CAPABILITIES.ALIGN_EXT_READ.id, BRIDGE_CAPABILITIES.PRINT_SETTINGS_READ.id, BRIDGE_CAPABILITIES.PIVOT_STATE_READ.id],
    ['range.alignmentExt.read/v1', 'sheet.printSettings.read/v1', 'pivot.state.read/v1']);
});

test('the MCP client refuses the new capabilities on the /3 runtime and when a capability is missing -- no fallback', () => {
  const call = (api, cap) => new Function('Api', `${BRIDGE_CLIENT_SRC}; return __eoBridgeRead(${JSON.stringify(cap.id)}, ${cap.version}, ${JSON.stringify(RUNTIMES.xlsx)}, { sheet: 'P', pivot: 'PT' }, ${JSON.stringify(cap.valueType)});`)(api);
  const api = pivotApi([pivot('PT', 'D!$A$1:$B$2')]);
  bridgeOf(api); // installs the shipped block's Api.GetEuroOfficeBridge on this fake
  assert.deepEqual(call(api, BRIDGE_CAPABILITIES.PIVOT_STATE_READ), { ok: true, value: { dataRef: 'D!$A$1:$B$2' } }, 'POSITIVE CONTROL on the accepted runtime (/5)');
  assert.equal(OTHER_PATCH, 'euro-office-bridge-v1/9.3.4.60/4');
  assert.deepEqual(call(withPatch(api), BRIDGE_CAPABILITIES.PIVOT_STATE_READ), { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, 'the previous revision (/4) refuses');
  const noCap = { GetEuroOfficeBridge: () => { const b = bridgeOf(api); const m = b.manifest(); return { manifest: () => ({ ...m, capabilities: m.capabilities.filter(c => c.id !== 'pivot.state.read/v1') }), read: b.read }; } };
  assert.deepEqual(call(noCap, BRIDGE_CAPABILITIES.PIVOT_STATE_READ), { ok: false, reason: 'CAPABILITY_MISSING' });
});
