'use strict';

/**
 * GH #356 / #357: read_print_rendering -- the rendered print page count of one sheet, through the DS Bridge /23
 * sheet.printPageCount.read/v1. The driver's printPageSheets read (xlsx-driver.cjs, the SAME per-sheet bridgeAxis convention
 * tableReads uses), run here against the REAL /22 and /23 bridge blocks over a fake engine in the served 9.3.5 shapes; the
 * projection (families/print-rendering-read.cjs) and the catalogue entry. ACCEPTED after the /23 engine gate G1 LIVE PASS
 * (docs/evidence/gh356-gh357-bridge23-live-20261101/g1-r1.json) -- this file proves the wiring, the gate proved the engine.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { OPERATIONS, STATUS, validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { projectPrintRenderingRead, AUTHORITY } = require('../../src/adapters/xlsx/families/print-rendering-read.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const blockOf = rev => fs.readFileSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js'), 'utf8');

/** a fake engine: one sheet `name`, a fit-to-page layout of `pages`; the engine writes the fit scale to the CURRENT options */
function engine({ name = 'P', pages = 1, hidden = false } = {}) {
  const H = { TurnOffHistory: 0, Points: 0, TurnOff() { this.TurnOffHistory += 1; }, TurnOn() { this.TurnOffHistory -= 1; }, Is_On() { return this.TurnOffHistory === 0; } };
  const opts = () => { const o = { scale: 100 }; o.asc_getPageSetup = () => ({ asc_getScale: () => o.scale, asc_getFitToWidth: () => 1, asc_getFitToHeight: () => 1 }); o.clone = () => opts(); return o; };
  const m = { getIndex: () => 0, getHidden: () => hidden, PagePrintOptions: opts() };
  const real = m.PagePrintOptions;
  const wb = { z: 1, getZoom() { return this.z; }, changeZoom(z) { this.z = z; },
    calcPagesPrint() { if (H.Is_On()) H.Points += 1; m.PagePrintOptions.scale = 40; return { arrPages: new Array(pages).fill(0) }; } };
  function asc_CAdjustPrint() { } asc_CAdjustPrint.prototype.asc_setPrintType = function () { }; asc_CAdjustPrint.prototype.asc_setActiveSheetsArray = function () { };
  const ws = { GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, worksheet: m };
  const api = { GetSheets: () => [ws], GetSheet: n => (n === name ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  const window = { Asc: { editor: { wb }, asc_CAdjustPrint, c_oAscPrintType: { ActiveSheets: 0 } }, AscCommonExcel: { c_kMaxPrintPages: 1500 } };
  return { api, window, AscCommon: { History: H }, H, m, real };
}
const install = (e, rev) => new Function('Api', 'Asc', 'window', 'AscCommon', blockOf(rev).replace(/\}\s*,\s*$/, '}'))(e.api, {}, e.window, e.AscCommon);
const run = (e, printPageSheets) => new Function('Api', readQuery([], { printPageSheets }))(e.api).printPageReads;

test('read_print_rendering is ACCEPTED (after the /23 G1 LIVE PASS), takes { sheet } only, asks the driver for that sheet only', () => {
  assert.strictEqual(OPERATIONS.xlsx.read_print_rendering.status, STATUS.ACCEPTED);
  assert.deepEqual(xlsxRequest('read_print_rendering', { sheet: 'P' }), { windows: [], printPageSheets: ['P'] });
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_rendering', sheet: 'P' }).ok, true);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_rendering' }).ok, false);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_rendering', sheet: 'P', range: 'A1' }).ok, false);
});

test('a /22 runtime does not advertise the read: UNOBSERVABLE:CAPABILITY_MISSING, never a number', () => {
  const e = engine();
  install(e, 22);
  assert.deepEqual(run(e, ['P']), { P: 'UNOBSERVABLE:CAPABILITY_MISSING' });
  assert.equal(e.H.Points, 0);
});

test('a /23 runtime: the count through the real handler, the live options untouched (scale, identity, no History point)', () => {
  const e = engine({ pages: 9 });
  install(e, 23);
  assert.deepEqual(run(e, ['P']), { P: { pages: 9 } });
  assert.equal(e.m.PagePrintOptions, e.real);
  assert.equal(e.real.scale, 100);
  assert.equal(e.H.Points, 0);
  assert.equal(e.H.TurnOffHistory, 0);
});

test('a /23 runtime: a hidden sheet is the bridge\'s own named reason; an unknown sheet is sheet-missing', () => {
  const e = engine({ hidden: true });
  install(e, 23);
  assert.deepEqual(run(e, ['P']), { P: 'UNOBSERVABLE:BRIDGE:sheet-hidden' });
  assert.deepEqual(run(e, ['NOPE']), { NOPE: 'UNOBSERVABLE:sheet-missing' });
});

test('readQuery emits the axis ONLY when asked (the existing query shapes stay byte-identical)', () => {
  assert.equal(readQuery([], {}).includes('printPageReads'), false);
  assert.equal(readQuery([], { printPageSheets: ['P'] }).includes('sheet.printPageCount.read/v1'), true);
});

test('projection: pages and reason exclude each other; capped is only the cap; never a 0 for a gap', () => {
  const inv = r => ({ sheets: [{ name: 'P' }], printPageReads: r === undefined ? undefined : { P: r } });
  assert.deepEqual(projectPrintRenderingRead(inv({ pages: 3 }), { sheet: 'P' }), { sheet: 'P', pages: 3, authority: AUTHORITY, capped: false, scope: 'sheet' });
  assert.deepEqual(projectPrintRenderingRead(inv({ pages: 0 }), { sheet: 'P' }).pages, 0, '0 is a legitimate count');
  assert.deepEqual(projectPrintRenderingRead(inv('UNOBSERVABLE:BRIDGE:page-count-capped'), { sheet: 'P' }), { sheet: 'P', pages: null, authority: 'none', capped: true, scope: 'sheet', reason: 'UNOBSERVABLE:BRIDGE:page-count-capped' });
  for (const [r, want] of [['UNOBSERVABLE:BRIDGE:sheet-hidden', 'UNOBSERVABLE:BRIDGE:sheet-hidden'], ['UNOBSERVABLE:CAPABILITY_MISSING', 'UNOBSERVABLE:CAPABILITY_MISSING'],
    [undefined, 'UNOBSERVABLE:not-read'], [{ pages: -1 }, 'UNOBSERVABLE:malformed-read'], [{ pages: 1.5 }, 'UNOBSERVABLE:malformed-read'], [{}, 'UNOBSERVABLE:malformed-read']]) {
    const p = projectPrintRenderingRead(inv(r), { sheet: 'P' });
    assert.deepEqual([p.pages, p.capped, p.authority, p.reason], [null, false, 'none', want], JSON.stringify(r));
  }
  assert.deepEqual(projectPrintRenderingRead(inv({ pages: 1 }), { sheet: 'X' }), { sheet: 'X', error: 'SHEET_NOT_FOUND' });
});
