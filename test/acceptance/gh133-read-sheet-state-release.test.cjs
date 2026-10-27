'use strict';

/**
 * GH #133: read_sheet_state on the product READ surface (src/server/read-tools.cjs) -- every sheet's visibility and the
 * workbook's active sheet. No new driver reader: the visibility axis (xlsx-driver.cjs readQuery sheetVisibility, the public
 * ApiWorksheet.GetVisible()) and the active-sheet name (the panes block's Api.GetActiveSheet().GetName()) already exist and
 * ride the tiered path (TIERED_EXTRAS); this operation is a new product-surface CONSUMER of them.
 *
 * Offline only: the fake driver already serves `visibility` (a sheet with hidden: true reads false) and `active` (store.active,
 * default the first sheet).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const rt = require('../../src/server/read-tools.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');
const { projectSheetStateRead } = require('../../src/adapters/xlsx/families/sheet-state-read.cjs');

const book = () => [
  { name: 'Alpha', cells: { A1: { v: 1 } } },
  { name: 'Hidden one', hidden: true, cells: { A1: { v: 2 } } },
  { name: 'Gamma', cells: { A1: { v: 3 } } },
];

function harness({ active, script = {} } = {}) {
  const store = createStore(book());
  if (active !== undefined) store.active = active;
  const b = fakeBroker('xlsx');
  const createDriver = () => createFakeXlsxDriver(store, script);
  return rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
}

test('GH #133: read_sheet_state is ACCEPTED, offered by the schema, takes no arguments, is not a write intent', () => {
  assert.ok(rt.OPERATIONS.xlsx.read_sheet_state);
  assert.equal(rt.OPERATIONS.xlsx.read_sheet_state.status, rt.STATUS.ACCEPTED);
  assert.ok(rt.OPERATIONS.xlsx.read_sheet_state.evidence.length > 3);
  assert.ok(rt.inputSchema('xlsx').properties.operation.enum.includes('read_sheet_state'));
  assert.equal(rt.writeIntents().has('read_sheet_state'), false);
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_sheet_state' }).ok, true);
  const withSheet = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_sheet_state', sheet: 'Alpha' });
  assert.equal(withSheet.ok, false, 'the operation reads the whole workbook: a sheet argument is refused, not silently ignored');
  assert.equal(withSheet.code, 'UNKNOWN_FIELD');
});

test('GH #133: xlsxRequest asks for the existing visibility axis and the panes block that carries the active sheet', () => {
  assert.deepEqual(rt.xlsxRequest('read_sheet_state', {}), { windows: [], panes: true, sheetVisibility: true });
});

test('GH #133: every sheet in order with its visibility, the hidden one NAMED hidden, the active sheet is the one the editor reports', async () => {
  const t = harness({ active: 'Gamma' });
  const r = await t.handler({ fileId: FILE, operation: 'read_sheet_state' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result, {
    sheets: [
      { name: 'Alpha', index: 0, visibility: 'visible' },
      { name: 'Hidden one', index: 1, visibility: 'hidden' },
      { name: 'Gamma', index: 2, visibility: 'visible' },
    ],
    active: 'Gamma',
  });
});

test('GH #133: with no saved active tab the first sheet is the active one (what the editor reports); a hidden active sheet is read as it is, never corrected', async () => {
  const first = await harness().handler({ fileId: FILE, operation: 'read_sheet_state' });
  assert.equal(first.result.active, 'Alpha');
  const hid = await harness({ active: 'Hidden one' }).handler({ fileId: FILE, operation: 'read_sheet_state' });
  assert.equal(hid.result.active, 'Hidden one');
  assert.equal(hid.result.sheets.find(s => s.name === 'Hidden one').visibility, 'hidden');
});

test('GH #133: an unreadable visibility is a NAMED reason on every sheet, never "visible"; the active sheet is still reported', async () => {
  const r = await harness({ script: { visibilityUnreadable: true } }).handler({ fileId: FILE, operation: 'read_sheet_state' });
  assert.equal(r.ok, true, JSON.stringify(r));
  for (const s of r.result.sheets) {
    assert.equal(s.visibility, null);
    assert.equal(s.reason, 'UNOBSERVABLE:scripted');
  }
  assert.equal(r.result.active, 'Alpha');
});

test('GH #133: the projection never invents a state -- a missing visibility axis, a gap string or a non-boolean is named; the active gap is its own reason, not a sheet name', () => {
  const sheets = [{ name: 'A', index: 0 }, { name: 'B', index: 1 }, { name: 'C', index: 2 }, { name: 'D', index: 3 }];
  const out = projectSheetStateRead({ sheets, visibility: { A: true, B: false, C: 'UNOBSERVABLE:visibility-not-boolean' }, active: 'UNOBSERVABLE:threw' });
  assert.deepEqual(out.sheets, [
    { name: 'A', index: 0, visibility: 'visible' },
    { name: 'B', index: 1, visibility: 'hidden' },
    { name: 'C', index: 2, visibility: null, reason: 'UNOBSERVABLE:visibility-not-boolean' },
    { name: 'D', index: 3, visibility: null, reason: 'UNOBSERVABLE:not-read' },
  ]);
  assert.equal(out.active, null);
  assert.equal(out.activeReason, 'UNOBSERVABLE:threw', 'a gap string must not pass for a sheet name');
  const noVis = projectSheetStateRead({ sheets, visibility: null, active: null });
  assert.ok(noVis.sheets.every(s => s.visibility === null && s.reason === 'UNOBSERVABLE:not-read'));
  assert.equal(noVis.activeReason, 'UNOBSERVABLE:not-read');
  const stray = projectSheetStateRead({ sheets, visibility: { A: true }, active: 'No such sheet' });
  assert.equal(stray.active, null);
  assert.equal(stray.activeReason, 'UNOBSERVABLE:active-sheet-not-in-workbook', 'a name that is not one of the workbook sheets is never reported as the active sheet');
});

test('GH #133: read-only -- no dispatch, no force-save, the stored workbook unchanged', async () => {
  const store = createStore(book());
  const b = fakeBroker('xlsx');
  let driver;
  const createDriver = () => { driver = createFakeXlsxDriver(store); return driver; };
  const t = rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
  const before = JSON.stringify(store.durable);
  const r = await t.handler({ fileId: FILE, operation: 'read_sheet_state' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(JSON.stringify(store.durable), before);
  assert.equal(store.saves, 0);
});
