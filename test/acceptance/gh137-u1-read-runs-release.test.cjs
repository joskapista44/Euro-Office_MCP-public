'use strict';

/**
 * GH #137 U1: read_runs on the product READ surface (src/server/read-tools.cjs) -- the MIRROR of format_run
 * (src/adapters/xlsx/families/format-run.cjs), released BEFORE format_run's own WRITE (the #137 issue's own
 * stated order: "productize BOTH sides together -- do not release WRITE first and leave normal product READ
 * unable to inspect it"). The run-map reader (xlsx-driver.cjs readQuery's runCells axis, runReads keyed
 * sheet!cell) is EXISTING code, already exercised by format_run's own offline tests -- this file adds a new
 * product-surface CONSUMER of it, not a new reader.
 *
 * Offline only: the fake driver (test/fakes/fake-xlsx-driver.cjs) already simulates the run-map build
 * character-by-character, the SAME consecutive-identical-tuple merge the real driver does ('s
 * correction, format-run.cjs's own header comment).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const rt = require('../../src/server/read-tools.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

const LONG = 'x'.repeat(300);
const book = () => [
  { name: 'S', cells: {
    // three runs, document order, NOT merged across the gap even though run 1 and run 3 share the same font
    A1: { v: 'AaBbCc', runs: [{ start: 1, length: 2, font: { bold: true } }, { start: 3, length: 2, font: { bold: false } }, { start: 5, length: 2, font: { bold: true } }] },
    // two ADJACENT script segments with the IDENTICAL font tuple -> the character-by-character merge
    // produces ONE run spanning both, regardless of how the fixture happened to be segmented
    A2: { v: 'AdjacentRun', runs: [{ start: 1, length: 4, font: { bold: true, color: 255 } }, { start: 5, length: 7, font: { bold: true, color: 255 } }] },
    A3: { v: '' },
    A4: { v: 42 },
    A5: { v: LONG },
    A6: { v: 'Hi' }, // no per-run font at all -> the uniform RUN_FONT_DEFAULT, one run over the whole text
  } },
];

function harness() {
  const store = createStore(book());
  const b = fakeBroker('xlsx');
  const createDriver = () => createFakeXlsxDriver(store);
  const t = rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
  return t;
}

test('GH #137 U1: read_runs is ACCEPTED, offered by the schema, not a write intent, evidence present', () => {
  assert.ok(rt.OPERATIONS.xlsx.read_runs);
  assert.equal(rt.OPERATIONS.xlsx.read_runs.status, rt.STATUS.ACCEPTED);
  assert.ok(rt.OPERATIONS.xlsx.read_runs.evidence.length > 3);
  assert.ok(rt.inputSchema('xlsx').properties.operation.enum.includes('read_runs'));
  assert.equal(rt.writeIntents().has('read_runs'), false);
  assert.ok(rt.inputSchema('xlsx').properties.cell, 'the schema offers the cell argument');
});

test('GH #137 U1: validateReadRequest -- sheet+cell required, cell must be a single A1 cell, normalized like range', () => {
  const ok = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_runs', sheet: 'S', cell: '$a$1' });
  assert.equal(ok.ok, true, JSON.stringify(ok));
  assert.equal(ok.args.cell, 'A1', 'the $ is stripped and the letter uppercased, the SAME normalization as range');
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_runs', sheet: 'S' }).code, 'INVALID_ARGS', 'cell is required');
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_runs', cell: 'A1' }).code, 'INVALID_ARGS', 'sheet is required');
  const rangeNotCell = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_runs', sheet: 'S', cell: 'A1:B2' });
  assert.equal(rangeNotCell.ok, false);
  assert.equal(rangeNotCell.code, 'INVALID_ARGS', 'a multi-cell range is refused, never silently narrowed to one cell');
  const badCell = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_runs', sheet: 'S', cell: 'ZZZZ99999999' });
  assert.equal(badCell.code, 'INVALID_ARGS');
});

test('GH #137 U1: xlsxRequest builds runCells { sheet, cell } -- the SAME key format-run.cjs liveExtras() already uses', () => {
  assert.deepEqual(rt.xlsxRequest('read_runs', { sheet: 'S', cell: 'A1' }), { windows: [], runCells: [{ sheet: 'S', cell: 'A1' }] });
});

test('GH #137 U1: a multi-run cell -- three runs, document order, NOT merged across the gap', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A1' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.runMap.text, 'AaBbCc');
  assert.deepEqual(r.result.runMap.runs, [
    { start: 1, length: 2, font: { bold: true, italic: null, name: 'Calibri', size: 11, color: 0 } },
    { start: 3, length: 2, font: { bold: false, italic: null, name: 'Calibri', size: 11, color: 0 } },
    { start: 5, length: 2, font: { bold: true, italic: null, name: 'Calibri', size: 11, color: 0 } },
  ]);
});

test('GH #137 U1: two adjacent script segments with the identical font tuple merge into ONE run', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A2' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.runMap.text, 'AdjacentRun');
  assert.deepEqual(r.result.runMap.runs, [{ start: 1, length: 11, font: { bold: true, italic: null, name: 'Calibri', size: 11, color: 255 } }]);
});

test('GH #137 U1: an empty cell is named UNOBSERVABLE:empty-cell, never an empty run list', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A3' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.runMap, 'UNOBSERVABLE:empty-cell');
});

test('GH #137 U1: a non-text (number) cell is NOT a gap -- the stringified value reads as one uniform run', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A4' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.runMap.text, '42');
  assert.deepEqual(r.result.runMap.runs, [{ start: 1, length: 2, font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 } }]);
});

test('GH #137 U1: a cell over the RUN_MAP_CHAR_CEILING is named UNOBSERVABLE:text-too-long, never a guessed/truncated run list', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A5' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.runMap, 'UNOBSERVABLE:text-too-long');
});

test('GH #137 U1: a cell with no per-character font at all -- one uniform run, the font defaults', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A6' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.runMap, { text: 'Hi', runs: [{ start: 1, length: 2, font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 } }] });
});

test('GH #137 U1: a non-existent sheet is named UNOBSERVABLE:sheet-missing, never a guessed empty run map', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'NoSuchSheet', cell: 'A1' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.runMap, 'UNOBSERVABLE:sheet-missing');
});

test('GH #137 U1: the read is non-mutating -- never dispatches, never force-saves', async () => {
  const store = createStore(book());
  const b = fakeBroker('xlsx');
  let driver;
  const createDriver = () => { driver = createFakeXlsxDriver(store); return driver; };
  const t = rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
  const r = await t.handler({ fileId: FILE, operation: 'read_runs', sheet: 'S', cell: 'A1' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(store.durable.find(x => x.name === 'S').cells.A1.v, 'AaBbCc', 'the store itself is untouched');
});
