'use strict';

/**
 * GH #298 D (owner decision, path A, order A: read_tables before read_hyperlinks): read_tables -- the XLSX READ of
 * every Excel Table of one sheet, through the DS Bridge sheet.tables.read/v1, served only on patchVersion /18.
 * The driver's tableSheets read (xlsx-driver.cjs, the SAME per-sheet bridgeAxis convention protectionReads already
 * uses, run here against the REAL bridge-block code -- not an imitation), the projection (families/tables-read.cjs),
 * the request it makes and the catalogue entry. A /16 (the shipped block, test/fakes/ds-bridge-runtime.cjs) or /17
 * (the protection-only candidate) runtime does not advertise the capability at all -- CAPABILITY_MISSING, never a
 * guessed value, the SAME convention test/acceptance/gh326-protection-read-client-wiring.test.cjs already proves
 * for sheet.protection.read/v1.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { OPERATIONS, STATUS, validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { projectTablesRead, autoFilterCriteriaOf, CRITERIA_NOT_READ } = require('../../src/adapters/xlsx/families/tables-read.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');
const V18 = fs.readFileSync(path.join(DIR, 'rev18-candidate', 'bridge-block.js'), 'utf8');

/** The SAME fake worksheet shape the ds935-bridge18 acceptance test uses: GetName (the driver's own per-sheet
 * existence check needs it -- the bridge handler itself only reads `.worksheet`). */
function sheet(name, worksheet = {}) { return { GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, worksheet }; }
function table({ name = 'Table1', ref = '$B$10:$E$13', header = true, totals = false, columns = ['Region', 'Q1'],
  style = { Name: 'TableStyleLight9', ShowRowStripes: true, ShowColumnStripes: false, ShowFirstColumn: false, ShowLastColumn: false },
  autoFilter = {}, extra } = {}) {
  const t = { DisplayName: name, Ref: { getAbsName: () => ref }, isHeaderRow: () => header, isTotalsRow: () => totals,
    TableColumns: columns.map(c => ({ Name: c })), TableStyleInfo: style, AutoFilter: autoFilter };
  if (extra !== undefined) t._sentinel = extra;
  return t;
}
const run = (api, tableSheets) => new Function('Api', readQuery([], { tableSheets }))(api).tableReads;

test('GH #298 D: read_tables is an ACCEPTED xlsx read operation taking { sheet }, asking the driver for that sheet\'s tables only', () => {
  assert.strictEqual(OPERATIONS.xlsx.read_tables.status, STATUS.ACCEPTED);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_tables', sheet: 'Data' }).ok, true);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_tables' }).ok, false);
  assert.strictEqual(validateReadRequest('xlsx', { fileId: '1', operation: 'read_tables', sheet: 'Data', range: 'A1' }).ok, false);
  assert.deepEqual(xlsxRequest('read_tables', { sheet: 'Data' }), { windows: [], tableSheets: ['Data'] });
});

test('a /16 runtime (the shipped block) does not advertise the capability at all -- UNOBSERVABLE:CAPABILITY_MISSING, never a guessed value', () => {
  const ws = sheet('Sheet1', { TableParts: [table()] });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'Sheet1' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  const runtimeId = installBridge(api);
  assert.equal(runtimeId, '9.3.5|euro-office-bridge-v1/9.3.5.0/16', 'calibration: the shipped block really is /16');
  assert.deepEqual(run(api, ['Sheet1']), { Sheet1: 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('a /17 runtime (the protection-only candidate) does not advertise sheet.tables.read/v1 either -- CAPABILITY_MISSING', () => {
  const ws = sheet('Sheet1', { TableParts: [table()] });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'Sheet1' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const m = api.GetEuroOfficeBridge().manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/17', 'calibration: the candidate block really is /17');
  assert.deepEqual(run(api, ['Sheet1']), { Sheet1: 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('a /18 runtime (the real, F1/F2-fixed candidate) reads every table through the SAME bridgeAxis convention', () => {
  const ws = sheet('T', { TableParts: [
    table({ name: 'Table1', ref: '$B$10:$E$13', columns: ['Region', 'Q1', 'Q2', 'Q3'], autoFilter: {}, extra: 'SENTINEL-TABLE' }),
    table({ name: 'Table2', ref: '$B$16:$C$18', columns: ['Name', 'Score'], style: null, autoFilter: null }),
  ] });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'T' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(api);
  const m = api.GetEuroOfficeBridge().manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/18', 'calibration: the candidate block really is /18');
  const out = run(api, ['T']);
  assert.equal(out.T.tables.length, 2);
  // GH #361 route probe (c): the driver ALSO tries the (candidate, not-installed) sheet.table.autoFilter.
  // criteria.read/v1 for every table with autoFilter:true -- on a /18 runtime (no such capability yet)
  // it degrades to the SAME CAPABILITY_MISSING convention every other not-yet-installed axis already uses.
  assert.deepEqual(out.T.tables[0], { name: 'Table1', ref: 'B10:E13', headerRow: true, totalsRow: false, columns: ['Region', 'Q1', 'Q2', 'Q3'], style: { name: 'TableStyleLight9', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false }, autoFilter: true, autoFilterDetail: 'UNOBSERVABLE:CAPABILITY_MISSING' });
  assert.deepEqual(out.T.tables[1], { name: 'Table2', ref: 'B16:C18', headerRow: true, totalsRow: false, columns: ['Name', 'Score'], style: null, autoFilter: false });
  // the owner condition (never an unnamed field in the output), re-checked at THIS client's own call site
  assert.ok(!JSON.stringify(out).includes('SENTINEL-TABLE'), 'an unnamed table field leaked');
});

test('not asked -- tableReads is left off the return entirely, no bridge call at all (same convention as commentReads)', () => {
  assert.doesNotMatch(readQuery(), /tableAsked/, 'the per-sheet loop (and its bridgeAxis call) is generated only inside the conditional block');
  const ws = sheet('T', { TableParts: [table()] });
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'T' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(api);
  assert.strictEqual(new Function('Api', readQuery())(api).tableReads, undefined);
});

test('"*" asks every sheet, the SAME all-sheets convention protectionSheets/read_charts already use', () => {
  const a = sheet('A', { TableParts: [] });
  const b = sheet('B', { TableParts: [table()] });
  const api = { GetSheets: () => [a, b], GetSheet: n => ({ A: a, B: b }[n] || null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(api);
  const out = run(api, ['*']);
  assert.deepEqual(Object.keys(out).sort(), ['A', 'B']);
  assert.deepEqual(out.A, { tables: [] });
  assert.equal(out.B.tables.length, 1);
});

test('a missing sheet names the gap, never a guessed empty object', () => {
  const api = { GetSheets: () => [], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  assert.deepEqual(run(api, ['NoSuchSheet']), { NoSuchSheet: 'UNOBSERVABLE:sheet-missing' });
});

const inv = (reads, sheets = ['T', 'Plain']) => ({ sheets: sheets.map(name => ({ name })), tableReads: reads });

test('the projection: a field-select over the already-shaped bridge row, nothing dropped -- PLUS GH #361\'s own named gap when autoFilter is true', () => {
  const reads = { T: { tables: [
    { name: 'Table1', ref: 'B10:E13', headerRow: true, totalsRow: false, columns: ['Region', 'Q1'], style: { name: 'TableStyleLight9', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false }, autoFilter: true },
  ] } };
  assert.deepEqual(projectTablesRead(inv(reads), { sheet: 'T' }), { sheet: 'T', tables: [
    { name: 'Table1', ref: 'B10:E13', headerRow: true, totalsRow: false, columns: ['Region', 'Q1'], style: { name: 'TableStyleLight9', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false }, autoFilter: true, autoFilterCriteria: 'UNOBSERVABLE:table-autofilter-criteria-not-read' },
  ] });
});

// GH #361 2. egyseg (card an internal branch): the projection names
// what it does NOT see -- the presence-only `autoFilter` flag never claims to be the criteria themselves.
test('GH #361: autoFilterCriteria is a named gap ONLY when autoFilter is true -- false/undefined carries no such field (nothing to name a gap for)', () => {
  const reads = { T: { tables: [
    { name: 'NoFilter', ref: 'A1:B2', headerRow: true, totalsRow: false, columns: ['A', 'B'], style: null, autoFilter: false },
    { name: 'Untouched', ref: 'D1:E2', headerRow: true, totalsRow: false, columns: ['D', 'E'], style: null }, // autoFilter key absent entirely
  ] } };
  const out = projectTablesRead(inv(reads), { sheet: 'T' });
  assert.ok(!('autoFilterCriteria' in out.tables[0]), JSON.stringify(out.tables[0]));
  assert.ok(!('autoFilterCriteria' in out.tables[1]), JSON.stringify(out.tables[1]));
});

// GH #361 route probe, route (c): autoFilterDetail (the candidate sheet.table.autoFilter.criteria.read/v1's
// raw per-table result, wired in xlsx-driver.cjs) upgrades autoFilterCriteria ONLY when it is a clean
// { present: true, range: <string>, fields: Array } -- anything else (missing, a string gap, present:false, a
// non-string range, a non-array fields) falls back to the EXACT SAME named string the field has always
// carried. "Byte-identical today" is the card's own requirement (the capability is not installed anywhere),
// proven by the first case below. When fields IS an array, each entry is projected field-by-field (see
// fieldOf() in the family module) to the candidate handler's own MEASURED per-column shape.
test('autoFilterCriteriaOf(): today (capability not installed) is BYTE-IDENTICAL to the pre-route-(c) string', () => {
  assert.equal(autoFilterCriteriaOf(undefined), CRITERIA_NOT_READ, 'no autoFilterDetail at all (today\'s shape)');
  assert.equal(autoFilterCriteriaOf('UNOBSERVABLE:CAPABILITY_MISSING'), CRITERIA_NOT_READ);
  assert.equal(autoFilterCriteriaOf('UNOBSERVABLE:sheet-missing'), CRITERIA_NOT_READ);
});

// a maintainer's own LIVE run of the candidate handler against a real, OOXML-preseeded fixture with a table-owned
// AutoFilter (one value filter, one custom filter) -- the handler's measured output, used verbatim below.
const MEASURED_DETAIL = { present: true, range: 'A1:C4', fields: [
  { colId: 0, values: ['b'] },
  { colId: 1, custom: { and: false, conditions: [{ operator: 2, operatorName: null, value: '1' }] } },
] };

test('autoFilterCriteriaOf(): a clean { present: true, range, fields } upgrades to the MEASURED field shape, field-by-field (not a blind spread)', () => {
  assert.deepEqual(autoFilterCriteriaOf(MEASURED_DETAIL), {
    range: 'A1:C4',
    fields: [
      { colId: 0, values: ['b'] },
      { colId: 1, custom: { and: false, conditions: [{ operator: 2, operatorName: null, value: '1' }] } },
    ],
  });
});

test('autoFilterCriteriaOf(): a field carrying a named unverified gap passes it through unchanged -- never dropped, never guessed', () => {
  const detail = { present: true, range: 'A1:E4', fields: [
    { colId: 0, values: ['x'], unverified: ['internal-shape-unverified:blank', 'internal-shape-unverified:dates'] },
  ] };
  assert.deepEqual(autoFilterCriteriaOf(detail), {
    range: 'A1:E4',
    fields: [{ colId: 0, values: ['x'], unverified: ['internal-shape-unverified:blank', 'internal-shape-unverified:dates'] }],
  });
});

test('autoFilterCriteriaOf(): an empty FilterColumns array survives as an empty fields array, not a gap', () => {
  assert.deepEqual(autoFilterCriteriaOf({ present: true, range: 'A1:C4', fields: [] }), { range: 'A1:C4', fields: [] });
});

test('autoFilterCriteriaOf(): present:false, a non-string range, a non-array fields, or a malformed object all fall back to the named string -- never guessed, never crashed', () => {
  assert.equal(autoFilterCriteriaOf({ present: false, range: null, reason: null }), CRITERIA_NOT_READ, 'flag mismatch with the outer autoFilter:true');
  assert.equal(autoFilterCriteriaOf({ present: true, range: null, fields: null, reason: 'internal-shape-unverified:ref' }), CRITERIA_NOT_READ, 'range itself unreadable');
  assert.equal(autoFilterCriteriaOf({ present: true, range: 42 }), CRITERIA_NOT_READ, 'range not a string');
  assert.equal(
    autoFilterCriteriaOf({ present: true, range: 'A1:C4', reason: 'internal-shape-unverified:fields' }),
    'UNOBSERVABLE:internal-shape-unverified:fields',
    'a string range with no fields array (the handler never produces this shape today) still falls back to its own named reason',
  );
  assert.equal(autoFilterCriteriaOf({ present: true, range: 'A1:C4' }), CRITERIA_NOT_READ, 'no fields array and no reason either');
  assert.equal(autoFilterCriteriaOf(null), CRITERIA_NOT_READ);
  assert.equal(autoFilterCriteriaOf(42), CRITERIA_NOT_READ);
});

test('projectTablesRead(): end to end, a table whose autoFilterDetail is the installed-candidate MEASURED shape gets the full field-by-field autoFilterCriteria object', () => {
  const reads = { T: { tables: [
    { name: 'Table1', ref: 'A1:C4', headerRow: true, totalsRow: false, columns: ['Item', 'Qty', 'Note'], style: null, autoFilter: true, autoFilterDetail: MEASURED_DETAIL },
  ] } };
  const out = projectTablesRead(inv(reads), { sheet: 'T' });
  assert.deepEqual(out.tables[0].autoFilterCriteria, {
    range: 'A1:C4',
    fields: [
      { colId: 0, values: ['b'] },
      { colId: 1, custom: { and: false, conditions: [{ operator: 2, operatorName: null, value: '1' }] } },
    ],
  });
});

test('the projection: no table -> []; a named axis gap -> tables null + the reason; not read -> NAMED; no such sheet -> SHEET_NOT_FOUND', () => {
  assert.deepEqual(projectTablesRead(inv({ Plain: { tables: [] } }), { sheet: 'Plain' }), { sheet: 'Plain', tables: [] });
  assert.deepEqual(projectTablesRead(inv({ T: 'UNOBSERVABLE:CAPABILITY_MISSING' }), { sheet: 'T' }), { sheet: 'T', tables: null, reason: 'UNOBSERVABLE:CAPABILITY_MISSING' });
  assert.deepEqual(projectTablesRead(inv(null), { sheet: 'T' }), { sheet: 'T', tables: null, reason: 'UNOBSERVABLE:not-read' });
  assert.deepEqual(projectTablesRead(inv({ T: { rows: [] } }), { sheet: 'T' }), { sheet: 'T', tables: null, reason: 'UNOBSERVABLE:malformed-read' });
  assert.deepEqual(projectTablesRead(inv({}), { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
});
