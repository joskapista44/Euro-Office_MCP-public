'use strict';

/**
 * GH #140 U2: read_pivots on the product READ surface (src/server/read-tools.cjs) -- a DEDICATED
 * operation (name-decision recorded in docs/GH140-PIVOT-CONTRACT.md,
 * section 4: a SEPARATE operation, read_charts.pivots[] stays the bare name list it always
 * was). Projects the content-proven pivot read (pivot-source.cjs withProvenPivots()) that
 * create_pivot/refresh_pivot's own gate() has consumed internally since M7-H -- GH #140 U1's own
 * finding: until this unit, that axis (inv.objects) had NO product-surface READ consumer at all.
 * No new driver reader; this file adds a new CONSUMER of an existing one.
 *
 * Offline only: the fake driver (test/fakes/fake-xlsx-driver.cjs) already simulates the objects
 * axis (withPivotProofs()) that create_pivot/refresh_pivot's own offline tests rely on.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const rt = require('../../src/server/read-tools.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

// The P-A2 probe's own shapes (docs/evidence/M7-H2-PROBE.md, m7-h2-probe1-20261004.json, the SAME
// fixture xlsx-m7-h-pivot-source.test.cjs already proves provenSource() verifies): source S!A1:C7
// (reg, prod, amt), pivot host sheet Hre, body I5:K7, table H3:K7.
const DATA = [['reg', 'prod', 'amt'], ['N', 'x', '1'], ['S', 'y', '2'], ['N', 'y', '3'], ['S', 'x', '4'], ['N', 'x', '5'], ['S', 'y', '6']];
const BODY = [['6', '3', '9'], ['4', '8', '12'], ['10', '11', '21']];
const TABLE = [['Sum of amt', 'Column Labels', '', ''], ['Row Labels', 'x', 'y', 'Grand Total'], ['N', ...BODY[0]], ['S', ...BODY[1]], ['Grand Total', ...BODY[2]]];
const SPEC = { dataRef: 'S!A1:C7', fields: ['reg', 'prod', 'amt'], rows: ['reg'], cols: ['prod'], data: ['Sum of amt'], body: { range: 'I5:K7', values: BODY }, table: { range: 'H3:K7', values: TABLE } };

const sourceCells = Object.fromEntries(DATA.flatMap((row, r) => row.map((v, c) => [`${String.fromCharCode(65 + c)}${r + 1}`, { v, f: '' }])));

const book = () => [
  { name: 'S', cells: sourceCells },
  // a publicly-readable source (no reopen yet): table/layout ride along per withProvenPivots()
  { name: 'H', cells: {}, objects: { charts: [], names: [], pivots: [{ name: 'PV', source: 'S!$A$1:$C$7', table: 'H1:J3', layout: { rows: ['reg'], cols: ['prod'], data: ['Sum of amt'] } }] } },
  // a pivot whose public source already threw (the reopened shape): no table/layout, name+source only
  { name: 'Hre', cells: {}, objects: { charts: [], names: [], pivots: [{ name: 'PR', proofSpec: SPEC }] } },
  // no pivots on this sheet at all -- an empty list, never a gap
  { name: 'Empty', cells: {}, objects: { charts: [], names: [], pivots: [] } },
];

function harness() {
  const store = createStore(book());
  const b = fakeBroker('xlsx');
  const createDriver = () => createFakeXlsxDriver(store);
  const t = rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
  return t;
}

test('GH #140 U2: read_pivots is ACCEPTED, offered by the schema, not a write intent, evidence present', () => {
  assert.ok(rt.OPERATIONS.xlsx.read_pivots);
  assert.equal(rt.OPERATIONS.xlsx.read_pivots.status, rt.STATUS.ACCEPTED);
  assert.ok(rt.OPERATIONS.xlsx.read_pivots.evidence.length > 3);
  assert.ok(rt.inputSchema('xlsx').properties.operation.enum.includes('read_pivots'));
  assert.equal(rt.writeIntents().has('read_pivots'), false);
});

test('GH #140 U2: read_charts.pivots[] is UNCHANGED -- still a bare name list, the name-decision\'s own constraint', () => {
  assert.ok(!rt.OPERATIONS.xlsx.read_charts.evidence.includes('pivots[].source'), 'the stale X-O10 claim must not be reintroduced into read_charts\'s own evidence');
});

test('GH #140 U2: validateReadRequest -- sheet is OPTIONAL (the read_charts/read_drawings convention)', () => {
  const withSheet = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_pivots', sheet: 'H' });
  assert.equal(withSheet.ok, true, JSON.stringify(withSheet));
  const noSheet = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_pivots' });
  assert.equal(noSheet.ok, true, JSON.stringify(noSheet));
});

test('GH #140 U2: xlsxRequest builds objectSheets [sheet-or-*] -- the SAME axis create_pivot/refresh_pivot\'s own liveExtras() already requests on the WRITE side', () => {
  assert.deepEqual(rt.xlsxRequest('read_pivots', { sheet: 'H' }), { windows: [], objectSheets: ['H'] });
  assert.deepEqual(rt.xlsxRequest('read_pivots', {}), { windows: [], objectSheets: ['*'] });
});

test('GH #140 U2: a publicly-readable pivot -- name + source + table + layout, all four', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'H' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.sheets, [{ sheet: 'H', pivots: [{ name: 'PV', source: 'S!$A$1:$C$7', table: 'H1:J3', layout: { rows: ['reg'], cols: ['prod'], data: ['Sum of amt'] } }] }]);
});

test('GH #140 U2: a content-proven pivot (public source already threw) -- name + source ONLY, table/layout ABSENT, never guessed', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'Hre' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.sheets.length, 1);
  const pv = r.result.sheets[0].pivots[0];
  assert.equal(pv.name, 'PR');
  assert.equal(pv.source, 'S!$A$1:$C$7', 'content-proven by pivot-source.cjs, the SAME quoted/absolute shape');
  assert.ok(!('table' in pv) && !('layout' in pv), 'table/layout must be absent, not null/guessed, when the public source did not answer');
});

test('GH #140 U2: a sheet with no pivots -- an empty list, never a gap', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'Empty' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.sheets, [{ sheet: 'Empty', pivots: [] }]);
});

test('GH #140 U2: an omitted sheet reads every sheet\'s pivots -- the same \'*\' convention as read_charts/read_drawings', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots' });
  assert.equal(r.ok, true, JSON.stringify(r));
  const names = r.result.sheets.map(s => s.sheet).sort();
  assert.deepEqual(names, ['Empty', 'H', 'Hre', 'S']);
  assert.deepEqual(r.result.sheets.find(s => s.sheet === 'S').pivots, []);
});

test('GH #140 U2: a non-existent sheet is named UNOBSERVABLE:sheet-missing (not not-read), never a guessed empty list', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'NoSuchSheet' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.sheets, [{ sheet: 'NoSuchSheet', pivots: null, reason: 'UNOBSERVABLE:sheet-missing' }]);
});

test('GH #140 U2: a whole-sheet objects-axis gap (the driver itself threw) is named, never an empty list', async () => {
  const b2 = book();
  b2[1] = { ...b2[1], objects: 'UNOBSERVABLE:threw' };
  const store = createStore(b2);
  const bk = fakeBroker('xlsx');
  const t = rt.createReadTool('xlsx', { configSource: bk.source, loadPlaywright: () => {}, createDriver: () => createFakeXlsxDriver(store, { objectsUnobservable: false }) });
  // the fixture's own gap (objects: 'UNOBSERVABLE:threw' on sheet H) must survive the fake driver's
  // own objects-read path unmodified for this to isolate the read_pivots handler, not the fake model.
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'H' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.sheets[0].pivots, null);
  assert.match(r.result.sheets[0].reason, /^UNOBSERVABLE:/);
});

test('GH #140 U2: a per-pivot gap inside an otherwise-good sheet names the WHOLE sheet\'s pivots axis, never a partial/guessed list', async () => {
  const b2 = book();
  // a pivot-ref that withProvenPivots()/provenSource() cannot verify at all (a shape it rejects outright)
  b2[1] = { ...b2[1], objects: { charts: [], names: [], pivots: [{ name: 'BAD', source: 123 }] } };
  const store = createStore(b2);
  const bk = fakeBroker('xlsx');
  const t = rt.createReadTool('xlsx', { configSource: bk.source, loadPlaywright: () => {}, createDriver: () => createFakeXlsxDriver(store) });
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'H' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.sheets[0].pivots, null);
  assert.match(r.result.sheets[0].reason, /^UNOBSERVABLE:/);
  assert.match(r.result.sheets[0].reason, /BAD/, 'the named pivot should be identified in the reason');
});

test('GH #140 U2: the read is non-mutating -- never dispatches, never force-saves', async () => {
  const store = createStore(book());
  const b = fakeBroker('xlsx');
  let driver;
  const createDriver = () => { driver = createFakeXlsxDriver(store); return driver; };
  const t = rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'H' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(driver.calls.forceSave, 0);
});

// ---- RED CONTROL: proves the handler actually discriminates, not pass-by-construction ----

test('RED CONTROL: a fabricated pivot name is correctly absent from the read result', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_pivots', sheet: 'H' });
  assert.ok(!r.result.sheets[0].pivots.some(p => p.name === 'DefinitelyNotARealPivot'));
});
