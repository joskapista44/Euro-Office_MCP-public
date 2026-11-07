'use strict';

/**
 * GH #336 (insert_table_rows / delete_table_rows) + GH #337 (insert_table_columns / delete_table_columns): the public Insert / Delete over the
 * table's FULL width / FULL height (docs/evidence/gh336-gh337-baseline-20261030.json: RI_IN, RD_IN, CI_IN, CD_IN are correct; a partial width,
 * a header delete, an all-rows delete and a delete under an outside formula are data hazards). The family emits ONLY the full range and refuses,
 * before the write and named, what the engine would get wrong. Offline: the fake driver models the measured behaviour.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/table-structure.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const STYLE = { name: 'TableStyleLight9', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false };
const cell = (v, f = '') => ({ v, f });
const book = (extra = {}) => [{
  name: 'A',
  cells: {
    A1: cell('Item'), B1: cell('Qty'), C1: cell('Note'),
    A2: cell('a'), B2: cell(1), C2: cell('x'), A3: cell('b'), B3: cell(2), C3: cell('y'),
    A4: cell('c'), B4: cell(3), C4: cell('z'), A5: cell('d'), B5: cell(4), C5: cell('w'),
    G1: cell('r1'), G2: cell('r2'), G3: cell('r3'), G4: cell('r4'), G5: cell('r5'),
    A7: cell('cA'), B7: cell('cB'), C7: cell('cC'),
    E8: cell(10, '=SUM(B2:B5)'),
    ...(extra.cells || {}),
  },
  tables: extra.tables || [{ name: 'Table1', ref: 'A1:C5', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: false, style: STYLE, autoFilter: true }],
  merges: extra.merges || [],
  objects: { charts: [], pivots: [], names: [] },
}, { name: 'Other', cells: { A1: cell('other') }, objects: { charts: [], pivots: [], names: [] } }];
// the formula that makes the delete-column case a #REF! (E8 sums the whole deleted column): the baseline's CD_IN
const NO_FORMULA = { cells: { E8: cell('') } };

const insRows = (extra = {}) => ({ intent: 'insert_table_rows', sheet: 'A', table: 'Table1', at: 2, count: 1, currentRows: 4, ...extra });
const delRows = (extra = {}) => ({ intent: 'delete_table_rows', sheet: 'A', table: 'Table1', at: 2, count: 1, currentRows: 4, ...extra });
const insCols = (extra = {}) => ({ intent: 'insert_table_columns', sheet: 'A', table: 'Table1', at: 2, count: 1, currentColumns: 3, ...extra });
const delCols = (extra = {}) => ({ intent: 'delete_table_columns', sheet: 'A', table: 'Table1', at: 2, count: 1, currentColumns: 3, ...extra });
const task = step => ({ id: 'gh336', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const ws0 = store => store.durable[0];
const tableOf = store => ws0(store).tables[0];
const valueAt = (store, a) => (ws0(store).cells[a] || {}).v;
const snapshot = store => JSON.parse(JSON.stringify(store.durable));
const why = r => JSON.stringify([r.execution.primitiveVerification, r.execution.appliedOperations[0].evidence]);

test('schema: the table by name, 1-based at, count 1..100 and the table\'s current size are all required; an unknown field is refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  for (const s of [insRows(), delRows(), insCols(), delCols()]) assert.equal(v(s).ok, true, JSON.stringify(v(s)));
  for (const [name, s] of [
    ['no table', insRows({ table: undefined })], ['no at', insRows({ at: undefined })], ['at 0', insRows({ at: 0 })], ['no count', insRows({ count: undefined })],
    ['count 0', insRows({ count: 0 })], ['count 101', insRows({ count: 101 })], ['no current size', insRows({ currentRows: undefined })],
    ['rows step with currentColumns', insRows({ currentRows: undefined, currentColumns: 3 })], ['unknown field', insRows({ range: 'A1:C5' })], ['bad table name', insRows({ table: '1bad name' })],
  ]) assert.equal(v(s).ok, false, name);
  assert.equal(fam.plan(insRows({ range: 'A1:C5' }), 0).ok, false);
  assert.equal(fam.plan(insCols({ currentColumns: undefined }), 0).ok, false);
});

test('INSERT ROWS: the table grows, the band under the cut moves down, the new row is blank, the side anchors and the other sheet stay; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  const { result, m } = await run(store, task(insRows()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).ref, 'A1:C6');
  assert.deepEqual([valueAt(store, 'A2'), valueAt(store, 'A3'), valueAt(store, 'A4'), valueAt(store, 'A6')], ['a', undefined, 'b', 'd']);
  assert.deepEqual([valueAt(store, 'A8'), valueAt(store, 'C8')], ['cA', 'cC']);
  assert.deepEqual([valueAt(store, 'G3'), valueAt(store, 'G5')], ['r3', 'r5'], 'the anchor column outside the table does not move');
  assert.deepEqual(store.durable[1], before[1]);
  const again = await run(store, task(insRows()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.execution.appliedOperations[0].evidence.observation, 'table-already-at-the-target-size');
});

test('INSERT ROWS: at:1 is UNAFFECTED by the GH #366 6th-unit follow-up fix -- the header row (the table\'s anchor) is excluded from `at` entirely for row ops, so at:1 never touches it', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(insRows({ at: 1 })));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).ref, 'A1:C6');
});

test('INSERT ROWS: count 2 inserts two rows in one call', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(insRows({ count: 2 })));
  assert.equal(result.ok, true, why(result));
  assert.equal(tableOf(store).ref, 'A1:C7');
  assert.deepEqual([valueAt(store, 'A3'), valueAt(store, 'A4'), valueAt(store, 'A5')], [undefined, undefined, 'b']);
});

test('DELETE ROWS: the row is gone, the band moves up, the table shrinks; a spanning formula (E8 = SUM(B2:B5)) does not block it; retry SATISFIED', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(delRows()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).ref, 'A1:C4');
  assert.deepEqual([valueAt(store, 'A2'), valueAt(store, 'A3'), valueAt(store, 'A4'), valueAt(store, 'A5')], ['a', 'c', 'd', undefined]);
  assert.deepEqual([valueAt(store, 'A6'), valueAt(store, 'C6')], ['cA', 'cC']);
  const again = await run(store, task(delRows()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('INSERT COLUMNS: the table is wider, the engine\'s own header (ColumnN) is REPORTED not set, the band right of the cut moves; retry SATISFIED', async () => {
  const store = createStore(book(NO_FORMULA));
  const { result, m } = await run(store, task(insCols()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).ref, 'A1:D5');
  assert.deepEqual(tableOf(store).columns, ['Item', 'Column1', 'Qty', 'Note']);
  assert.equal(result.execution.primitiveVerification[0].evidence.evidence.newColumnHeaders[0], 'Column1');
  assert.deepEqual([valueAt(store, 'B2'), valueAt(store, 'C2'), valueAt(store, 'D2')], [undefined, 1, 'x']);
  assert.deepEqual([valueAt(store, 'H1'), valueAt(store, 'H5')], ['r1', 'r5'], 'the cells right of the table in the table\'s rows move with the band');
  assert.deepEqual([valueAt(store, 'A7'), valueAt(store, 'B7')], ['cA', 'cB'], 'rows under the table do not move');
  const again = await run(store, task(insCols()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('DELETE COLUMNS: the column is gone, the table is narrower with its other column names in order (no formula reaches it)', async () => {
  const store = createStore(book(NO_FORMULA));
  const { result, m } = await run(store, task(delCols()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).ref, 'A1:B5');
  assert.deepEqual(tableOf(store).columns, ['Item', 'Note']);
  assert.deepEqual([valueAt(store, 'B2'), valueAt(store, 'B5')], ['x', 'w']);
  assert.deepEqual([valueAt(store, 'F1'), valueAt(store, 'F5')], ['r1', 'r5']);
});

// ---- the refusal matrix: each named, BEFORE the write (0 commands), the store untouched ----
const refusals = [
  ['a structured reference anywhere in the workbook', book({ cells: { E9: cell(6, '=SUM(Table1[Qty])') } }), insRows(), 'xlsx-table-structured-reference'],
  ['a structured reference on ANOTHER sheet', (() => { const b = book(); b[1].cells.B2 = cell(6, '=SUM(Table1[Qty])'); return b; })(), delRows(), 'xlsx-table-structured-reference'],
  ['a this-row reference', book({ cells: { D2: cell(1, '=[@Qty]*2') } }), insRows(), 'xlsx-table-structured-reference'],
  ['a totals row', book({ tables: [{ name: 'Table1', ref: 'A1:C5', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: true, style: STYLE, autoFilter: true }] }), insRows(), 'xlsx-table-totals-row-unmeasured'],
  ['a table without a header row', book({ tables: [{ name: 'Table1', ref: 'A1:C5', columns: ['Item', 'Qty', 'Note'], headerRow: false, totalsRow: false, style: STYLE, autoFilter: true }] }), insRows(), 'xlsx-table-no-header-row'],
  ['deleting EVERY data row', book(), delRows({ at: 1, count: 4 }), 'xlsx-table-delete-all-data-rows'],
  ['deleting EVERY column', book(NO_FORMULA), delCols({ at: 1, count: 3 }), 'xlsx-table-delete-all-columns'],
  ['an insert AFTER the last data row (the table cannot grow by an append)', book(), insRows({ at: 5 }), 'xlsx-table-at-out-of-range'],
  // GH #366 6th-unit follow-up (LIVE t6, prod): insert_table_columns at:1 shifts the whole table
  // instead of growing it -- the table's own anchor cell IS column 1, unlike a row insert at:1
  // (the header row is excluded from `at` entirely, so it never touches the anchor).
  ['inserting a column BEFORE the first column (the table\'s own anchor) -- MEASURED LIVE: the engine shifts the whole table instead of growing it', book(NO_FORMULA), insCols({ at: 1 }), 'xlsx-table-insert-before-first-column-unsupported'],
  ['inserting a column at count:2, at:1 -- still refused (the anchor-column hazard does not depend on count)', book(NO_FORMULA), insCols({ at: 1, count: 2 }), 'xlsx-table-insert-before-first-column-unsupported'],
  ['a delete that runs past the last data row', book(), delRows({ at: 3, count: 3 }), 'xlsx-table-at-out-of-range'],
  ['a wrong current size (the retry guard)', book(), insRows({ currentRows: 7 }), 'xlsx-table-size-unexpected'],
  ['a formula the column delete would break (E8 sums the whole deleted column)', book(), delCols(), 'xlsx-table-block-referenced'],
  ['a single-cell reference into the deleted row', book({ cells: { E9: cell(2, '=B3') } }), delRows(), 'xlsx-table-block-referenced'],
  ['a merge in the shifted band', book({ merges: ['A7:B7'] }), insRows(), 'xlsx-table-band-has-merge'],
  ['another table in the shifted band', book({ tables: [{ name: 'Table1', ref: 'A1:C5', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: false, style: STYLE, autoFilter: true }, { name: 'Table2', ref: 'A8:B9', columns: ['p', 'q'], headerRow: true, totalsRow: false, style: STYLE, autoFilter: true }] }), insRows(), 'xlsx-table-band-has-table'],
  ['a table that does not exist', book(), insRows({ table: 'NoSuch' }), 'xlsx-table-not-found'],
  ['a sheet that does not exist', book(), insRows({ sheet: 'Nope' }), 'xlsx-sheet-not-found'],
];
for (const [name, b, step, reason] of refusals) {
  test(`REFUSED before the write: ${name} -> ${reason}`, async () => {
    const store = createStore(b);
    const before = snapshot(store);
    const { result, m } = await run(store, task(step));
    assert.equal(result.ok, false);
    assert.equal(m.commands, 0);
    assert.match(JSON.stringify(result.execution), new RegExp(reason));
    assert.deepEqual(snapshot(store), before);
  });
}

test('REFUSED: the tables cannot be read -> fail-closed, 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(insRows()), { tablesUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-tables-unobservable/);
});

// ---- the verify can go RED (fake-driver faults) ----
for (const [fault, step, b] of [
  ['ts-ref-not-grown', insRows(), book()], ['ts-partial-width', insRows(), book()], ['ts-band-stays', insRows(), book()], ['ts-band-stays', delRows(), book()],
  ['ts-ref-not-grown', delRows(), book()], ['ts-columns-lost', insCols(), book(NO_FORMULA)], ['ts-band-stays', insCols(), book(NO_FORMULA)],
  ['ts-leaves-ref-error', delRows(), book({ cells: { E9: cell(5, '=SUM(B2:B5)') } })], ['side-effect', insRows(), book()], ['ignore', insRows(), book()],
]) {
  test(`VERIFY goes RED: ${step.intent} with fault ${fault}`, async () => {
    const store = createStore(b);
    const { result } = await run(store, task(step), { fault });
    assert.equal(result.ok, false, fault);
  });
}

test('each is the only step of its task; liveScope asks the tables, the objects and the dependents', () => {
  for (const [a, b] of [[insRows(), insRows()], [delRows(), delCols()], [insCols(), insRows()]]) {
    const p = planOperations({ id: 't', steps: [a, b] });
    assert.equal(p.ok, false);
    assert.match(p.detail, /must be the only step of its task/);
  }
  const s = liveScope(task(insRows()));
  assert.deepEqual(s.tableSheets, ['A']);
  assert.deepEqual(s.objectSheets, ['A']);
  assert.equal(s.dependents, true);
});

test('the lowered commands are the FULL-width / FULL-height range (header included for columns), never a partial range', () => {
  const lo = (step, ref = 'A1:C5') => fam.lower(fam.plan(step, 0).operation, { ref });
  assert.match(lo(insRows()).body, /GetRange\("A3:C3"\)\.Insert\('down'\)/);
  assert.match(lo(insRows({ count: 2 })).body, /GetRange\("A3:C4"\)\.Insert\('down'\)/);
  assert.match(lo(delRows({ at: 2, count: 2 })).body, /GetRange\("A3:C4"\)\.Delete\('up'\)/);
  assert.match(lo(insCols()).body, /GetRange\("B1:B5"\)\.Insert\('right'\)/);
  assert.match(lo(delCols({ count: 2 })).body, /GetRange\("B1:C5"\)\.Delete\('left'\)/);
  assert.match(lo(insRows(), 'B2:E9').body, /GetRange\("B4:E4"\)\.Insert\('down'\)/, 'a table that does not start at A1');
  for (const s of [insRows(), delRows(), insCols(), delCols()]) {
    const c = lo(s);
    assert.equal(c.meta.family, 'table-structure');
    assert.doesNotMatch(c.body, /AddComment|FormatAsTable/);
  }
});

// ---- the guards the fake cannot model, on a synthetic observation (fam.classify directly) ----
const obs = (over = {}) => ({
  format: 'xlsx', complete: true, tier: 'full', sheets: [{ name: 'A' }],
  tableReads: { A: { tables: [{ name: 'Table1', ref: 'A1:C5', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: false, style: STYLE, autoFilter: true }] } },
  cells: { 'A\u0000A1': { v: 'Item', f: '', merge: 'A1' } }, formulas: {}, objects: { A: { pivots: [] } }, dependencies: { complete: true, refs: [], unobservable: [] },
  ...over,
});
const opOf = step => fam.plan(step, 0).operation;

test('LIVE finding: the target table\'s OWN name in the defined-name read is not a dependent (every table carries one, the seed was refused); another name still is', () => {
  const own = obs({ dependencies: { complete: true, unobservable: [], refs: [{ kind: 'name', host: null, name: 'table1', sheets: ['A'] }] } });
  assert.equal(fam.classify(own, opOf(insRows())).state, 'APPLY');
  const other = obs({ dependencies: { complete: true, unobservable: [], refs: [{ kind: 'name', host: null, name: 'Table1', sheets: ['A'] }, { kind: 'name', host: null, name: 'Budget', sheets: ['A'] }] } });
  const c = fam.classify(other, opOf(insRows()));
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'xlsx-table-structure-dependent-not-measured');
  assert.deepEqual(c.evidence.dependents.map(r => r.name), ['Budget']);
});

test('classify on a clean synthetic observation is APPLY; the same with each unmeasured reader gap is refused, named', () => {
  assert.equal(fam.classify(obs(), opOf(insRows())).state, 'APPLY');
  const refused = [
    ['a pivot in the shifted band', obs({ objects: { A: { pivots: [{ name: 'P1', table: '$A$8:$C$12' }] } } }), 'xlsx-table-band-has-pivot'],
    ['an unparsable pivot location', obs({ objects: { A: { pivots: [{ name: 'P1', table: 'weird' }] } } }), 'xlsx-pivots-unobservable'],
    ['pivots not read', obs({ objects: {} }), 'xlsx-pivots-unobservable'],
    ['a validation reaching the sheet (conservative)', obs({ dependencies: { complete: true, unobservable: [], refs: [{ kind: 'validation', host: 'A', sheets: ['A'] }] } }), 'xlsx-table-structure-dependent-not-measured'],
    ['a chart reaching the sheet (conservative)', obs({ dependencies: { complete: true, unobservable: [], refs: [{ kind: 'chart', host: 'B', sheets: ['A'] }] } }), 'xlsx-table-structure-dependent-not-measured'],
    ['the dependency reader incomplete', obs({ dependencies: { complete: false, unobservable: ['names'], refs: [] } }), 'xlsx-structure-dependency-observation-incomplete'],
    ['the dependency reader absent', obs({ dependencies: null }), 'xlsx-structure-dependency-observation-incomplete'],
    ['a windowed read', obs({ tier: 'window' }), 'xlsx-table-structure-needs-a-full-read'],
    ['a tiered read', obs({ tier: 'tiered' }), 'xlsx-table-structure-needs-a-full-read'],
    ['an incomplete read', obs({ complete: false, truncated: true }), 'xlsx-observation-truncated'],
    ['an unparsable table ref', obs({ tableReads: { A: { tables: [{ name: 'Table1', ref: 'nonsense', columns: [], headerRow: true, totalsRow: false }] } } }), 'xlsx-table-ref-unparsable'],
    ['two tables of the same name', obs({ tableReads: { A: { tables: [{ name: 'Table1', ref: 'A1:C5', columns: [], headerRow: true, totalsRow: false }, { name: 'table1', ref: 'E1:F5', columns: [], headerRow: true, totalsRow: false }] } } }), 'xlsx-table-name-ambiguous'],
  ];
  for (const [name, p, reason] of refused) {
    const c = fam.classify(p, opOf(insRows()));
    assert.equal(c.state, 'FAIL_CLOSED', name);
    assert.equal(c.evidence.reason, reason, name);
  }
  // a pivot BESIDE the band (rows above the cut, same columns) does not block
  assert.equal(fam.classify(obs({ objects: { A: { pivots: [{ name: 'P1', table: '$A$1:$C$2' }] } } }), opOf(insRows({ at: 3 }))).state, 'APPLY');
});

test('a satisfied step names its observation; a table at a size that is neither the guard nor the target is refused (a different run changed it)', () => {
  const grown = obs({ tableReads: { A: { tables: [{ name: 'Table1', ref: 'A1:C6', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: false, style: STYLE, autoFilter: true }] } } });
  const c = fam.classify(grown, opOf(insRows()));
  assert.equal(c.state, 'SATISFIED');
  assert.equal(c.evidence.observation, 'table-already-at-the-target-size');
  const other = fam.classify(grown, opOf(insRows({ currentRows: 2 })));
  assert.equal(other.state, 'FAIL_CLOSED');
  assert.equal(other.evidence.reason, 'xlsx-table-size-unexpected');
});
