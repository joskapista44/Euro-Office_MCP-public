'use strict';

/**
 * GH #331 (T06): insert / delete rows / columns on a sheet holding an Excel Table. The table is read as a defined name
 * over its range, which the structure family refused as a range-shaped name on EVERY structural write of that sheet
 * (baseline LIVE, test/live/xlsx-gh331-t06-table-aware-baseline-probe.live.cjs). Now the table is told by read_tables:
 * a table entirely before the cut is kept (APPLY, L3 the table unchanged); a cut through or before the table is refused,
 * named (xlsx-structure-overlaps-table); a range name that is not a readable table keeps the range-shape refusal.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

const TABLE = { name: 'Table1', ref: 'B1:C3', headerRow: true, totalsRow: false, columns: ['K', 'V'], style: null, autoFilter: true };

function book({ table = true } = {}) {
  const store = createStore([{ name: 'P', cells: {
    A1: { v: 'r1' }, A2: { v: 'r2' }, A3: { v: 'r3' }, A4: { v: 'r4' }, A5: { v: 'r5' }, A6: { v: 'r6' }, A7: { v: 'r7' },
    B1: { v: 'K' }, C1: { v: 'V' }, B2: { v: 1 }, C2: { v: 'x' }, B3: { v: 2 }, C3: { v: 'y' },
    E1: { v: 'cE' }, F1: { v: 'cF' }, D1: { v: 'cD' },
  }, objects: { charts: [], pivots: [], names: [] } }]);
  // the table's name over its DATA BODY, the measured form (header row excluded)
  store.names = [{ name: 'Table1', ref: 'P!$B$2:$C$3' }];
  if (table) for (const s of [store.durable, store.live].filter(Boolean)) { const ws = s.find(x => x.name === 'P'); if (ws) ws.tables = [{ ...TABLE }]; }
  return store;
}

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const step = (intent, extra) => ({ id: 'gh331', steps: [{ intent, sheet: 'P', ...extra }] });
const op = result => result.execution.appliedOperations[0];
const reasonOf = result => { const e = op(result).evidence; return (e.evidence && e.evidence.reason) || e.reason; };

test('a table entirely BEFORE the cut: insert / delete rows below it and insert columns to its right are APPLIED, the table kept (L3)', async () => {
  const cases = [
    step('insert_rows', { at: 5, count: 1, anchor: { column: 'A', value: 'r5' } }),
    step('delete_rows', { at: 5, count: 1, anchor: { column: 'A', value: 'r6' }, witness: { column: 'A', value: 'r5' } }),
    step('insert_columns', { at: 'E', count: 1, anchor: { row: 1, value: 'cE' } }),
  ];
  for (const t of cases) {
    const store = book();
    const { result } = await run(store, t);
    assert.equal(result.ok, true, JSON.stringify([t.steps[0].intent, op(result) && op(result).evidence, result.execution.preservationVerification]));
    assert.equal(op(result).status, OPERATION_STATUS.APPLIED);
    assert.equal(result.execution.preservationVerification[0].evidence.evidence.tablesKept, true);
  }
});

test('a cut through the table or before it is refused before the write, named, with no mutation', async () => {
  const cases = [
    step('insert_rows', { at: 2, count: 1, anchor: { column: 'A', value: 'r2' } }),
    step('insert_rows', { at: 1, count: 1, anchor: { column: 'A', value: 'r1' } }),
    step('delete_rows', { at: 3, count: 1, anchor: { column: 'A', value: 'r4' }, witness: { column: 'A', value: 'r3' } }),
    step('insert_columns', { at: 'C', count: 1, anchor: { row: 1, value: 'V' } }),
    step('insert_columns', { at: 'A', count: 1, anchor: { row: 1, value: 'r1' } }),
  ];
  for (const t of cases) {
    const store = book();
    const { result, m } = await run(store, t);
    assert.equal(result.ok, false, t.steps[0].intent);
    assert.equal(reasonOf(result), 'xlsx-structure-overlaps-table', JSON.stringify(t.steps[0]));
    assert.equal(m.writes, 0);
    assert.equal(m.commands, 0);
  }
});

test('NEG: the same range name with no table behind it keeps the range-shape refusal, also when the tables cannot be read', async () => {
  const plain = await run(book({ table: false }), step('insert_rows', { at: 5, count: 1, anchor: { column: 'A', value: 'r5' } }));
  assert.equal(reasonOf(plain.result), 'xlsx-structure-name-shape-unsupported');
  const unread = await run(book(), step('insert_rows', { at: 5, count: 1, anchor: { column: 'A', value: 'r5' } }), { tablesUnobservable: true });
  assert.equal(reasonOf(unread.result), 'xlsx-structure-name-shape-unsupported');
});

test('liveScope: a structural step asks for its own sheet\'s tables', () => {
  const s = liveScope(step('insert_rows', { at: 5, count: 1, anchor: { column: 'A', value: 'r5' } }));
  assert.deepEqual(s.tableSheets, ['P']);
});

const merge = range => ({ id: 'gh331-mg', steps: [{ intent: 'merge_range', sheet: 'P', range }] });
const sort = range => ({ id: 'gh331-sr', steps: [{ intent: 'sort_range', sheet: 'P', range, keyColumn: 'B', order: 'desc', hasHeaders: true }] });

test('merge_range: a merge overlapping the table is refused, named, before the write; one beside it applies', async () => {
  // across the table's right edge, and inside it (the table check comes before the content guard)
  for (const range of ['C2:D2', 'B3:C3']) {
    const { result, m } = await run(book(), merge(range));
    assert.equal(result.ok, false, range);
    assert.equal(reasonOf(result), 'xlsx-merge-overlaps-table');
    assert.deepEqual([m.writes, m.commands], [0, 0]);
  }
  const { result } = await run(book(), merge('E5:F5'));
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
});

test('sort_range: a range only partly over the table is refused, named; exactly the table range, and a range beside it, sort', async () => {
  const part = await run(book(), sort('B1:D3'));
  assert.equal(part.result.ok, false);
  assert.equal(reasonOf(part.result), 'xlsx-sort-overlaps-table');
  assert.deepEqual([part.m.writes, part.m.commands], [0, 0]);
  const exact = await run(book(), sort('B1:C3'));
  assert.equal(exact.result.ok, true, JSON.stringify(op(exact.result).evidence));
});

test('tables unreadable (no Bridge /18 read): merge and sort keep their behaviour from before T06 -- never read as a refusal or as "no table"', async () => {
  const mg = await run(book(), merge('E5:F5'), { tablesUnobservable: true });
  assert.equal(mg.result.ok, true, JSON.stringify(op(mg.result).evidence));
  const sr = await run(book(), sort('B1:D3'), { tablesUnobservable: true });
  assert.notEqual(reasonOf(sr.result), 'xlsx-sort-overlaps-table');
});
