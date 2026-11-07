'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #361 2. egyseg (card an internal branch, kiosztó zero): the
 * PRE-MUTATION named refusal, filter.cjs's own table-owned-AutoFilter guard. The mert alap: the READ-only
 * measurement unit (test/live/xlsx-gh361-table-autofilter-criteria-read.live.cjs, docs/evidence/
 * gh361-table-autofilter-live-20261104/, GH #361 komment 6643346450) MEASURED, on the installed product, that
 * WITHOUT a guard: filter_range / set_autofilter_range on a Table's own range dispatch into
 * DISPATCHED_OUTCOME_UNKNOWN (T_FILTER_RANGE_WRITE / T_SET_AUTOFILTER_RANGE, both
 * {"reason":"operation ...: VERIFY_CANNOT_PROVE_TERMINAL_STATE"}), and remove_autofilter reports a FALSE
 * SATISFIED while the table's own AutoFilter stays completely intact (R_EFFECT:
 * SHEET_LEVEL_NOOP_TABLE_FILTER_KEPT). This unit refuses all four intents PRE-MUTATION instead, named
 * 'xlsx-filter-table-owned-autofilter'.
 *
 * RED/GREEN: every test below FAILS against the pre-patch filter.cjs (no table-guard import, gate() never
 * consults p.tableReads) -- confirmed by stashing this patch and re-running (see the patch's own commit
 * message for the exact RED output). The guard itself lives entirely in gate(), so a mutation that removes
 * the `tableBlock` check (or the `tableOverlap`/`anyAutoFilterTable` import) makes every "refused, named"
 * assertion below fail -- the dispatch count assertions (driver.calls.dispatch === 0) are the mutation-control:
 * without the guard every one of these cases reaches the engine.
 */

const TABLE_REF = 'A1:C4';
const bookCells = () => ({
  A1: { v: 'id' }, B1: { v: 'tag' }, C1: { v: 'note' },
  A2: { v: 1 }, B2: { v: 'x' }, C2: { v: 'n1' },
  A3: { v: 2 }, B3: { v: 'y' }, C3: { v: 'n2' },
  A4: { v: 3 }, B4: { v: 'z' }, C4: { v: 'n3' },
});
const TABLE = (ref = TABLE_REF, autoFilter = true) => ({ name: 'Table1', ref, headerRow: true, totalsRow: false, columns: ['id', 'tag', 'note'], style: null, autoFilter });

// F: table on A1:C4 WITH its own AutoFilter -- the measured risk shape (filter_range/set_autofilter_range's
// target range IS the table's range; remove_autofilter/clear_filter's range is also the table's range).
const bookTableOwned = () => {
  const sheets = [{ name: 'F', cells: bookCells() }, { name: 'O', cells: { A1: { v: 'keep' } } }];
  sheets[0].tables = [TABLE(TABLE_REF, true)];
  return sheets;
};
// F2: table on A1:C4 WITHOUT its own AutoFilter (autoFilter:false) -- the guard must NOT fire (nothing of its
// own to protect), same as the control.
const bookTableNoAutoFilter = () => {
  const sheets = [{ name: 'F', cells: bookCells() }, { name: 'O', cells: { A1: { v: 'keep' } } }];
  sheets[0].tables = [TABLE(TABLE_REF, false)];
  return sheets;
};
// F3: a table WITH its own AutoFilter, but on a DIFFERENT, non-overlapping range (E1:F4) -- the boundary
// case: filter_range/set_autofilter_range target A1:C4, which the table does not touch at all.
const bookTableElsewhere = () => {
  const sheets = [{ name: 'F', cells: { ...bookCells(), E1: { v: 'k' }, F1: { v: 'v' }, E2: { v: 'k1' }, F2: { v: 'v1' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
  sheets[0].tables = [TABLE('E1:F4', true)];
  return sheets;
};
// F4: no table at all -- the control case, must be byte-for-byte unaffected by this unit's guard.
const bookNoTable = () => [{ name: 'F', cells: bookCells() }, { name: 'O', cells: { A1: { v: 'keep' } } }];

// remove_autofilter/clear_filter are SATISFIED (no dispatch) on a sheet with no SHEET-LEVEL AutoFilter at all,
// regardless of any table -- that is the pre-existing, already-released behaviour (xlsx-fa7244ac-autofilter-
// product.test.cjs), not this unit's guard. To prove the guard does not interfere with the cases it must NOT
// block, these fixtures carry a PRE-EXISTING sheet-level AutoFilter so there is something to actually remove/
// clear, and an un-blocked run is observable as a real dispatch (driver.calls.dispatch === 1).
const withSheetLevelAutoFilter = sheets => { sheets[0].autoFilter = { range: TABLE_REF, filters: [{ field: 2, on: true, operator: 'xlOr', criteria1: '=x', criteria2: null }] }; return sheets; };

const flt = (extra = {}) => ({ id: 'gh361r', steps: [{ intent: 'filter_range', sheet: 'F', range: TABLE_REF, field: 2, value: 'x', ...extra }] });
const setRange = (range = TABLE_REF) => ({ id: 'gh361r', steps: [{ intent: 'set_autofilter_range', sheet: 'F', range }] });
const removeAf = (range = TABLE_REF) => ({ id: 'gh361r', steps: [{ intent: 'remove_autofilter', sheet: 'F', range }] });
const clearF = (range = TABLE_REF) => ({ id: 'gh361r', steps: [{ intent: 'clear_filter', sheet: 'F', range }] });

async function run(store, task, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(task, { observation: { live: liveScope(task) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;
const op = result => result.execution.appliedOperations[0];

// ---------------------------------------------------------------------------------------------------- RANGE-SCOPED
for (const [label, task] of [['filter_range', flt()], ['set_autofilter_range', setRange()]]) {
  test(`${label}: refused PRE-MUTATION, named, when the target range IS a table carrying its own AutoFilter`, async () => {
    const { result, driver } = await run(createStore(bookTableOwned()), task);
    assert.equal(result.ok, false, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(reason(result), 'xlsx-filter-table-owned-autofilter');
    assert.equal(op(result).evidence.evidence.table, 'Table1');
    assert.equal(op(result).evidence.evidence.ref, TABLE_REF);
    assert.equal(driver.calls.dispatch, 0, 'mutation-control: the engine must never be called once the guard fires');
    assert.equal(result.metrics.commands, 0);
  });

  test(`${label}: DISCRIMINATES -- a table on the SAME sheet WITHOUT its own AutoFilter does not block (nothing of its own to protect)`, async () => {
    const { result, driver } = await run(createStore(bookTableNoAutoFilter()), task);
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(driver.calls.dispatch, 1);
  });

  test(`${label}: DISCRIMINATES -- a table WITH its own AutoFilter, but NOT overlapping the target range, does not block`, async () => {
    const { result, driver } = await run(createStore(bookTableElsewhere()), task);
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.notEqual(result.ok ? null : reason(result), 'xlsx-filter-table-owned-autofilter');
    assert.equal(driver.calls.dispatch, 1);
  });

  test(`${label}: NEG. KONTROLL -- a sheet with no table at all behaves byte-for-byte as before this guard`, async () => {
    const { result, driver } = await run(createStore(bookNoTable()), task);
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(driver.calls.dispatch, 1);
  });

  test(`${label}: tables unreadable on the runtime (no Bridge /18) keeps the pre-GH#361 behaviour -- never read as "no table"`, async () => {
    const { result, driver } = await run(createStore(bookTableOwned()), task, { tablesUnobservable: true });
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.notEqual(result.ok ? null : reason(result), 'xlsx-filter-table-owned-autofilter');
    assert.equal(driver.calls.dispatch, 1);
  });
}

// ---------------------------------------------------------------------------------------------------- SHEET-WIDE
for (const [label, task] of [['remove_autofilter', removeAf()], ['clear_filter', clearF()]]) {
  test(`${label}: refused PRE-MUTATION, named, when the target range IS a table carrying its own AutoFilter`, async () => {
    const { result, driver } = await run(createStore(bookTableOwned()), task);
    assert.equal(result.ok, false, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(reason(result), 'xlsx-filter-table-owned-autofilter');
    assert.equal(op(result).evidence.evidence.table, 'Table1');
    assert.equal(driver.calls.dispatch, 0, 'mutation-control: the engine must never be called once the guard fires');
  });

  test(`${label}: DISCRIMINATES (sheet-wide) -- ALSO refused when the table with its own AutoFilter is elsewhere on the sheet, NOT overlapping the asked range -- remove/clear address the ONE sheet-level AutoFilter, not a sub-range`, async () => {
    const { result, driver } = await run(createStore(bookTableElsewhere()), task);
    assert.equal(result.ok, false, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(reason(result), 'xlsx-filter-table-owned-autofilter');
    assert.equal(driver.calls.dispatch, 0);
  });

  test(`${label}: DISCRIMINATES -- a table on the sheet WITHOUT its own AutoFilter does not block`, async () => {
    const { result, driver } = await run(createStore(withSheetLevelAutoFilter(bookTableNoAutoFilter())), task);
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(driver.calls.dispatch, 1);
  });

  test(`${label}: NEG. KONTROLL -- a sheet with no table at all behaves byte-for-byte as before this guard`, async () => {
    const { result, driver } = await run(createStore(withSheetLevelAutoFilter(bookNoTable())), task);
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.equal(driver.calls.dispatch, 1);
  });

  test(`${label}: tables unreadable on the runtime keeps the pre-GH#361 behaviour -- never read as "no table"`, async () => {
    const { result, driver } = await run(createStore(withSheetLevelAutoFilter(bookTableOwned())), task, { tablesUnobservable: true });
    assert.equal(result.ok, true, JSON.stringify(op(result) && op(result).evidence));
    assert.notEqual(result.ok ? null : reason(result), 'xlsx-filter-table-owned-autofilter');
    assert.equal(driver.calls.dispatch, 1);
  });
}

// ---------------------------------------------------------------------------------------------------- liveScope wiring
test('liveScope: every filter-family intent now also asks for its own sheet\'s tables (read_tables\' tableSheets axis)', () => {
  for (const task of [flt(), setRange(), removeAf(), clearF()]) {
    const s = liveScope(task);
    assert.deepEqual(s.tableSheets, ['F'], JSON.stringify([task.steps[0].intent, s]));
  }
});

// ---------------------------------------------------------------------------------------------------- table-guard.cjs unit
const { tableOverlap, anyAutoFilterTable } = require('../../src/adapters/xlsx/table-guard.cjs');

test('tableOverlap: additively carries the autoFilter flag (merge.cjs/sort.cjs ignore it, filter.cjs reads it)', () => {
  const p = { tableReads: { F: { tables: [{ name: 'Table1', ref: 'A1:C4', autoFilter: true }] } } };
  assert.deepEqual(tableOverlap(p, 'F', 'A1:C4'), { table: { name: 'Table1', ref: 'A1:C4', autoFilter: true }, exact: true });
  const p2 = { tableReads: { F: { tables: [{ name: 'Table1', ref: 'A1:C4' }] } } }; // autoFilter field absent -- never guessed true
  assert.equal(tableOverlap(p2, 'F', 'A1:C4').table.autoFilter, false);
});

test('anyAutoFilterTable: finds the first table (document order) carrying its own AutoFilter, sheet-wide', () => {
  const p = { tableReads: { F: { tables: [{ name: 'T0', ref: 'E1:F4', autoFilter: false }, { name: 'T1', ref: 'A1:C4', autoFilter: true }] } } };
  assert.deepEqual(anyAutoFilterTable(p, 'F'), { table: { name: 'T1', ref: 'A1:C4', autoFilter: true } });
});

test('anyAutoFilterTable: DISCRIMINATES -- no table carries autoFilter -> table: null, not unobservable', () => {
  const p = { tableReads: { F: { tables: [{ name: 'T0', ref: 'E1:F4', autoFilter: false }] } } };
  assert.deepEqual(anyAutoFilterTable(p, 'F'), { table: null });
});

test('anyAutoFilterTable: an unread sheet is named unobservable, never guessed as "no table"', () => {
  assert.deepEqual(anyAutoFilterTable({ tableReads: {} }, 'F'), { unobservable: 'not-read' });
  assert.deepEqual(anyAutoFilterTable({ tableReads: { F: 'UNOBSERVABLE:threw' } }, 'F'), { unobservable: 'UNOBSERVABLE:threw' });
});
