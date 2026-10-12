'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const names = require('../../src/adapters/xlsx/families/names.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M7 F Q1 / Q2 / Q4 / Q5 (tier F probe m7-f-probe1-20261003.json): print area / row titles as sheet-level names. */

const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'Q 1', cells: { A1: { v: 2 } } }];
const task = step => ({ id: 'm7-f', steps: [{ sheet: 'P', ...step }] });
const AREA = task({ intent: 'set_print_area', range: 'A1:C5' });
const TITLES = task({ intent: 'set_print_titles', rows: '1:2' });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];
const sheetNames = (store, s = 'P') => ((store.durable.find(x => x.name === s).objects || {}).names || []).map(n => `${n.name}=${n.refersTo}`);

test('Q1 / Q4: set the print area and the row titles -> sheet-level names; retries SATISFIED 0/0; Q2 / Q5: clear both; retries 0/0', async () => {
  const store = createStore(book());
  for (const t of [AREA, TITLES]) {
    const a = await run(store, t);
    assert.equal(a.result.ok, true, JSON.stringify([op(a.result).evidence, a.result.execution.primitiveVerification]));
    const again = await run(store, t);
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  }
  assert.deepEqual(sheetNames(store), ['Print_Area==P!$A$1:$C$5', 'Print_Titles==P!$1:$2']);
  assert.deepEqual(store.names, []);
  for (const t of [task({ intent: 'clear_print_area' }), task({ intent: 'clear_print_titles' })]) {
    const c = await run(store, t);
    assert.equal(c.result.ok, true, JSON.stringify(c.result.execution.primitiveVerification));
    assert.equal((await run(store, t)).m.commands, 0);
  }
  assert.deepEqual(sheetNames(store), []);
});

test('Q1: the command is the measured public sheet-level call with a bare reference; a quoted sheet name is quoted', () => {
  const body = t => names.lower(planOperations(t).operations[0]).body;
  assert.match(body(AREA), /const ws = Api\.GetSheet\("P"\);[\s\S]*ws\.AddDefName\("Print_Area", "P!\$A\$1:\$C\$5"\)/);
  assert.match(body(TITLES), /ws\.AddDefName\("Print_Titles", "P!\$1:\$2"\)/);
  assert.match(body({ id: 'x', steps: [{ intent: 'set_print_area', sheet: 'Q 1', range: 'B2' }] }), /AddDefName\("Print_Area", "'Q 1'!\$B\$2:\$B\$2"\)/);
  assert.match(body(task({ intent: 'clear_print_area' })), /const d = ws\.GetDefName\("Print_Area"\);[\s\S]*d\.Delete\(\)/);
});

test('Q1 RED: the editor stores the area at workbook level -> L1 FAIL, not APPLIED', async () => {
  const { result } = await run(createStore(book()), AREA, { fault: 'name-wrong-scope' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('Q1 refused: an area that already exists with another reference (no silent re-point)', async () => {
  const b = book();
  b[0].objects = { charts: [], pivots: [], names: [{ name: 'Print_Area', refersTo: '=P!$A$1:$B$2' }] };
  const r = await run(createStore(b), AREA);
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(op(r.result).evidence.evidence.reason, 'xlsx-name-exists-with-another-reference');
});

test('Q1: another sheet\'s print area does not block (one per sheet, measured); a workbook-level Print_Area does', async () => {
  const b = book();
  b[1].objects = { charts: [], pivots: [], names: [{ name: 'Print_Area', refersTo: "='Q 1'!$A$1:$B$2" }] };
  const store = createStore(b);
  const a = await run(store, AREA);
  assert.equal(a.result.ok, true, JSON.stringify(op(a.result).evidence));
  const wb = createStore(book());
  wb.names = [{ name: 'Print_Area', ref: 'P!$A$1:$B$2' }];
  const r = await run(wb, AREA);
  assert.equal(r.driver.calls.dispatch, 0);
  assert.equal(op(r.result).evidence.evidence.reason, 'xlsx-name-exists-in-another-scope');
});

test('refused at planning: a malformed row span, a multi-area print area, set_defined_name of a reserved name (Q17)', () => {
  for (const t of [task({ intent: 'set_print_titles', rows: '3:1' }), task({ intent: 'set_print_area', range: 'A1:B2,D1:E2' }),
    { id: 'x', steps: [{ intent: 'set_defined_name', name: 'Print_Area', refersTo: '=P!$A$1' }] }]) {
    assert.equal(planOperations(t).ok, false, JSON.stringify(t));
  }
});

// M7 X-P03 (measured: docs/evidence/m7-p57-print-titles-columns-probe-20261005.json):
// column-only and row+column print titles, the stored shape COLUMNS FIRST when both are given.
const TITLES_COL = task({ intent: 'set_print_titles', columns: 'A:B' });
const TITLES_BOTH = task({ intent: 'set_print_titles', rows: '1:2', columns: 'A:B' });

test('X-P03: column-only and row+column print titles -> sheet-level name, columns FIRST; retries SATISFIED 0/0', async () => {
  for (const [t, want] of [[TITLES_COL, 'Print_Titles==P!$A:$B'], [TITLES_BOTH, 'Print_Titles==P!$A:$B,P!$1:$2']]) {
    const store = createStore(book());
    const a = await run(store, t);
    assert.equal(a.result.ok, true, JSON.stringify([op(a.result).evidence, a.result.execution.primitiveVerification]));
    assert.deepEqual(sheetNames(store), [want]);
    const again = await run(store, t);
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  }
});

test('X-P03: the command is the measured public call, columns-only and row+column (columns first, comma-joined)', () => {
  const body = t => names.lower(planOperations(t).operations[0]).body;
  assert.match(body(TITLES_COL), /ws\.AddDefName\("Print_Titles", "P!\$A:\$B"\)/);
  assert.match(body(TITLES_BOTH), /ws\.AddDefName\("Print_Titles", "P!\$A:\$B,P!\$1:\$2"\)/);
});

test('X-P03 refused at planning: reversed columns, columns in rows\' own shape, neither given, out-of-range columns', () => {
  for (const s of [{ columns: 'B:A' }, { columns: '1:2' }, { columns: 'a:b' }, {}, { columns: 'XFE:XFE' }]) {
    const t = task({ intent: 'set_print_titles', ...s });
    assert.equal(planOperations(t).ok, false, JSON.stringify(s));
  }
});

// M7 X-P02 (measured: docs/evidence/m7-p61-print-area-multi-probe-20261005.json): the already-public
// AddDefName accepts a comma-joined multi-area Print_Area the same way it does Print_Titles' union, byte-stable
// across reopens, order preserved (not reordered by the engine either way).
const AREA_MULTI = task({ intent: 'set_print_area', areas: ['A1:B2', 'D1:E2'] });

test('X-P02: multi-area print area -> sheet-level name, comma-joined, given order preserved; retries SATISFIED 0/0', async () => {
  const store = createStore(book());
  const a = await run(store, AREA_MULTI);
  assert.equal(a.result.ok, true, JSON.stringify([op(a.result).evidence, a.result.execution.primitiveVerification]));
  assert.deepEqual(sheetNames(store), ['Print_Area==P!$A$1:$B$2,P!$D$1:$E$2']);
  const again = await run(store, AREA_MULTI);
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  const c = await run(store, task({ intent: 'clear_print_area' }));
  assert.equal(c.result.ok, true, JSON.stringify(c.result.execution.primitiveVerification));
  assert.deepEqual(sheetNames(store), []);
});

test('X-P02: the command is the measured public call, areas comma-joined in the GIVEN order (no reordering)', () => {
  const body = t => names.lower(planOperations(t).operations[0]).body;
  assert.match(body(AREA_MULTI), /ws\.AddDefName\("Print_Area", "P!\$A\$1:\$B\$2,P!\$D\$1:\$E\$2"\)/);
  const reversed = task({ intent: 'set_print_area', areas: ['D1:E2', 'A1:B2'] });
  assert.match(body(reversed), /ws\.AddDefName\("Print_Area", "P!\$D\$1:\$E\$2,P!\$A\$1:\$B\$2"\)/);
});

test('X-P02 refused at planning: range and areas together, neither, a single-element areas, overlapping areas, a malformed element', () => {
  for (const s of [
    { range: 'A1:B2', areas: ['A1:B2', 'D1:E2'] },
    {},
    { areas: ['A1:B2'] },
    { areas: ['A1:C3', 'B2:D4'] },
    { areas: ['A1:B2', 'not-a-range'] },
  ]) {
    const t = task({ intent: 'set_print_area', ...s });
    assert.equal(planOperations(t).ok, false, JSON.stringify(s));
  }
});
