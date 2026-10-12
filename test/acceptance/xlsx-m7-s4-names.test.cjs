'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 S4 R1 (owner GO 2026-10-03, owner channel): defined names under insert / delete rows / columns. Measured on the
 * engine (W5 probe + the S4 probe, run m7-s4-probe2-20261003: 40 names x {session, reopen} = 80 / 80 as shiftFormula
 * predicts, workbook AND sheet level, the scope kept). Supported: a name whose text is ONE absolute cell; refused before
 * the write: any other shape reaching the shifted sheet, a name the delete would break (#REF!), an unreadable name list.
 * L3: the same (name, scope) set, every text where the shift puts it, the untouched names unchanged.
 */

const book = (names = []) => [
  { name: 'P', cells: {
    A1: { v: 'r1' }, A2: { v: 'r2' }, A3: { v: 'r3' }, A4: { v: 'r4' }, A5: { v: 'r5' }, A6: { v: 'r6' },
    B1: { v: 1 }, B2: { v: 2 }, B3: { v: 3 }, B4: { v: 4 }, B5: { v: 5 }, B6: { v: 6 },
  }, objects: { charts: [], pivots: [], names: names.filter(n => n.on === 'P').map(n => ({ name: n.name, refersTo: `=${n.ref}` })) } },
  { name: 'Q', cells: { B5: { v: 77 }, ...(names.some(n => n.name === 'Below' && !n.on) ? { A1: { v: 10, f: '=Below*2' } } : {}) },
    objects: { charts: [], pivots: [], names: names.filter(n => n.on === 'Q').map(n => ({ name: n.name, refersTo: `=${n.ref}` })) } },
];

const SUPPORTED = [
  { name: 'Below', ref: 'P!$B$5' }, { name: 'Above', ref: 'P!$B$2' }, { name: 'Other', ref: 'Q!$B$5' },
  { name: 'ShBelow', ref: 'P!$B$5', on: 'P' }, { name: 'ShOther', ref: 'Q!$B$5', on: 'P' }, { name: 'ShOnQ', ref: 'P!$B$6', on: 'Q' },
];
const workbook = list => list.filter(n => !n.on).map(n => ({ name: n.name, ref: n.ref }));

const step = (intent, extra) => ({ id: 'm7-s4-r1', steps: [{ intent, sheet: 'P', ...extra }] });
const INSERT = step('insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r3' } });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;
const pres = result => result.execution.preservationVerification[0];
const presEv = result => pres(result).evidence.evidence;
const sheetNames = (store, s) => store.durable.find(x => x.name === s).objects.names.map(n => [n.name, n.refersTo]);

test('insert 2 rows at 3: APPLIED, every name follows (workbook + sheet level, either scope host), the untouched ones unchanged; retry SATISFIED 0/0', async () => {
  const store = createStore(book(SUPPORTED));
  store.names = workbook(SUPPORTED);
  const { result, m } = await run(store, INSERT);
  assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.preservationVerification]));
  assert.equal(m.commands, 1);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.deepEqual(store.names, [{ name: 'Below', ref: 'P!$B$7' }, { name: 'Above', ref: 'P!$B$2' }, { name: 'Other', ref: 'Q!$B$5' }]);
  assert.deepEqual(sheetNames(store, 'P'), [['ShBelow', '=P!$B$7'], ['ShOther', '=Q!$B$5']]);
  assert.deepEqual(sheetNames(store, 'Q'), [['ShOnQ', '=P!$B$8']]);
  assert.equal(presEv(result).namesKept, true);
  assert.equal(presEv(result).namesCompared, 6);
  const again = await run(store, INSERT);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
});

test('delete rows 3:4 and the columns axis: the names follow (delete moves up, insert / delete columns move along)', async () => {
  const del = createStore(book(SUPPORTED));
  del.names = workbook(SUPPORTED);
  const d = await run(del, step('delete_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r5' } }));
  assert.equal(d.result.ok, true, JSON.stringify([op(d.result).evidence, d.result.execution.preservationVerification]));
  assert.deepEqual(del.names.map(n => n.ref), ['P!$B$3', 'P!$B$2', 'Q!$B$5']);
  const cols = createStore(book(SUPPORTED));
  cols.names = workbook(SUPPORTED);
  const c = await run(cols, step('insert_columns', { at: 'B', count: 1, anchor: { row: 1, value: 1 } }));
  assert.equal(c.result.ok, true, JSON.stringify([op(c.result).evidence, c.result.execution.preservationVerification]));
  assert.deepEqual(cols.names.map(n => n.ref), ['P!$C$5', 'P!$C$2', 'Q!$B$5']);
  assert.deepEqual(sheetNames(cols, 'Q'), [['ShOnQ', '=P!$C$6']]);
});

for (const [title, names, t, why] of [
  ['a range name', [{ name: 'Rg', ref: 'P!$B$5:$B$6' }], INSERT, 'xlsx-structure-name-shape-unsupported'],
  ['a mixed-reference name', [{ name: 'Mx', ref: 'P!$B5' }], INSERT, 'xlsx-structure-name-shape-unsupported'],
  ['a name through another name', [{ name: 'Base', ref: 'Q!$B$5' }, { name: 'Via', ref: 'Below' }, { name: 'Below', ref: 'P!$B$5' }], INSERT, 'xlsx-structure-name-shape-unsupported'],
  ['a name the delete breaks (#REF!)', [{ name: 'In', ref: 'P!$B$4' }], step('delete_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r5' } }), 'xlsx-structure-name-would-break'],
  ['a sheet-level name the delete breaks', [{ name: 'ShIn', ref: 'P!$B$3', on: 'P' }], step('delete_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r5' } }), 'xlsx-structure-name-would-break'],
]) {
  test(`refused BEFORE the write, 0/0: ${title}`, async () => {
    const store = createStore(book(names));
    store.names = workbook(names);
    const { result, m } = await run(store, t);
    assert.equal(result.ok, false);
    assert.deepEqual([m.writes, m.commands], [0, 0]);
    assert.equal(reason(result), why, JSON.stringify(op(result).evidence));
    assert.equal(store.saves, 0);
  });
}

test('refused 0/0: the name list unreadable (the reader a reason string)', async () => {
  const store = createStore(book());
  store.names = workbook(SUPPORTED);
  const { result, m } = await run(store, INSERT, { namesFault: 'UNOBSERVABLE:asc_getDefinedNames-threw' });
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.equal(reason(result), 'xlsx-structure-dependency-observation-incomplete');
});

test('positive control: a name on another sheet only (nothing reaches P) still APPLIES, and L3 compares it', async () => {
  const names = [{ name: 'Other', ref: 'Q!$B$5' }];
  const store = createStore(book(names));
  store.names = workbook(names);
  const { result } = await run(store, INSERT);
  assert.equal(result.ok, true);
  assert.equal(presEv(result).namesCompared, 1);
});

for (const fault of ['name-no-follow', 'name-wrong-shift', 'name-scope-lost', 'name-dropped', 'name-appears', 'name-other-changes']) {
  test(`L3 RED (mutant '${fault}'): the names L3 fails, the task is not ok`, async () => {
    const store = createStore(book(SUPPORTED));
    store.names = workbook(SUPPORTED);
    const { result } = await run(store, INSERT, { fault });
    assert.equal(result.ok, false);
    assert.equal(pres(result).verdict, 'FAIL');
    assert.equal(presEv(result).namesKept, false, JSON.stringify(presEv(result).nameDiffs));
  });
}
