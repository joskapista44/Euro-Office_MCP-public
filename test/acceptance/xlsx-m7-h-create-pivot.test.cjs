'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const pivots = require('../../src/adapters/xlsx/families/pivots.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 tier H P1 (owner GO 2026-10-03, owner channel point 1): create_pivot on the probe-measured path
 * (docs/evidence/M7-H3-PROBE.md). The fake builds the table exactly as P-A3 measured it: (2 + row keys + 1) x
 * (1 + column keys + 1) from `at`, keys ascending, sums and grand totals; the record carries its in-session public
 * source; a reopen is simulated by replacing it with the content-proof shape (pivot-source.cjs).
 */

const DATA = [['reg', 'prod', 'amt'], ['N', 'x', 1], ['S', 'y', 2], ['N', 'y', 3], ['S', 'x', 4], ['N', 'x', 5], ['S', 'y', 6]];
const cellsOf = rows => Object.fromEntries(rows.flatMap((r, i) => r.map((v, j) => [`${'ABC'[j]}${i + 1}`, { v }])));
const book = (rows = DATA) => [
  { name: 'Src', cells: cellsOf(rows) },
  { name: 'Host', cells: { A1: { v: 'keep' } } },
  { name: 'Other', cells: { A1: { v: 'x' } } },
];
const STEP = { intent: 'create_pivot', name: 'PE', sheet: 'Src', range: 'A1:C7', rows: 'reg', columns: 'prod', sum: 'amt', pivotSheet: 'Host', at: 'H3' };
const one = (over = {}) => ({ id: 'm7-p1', steps: [{ ...STEP, ...over }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const host = store => store.durable.find(s => s.name === 'Host');
const refused = async (store, t, reason, script) => {
  const before = JSON.stringify(store.durable);
  const r = await run(store, t, script);
  assert.equal(r.result.ok, false, reason);
  assert.equal(r.m.writes, 0, reason);
  assert.equal(r.driver.calls.dispatch, 0, reason);
  assert.match(JSON.stringify(r.result), new RegExp(reason));
  assert.equal(JSON.stringify(store.durable), before, `${reason}: nothing changed`);
};

/** The record a reopened session reads: the public source throws, the content proof stands in (P9). */
function reopen(store) {
  const h = host(store);
  const pv = h.objects.pivots.find(x => x.name === 'PE');
  const v = a => String((h.cells[a] || {}).v ?? '');
  const grid = (cs, rs) => rs.map(r => cs.map(c => v(`${c}${r}`)));
  h.objects.pivots = h.objects.pivots.map(x => (x !== pv ? x : { name: 'PE', table: pv.table, layout: pv.layout,
    proofSpec: { dataRef: 'Src!A1:C7', fields: ['reg', 'prod', 'amt'], rows: ['reg'], cols: ['prod'], data: ['Sum of amt'],
      body: { range: 'I5:K7', values: grid(['I', 'J', 'K'], [5, 6, 7]) }, table: { range: 'H3:K7', values: grid(['H', 'I', 'J', 'K'], [3, 4, 5, 6, 7]) } } }));
}

test('create_pivot: APPLIED (L1 identity + content, L3), the table as P-A3 measured, the retry SATISFIED 0/0 -- in session and after a reopen', async () => {
  const store = createStore(book());
  const a = await run(store, one());
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.preservationVerification, a.result.execution.appliedOperations]));
  assert.equal(a.m.commands, 1);
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  const h = host(store);
  assert.deepEqual(['H3', 'I3', 'H4', 'I4', 'J4', 'K4', 'H5', 'I5', 'J5', 'K5', 'H7', 'K7'].map(x => h.cells[x].v), ['Sum of amt', 'Column labels', 'Row Labels', 'x', 'y', 'Grand Total', 'N', '6', '3', '9', 'Grand Total', '21']);
  assert.equal(h.cells.A1.v, 'keep');
  const r = await run(store, one());
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  reopen(store);
  const z = await run(store, one());
  assert.equal(z.result.ok, true, JSON.stringify(z.result.execution.appliedOperations));
  assert.equal(z.m.commands, 0);
  assert.equal(z.m.writes, 0);
  assert.equal(z.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('refused before the write: a non-empty destination (P-A3: the insert overwrites it), an overlap with the source, a name taken by another identity, an unobservable pivots axis', async () => {
  const full = book();
  full[1].cells.J6 = { v: 'busy' };
  await refused(createStore(full), one(), 'xlsx-pivot-destination-not-empty');
  await refused(createStore(book()), one({ pivotSheet: 'Src', at: 'B2' }), 'xlsx-pivot-destination-overlaps-source');
  const taken = book();
  taken[2].objects = { charts: [], pivots: [{ name: 'PE', source: 'Other!$A$1:$C$7', table: 'H3:K7', layout: { rows: ['reg'], cols: ['prod'], data: ['Sum of amt'] } }], names: [] };
  await refused(createStore(taken), one(), 'xlsx-pivot-exists-with-another-identity');
  const twice = book();
  twice[2].objects = { charts: [], pivots: [{ name: 'PE', source: 'x' }, { name: 'PE', source: 'y' }], names: [] };
  await refused(createStore(twice), one(), 'xlsx-pivot-name-ambiguous');
  const blind = book();
  blind[2].objects = { charts: [], pivots: [{ name: 'Q', proofSpec: { dataRef: 'Src!A1:C7', fields: ['reg', 'prod', 'amt'], rows: ['reg'], cols: ['prod'], data: ['Sum of amt'], body: { range: 'C2:D3', values: [['1', '1'], ['1', '1']] }, table: { range: 'B1:D3', values: [['', '', ''], ['', '', ''], ['', '', '']] } } }], names: [] };
  await refused(createStore(blind), one(), 'xlsx-pivot-observation-incomplete');
});

test('refused before the write: an unmeasured source shape', async () => {
  const formula = book();
  formula[0].cells.C3 = { v: 2, f: '=1+1' };
  await refused(createStore(formula), one(), 'xlsx-pivot-source-shape-not-measured');
  await refused(createStore(book(DATA.map((r, i) => (i === 2 ? ['S', 'y', 'two'] : r)))), one(), 'xlsx-pivot-source-value-not-measured');
  await refused(createStore(book(DATA.map((r, i) => (i === 3 ? ['', 'y', 3] : r)))), one(), 'xlsx-pivot-source-blank-key-not-measured');
  await refused(createStore(book([['reg', 'reg', 'amt'], ...DATA.slice(1)])), one({ rows: 'reg', columns: 'prod' }), 'xlsx-pivot-source-header-not-unique-text');
  await refused(createStore(book()), one({ columns: 'region' }), 'xlsx-pivot-field-not-in-header');
});

test('planning: only the measured shapes', () => {
  const bad = (over, re) => { const p = planOperations(one(over)); assert.equal(p.ok, false, JSON.stringify(over)); assert.match(p.detail, re); };
  bad({ pivotSheet: 'My host' }, /need no quoting/);
  bad({ name: '1x' }, /pivot name/);
  bad({ columns: 'reg' }, /three different header names/);
  bad({ range: 'A1:C1' }, /data rows/);
  bad({ range: `A1:C${pivots.MAX_DATA_ROWS + 2}` }, /data rows/);
  bad({ range: 'A1:B7' }, /three columns/);
  bad({ at: 'H3:H4' }, /one cell/);
  bad({ data: 'Count' }, /does not support field/);
  bad({ host: 'Other' }, /does not support field/); // the pre-rename name does not silently keep working
  assert.equal(planOperations({ id: 'x', steps: [STEP, { ...STEP, at: 'Q3' }] }).ok, false, 'two steps on one pivot name clash');
});

test('RED: a wrong sum / the name not set / ignored -> L1 FAIL; a cell beside the table / another object touched -> L3 FAIL', async () => {
  for (const fault of ['pivot-wrong-sum', 'pivot-not-named', 'ignore']) {
    const { result, driver } = await run(createStore(book()), one(), { fault });
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  for (const fault of ['pivot-overwrites-neighbour', 'pivot-touches-other-object', 'side-effect']) {
    const { result } = await run(createStore(book()), one(), { fault });
    assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', fault);
  }
});

test('the command: name check, the destination re-checked at dispatch, then the measured call sequence', () => {
  const op = planOperations(one()).operations[0];
  const cmd = pivots.lower(op, { extent: 'H3:K7' });
  const log = [];
  const pv = { SetName: x => log.push(['name', x]), AddFields: x => log.push(['fields', x]), AddDataField: x => log.push(['data', x]), RefreshTable: () => log.push(['refresh']) };
  const sheet = (n, extent) => ({ GetName: () => n, GetAllPivotTables: () => [], GetRange: a => ({ a, GetValue: () => extent }) });
  const api = extent => ({ GetSheet: n => sheet(n, extent), GetSheets: () => [sheet('Src'), sheet('Host')], InsertPivotExistingWorksheet: (s, d) => { log.push(['insert', s.a, d.a]); return pv; } });
  const body = new Function('Api', cmd.body);
  assert.deepEqual(body(api([['', ''], ['', '']])), { ok: true });
  assert.deepEqual(log, [['insert', 'A1:C7', 'H3'], ['name', 'PE'], ['fields', { rows: 'reg', columns: 'prod' }], ['data', 'amt'], ['refresh']]);
  log.length = 0;
  assert.deepEqual(body(api([['', 'late'], ['', '']])), { ok: false, stage: 'destination-not-empty' });
  assert.deepEqual(log, [], 'nothing inserted over a cell that filled since the read');
});

test('identity is source + extent + content: a pivot of the name on the host that differs in ONE of them is refused, not satisfied', async () => {
  // build the real one first, then change one axis of its record or its cells
  const variants = [
    ['another source', pv => ({ ...pv, source: 'Other!$A$1:$C$7' })],
    ['another extent', pv => ({ ...pv, table: 'H3:K8' })],
  ];
  for (const [label, edit] of variants) {
    const store = createStore(book());
    await run(store, one());
    const h = host(store);
    h.objects.pivots = h.objects.pivots.map(x => (x.name === 'PE' ? edit(x) : x));
    await refused(store, one(), 'xlsx-pivot-exists-with-another-identity');
    assert.ok(true, label);
  }
  const store = createStore(book());
  await run(store, one());
  host(store).cells.I5 = { v: '7' };
  await refused(store, one(), 'xlsx-pivot-exists-with-another-identity');
});

test('the command refuses a name that appeared since the read', () => {
  const cmd = pivots.lower(planOperations(one()).operations[0], { extent: 'H3:K7' });
  const sheet = n => ({ GetName: () => n, GetAllPivotTables: () => (n === 'Other' ? [{ GetName: () => 'PE' }] : []), GetRange: () => ({ GetValue: () => '' }) });
  const api = { GetSheet: sheet, GetSheets: () => ['Src', 'Host', 'Other'].map(sheet), InsertPivotExistingWorksheet: () => { throw new Error('must not insert'); } };
  assert.deepEqual(new Function('Api', cmd.body)(api), { ok: false, stage: 'name-taken' });
});
