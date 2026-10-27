'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const refresh = require('../../src/adapters/xlsx/families/pivot-refresh.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 tier H P3 (owner GO 2026-10-04, owner channel): refresh_pivot { pivot, sourceSheet, sourceRange } on the
 * public SetSource + RefreshTable path (m7-p3c-probe4). The pivot is made by create_pivot (the fake's measured table);
 * the fake's refresh rebinds to the NAMED range and rebuilds the table at the same anchor. A reopen replaces the record
 * with the content-proof shape; a STALE reopened pivot fails that proof, and refresh still finds it by its raw identity.
 */

const DATA = [['reg', 'prod', 'amt'], ['N', 'x', 1], ['S', 'y', 2], ['N', 'y', 3], ['S', 'x', 4], ['N', 'x', 5], ['S', 'y', 6]];
const MORE = [['reg', 'prod', 'amt'], ['N', 'x', 10], ['S', 'y', 20], ['W', 'z', 30], ['N', 'z', 40], ['S', 'x', 50], ['W', 'y', 60]];
const FEW = [['reg', 'prod', 'amt'], ['N', 'x', 7], ['N', 'x', 8]];
const cellsOf = rows => Object.fromEntries(rows.flatMap((r, i) => r.map((v, j) => [`${'ABC'[j]}${i + 1}`, { v }])));
const book = () => [
  { name: 'Src', cells: cellsOf(DATA) },
  { name: 'Src2', cells: cellsOf(MORE) },
  { name: 'Host', cells: { A1: { v: 'keep' } } },
  { name: 'Other', cells: { A1: { v: 'x' } } },
];
const CREATE = { intent: 'create_pivot', name: 'PE', sheet: 'Src', range: 'A1:C7', rows: 'reg', columns: 'prod', sum: 'amt', pivotSheet: 'Host', at: 'H3' };
const create = (over = {}) => ({ id: 'm7-p3-seed', steps: [{ ...CREATE, ...over }] });
const STEP = { intent: 'refresh_pivot', pivot: 'PE', sourceSheet: 'Src', sourceRange: 'A1:C7' };
const one = (over = {}) => ({ id: 'm7-p3', steps: [{ ...STEP, ...over }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const sheet = (store, name) => store.durable.find(s => s.name === name);
const status = r => r.result.execution.appliedOperations[0].status;
const why = r => JSON.stringify([r.result.execution.primitiveVerification, r.result.execution.preservationVerification, r.result.execution.appliedOperations]);
async function seeded(over) {
  const store = createStore(book());
  assert.equal((await run(store, create(over))).result.ok, true);
  return store;
}
const setRows = (store, name, rows) => { sheet(store, name).cells = { ...sheet(store, name).cells, ...cellsOf(rows) }; };
const tableOf = store => sheet(store, 'Host').objects.pivots.find(p => p.name === 'PE');
const refused = async (store, t, reason, script) => {
  const before = JSON.stringify(store.durable);
  const r = await run(store, t, script);
  assert.equal(r.result.ok, false, reason);
  assert.equal(r.m.writes, 0, reason);
  assert.equal(r.driver.calls.dispatch, 0, reason);
  assert.match(JSON.stringify(r.result), new RegExp(reason));
  assert.equal(JSON.stringify(store.durable), before, `${reason}: nothing changed`);
};

/** A reopened session's record: the public source throws, the proof's dataRef and fields are the raw identity. */
function reopen(store, dataRef = 'Src!$A$1:$C$7') {
  const h = sheet(store, 'Host');
  h.objects.pivots = h.objects.pivots.map(x => {
    const t = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(x.table);
    const cols = []; for (let c = t[1].charCodeAt(0); c <= t[3].charCodeAt(0); c += 1) cols.push(String.fromCharCode(c));
    const v = a => String((h.cells[a] || {}).v ?? '');
    const grid = (cs, r1, r2) => { const out = []; for (let r = r1; r <= r2; r += 1) out.push(cs.map(c => v(`${c}${r}`))); return out; };
    return { name: x.name, table: x.table, layout: x.layout, proofSpec: { dataRef: x.dataRef || dataRef, fields: x.fields, rows: x.layout.rows, cols: x.layout.cols, data: x.layout.data,
      body: { range: `${cols[1]}${Number(t[2]) + 2}:${t[3]}${t[4]}`, values: grid(cols.slice(1), Number(t[2]) + 2, Number(t[4])) }, table: { range: x.table, values: grid(cols, Number(t[2]), Number(t[4])) } } };
  });
}

test('refresh_pivot: a stale pivot (its source edited) is refreshed APPLIED (exact dataRef + table content, L3); the retry SATISFIED 0/0, also reopened', async () => {
  const store = await seeded();
  setRows(store, 'Src', [DATA[0], ['N', 'x', 100]]);
  const a = await run(store, one());
  assert.equal(a.result.ok, true, why(a));
  assert.deepEqual([a.m.writes, a.m.commands, status(a)], [1, 1, OPERATION_STATUS.APPLIED]);
  assert.equal(sheet(store, 'Host').cells.I5.v, '105', 'N/x = 100 + 5');
  assert.equal(sheet(store, 'Host').cells.A1.v, 'keep');
  const r = await run(store, one());
  assert.deepEqual([r.m.writes, r.m.commands, status(r)], [0, 0, OPERATION_STATUS.SATISFIED]);
  reopen(store);
  const z = await run(store, one());
  assert.equal(z.result.ok, true, why(z));
  assert.deepEqual([z.m.writes, z.m.commands, status(z)], [0, 0, OPERATION_STATUS.SATISFIED]);
});

test('a STALE REOPENED pivot (it fails the content proof) is still found by its raw identity and refreshed; the retry 0/0', async () => {
  const store = await seeded();
  reopen(store);
  setRows(store, 'Src', [DATA[0], ['S', 'y', 1000]]);
  const a = await run(store, one());
  assert.equal(a.result.ok, true, why(a));
  assert.equal(status(a), OPERATION_STATUS.APPLIED);
  reopen(store);
  const z = await run(store, one());
  assert.deepEqual([z.m.writes, z.m.commands, status(z)], [0, 0, OPERATION_STATUS.SATISFIED]);
});

test('rebind to ANOTHER named sheet: the table grows into empty cells, the dataRef names Src2; next to a same-named-content impostor the named sheet wins', async () => {
  const store = await seeded();
  sheet(store, 'Other').cells = { ...sheet(store, 'Other').cells, ...cellsOf(MORE) }; // an impostor of Src2's content
  const a = await run(store, one({ sourceSheet: 'Src2', sourceRange: 'A1:C7' }));
  assert.equal(a.result.ok, true, why(a));
  assert.equal(tableOf(store).dataRef, 'Src2!$A$1:$C$7');
  assert.equal(tableOf(store).table, 'H3:L8', '3 row keys x 3 column keys');
  const r = await run(store, one({ sourceSheet: 'Src2', sourceRange: 'A1:C7' }));
  assert.deepEqual([r.m.commands, status(r)], [0, OPERATION_STATUS.SATISFIED]);
  const back = await run(store, one());
  assert.equal(back.result.ok, true, why(back));
  assert.equal(tableOf(store).table, 'H3:K7', 'shrunk back');
  assert.equal(sheet(store, 'Host').cells.L8, undefined, 'the vacated cells are blank');
});

test('a shrink leaves the vacated cells blank -- the fake that does not clear them (refresh-leaves-stale) fails L1', async () => {
  const store = await seeded();
  setRows(store, 'Src', FEW);
  const r = await run(store, one({ sourceRange: 'A1:C3' }), { fault: 'refresh-leaves-stale' });
  assert.equal(r.result.ok, false, 'stale cells left');
  assert.match(why(r), /vacatedNotBlank":\["[A-Z]/);
  const store2 = await seeded();
  setRows(store2, 'Src', FEW);
  const ok = await run(store2, one({ sourceRange: 'A1:C3' }));
  assert.equal(ok.result.ok, true, why(ok));
  assert.equal(tableOf(store2).table, 'H3:J6', 'one row key x one column key: title, labels, the key, the total');
  assert.equal(sheet(store2, 'Host').cells.K7, undefined);
});

test('another pivot (another cache) is unchanged: refreshing PE leaves PQ\'s identity and table', async () => {
  const store = await seeded();
  assert.equal((await run(store, create({ name: 'PQ', at: 'H20' }))).result.ok, true);
  const pq = JSON.stringify(sheet(store, 'Host').objects.pivots.find(p => p.name === 'PQ'));
  setRows(store, 'Src', [DATA[0], ['N', 'x', 100]]);
  const a = await run(store, one());
  assert.equal(a.result.ok, true, why(a));
  assert.equal(JSON.stringify(sheet(store, 'Host').objects.pivots.find(p => p.name === 'PQ')), pq);
});

test('refused before the write: no such pivot, an ambiguous name, a header that is not the pivot\'s fields, a grown extent over a non-empty cell, an extent over the source, an unread raw identity', async () => {
  await refused(await seeded(), one({ pivot: 'Nope' }), 'xlsx-pivot-not-found');
  const amb = await seeded();
  assert.equal((await run(amb, create({ pivotSheet: 'Other', at: 'H3' }))).result.ok, false, 'create refuses the taken name');
  sheet(amb, 'Other').objects = { charts: [], names: [], pivots: [{ ...tableOf(amb) }] };
  await refused(amb, one(), 'xlsx-pivot-name-ambiguous');
  const hdr = await seeded();
  setRows(hdr, 'Src2', [['reg', 'product', 'amt']]);
  await refused(hdr, one({ sourceSheet: 'Src2' }), 'xlsx-pivot-source-header-mismatch');
  const full = await seeded();
  sheet(full, 'Host').cells.L8 = { v: 'busy' };
  await refused(full, one({ sourceSheet: 'Src2' }), 'xlsx-pivot-destination-not-empty');
  const over = await seeded({ pivotSheet: 'Src2', at: 'A20' }); // table A20:D24
  sheet(over, 'Src2').cells = { ...sheet(over, 'Src2').cells, ...Object.fromEntries(MORE.flatMap((r, i) => r.map((v, j) => [`${'CDE'[j]}${22 + i}`, { v }]))) };
  await refused(over, one({ sourceSheet: 'Src2', sourceRange: 'C22:E28' }), 'xlsx-pivot-destination-overlaps-source');
  const raw = await seeded();
  delete tableOf(raw).dataRef;
  await refused(raw, one(), 'xlsx-pivot-observation-incomplete');
});

test('planning: only the measured shapes -- a name that needs quoting (the public setter\'s silent no-op), a bad range, an extra field', () => {
  const bad = [
    [{ sourceSheet: 'My sheet' }, /needs no quoting/],
    [{ sourceSheet: "O'Brien" }, /needs no quoting/],
    [{ sourceRange: 'A1:B7' }, /at least three columns/],
    [{ sourceRange: 'A1:C1' }, /header row/],
    [{ pivot: '1x' }, /pivot name/],
    [{ host: 'Host' }, /does not support field/],
  ];
  for (const [over, re] of bad) {
    const p = planOperations(one(over));
    assert.equal(p.ok, false, JSON.stringify(over));
    assert.match(p.detail, re);
  }
  assert.equal(planOperations(one()).ok, true);
});

test('RED: the setter a no-op / a wrong sum -> L1 FAIL; another object / another sheet touched -> L3 FAIL', async () => {
  for (const [fault, re] of [['refresh-no-rebind', /dataRef/], ['pivot-wrong-sum', /pivot-body-differs|reproduces/], ['pivot-touches-other-object', /otherObjectsKept/], ['side-effect', /Z99|contentChanged|changed/]]) {
    const store = await seeded();
    const r = await run(store, one({ sourceSheet: 'Src2' }), { fault });
    assert.equal(r.result.ok, false, fault);
    assert.match(why(r), re, fault);
  }
});

test('the command: the pivot found by name on its host, the source sheet checked, the growth re-checked, then SetSource + RefreshTable', () => {
  const op = planOperations(one({ sourceSheet: 'Src2' })).operations[0];
  const target = { sheet: 'Src2', pivot: 'PE', host: 'Host', extent: 'H3:L8', old: 'H3:K7', growth: ['L3', 'H8'] };
  const body = refresh.lower(op, target).body;
  const order = ['pivot-identity', 'sheet-identity', 'destination-not-empty', 'SetSource(src.GetRange("A1:C7"))', 'RefreshTable()'].map(x => body.indexOf(x));
  assert.ok(order.every(i => i > 0), JSON.stringify(order));
  assert.deepEqual([...order].sort((a, b) => a - b), order);
  const calls = [];
  const range = v => ({ GetValue: () => v });
  const pv = { GetName: () => 'PE', SetSource: r => calls.push(['SetSource', r]), RefreshTable: () => calls.push(['RefreshTable']) };
  const ws = (name, extra = {}) => ({ GetName: () => name, GetAllPivotTables: () => (name === 'Host' ? [pv] : []), GetRange: a => ({ ...range(extra[a] ?? ''), a }) });
  const sheets = [ws('Src2'), ws('Host')];
  const Api = { GetSheets: () => sheets, GetSheet: n => sheets.find(s => s.GetName() === n) };
  assert.deepEqual(new Function('Api', body)(Api), { ok: true });
  assert.deepEqual(calls.map(c => c[0]), ['SetSource', 'RefreshTable']);
  assert.equal(calls[0][1].a, 'A1:C7');
  sheets[1] = ws('Host', { H8: 'busy' });
  assert.deepEqual(new Function('Api', body)(Api), { ok: false, stage: 'destination-not-empty', cell: 'H8' });
});

test('identity is the EXACT dataRef, never the content: a sheet of IDENTICAL content is APPLIED (not SATISFIED) and bound by name; a setter that silently keeps the old reference fails L1', async () => {
  const store = await seeded();
  sheet(store, 'Other').cells = cellsOf(DATA);
  const a = await run(store, one({ sourceSheet: 'Other' }));
  assert.equal(a.result.ok, true, why(a));
  assert.deepEqual([a.m.commands, status(a)], [1, OPERATION_STATUS.APPLIED]);
  assert.equal(tableOf(store).dataRef, 'Other!$A$1:$C$7');
  const noop = await seeded();
  sheet(noop, 'Other').cells = cellsOf(DATA);
  const r = await run(noop, one({ sourceSheet: 'Other' }), { fault: 'refresh-no-rebind' });
  assert.equal(r.result.ok, false, 'the table matches, the reference does not');
  assert.match(why(r), /"dataRef":"Src!\$A\$1:\$C\$7","want":"Other!\$A\$1:\$C\$7"/);
});

test('pivotRefsOf: the raw identity per sheet; any unread part makes the sheet a reason string, never a shorter list', () => {
  const { pivotRefsOf } = require('../../src/adapters/xlsx/pivot-refs.cjs');
  const pv = { name: 'P', dataRef: 'S!$A$1:$C$7', fields: ['a', 'b', 'c'], table: 'H3:K7', layout: { rows: ['a'], cols: ['b'], data: ['Sum of c'] }, source: 'x' };
  assert.deepEqual(pivotRefsOf({ H: { pivots: [pv] } }), { H: [{ name: 'P', dataRef: 'S!$A$1:$C$7', fields: ['a', 'b', 'c'], table: 'H3:K7', layout: pv.layout }] });
  for (const bad of [{ dataRef: 'UNOBSERVABLE:threw' }, { dataRef: undefined }, { fields: 'UNOBSERVABLE:x' }, { table: 'UNOBSERVABLE:x' }, { layout: { rows: ['a'], cols: 'UNOBSERVABLE:x', data: [] } }]) {
    assert.match(String(pivotRefsOf({ H: { pivots: [pv, { ...pv, name: 'Q', ...bad }] } }).H), /^UNOBSERVABLE:pivot-ref-unread \(pivot Q\)$/, JSON.stringify(bad));
  }
  assert.equal(pivotRefsOf({ H: { pivots: 'UNOBSERVABLE:threw' } }).H, 'UNOBSERVABLE:threw');
  assert.deepEqual(pivotRefsOf(undefined), {});
});

test('the driver reads the raw identity (pivot.state.read/v1 + GetPivotFields) for every pivot, through the DS Bridge only (B13 cutover)', () => {
  const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  const q = readQuery([], { objectSheets: ['*'] });
  assert.match(q, /"pivot\.state\.read\/v1"/);
  assert.match(q, /out\.dataRef = typeof st === 'string' \? st : st\.dataRef/);
  assert.match(q, /out\.fields = fieldNames\(\(\) => p\.GetPivotFields\(\)\)/);
  assert.doesNotMatch(q, /asc_getDataRef|p\.pivot\b/, 'the direct internal read is gone, no fallback');
});
