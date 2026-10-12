'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { provenSource } = require('../../src/adapters/xlsx/pivot-source.cjs');

/**
 * M7 tier H P-A / P9 (owner GO 2026-10-03, owner channel, point 1): a reopened pivot's source proven by content.
 * The shapes are the P-A2 probe's (docs/evidence/M7-H2-PROBE.md, m7-h2-probe1-20261004.json): source HSA!A1:C7 (reg,
 * prod, amt), pivot PE on HHX!H3, body I5:K7 = [[6,3,9],[4,8,12],[10,11,21]]; after a rename of the source the data
 * reference still names the OLD sheet, and an impostor of that name then answers for it. The label cells around the
 * body (row labels left of it, column labels above it) were NOT read by the probe -- the LIVE acceptance measures them.
 */

// the DS runtime identity (manifest version|patchVersion), never the URL cache key (owner decision (B), 9580767b)
const BUILD = '9.3.4|euro-office-bridge-v1/9.3.4.60/5';
const DATA = [['reg', 'prod', 'amt'], ['N', 'x', '1'], ['S', 'y', '2'], ['N', 'y', '3'], ['S', 'x', '4'], ['N', 'x', '5'], ['S', 'y', '6']];
const IMPOSTOR = DATA.map((r, i) => (i ? [r[0], r[1], String(Number(r[2]) * 100)] : r));
const BODY = [['6', '3', '9'], ['4', '8', '12'], ['10', '11', '21']];
const TABLE = [['Sum of amt', 'Column Labels', '', ''], ['Row Labels', 'x', 'y', 'Grand Total'], ['N', ...BODY[0]], ['S', ...BODY[1]], ['Grand Total', ...BODY[2]]];
const SPEC = { fields: ['reg', 'prod', 'amt'], rows: ['reg'], cols: ['prod'], data: ['Sum of amt'], body: { range: 'I5:K7', values: BODY }, table: { range: 'H3:K7', values: TABLE } };
const EMPTY = [['', '', ''], ['', '', ''], ['', '', ''], ['', '', ''], ['', '', ''], ['', '', ''], ['', '', '']];

/** A proof as the driver reads it after a reopen: the source and every other sheet at the same address. */
const proof = (sheets, over = {}) => ({ ...SPEC, dataRef: 'HSA!A1:C7', source: { sheet: 'HSA', range: 'A1:C7', values: sheets.HSA }, others: Object.keys(sheets).filter(n => n !== 'HSA').map(n => ({ sheet: n, values: sheets[n] })), ...over });
const names = sheets => Object.keys(sheets);
const why = (p, sheets, build = BUILD) => provenSource(p, build, names(sheets));

test('the measured shape: the source reproduces the cached body, no other sheet does -> HSA!$A$1:$C$7', () => {
  const book = { HSA: DATA, HSB: EMPTY, HHX: EMPTY };
  assert.equal(why(proof(book), book), 'HSA!$A$1:$C$7');
  // the label order does not matter, only that body and labels agree (rows S, N; columns y, x)
  const swapped = { ...SPEC.table, values: [TABLE[0], ['Row Labels', 'y', 'x', 'Grand Total'], ['S', '8', '4', '12'], ['N', '3', '6', '9'], ['Grand Total', '11', '10', '21']] };
  assert.equal(why(proof(book, { table: swapped, body: { range: 'I5:K7', values: [['8', '4', '12'], ['3', '6', '9'], ['11', '10', '21']] } }), book), 'HSA!$A$1:$C$7');
  // a quoted sheet name round-trips quoted
  const q = { 'Bev ét': DATA, HHX: EMPTY };
  assert.equal(provenSource({ ...proof(q), dataRef: "'Bev ét'!$A$1:$C$7", source: { sheet: 'Bev ét', range: 'A1:C7', values: DATA }, others: [{ sheet: 'HHX', values: EMPTY }] }, BUILD, names(q)), "'Bev ét'!$A$1:$C$7");
});

test('P-A2 shapes after a source rename are NOT the source: an impostor of the old name, no sheet of the old name', () => {
  // HSA renamed to HSA2, a new HSA with other values: the data reference still names HSA -> the body differs
  const impostor = { HSA: IMPOSTOR, HSA2: DATA, HHX: EMPTY };
  assert.equal(why(proof(impostor), impostor), 'UNOBSERVABLE:pivot-body-differs');
  // renamed, nothing named HSA any more
  const gone = { HSA2: DATA, HHX: EMPTY };
  assert.equal(provenSource({ ...proof({ HSA: DATA }), others: [] }, BUILD, names(gone)), 'UNOBSERVABLE:pivot-source-sheet-missing:HSA');
});

test('ambiguity refuses: a renamed source next to an impostor of the SAME content, or a copy of the source sheet', () => {
  const twin = { HSA: DATA, HSA2: DATA, HHX: EMPTY };
  assert.equal(why(proof(twin), twin), 'UNOBSERVABLE:pivot-source-ambiguous:HSA2');
});

test('fail-closed: unverified build, a stale cache, unmeasured layouts, unread parts', () => {
  const book = { HSA: DATA, HSB: EMPTY, HHX: EMPTY };
  assert.equal(why(proof(book), book, '9.3.4|euro-office-bridge-v1/9.3.4.60/2'), 'UNOBSERVABLE:pivot-runtime-not-verified');
  assert.equal(why('UNOBSERVABLE:asc_getDataRef-threw', book), 'UNOBSERVABLE:asc_getDataRef-threw');
  // the source edited after the last refresh: the cache no longer matches -> not proven
  const stale = { ...book, HSA: DATA.map((r, i) => (i === 1 ? ['N', 'x', '7'] : r)) };
  assert.equal(why(proof(stale), stale), 'UNOBSERVABLE:pivot-body-differs');
  assert.equal(why(proof(book, { rows: ['reg', 'prod'] }), book), 'UNOBSERVABLE:pivot-layout-not-measured');
  assert.equal(why(proof(book, { data: ['Count of amt'] }), book), 'UNOBSERVABLE:pivot-layout-not-measured');
  assert.equal(why(proof(book, { cols: [] }), book), 'UNOBSERVABLE:pivot-layout-not-measured');
  assert.equal(why(proof(book, { cols: ['reg'] }), book), 'UNOBSERVABLE:pivot-layout-not-measured');
  assert.equal(why(proof(book, { cols: ['region'] }), book), 'UNOBSERVABLE:pivot-layout-not-measured');
  const renamedHeader = { ...book, HSA: [['reg', 'prod', 'sum'], ...DATA.slice(1)] };
  assert.equal(why(proof(renamedHeader), renamedHeader), 'UNOBSERVABLE:pivot-source-header-mismatch');
  const blank = { ...book, HSA: DATA.map((r, i) => (i === 2 ? ['', 'y', '2'] : r)) };
  assert.equal(why(proof(blank), blank), 'UNOBSERVABLE:pivot-source-blank-key-not-measured');
  const text = { ...book, HSA: DATA.map((r, i) => (i === 2 ? ['S', 'y', 'two'] : r)) };
  assert.equal(why(proof(text), text), 'UNOBSERVABLE:pivot-source-value-not-measured');
  assert.equal(why(proof(book, { source: { sheet: 'HSB', range: 'A1:C7', values: DATA } }), book), 'UNOBSERVABLE:pivot-source-read-differs-from-dataref');
  assert.equal(why(proof(book, { dataRef: 'HSA!A:C' }), book), 'UNOBSERVABLE:pivot-dataref-shape');
  assert.equal(why(proof(book, { others: [{ sheet: 'HSB', values: EMPTY }] }), book), 'UNOBSERVABLE:pivot-candidates-incomplete');
  assert.equal(why(proof(book, { others: [{ sheet: 'HSB', values: 'UNOBSERVABLE:threw' }, { sheet: 'HHX', values: EMPTY }] }), book), 'UNOBSERVABLE:pivot-candidate-unread:HSB');
  assert.equal(why(proof(book, { body: { range: 'I5:K7', values: 'THREW' } }), book), 'UNOBSERVABLE:pivot-body-unread');
  assert.equal(why(proof(book, { body: { range: 'G5:K7', values: BODY } }), book), 'UNOBSERVABLE:pivot-extent-shape');
});

test('every total is checked: a row total, a column total, the grand total, an extra label', () => {
  const book = { HSA: DATA, HHX: EMPTY };
  const tamper = (i, j, v) => { const b = BODY.map(r => r.slice()); b[i][j] = v; const t = TABLE.map(r => r.slice()); t[i + 2][j + 1] = v; return { body: { range: 'I5:K7', values: b }, table: { range: 'H3:K7', values: t } }; };
  assert.equal(why(proof(book, tamper(0, 2, '10')), book), 'UNOBSERVABLE:pivot-row-total-differs');
  assert.equal(why(proof(book, tamper(2, 0, '11')), book), 'UNOBSERVABLE:pivot-column-total-differs');
  assert.equal(why(proof(book, tamper(2, 2, '22')), book), 'UNOBSERVABLE:pivot-grand-total-differs');
  // a combination absent from the source is an EMPTY body cell -- a value there is not the source's
  const sparse = { HSA: DATA.filter(r => !(r[0] === 'S' && r[1] === 'x')), HHX: EMPTY };
  const sb = [['6', '3', '9'], ['', '8', '8'], ['6', '11', '17']];
  const st = [TABLE[0], TABLE[1], ['N', ...sb[0]], ['S', ...sb[1]], ['Grand Total', ...sb[2]]];
  const sp = (b, t2) => ({ ...proof(sparse), dataRef: 'HSA!A1:C5', source: { sheet: 'HSA', range: 'A1:C5', values: sparse.HSA }, others: [{ sheet: 'HHX', values: EMPTY.slice(0, 5) }], body: { range: 'I5:K7', values: b }, table: { range: 'H3:K7', values: t2 } });
  assert.equal(why(sp(sb, st), sparse), 'HSA!$A$1:$C$5');
  const zb = sb.map(r => r.slice());
  zb[1][0] = '0';
  assert.equal(why(sp(zb, [st[0], st[1], st[2], ['S', ...zb[1]], st[4]]), sparse), 'UNOBSERVABLE:pivot-body-differs');
  const t = TABLE.map(r => r.slice());
  t[3][0] = 'W';
  assert.equal(why(proof(book, { table: { range: 'H3:K7', values: t } }), book), 'UNOBSERVABLE:pivot-labels-differ');
});

/* --- the product path: an UNRELATED pivot no longer blocks a delete; the source sheet still does --- */

const store = (hsa = DATA) => createStore([
  { name: 'HSA', cells: Object.fromEntries(hsa.flatMap((r, i) => r.map((v, j) => [`${'ABC'[j]}${i + 1}`, { v }]))) },
  { name: 'HHX', cells: { A1: { v: 'host' } }, objects: { charts: [], pivots: [{ name: 'PE', proofSpec: { ...SPEC, dataRef: 'HSA!A1:C7' } }], names: [] } },
  { name: 'Notes', cells: { A1: { v: 'free text' } } },
]);

async function run(s, step, script = {}) {
  let n = 0;
  // M11.5c.4: copy_sheet writes only through the /8 Bridge WRITE -- a copy task runs on /8 unless the test names a runtime
  const driver = createFakeXlsxDriver(s, step.intent === 'copy_sheet' && script.runtime === undefined ? { ...script, runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/8' } : script);
  const t = { id: 'm7-p9', steps: [step] };
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

test('P9: delete_sheet of a sheet the reopened pivot neither reads nor lives on -> APPLIED (was refused: the source unobservable)', async () => {
  const s = store();
  const { result, m } = await run(s, { intent: 'delete_sheet', sheet: 'Notes' });
  assert.equal(result.ok, true, JSON.stringify(result.execution && (result.execution.refusal || result.execution.primitiveVerification)));
  assert.equal(m.writes, 1);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.deepEqual(s.durable.map(x => x.name), ['HSA', 'HHX']);
});

test('P9: the proven source sheet is a dependent -> delete_sheet HSA refused before any write', async () => {
  const { result, driver, m } = await run(store(), { intent: 'delete_sheet', sheet: 'HSA' });
  assert.equal(result.ok, false);
  assert.equal(m.writes, 0);
  assert.equal(driver.calls.dispatch, 0);
  assert.match(JSON.stringify(result), /xlsx-delete-sheet-referenced-by-dependents/);
});

test('P9: an unproven source keeps every refusal -- an impostor-shaped source, an unread proof, an unverified build', async () => {
  for (const [label, s, script, reason] of [
    ['source values do not reproduce the cache', store(IMPOSTOR), {}, /pivot-body-differs/],
    ['the proof reader failed', store(), { pivotProofUnobservable: 'UNOBSERVABLE:asc_getDataRef-threw' }, /asc_getDataRef-threw/],
    ['another build', store(), { runtime: '9.3.5-0000000000000000' }, /pivot-runtime-not-verified/],
  ]) {
    const { result, driver, m } = await run(s, { intent: 'delete_sheet', sheet: 'Notes' }, script);
    assert.equal(result.ok, false, label);
    assert.equal(m.writes, 0, label);
    assert.equal(driver.calls.dispatch, 0, label);
    assert.match(JSON.stringify(result), /xlsx-delete-dependency-observation-incomplete/, label);
    assert.match(JSON.stringify(result), reason, label);
  }
});

test('P9: an unproven pivot keeps the whole pivots axis a gap -- a copy / rename of ANY sheet still refuses (W3 LIVE regression m7-h-regr-w3-20261004: the copy APPLIED)', async () => {
  for (const [label, step, reason] of [
    ['copy of the source sheet', { intent: 'copy_sheet', sheet: 'HSA', name: 'HSA_copy' }, /xlsx-copy-source-dependency-unobservable/],
    ['copy of an unrelated sheet', { intent: 'copy_sheet', sheet: 'Notes', name: 'Notes_copy' }, /xlsx-copy-/],
    ['rename of an unrelated sheet', { intent: 'rename_sheet', sheet: 'Notes', name: 'Notes2' }, /xlsx-rename-dependency-observation-incomplete/],
  ]) {
    const { result, driver, m } = await run(store(IMPOSTOR), step);
    assert.equal(result.ok, false, label);
    assert.equal(m.writes, 0, label);
    assert.equal(driver.calls.dispatch, 0, label);
    assert.match(JSON.stringify(result), reason, label);
  }
  // and the proven case lets the unrelated copy through
  const { result } = await run(store(), { intent: 'copy_sheet', sheet: 'Notes', name: 'Notes_copy' });
  assert.equal(result.ok, true, JSON.stringify(result.execution && result.execution.appliedOperations));
});
