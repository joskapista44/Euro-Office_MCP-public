'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const column = require('../../src/adapters/docx/families/table-column.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 4c offline acceptance: table-column (insert_table_column_before / _after,
 * delete_table_column) through the TARGET path on the fake transport whose column model follows
 * column probe 1 (2026-09-24): AddColumn puts one empty column at the boundary in every row, a
 * colspan straddling the boundary gets the new cell before it (drift); RemoveColumn removes the
 * column, a colspan covering it is lost (content loss). The adapter refuses both before dispatch.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (text, span = 1) => ({ span, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, rows, cols = rows) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (__, c) => cell(`${tag} r${r}c${c}`))) });
/** 3x5, row 1 merges columns 1..3 (colspan 3) -- the probe's K table */
const kTable = tag => { const t = grid(tag, 3, 5); t.rows[1] = [cell(`${tag} r1c0`), cell(`${tag} r1c1 ${tag} r1c2 ${tag} r1c3`, 3), cell(`${tag} r1c4`)]; return t; };
/** the public Split's shape (probe G): rows 0 and 2 carry a colspan-2 cell over row 1's two halves */
const gTable = tag => { const t = { rows: [[cell(`${tag} r0c0`), cell(`${tag} r0c1`, 2), cell(`${tag} r0c2`)], [cell(`${tag} r1c0`), cell(`${tag} r1c1`), cell(''), cell(`${tag} r1c2`)], [cell(`${tag} r2c0`), cell(`${tag} r2c1`, 2), cell(`${tag} r2c2`)]] }; return t; };
const paragraphs = () => [{ text: 'Column heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Between.', props: { ...PLAIN } }];
const store = (tables = [grid('A', 3)], paras = paragraphs()) => createStore(paras, { doc: { tables } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1);
  assert.equal(driver.calls.close, 1);
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w4c', steps });
const ins = (side, text, terminalColumnCount, tableAnchor) => ({ intent: `insert_table_column_${side}`, selector: tableAnchor ? { tableCell: { text, tableAnchor } } : { text }, terminalColumnCount });
const del = (text, terminalColumnCount, tableAnchor) => ({ intent: 'delete_table_column', selector: { tableCell: { text, tableAnchor } }, terminalColumnCount });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const texts = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.paras.map(p => p.text).join(' ')));
const spans = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.span || 1));

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per column operation');
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.equal(first.x.preservationVerification.length, t.steps.length);
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(s, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

const row = (r, cols) => cols.map(c => (c === null ? '' : `A r${r}c${c}`));
const rows3 = cols => [0, 1, 2].map(r => row(r, cols));

test('rectangular 3x3: before/after on the FIRST, MIDDLE and LAST column; delete of each -- FIRST applies, SECOND is 0/0; order and direction exact', async () => {
  const cases = [
    [ins('before', 'A r0c0', 4), rows3([null, 0, 1, 2])],
    [ins('before', 'A r1c1', 4), rows3([0, null, 1, 2])],
    [ins('before', 'A r2c2', 4), rows3([0, 1, null, 2])],
    [ins('after', 'A r0c0', 4), rows3([0, null, 1, 2])],
    [ins('after', 'A r1c1', 4), rows3([0, 1, null, 2])],
    [ins('after', 'A r2c2', 4), rows3([0, 1, 2, null])],
    [del('A r0c0', 2, 'A r1c1'), rows3([1, 2])],
    [del('A r1c1', 2, 'A r0c0'), rows3([0, 2])],
    [del('A r2c2', 2, 'A r0c0'), rows3([0, 1])],
  ];
  for (const [st, want] of cases) {
    const s = store();
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), want, JSON.stringify(st));
  }
});

test('colspan matrix (probe K): OUTSIDE and span-EDGE inserts, an OUTSIDE delete -> SUPPORTED, the colspan and its content kept', async () => {
  const merged = t => `${t} r1c1 ${t} r1c2 ${t} r1c3`;
  for (const [tag, st, row1] of [
    ['K1', ins('before', 'K1 r0c0', 6), ['', 'K1 r1c0', merged('K1'), 'K1 r1c4']],
    ['K2', ins('after', 'K2 r0c4', 6), ['K2 r1c0', merged('K2'), 'K2 r1c4', '']],
    ['K3', del('K3 r0c4', 4, 'K3 r0c0'), ['K3 r1c0', merged('K3')]],
    ['K4', ins('before', 'K4 r0c1', 6), ['K4 r1c0', '', merged('K4'), 'K4 r1c4']],
    ['K5', ins('after', 'K5 r0c3', 6), ['K5 r1c0', merged('K5'), '', 'K5 r1c4']],
  ]) {
    const s = store([kTable(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s)[1], row1, tag);
    assert.ok(spans(s)[1].includes(3), `${tag}: the colspan 3 is kept`);
  }
});

test('colspan matrix: an insert INSIDE the span, a delete of a column the span covers, a DELETE on a merged target -> FAIL_CLOSED_UNSUPPORTED_SHAPE, 0 dispatch', async () => {
  for (const [tag, st, want] of [
    ['K6', del('K6 r0c1', 4, 'K6 r0c0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-covers-deleted-column'],
    ['K7', ins('after', 'K7 r0c1', 6), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-straddles-insert-boundary'],
    ['K8', ins('before', 'K8 r0c2', 6), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-straddles-insert-boundary'],
    ['K9', del('K9 r0c2', 4, 'K9 r0c0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-covers-deleted-column'],
    ['KM3', del('KM3 r1c1 KM3 r1c2 KM3 r1c3', 4, 'KM3 r0c0'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns'],
  ]) {
    const s = store([kTable(tag)]);
    const before = texts(s);
    const { x, driver } = await run(s, task(st));
    assert.equal(status(x), OPERATION_STATUS.FAILED, tag);
    assert.equal(reason(x), want, tag);
    assert.equal(driver.calls.dispatch, 0, tag);
    assert.deepEqual(texts(s), before, `${tag}: nothing written`);
  }
});

test('GH #291 H3: an INSERT on a target that spans columns lands at the span\'s own edge in every row (LIVE, gh291-h-colspan-probe-20261022.json); the span is kept', async () => {
  for (const [tag, side, row0, row1] of [
    ['KM1', 'before', ['KM1 r0c0', '', 'KM1 r0c1', 'KM1 r0c2', 'KM1 r0c3', 'KM1 r0c4'], ['KM1 r1c0', '', 'KM1 r1c1 KM1 r1c2 KM1 r1c3', 'KM1 r1c4']],
    ['KM2', 'after', ['KM2 r0c0', 'KM2 r0c1', 'KM2 r0c2', 'KM2 r0c3', '', 'KM2 r0c4'], ['KM2 r1c0', 'KM2 r1c1 KM2 r1c2 KM2 r1c3', '', 'KM2 r1c4']],
  ]) {
    const s = store([kTable(tag)]);
    await firstAndSecond(s, task(ins(side, `${tag} r1c1 ${tag} r1c2 ${tag} r1c3`, 6)));
    assert.deepEqual(texts(s)[0], row0, tag);
    assert.deepEqual(texts(s)[1], row1, tag);
    assert.deepEqual(texts(s)[2], row0.map(t => t.replace('r0', 'r2')), tag);
    assert.ok(spans(s)[1].includes(3), `${tag}: the colspan 3 is kept`);
  }
});

test('GH #291 H3v: an INSERT on a target that spans BOTH rows and columns lands at the span\'s own edge in every row (LIVE, gh291-h3v-colspan-rowspan-probe-20261022.json); the span is kept', async () => {
  const kv = tag => {
    const t = kTable(tag);
    t.rows[1][1].rowspan = 2;
    t.rows[2] = [cell(`${tag} r2c0`), { span: 3, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell(`${tag} r2c4`)];
    return t;
  };
  for (const [tag, side, row0, row1, row2] of [
    ['KVb', 'before', ['KVb r0c0', '', 'KVb r0c1', 'KVb r0c2', 'KVb r0c3', 'KVb r0c4'], ['KVb r1c0', '', 'KVb r1c1 KVb r1c2 KVb r1c3', 'KVb r1c4'], ['KVb r2c0', '', '', 'KVb r2c4']],
    ['KVa', 'after', ['KVa r0c0', 'KVa r0c1', 'KVa r0c2', 'KVa r0c3', '', 'KVa r0c4'], ['KVa r1c0', 'KVa r1c1 KVa r1c2 KVa r1c3', '', 'KVa r1c4'], ['KVa r2c0', '', '', 'KVa r2c4']],
  ]) {
    const s = store([kv(tag)]);
    await firstAndSecond(s, task(ins(side, `${tag} r1c1 ${tag} r1c2 ${tag} r1c3`, 6)));
    assert.deepEqual(texts(s), [row0, row1, row2], tag);
    assert.ok(spans(s)[1].includes(3) && spans(s)[2].includes(3), `${tag}: the colspan 3 is kept in both merged rows`);
  }
});

test('GH #291 H3v: a DELETE on a target that spans BOTH rows and columns stays FAIL_CLOSED target-spans-columns, 0 dispatch', async () => {
  const t = kTable('KV');
  t.rows[1][1].rowspan = 2;
  t.rows[2] = [cell('KV r2c0'), { span: 3, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('KV r2c4')];
  const s = store([t]);
  const { x, driver } = await run(s, task(del('KV r1c1 KV r1c2 KV r1c3', 4, 'KV r0c0')));
  assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns', JSON.stringify(x.appliedOperations.map(e => e.evidence)));
  assert.equal(driver.calls.dispatch, 0);
});

test('the Split shape (probe G) is a colspan grid: span-edge inserts (also on the colspan-2 cell, GH #291 H3) and a span-free delete SUPPORTED; a DELETE on the colspan-2 cell FAIL_CLOSED', async () => {
  const g3 = store([gTable('G')]);
  await firstAndSecond(g3, task(ins('after', 'G r1c2', 5)));
  assert.deepEqual(texts(g3).map(r => r[r.length - 1]), ['', '', '']);
  const g4 = store([gTable('G')]);
  await firstAndSecond(g4, task(del('G r1c0', 3, 'G r0c2')));
  assert.deepEqual(texts(g4), [['G r0c1', 'G r0c2'], ['G r1c1', '', 'G r1c2'], ['G r2c1', 'G r2c2']]);
  const g5 = store([gTable('G')]);
  await firstAndSecond(g5, task(ins('before', 'G r0c2', 5)));
  assert.deepEqual(texts(g5)[1], ['G r1c0', 'G r1c1', '', '', 'G r1c2']);
  const g6 = store([gTable('G')]);
  await firstAndSecond(g6, task(ins('after', 'G r0c1', 5)));
  assert.deepEqual(texts(g6)[1], ['G r1c0', 'G r1c1', '', '', 'G r1c2'], 'GH #291 H3: the new column at the span\'s right edge');
  for (const st of [del('G r0c1', 2, 'G r0c0')]) {
    const { x, driver } = await run(store([gTable('G')]), task(st));
    assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns');
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('LOGICAL vs PHYSICAL column: deleting logical column 3 of the Split shape removes row 1\'s 4th physical cell; a physical-index delete is RED at L1', async () => {
  const s = store([gTable('G')]);
  await firstAndSecond(s, task(del('G r0c2', 3, 'G r0c0')));
  assert.deepEqual(texts(s)[1], ['G r1c0', 'G r1c1', ''], 'the physical index 2 of row 1 (an empty half) survives');
  const { x } = await run(store([gTable('G')]), task(del('G r0c2', 3, 'G r0c0')), { fault: 'col-physical-index' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL');
  assert.notEqual(status(x), OPERATION_STATUS.APPLIED);
});

test('MULTI_TABLE_TARGETING: the same cell text in two tables -- the anchor picks one, the other is unchanged; no anchor -> FAIL_CLOSED', async () => {
  const two = () => { const b = grid('B', 3); b.rows[1][1] = cell('ZQW SHARED'); const a = grid('A', 3); a.rows[1][1] = cell('ZQW SHARED'); return [a, b]; };
  const s = store(two());
  const other = texts(s, 1);
  await firstAndSecond(s, task(ins('after', 'ZQW SHARED', 4, 'A r0c0')));
  assert.deepEqual(texts(s, 1), other, 'the other table is unchanged');
  assert.deepEqual(texts(s, 0)[1], ['A r1c0', 'ZQW SHARED', '', 'A r1c2']);
  const s2 = store(two());
  const other2 = texts(s2, 0);
  await firstAndSecond(s2, task(del('ZQW SHARED', 2, 'B r0c0')));
  assert.deepEqual(texts(s2, 0), other2);
  assert.deepEqual(texts(s2, 1), [['B r0c0', 'B r0c2'], ['B r1c0', 'B r1c2'], ['B r2c0', 'B r2c2']]);
  const { x, driver } = await run(store(two()), task(ins('after', 'ZQW SHARED', 4)));
  assert.equal(reason(x), 'table-column-target-ambiguous');
  assert.equal(driver.calls.dispatch, 0);
});

test('mixed one session: insert_before -> insert_after -> delete, each resolved from the FRESH state; retry 0/0', async () => {
  const s = store([grid('M', 3, 4)]);
  const t = task(ins('before', 'M r0c1', 5, 'M r2c0'), ins('after', 'M r0c2', 6, 'M r2c0'), del('M r0c3', 5, 'M r2c0'));
  const { first } = await firstAndSecond(s, t);
  // FRESH_STRUCTURAL_RETARGETING: every later target column moved because of an earlier insert
  assert.deepEqual(first.x.primitiveVerification.map(v => v.evidence && v.evidence.evidence && v.evidence.evidence.targetColumnBefore), [1, 3, 5]);
  assert.deepEqual(texts(s)[0], ['M r0c0', '', 'M r0c1', 'M r0c2', '']);
});

test('mixed: a partial first run (the 2nd operation did not happen) -> the retry SATISFIES the 1st and APPLIES the rest', async () => {
  const s = store([grid('M', 3, 4)]);
  const t = task(ins('before', 'M r0c1', 5, 'M r2c0'), ins('after', 'M r0c2', 6, 'M r2c0'), del('M r0c3', 5, 'M r2c0'));
  const first = await run(s, t, { fault: 'ignore', faultOn: [2] });
  assert.equal(status(first.x, 0), OPERATION_STATUS.APPLIED);
  assert.notEqual(status(first.x, 1), OPERATION_STATUS.APPLIED);
  const second = await run(s, t);
  assert.deepEqual(second.x.appliedOperations.map(e => e.status), [OPERATION_STATUS.SATISFIED, OPERATION_STATUS.APPLIED, OPERATION_STATUS.APPLIED]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  assert.deepEqual(texts(s)[0], ['M r0c0', '', 'M r0c1', 'M r0c2', '']);
});

test('delete retry identity: target absent + terminal width in the anchor table -> SATISFIED; recovery -> FAIL_CLOSED', async () => {
  const s = store();
  await run(s, task(del('A r1c1', 2, 'A r0c0')));
  const r = await run(s, task(del('A r1c1', 2, 'A r0c0')));
  assert.equal(status(r.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([r.m.writes, r.m.commands], [0, 0]);
  const rec = await run(s, { ...task(del('A r1c1', 2, 'A r0c0')), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(reason(rec.x), 'preservation-baseline-unavailable');
  assert.equal(rec.driver.calls.dispatch, 0);
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch', async () => {
  const dup = grid('A', 3); dup.rows[2][2] = cell('A r1c1');
  const uneven = grid('A', 3); uneven.rows[1].pop();
  const twoAnchor = () => { const b = grid('B', 2); b.rows[0][0] = cell('A r0c0'); return [grid('A', 3), b]; };
  const failed = [
    [task(ins('after', 'A r1c1', 4)), store([dup]), 'table-column-target-ambiguous'],
    [task(ins('after', 'A r1c1', 4, 'A r0c0')), store(twoAnchor()), 'table-anchor-ambiguous'],
    [task(ins('after', 'nowhere', 4)), store(), 'table-column-target-not-found'],
    [task(ins('after', 'A r1c1', 9)), store(), 'table-column-terminal-unprovable'],
    [task(del('A r1c1', 1, 'A r0c0')), store(), 'delete-target-present-but-column-count-not-terminal-plus-one'],
    [task(del('A r1c1', 2, 'A r0c1')), store(), 'table-anchor-in-target-column'],
    // GH #291 T-04: the ADD side on a vMerge-bearing table is no longer a failure either (its own positive test below)
    // GH #291 H5d: deleting one of an implicit gap's missing columns APPLIES now (its own positive test below)
    [task(ins('before', 'A r0c0', 4, 'A r2c2'), ins('after', 'A r2c2', 5)), store(), 'column-operations-without-table-anchor-cannot-be-grouped'],
  ];
  for (const [t, s, want] of failed) {
    const { x, driver } = await run(s, t);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
  for (const st of [
    { intent: 'delete_table_column', selector: { text: 'A r1c1' }, terminalColumnCount: 2 },
    { ...ins('after', 'A r1c1', 4), terminalColumnCount: undefined },
    { ...ins('after', 'A r1c1', 1) },
    { intent: 'insert_table_column_after', selector: { tableCell: { text: 'A r1c1', tableAnchor: 'A r1c1' } }, terminalColumnCount: 4 },
    { ...del('A r1c1', 2, 'A r0c0'), terminalTableAnchor: 'A r2c2' },
    { intent: 'insert_table_column_after', selector: { text: 'A r1c1', index: 1 }, terminalColumnCount: 4 },
  ]) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
  }
});

// W06-3 (#ref point 2, LIVE-measured #ref READ_DIVERGENCE): the vertical-merge refusal is
// PER-TABLE, not document-wide -- a merge in ONE table no longer refuses a column insert/delete
// on an UNRELATED, merge-free table in the same document; the merge-containing table itself stays
// refused, same reason as before this unit.
test('W06-3 / GH #291: a vertical merge in one table does not refuse a column op on an unrelated, merge-free table; the merge table itself is now written too, its merge kept', async () => {
  const vm = grid('A', 3); vm.rows[0][2] = { ...vm.rows[0][2], rowspan: 2 }; vm.rows[1][2] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  const two = () => [grid('B', 3), vm];
  const s = store(two());
  await firstAndSecond(s, task(ins('after', 'B r1c1', 4, 'B r0c0')));
  assert.equal(texts(s, 0)[1].length, 4, 'B (merge-free) got its new column');
  assert.equal(s.durableDoc.tables[1].rows[0][2].rowspan, 2, 'A (the merge table) untouched');
  const m = store(two());
  await firstAndSecond(m, task(ins('after', 'A r1c1', 4, 'A r0c0')));
  assert.deepEqual([m.durableDoc.tables[1].rows[0][3].rowspan, m.durableDoc.tables[1].rows[1][3].vContinue], [2, true], 'the merge kept, shifted with its column');
});

test('GH #291 T-04: a column insert before / after the merge column and away from it -> APPLIED, the vMerge stays on its own column, the new column has none; retry SATISFIED 0/0', async () => {
  const vmAt1 = () => { const t = grid('A', 3); t.rows[0][1] = { ...t.rows[0][1], rowspan: 2 }; t.rows[1][1] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }; return t; };
  for (const [side, target, ownerCol] of [['before', 'A r2c0', 2], ['after', 'A r2c2', 1], ['before', 'A r2c1', 2], ['after', 'A r2c1', 1]]) {
    const s = store([vmAt1()]);
    await firstAndSecond(s, task(ins(side, target, 4, 'A r0c2')));
    const rows = s.durableDoc.tables[0].rows;
    assert.equal(rows[0][ownerCol].rowspan, 2, `${side} ${target}: the owner keeps its rowspan`);
    assert.equal(rows[1][ownerCol].vContinue, true, `${side} ${target}: the continuation stays under it`);
    assert.equal(rows.flat().filter(c => c.vContinue).length, 1, `${side} ${target}: the new column has no continuation`);
  }
});

/**
 * W-06/2 (the vMerge delete follow-on to W-06/1's LIVE measurement): `delete_table_column` on a
 * vMerge-bearing TARGET table, the 3 measured shapes -- a 3x3 table, the merge at rows 1-2
 * (restart@r1, continue@r2, column probe 1's `vContinue` continuation placeholder, column
 * probe 1's own rowspan reading). `add_table_column` on the SAME tables is untouched by this unit
 * (covered above: still refused).
 */
const vmergeTable = (tag, mergeCol) => {
  const t = grid(tag, 3);
  t.rows[1][mergeCol] = { ...t.rows[1][mergeCol], rowspan: 2 };
  t.rows[2][mergeCol] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  return t;
};
const rowspanOf = (s, t, r, c) => s.durableDoc.tables[t].rows[r][c].rowspan || 1;
const vContinueOf = (s, t, r, c) => s.durableDoc.tables[t].rows[r][c].vContinue === true;

test('W-06/2: delete_table_column UNAFF -- a column the vMerge does not touch, the merge survives at the same column', async () => {
  const s = store([vmergeTable('UNAFF', 0)]);
  await firstAndSecond(s, task(del('UNAFF r0c2', 2, 'UNAFF r0c0')));
  assert.deepEqual(texts(s)[0], ['UNAFF r0c0', 'UNAFF r0c1']);
  assert.deepEqual(texts(s)[1], ['UNAFF r1c0', 'UNAFF r1c1']);
  assert.deepEqual(texts(s)[2], ['', 'UNAFF r2c1']);
  assert.equal(rowspanOf(s, 0, 1, 0), 2, 'the merge (restart) is kept, at the same column');
  assert.equal(vContinueOf(s, 0, 2, 0), true, 'the continuation is kept, at the same column');
});

test('W-06/2: delete_table_column VMCOL -- the vMerge\'s OWN column: the merged cell vanishes with it, no vMerge remnant in the remaining columns', async () => {
  const s = store([vmergeTable('VMCOL', 0)]);
  const { result, x } = await run(s, task(del('VMCOL r0c0', 2, 'VMCOL r0c2')));
  assert.equal(result.editOk, true, JSON.stringify(x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(status(x), OPERATION_STATUS.APPLIED);
  assert.deepEqual(texts(s)[0], ['VMCOL r0c1', 'VMCOL r0c2']);
  assert.deepEqual(texts(s)[1], ['VMCOL r1c1', 'VMCOL r1c2'], 'the merged cell\'s own text is gone with its column');
  assert.deepEqual(texts(s)[2], ['VMCOL r2c1', 'VMCOL r2c2']);
  assert.equal(rowspanOf(s, 0, 1, 0), 1, 'no rowspan anywhere in the remaining columns');
  assert.ok(!s.durableDoc.tables[0].rows[1].some(c => (c.rowspan || 1) > 1) && !s.durableDoc.tables[0].rows[2].some(c => c.vContinue), 'no vMerge remnant at all');
  // requirement 3 (the receipt mechanism: the executor keeps verifyPrimitive's evidence, not
  // classify()'s, for an APPLIED op): the deleted column's own text, per row -- including the
  // merged cell's -- is reported, not silently dropped.
  const ev = x.appliedOperations[0].evidence.evidence;
  assert.deepEqual(ev.deletedColumnTexts, ['VMCOL r0c0', 'VMCOL r1c0', null], 'row2 has no real cell there (the gap) -- null, not a lost string');
  // requirement 4: a retry on the already-deleted column -- SATISFIED, 0 writes, the existing identity rule
  const retry = await run(s, task(del('VMCOL r0c0', 2, 'VMCOL r0c2')));
  assert.equal(status(retry.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([retry.m.writes, retry.m.commands], [0, 0]);
});

test('W-06/2: delete_table_column NEIGH -- a NEIGHBOR of the vMerge column: the merge survives, shifted left', async () => {
  const s = store([vmergeTable('NEIGH', 1)]);
  await firstAndSecond(s, task(del('NEIGH r0c0', 2, 'NEIGH r0c2')));
  assert.deepEqual(texts(s)[0], ['NEIGH r0c1', 'NEIGH r0c2']);
  assert.deepEqual(texts(s)[1], ['NEIGH r1c1', 'NEIGH r1c2'], 'the merged cell\'s own text is preserved');
  assert.deepEqual(texts(s)[2], ['', 'NEIGH r2c2'], 'the continuation placeholder itself is kept (empty text), shifted to column 0');
  assert.equal(rowspanOf(s, 0, 1, 0), 2, 'the merge (restart) shifted to the new column 0');
  assert.equal(vContinueOf(s, 0, 2, 0), true, 'the continuation shifted to the new column 0');
});

test('the source form is accepted: combined selector + terminalColumnCount (+ terminalTableAnchor for delete)', async () => {
  await firstAndSecond(store(), task({ intent: 'insert_table_column_after', selector: { tableCell: { text: 'A r1c1' }, text: 'A r1c1' }, terminalColumnCount: 4 }));
  await firstAndSecond(store(), task({ intent: 'delete_table_column', selector: { tableCell: { text: 'A r1c1' }, text: 'A r1c1' }, terminalColumnCount: 2, terminalTableAnchor: 'A r0c0' }));
});

test('faults at L1: wrong side / twice / not empty / a new cell of span 2 / reordered / delete ignored / other column deleted / stale index -> never APPLIED', async () => {
  for (const [fault, st] of [
    ['col-wrong-side', ins('before', 'A r1c1', 4)], ['col-wrong-side', ins('after', 'A r1c1', 4)], ['col-twice', ins('after', 'A r1c1', 4)],
    ['col-new-not-empty', ins('after', 'A r1c1', 4)], ['col-new-span2', ins('after', 'A r1c1', 4)], ['col-reorder', ins('before', 'A r0c0', 4)],
    ['ignore', del('A r1c1', 2, 'A r0c0')], ['col-delete-other', del('A r1c1', 2, 'A r0c0')], ['col-stale-index', ins('after', 'A r1c1', 4)],
    ['col-stale-index', del('A r2c2', 2, 'A r0c0')],
  ]) {
    const { x, driver } = await run(store(), task(st), { fault });
    const tag = `${fault} ${st.intent}`;
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, tag);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', tag);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, tag);
    assert.equal(driver.calls.forceSave, 0, tag);
  }
});

test('faults at L3: an unrelated cell / another table / text outside the tables -> NOT APPLIED (P1)', async () => {
  for (const [fault, st, tables] of [
    ['col-touch-cell', ins('after', 'A r1c1', 4), undefined],
    ['col-touch-other-table', ins('after', 'A r1c1', 4, 'A r0c0'), [grid('A', 3), grid('B', 2)]],
    ['col-change-outside', del('A r1c1', 2, 'A r0c0'), undefined],
  ]) {
    const { x } = await run(store(tables), task(st), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the column itself is right`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
  }
});

test('a colspan changed by the write (3 -> 1) -> the row grid turns uneven: RED already at L1, never APPLIED', async () => {
  const { x, driver } = await run(store([kTable('K')]), task(ins('after', 'K r0c4', 6)), { fault: 'col-span-change' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL');
  // GH #291 H5: the short row reads as an implicit trailing gap -- the lost columns show up as a gap nobody asked for
  assert.deepEqual(x.primitiveVerification[0].evidence.evidence.gapsAfter, [0, 2, 0]);
  assert.equal(driver.calls.forceSave, 0);
});

test('DEFENCE IN DEPTH: with the shape guard bypassed, the measured drift (K7) and content loss (K6) are RED at L1', async () => {
  for (const [tag, st] of [['K7', ins('after', 'K7 r0c1', 6)], ['K6', del('K6 r0c1', 4, 'K6 r0c0')]]) {
    const s = store([kTable(tag)]);
    const driver = createFakeDocxDriver(s);
    await driver.open({ fileId: 'd' });
    const inv = async () => normalizeInventory((await driver.readLive()).payload);
    const op = planOperations(task(st)).operations[0];
    const before = await inv();
    assert.match(column.classify(before, op).evidence.reason, /^FAIL_CLOSED_UNSUPPORTED_SHAPE/, `${tag}: the guard refuses it`);
    const target = column.resolve(before, op.selector, { operation: op }).target;
    await driver.dispatch(null, column.lower(op, target));
    const after = await inv();
    assert.equal(column.verifyPrimitive(op, { primary: before }, { primary: after }).ok, false, `${tag}: L1 would still catch it`);
  }
});

test('the command finds the table by its anchor and the cell by its full text, then AddColumn / RemoveColumn -- never Search, never frame.evaluate', () => {
  const body = column.commandBody({ family: 'table', action: 'column', columnAction: 'add', before: true, targetText: 'x', tableAnchor: 'a' });
  assert.match(body, /\.AddColumn\(hits\[0\]\.cell, m\.before === true\)/);
  assert.match(column.commandBody({ family: 'table', action: 'column', columnAction: 'remove', before: null, targetText: 'x', tableAnchor: 'a' }), /\.RemoveColumn\(/);
  assert.doesNotMatch(body, /\.Search\(|evaluate|asc_|addColumnLeft|remColumn/);
});

test('group final state reached but THIS operation\'s claim missing -> FAIL_CLOSED, not a blind re-apply', async () => {
  const planned = planOperations(task(ins('after', 'M r0c0', 4, 'M r0c2'), del('M r0c1', 3, 'M r0c2')));
  assert.equal(planned.ok, true);
  // 3 columns, M r0c1 gone (the delete's claim holds), and no empty column after M r0c0 (the insert's does not)
  const t = { rows: [[cell('M r0c0'), cell('M r0c2'), cell('M r0c3')], [cell('M r1c0'), cell('M r1c2'), cell('M r1c3')]] };
  const driver = createFakeDocxDriver(store([t]));
  await driver.open({ fileId: 'd' });
  const primary = normalizeInventory((await driver.readLive()).payload);
  const c = column.classify(primary, planned.operations[0]);
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'group-final-state-without-this-operation-claim');
});

/** M5 row 14: the gridAfter probe's S table (row 0 ends in a one-column grid gap) and its R control (the gap next to a colspan) */
const gap = () => ({ gap: true, span: 1, paras: [] });
const sTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), cell(`${tag} a2`), gap()], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} anchor`)]] });
/** GH #291 H4 G2: a TWO-column gap -- ONE placeholder of span 2 (LIVE): row 0 [a0, a1, gap(2)], row 1 four real cells */
const wTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), { ...gap(), span: 2 }], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} b3`)]] });
/** GH #291 H4 G3: a LEADING (gridBefore) gap: row 0 [gap, a1, a2, a3], row 1 four real cells */
const lTable = tag => ({ rows: [[gap(), cell(`${tag} a1`), cell(`${tag} a2`), cell(`${tag} a3`)], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} b3`)]] });
const rTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), cell(`${tag} a2`), gap()], [cell(`${tag} b0`), cell(`${tag} b1`, 2), cell(`${tag} b2`)]] });
const gapFlags = (s, t = 0) => s.durableDoc.tables[t].rows.map(r => r.map(c => c.gap === true));

test('M5 row 14, grid gap (gridAfter): the measured inserts APPLY -- the new column at the boundary in every row, the gap stays trailing; retry 0/0', async () => {
  for (const [tag, st, row0, row1] of [
    ['S1', ins('after', 'S1 a2', 5), ['S1 a0', 'S1 a1', 'S1 a2', '', ''], ['S1 b0', 'S1 b1', 'S1 b2', '', 'S1 anchor']],
    ['S2', ins('before', 'S2 anchor', 5), ['S2 a0', 'S2 a1', 'S2 a2', '', ''], ['S2 b0', 'S2 b1', 'S2 b2', '', 'S2 anchor']],
    ['S3', ins('after', 'S3 a0', 5), ['S3 a0', '', 'S3 a1', 'S3 a2', ''], ['S3 b0', '', 'S3 b1', 'S3 b2', 'S3 anchor']],
    ['S4', ins('after', 'S4 b0', 5), ['S4 a0', '', 'S4 a1', 'S4 a2', ''], ['S4 b0', '', 'S4 b1', 'S4 b2', 'S4 anchor']],
    ['S5', ins('before', 'S5 a0', 5), ['', 'S5 a0', 'S5 a1', 'S5 a2', ''], ['', 'S5 b0', 'S5 b1', 'S5 b2', 'S5 anchor']],
  ]) {
    const s = store([sTable(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), [row0, row1], tag);
    assert.deepEqual(gapFlags(s), [[false, false, false, false, true], [false, false, false, false, false]], `${tag}: the gap is kept, trailing`);
  }
});

test('M5 row 14, grid gap: a delete of a column that is a real cell in the gap row APPLIES -- the gap kept; retry 0/0 (probe DEL_S_a1, source ...a145c7)', async () => {
  for (const [tag, st, row0, row1] of [
    ['SX1', del('SX1 a1', 3, 'SX1 a0'), ['SX1 a0', 'SX1 a2', ''], ['SX1 b0', 'SX1 b2', 'SX1 anchor']],
    ['SX2', del('SX2 b0', 3, 'SX2 a1'), ['SX2 a1', 'SX2 a2', ''], ['SX2 b1', 'SX2 b2', 'SX2 anchor']],
  ]) {
    const s = store([sTable(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), [row0, row1], tag);
    assert.deepEqual(gapFlags(s), [[false, false, true], [false, false, false]], `${tag}: the gap is kept, trailing`);
  }
});

test('M5 row 14, grid gap: a delete of a ONE-column gap\'s own column APPLIES -- the gap row loses its gap; retry 0/0 (probe DEL_S_anchor, source ...ff8ea3)', async () => {
  const s = store([sTable('SY')]);
  await firstAndSecond(s, task(del('SY anchor', 3, 'SY a0')));
  assert.deepEqual(texts(s), [['SY a0', 'SY a1', 'SY a2'], ['SY b0', 'SY b1', 'SY b2']]);
  assert.deepEqual(gapFlags(s), [[false, false, false], [false, false, false]], 'the gap is consumed, no cell of row 0 lost');
});

test('GH #291 H4: a two-column gap, a leading gap and a gap next to a colspan -- the measured writes APPLY (LIVE, gh291-h4-grid-gap-probe-20261022.json); retry 0/0', async () => {
  for (const [tag, mk, st, row0, row1, flags0] of [
    // H4b: the boundary at the short row's real-cell edge -- the new cell there in every row, the two-column gap kept
    ['WA', wTable, ins('after', 'WA b1', 5), ['WA a0', 'WA a1', '', ''], ['WA b0', 'WA b1', '', 'WA b2', 'WA b3'], [false, false, false, true]],
    // H4d / H4e: a gap column deleted -- the gap shrinks to one column, row 1 loses its cell
    ['WD', wTable, del('WD b2', 3, 'WD a0'), ['WD a0', 'WD a1', ''], ['WD b0', 'WD b1', 'WD b3'], [false, false, true]],
    ['WE', wTable, del('WE b3', 3, 'WE a0'), ['WE a0', 'WE a1', ''], ['WE b0', 'WE b1', 'WE b2'], [false, false, true]],
    // H4f / H4g: a leading gap -- an insert at its end, a delete of a real column; the leading gap kept
    ['LF', lTable, ins('after', 'LF b0', 5), ['', '', 'LF a1', 'LF a2', 'LF a3'], ['LF b0', '', 'LF b1', 'LF b2', 'LF b3'], [true, false, false, false, false]],
    ['LG', lTable, del('LG b1', 3, 'LG b0'), ['', 'LG a2', 'LG a3'], ['LG b0', 'LG b2', 'LG b3'], [true, false, false]],
    // H4h / H4i: a gap next to a colspan, the write away from the span; the span and the gap kept
    ['RH', rTable, ins('before', 'RH b0', 5), ['', 'RH a0', 'RH a1', 'RH a2', ''], ['', 'RH b0', 'RH b1', 'RH b2'], [false, false, false, false, true]],
    ['RI', rTable, del('RI b0', 3, 'RI a1'), ['RI a1', 'RI a2', ''], ['RI b1', 'RI b2'], [false, false, true]],
    // composed: the boundary at the gap row's real-cell edge AND the colspan's right edge
    ['RJ', rTable, ins('after', 'RJ a2', 5), ['RJ a0', 'RJ a1', 'RJ a2', '', ''], ['RJ b0', 'RJ b1', '', 'RJ b2'], [false, false, false, false, true]],
  ]) {
    const s = store([mk(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), [row0, row1], tag);
    assert.deepEqual(gapFlags(s)[0], flags0, `${tag}: the gap kept where measured`);
  }
  const wd = store([wTable('WX')]);
  await firstAndSecond(wd, task(del('WX b2', 3, 'WX a0')));
  assert.equal(wd.durableDoc.tables[0].rows[0][2].span, 1, 'H4d: the two-column gap became ONE column');
});

test('GH #291 H5: a row short WITHOUT a gap marker is an implicit trailing gap -- the measured writes APPLY, the editor writes the gap marker (LIVE, gh291-v13-h5-probe-20261022.json); retry 0/0', async () => {
  const uTable = tag => ({ rows: [[cell(`${tag} a0`), cell(`${tag} a1`), cell(`${tag} a2`)], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} b3`)]] });
  for (const [tag, st, row0, row1] of [
    ['UA', ins('after', 'UA a0', 5), ['UA a0', '', 'UA a1', 'UA a2'], ['UA b0', '', 'UA b1', 'UA b2', 'UA b3']],
    ['UB', del('UB a1', 3, 'UB a0'), ['UB a0', 'UB a2'], ['UB b0', 'UB b2', 'UB b3']],
    ['UD', ins('after', 'UD b2', 5), ['UD a0', 'UD a1', 'UD a2', ''], ['UD b0', 'UD b1', 'UD b2', '', 'UD b3']],
  ]) {
    const s = store([uTable(tag)]);
    // the editor writes gridAfter on the column write (measured) -- the fake does not, so the gap stays implicit here
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), [row0, row1], tag);
  }
  const s = store([uTable('UX')]);
  const { x, driver } = await run(s, task(ins('after', 'UX b3', 5)));
  assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-grid-gap', 'a boundary past the short row\'s cells stays refused');
  assert.equal(driver.calls.dispatch, 0);
});

test('GH #291 V13: a vertical merge next to a trailing grid gap -- the measured inserts / deletes APPLY, the merge and the gap kept (LIVE, gh291-v13-h5-probe-20261022.json); retry 0/0', async () => {
  const cont = () => ({ vContinue: true, span: 1, paras: [{ text: '', props: { ...PLAIN } }] });
  // V1: rows 0-1 of column 0 merged, row 2 three cells + a one-column gap
  const v1 = tag => ({ rows: [[{ ...cell(`${tag} m`), rowspan: 2 }, cell(`${tag} r0c1`), cell(`${tag} r0c2`), cell(`${tag} r0c3`)], [cont(), cell(`${tag} r1c1`), cell(`${tag} r1c2`), cell(`${tag} r1c3`)], [cell(`${tag} r2c0`), cell(`${tag} r2c1`), cell(`${tag} r2c2`), gap()]] });
  // V2: rows 1-2 of column 1 merged, row 2 (the continuation row) three cells + a one-column gap
  const v2 = tag => ({ rows: [[cell(`${tag} r0c0`), cell(`${tag} r0c1`), cell(`${tag} r0c2`), cell(`${tag} r0c3`)], [cell(`${tag} r1c0`), { ...cell(`${tag} m`), rowspan: 2 }, cell(`${tag} r1c2`), cell(`${tag} r1c3`)], [cell(`${tag} r2c0`), cont(), cell(`${tag} r2c2`), gap()]] });
  for (const [tag, mk, st, want] of [
    ['VA', v1, ins('after', 'VA r0c1', 5), [['VA m', 'VA r0c1', '', 'VA r0c2', 'VA r0c3'], ['', 'VA r1c1', '', 'VA r1c2', 'VA r1c3'], ['VA r2c0', 'VA r2c1', '', 'VA r2c2', '']]],
    ['VB', v1, del('VB r0c1', 3, 'VB r0c3'), [['VB m', 'VB r0c2', 'VB r0c3'], ['', 'VB r1c2', 'VB r1c3'], ['VB r2c0', 'VB r2c2', '']]],
    ['VD', v1, del('VD r0c3', 3, 'VD r0c1'), [['VD m', 'VD r0c1', 'VD r0c2'], ['', 'VD r1c1', 'VD r1c2'], ['VD r2c0', 'VD r2c1', 'VD r2c2']]],
    ['VE', v2, ins('after', 'VE r0c2', 5), [['VE r0c0', 'VE r0c1', 'VE r0c2', '', 'VE r0c3'], ['VE r1c0', 'VE m', 'VE r1c2', '', 'VE r1c3'], ['VE r2c0', '', 'VE r2c2', '', '']]],
    ['VF', v2, del('VF r0c0', 3, 'VF r0c3'), [['VF r0c1', 'VF r0c2', 'VF r0c3'], ['VF m', 'VF r1c2', 'VF r1c3'], ['', 'VF r2c2', '']]],
  ]) {
    const s = store([mk(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), want, tag);
    assert.ok(s.durableDoc.tables[0].rows.flat().some(c => c.rowspan === 2), `${tag}: the merge kept`);
  }
});

test('GH #291 V13l / H4q / H5d: a leading gap next to a vertical merge or a colspan, a row gapped on both sides, an implicit gap column delete -- the measured writes APPLY (LIVE, gh291-gap-remaining-probe-20261023.json); retry 0/0', async () => {
  const cont = () => ({ vContinue: true, span: 1, paras: [{ text: '', props: { ...PLAIN } }] });
  // L1: rows 0-1 of column 1 merged, row 2 a leading gap (its column-0 cell missing)
  const l1 = tag => ({ rows: [[cell(`${tag} r0c0`), { ...cell(`${tag} m`), rowspan: 2 }, cell(`${tag} r0c2`), cell(`${tag} r0c3`)], [cell(`${tag} r1c0`), cont(), cell(`${tag} r1c2`), cell(`${tag} r1c3`)], [gap(), cell(`${tag} r2c1`), cell(`${tag} r2c2`), cell(`${tag} r2c3`)]] });
  // L2: rows 1-2 of column 1 merged, row 2 (the continuation row) a leading gap
  const l2 = tag => ({ rows: [[cell(`${tag} r0c0`), cell(`${tag} r0c1`), cell(`${tag} r0c2`), cell(`${tag} r0c3`)], [cell(`${tag} r1c0`), { ...cell(`${tag} m`), rowspan: 2 }, cell(`${tag} r1c2`), cell(`${tag} r1c3`)], [gap(), cont(), cell(`${tag} r2c2`), cell(`${tag} r2c3`)]] });
  // Q2: row 0 a leading gap, row 1 columns 1-2 spanned
  const q2 = tag => ({ rows: [[gap(), cell(`${tag} r0c1`), cell(`${tag} r0c2`), cell(`${tag} r0c3`)], [cell(`${tag} r1c0`), cell(`${tag} r1c1`, 2), cell(`${tag} r1c3`)]] });
  // Q3: row 0 gapped on both sides
  const q3 = tag => ({ rows: [[gap(), cell(`${tag} r0c1`), cell(`${tag} r0c2`), gap()], [cell(`${tag} r1c0`), cell(`${tag} r1c1`), cell(`${tag} r1c2`), cell(`${tag} r1c3`)]] });
  // U1: row 0 one column short WITHOUT a marker
  const u1 = tag => ({ rows: [[cell(`${tag} r0c0`), cell(`${tag} r0c1`), cell(`${tag} r0c2`)], [cell(`${tag} r1c0`), cell(`${tag} r1c1`), cell(`${tag} r1c2`), cell(`${tag} r1c3`)]] });
  for (const [tag, mk, st, want] of [
    // L1a / L1b / L1c: right of the gap, a real column deleted, a boundary at the gap's end (before the merged cell)
    ['LA', l1, ins('after', 'LA r0c2', 5), [['LA r0c0', 'LA m', 'LA r0c2', '', 'LA r0c3'], ['LA r1c0', '', 'LA r1c2', '', 'LA r1c3'], ['', 'LA r2c1', 'LA r2c2', '', 'LA r2c3']]],
    ['LB', l1, del('LB r0c2', 3, 'LB r0c3'), [['LB r0c0', 'LB m', 'LB r0c3'], ['LB r1c0', '', 'LB r1c3'], ['', 'LB r2c1', 'LB r2c3']]],
    ['LC', l1, ins('before', 'LC m', 5), [['LC r0c0', '', 'LC m', 'LC r0c2', 'LC r0c3'], ['LC r1c0', '', '', 'LC r1c2', 'LC r1c3'], ['', '', 'LC r2c1', 'LC r2c2', 'LC r2c3']]],
    // L2a / L2b / L2d: the continuation row holds the leading gap
    ['LE', l2, ins('after', 'LE r0c2', 5), [['LE r0c0', 'LE r0c1', 'LE r0c2', '', 'LE r0c3'], ['LE r1c0', 'LE m', 'LE r1c2', '', 'LE r1c3'], ['', '', 'LE r2c2', '', 'LE r2c3']]],
    ['LF', l2, del('LF r0c3', 3, 'LF r0c0'), [['LF r0c0', 'LF r0c1', 'LF r0c2'], ['LF r1c0', 'LF m', 'LF r1c2'], ['', '', 'LF r2c2']]],
    ['LG', l2, ins('before', 'LG r0c1', 5), [['LG r0c0', '', 'LG r0c1', 'LG r0c2', 'LG r0c3'], ['LG r1c0', '', 'LG m', 'LG r1c2', 'LG r1c3'], ['', '', '', 'LG r2c2', 'LG r2c3']]],
    // Q2a / Q2b: a leading gap next to a colspan, the write away from both
    ['QA', q2, ins('after', 'QA r0c3', 5), [['', 'QA r0c1', 'QA r0c2', 'QA r0c3', ''], ['QA r1c0', 'QA r1c1', 'QA r1c3', '']]],
    ['QB', q2, del('QB r0c3', 3, 'QB r1c0'), [['', 'QB r0c1', 'QB r0c2'], ['QB r1c0', 'QB r1c1']]],
    // Q3a / Q3b / Q3d: a row gapped on both sides -- inside its real cells, and its TRAILING gap column (the gap consumed)
    ['QC', q3, ins('after', 'QC r0c1', 5), [['', 'QC r0c1', '', 'QC r0c2', ''], ['QC r1c0', 'QC r1c1', '', 'QC r1c2', 'QC r1c3']]],
    ['QD', q3, del('QD r0c1', 3, 'QD r1c0'), [['', 'QD r0c2', ''], ['QD r1c0', 'QD r1c2', 'QD r1c3']]],
    ['QE', q3, del('QE r1c3', 3, 'QE r1c0'), [['', 'QE r0c1', 'QE r0c2'], ['QE r1c0', 'QE r1c1', 'QE r1c2']]],
    // H5d: the column an implicit short row lacks -- only the full rows lose a cell
    ['UE', u1, del('UE r1c3', 3, 'UE r1c0'), [['UE r0c0', 'UE r0c1', 'UE r0c2'], ['UE r1c0', 'UE r1c1', 'UE r1c2']]],
  ]) {
    const s = store([mk(tag)]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(texts(s), want, tag);
  }
  // the leading gap stays where measured; the merge stays a merge
  const lc = store([l1('LH')]);
  await firstAndSecond(lc, task(ins('before', 'LH m', 5)));
  assert.equal(gapFlags(lc)[2][0], true, 'L1c: row 2 keeps its leading gap, the new cell after it');
  assert.ok(lc.durableDoc.tables[0].rows.flat().some(c => c.rowspan === 2), 'L1c: the merge kept');
  // measured misaligning, so refused (0 dispatch): the leading gap column next to a merge, a boundary before the gap
  for (const [tag, mk, st, want] of [
    ['LX', l1, del('LX r0c0', 3, 'LX r0c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:leading-grid-gap-column-delete'],
    ['LY', l2, del('LY r0c0', 3, 'LY r0c3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:leading-grid-gap-column-delete'],
    ['LZ', l1, ins('before', 'LZ r0c0', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-leading-grid-gap'],
  ]) {
    const s = store([mk(tag)]);
    const { x, driver } = await run(s, task(st));
    assert.equal(reason(x), want, tag);
    assert.equal(driver.calls.dispatch, 0, tag);
  }
});

test('M5 row 14 + GH #291 H4 / H4q: a boundary past a short row\'s real cells, a boundary inside or before a leading gap, a delete of a leading gap column (all measured: misalign) -> FAIL_CLOSED_UNSUPPORTED_SHAPE, 0 dispatch', async () => {
  const both = tag => ({ rows: [[gap(), cell(`${tag} a1`), cell(`${tag} a2`), gap()], [cell(`${tag} b0`), cell(`${tag} b1`), cell(`${tag} b2`), cell(`${tag} b3`)]] });
  const lSpan = tag => ({ rows: [[gap(), cell(`${tag} a1`), cell(`${tag} a2`), cell(`${tag} a3`)], [cell(`${tag} b0`), cell(`${tag} b1`, 2), cell(`${tag} b3`)]] });
  for (const [tag, mk, st, want] of [
    ['SG1', sTable, ins('after', 'SG1 anchor', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-grid-gap'],
    ['WC', wTable, ins('after', 'WC b2', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-grid-gap'],
    ['LB', lTable, ins('before', 'LB b0', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-leading-grid-gap'],
    ['LD', lTable, del('LD b0', 3, 'LD b1'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:leading-grid-gap-column-delete'],
    // GH #291 H4q (LIVE, gh291-gap-remaining-probe-20261023.json): measured misaligning -- refused next to a colspan and in a row
    // gapped on both sides too (the shapes themselves APPLY outside the leading gap, their own positive test below)
    ['LSB', lSpan, ins('before', 'LSB b0', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-leading-grid-gap'],
    ['LSD', lSpan, del('LSD b0', 3, 'LSD b3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:leading-grid-gap-column-delete'],
    ['BSD', both, del('BSD b0', 3, 'BSD b3'), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:leading-grid-gap-column-delete'],
    ['BSB', both, ins('before', 'BSB b0', 5), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-leading-grid-gap'],
  ]) {
    const s = store([mk(tag)]);
    const before = JSON.stringify(s.durableDoc.tables);
    const { x, driver } = await run(s, task(st));
    assert.equal(status(x), OPERATION_STATUS.FAILED, tag);
    assert.equal(reason(x), want, tag);
    assert.equal(driver.calls.dispatch, 0, tag);
    assert.equal(JSON.stringify(s.durableDoc.tables), before, `${tag}: nothing written`);
  }
});

test('M5 row 14, columnRead: a placeholder between real cells, a non-empty placeholder, a row-count or cell-count mismatch -> grid-gap-unreconciled', () => {
  const td = (text, extra = {}) => ({ text, span: 1, rowspan: 1, ...extra });
  const ph = (extra = {}) => td('', { placeholder: true, ...extra });
  const rd = t => ({ text: t });
  const primary = (hrows, drows) => ({ blocks: [{ kind: 'table', rows: hrows }], document: { tables: [{ rows: drows }] } });
  const ok = column.columnRead(primary([[td('a'), td('b'), ph()], [td('c'), td('d'), td('e')]], [[rd('a'), rd('b')], [rd('c'), rd('d'), rd('e')]]));
  assert.equal(ok.ok, true, 'POSITIVE CONTROL: a trailing empty placeholder reconciles');
  assert.deepEqual(ok.gaps, [[1, 0]]);
  assert.deepEqual(ok.html[0].rows.map(r => r.length), [2, 3]);
  const nbsp = column.columnRead(primary([[td('a'), ph({ text: ' ' })], [td('b'), td('c')]], [[rd('a')], [rd('b'), rd('c')]]));
  assert.equal(nbsp.ok, true, 'a whitespace-only placeholder (the projection\'s &nbsp; paragraph) is still a gap');
  const named = column.columnRead(primary([[td('a'), ph({ text: 'x' })]], [[rd('a')]]));
  assert.equal(named.detail.check, 'gap-shape', 'the refusal names the failed check');
  assert.deepEqual([named.detail.table, named.detail.row, named.detail.driverCells], [0, 0, 1]);
  assert.equal(column.columnRead(primary([[td('a'), td('b'), ph()]], [[rd('a')]])).detail.check, 'cell-count');
  assert.equal(column.columnRead(primary([[td('a'), ph()], [td('b')]], [[rd('a')]])).detail.check, 'row-count');
  // GH #291 H4 (LIVE): a two-column gap is ONE placeholder of span 2; a gridBefore gap is a LEADING placeholder
  const wide = column.columnRead(primary([[td('a'), ph({ span: 2 })], [td('b'), td('c'), td('d')]], [[rd('a')], [rd('b'), rd('c'), rd('d')]]));
  assert.deepEqual([wide.ok, wide.gaps, wide.leads], [true, [[2, 0]], [[0, 0]]], 'a spanning placeholder is a gap of its span');
  const lead = column.columnRead(primary([[ph(), td('a'), td('b')], [td('c'), td('d'), td('e')]], [[rd('a'), rd('b')], [rd('c'), rd('d'), rd('e')]]));
  assert.deepEqual([lead.ok, lead.gaps, lead.leads, lead.html[0].rows[0].map(c => c.text)], [true, [[0, 0]], [[1, 0]], ['a', 'b']], 'a leading placeholder is a leading gap');
  for (const [name, p] of [
    ['placeholder in the middle', primary([[td('a'), ph(), td('b')]], [[rd('a'), rd('b')]])],
    ['non-empty placeholder', primary([[td('a'), ph({ text: 'x' })]], [[rd('a')]])],
    ['row-count mismatch', primary([[td('a'), ph()], [td('b')]], [[rd('a')]])],
    ['cell-count mismatch', primary([[td('a'), td('b'), ph()]], [[rd('a')]])],
  ]) assert.equal(column.columnRead(p).reason, 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-unreconciled', name);
});

// GH #291 V11 (LIVE, docs/evidence/gh291-v11-vmerge-colspan-probe-20261022.json): a merged cell spanning rows 0-1 AND columns 1-2 of a
// 3x4 table. Outside / edge inserts and an outside delete keep the block; a boundary inside it, and deleting one of its columns, stay refused.
const blockTable = tag => {
  const t = grid(tag, 3, 4);
  t.rows[0] = [cell(`${tag} r0c0`), { ...cell(`${tag} block`, 2), rowspan: 2 }, cell(`${tag} r0c3`)];
  t.rows[1] = [cell(`${tag} r1c0`), { span: 2, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell(`${tag} r1c3`)];
  return t;
};
test('GH #291 V11: a column insert outside / at the edge of a rows+columns merged block, and an outside delete -> APPLIED, the block kept (rowspan 2, span 2, continuation under it); retry SATISFIED 0/0', async () => {
  for (const [label, step, ownerAt] of [
    ['insert before col 0', ins('before', 'B r2c0', 5, 'B r0c3'), 2], ['insert after col 3', ins('after', 'B r2c3', 5, 'B r0c3'), 1],
    ['insert before col 1 (left edge)', ins('before', 'B r2c1', 5, 'B r0c3'), 2], ['insert after col 2 (right edge)', ins('after', 'B r2c2', 5, 'B r0c3'), 1],
    ['delete col 0', del('B r2c0', 3, 'B r0c3'), 0]]) {
    const s = store([blockTable('B')]);
    const { x } = await run(s, task(step));
    assert.equal(status(x), OPERATION_STATUS.APPLIED, `${label}: ${JSON.stringify(x.appliedOperations[0].evidence).slice(0, 300)}`);
    const rows = s.durableDoc.tables[0].rows;
    assert.deepEqual([rows[0][ownerAt].span, rows[0][ownerAt].rowspan, rows[1][ownerAt].vContinue, rows[1][ownerAt].span], [2, 2, true, 2], `${label}: the block kept`);
    const saves = s.saves;
    const { x: x2 } = await run(s, task(step));
    assert.equal(status(x2), OPERATION_STATUS.SATISFIED, `${label}: retry`);
    assert.equal(s.saves, saves, `${label}: no write on retry`);
  }
});
test('GH #291 V11: a column insert INSIDE the block, and deleting one of its columns -> FAILED before any write (measured: the grid misaligns / the whole block is lost)', async () => {
  for (const [step, re] of [[ins('after', 'B r2c1', 5, 'B r0c3'), /colspan-straddles-insert-boundary/], [del('B r2c1', 3, 'B r0c3'), /colspan-covers-deleted-column/], [del('B r2c2', 3, 'B r0c3'), /colspan-covers-deleted-column/]]) {
    const s = store([blockTable('B')]);
    const saves = s.saves;
    const { x } = await run(s, task(step));
    assert.equal(status(x), OPERATION_STATUS.FAILED);
    assert.match(String(reason(x)), re);
    assert.equal(s.saves, saves, 'no write');
  }
});
