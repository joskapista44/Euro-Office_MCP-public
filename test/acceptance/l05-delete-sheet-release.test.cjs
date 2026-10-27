'use strict';

/**
 * L-05 (owner decision 2026-10-15, inventory row L-05): delete_sheet on the PRODUCT surface. The sheets family already routes it; the LIVE measurement on the
 * prod DS (test/live/l05-sheets-measure.live.cjs, r3, 10/10 classified) found it WRITABLE for an empty sheet, a sheet with data and a sheet with a chart, the
 * retry SATISFIED with 0 commands, an absent sheet SATISFIED, and REFUSED, named, a sheet a formula elsewhere names. Offline, the product WRITE tool with the fake
 * XLSX editor; the LIVE proof is test/live/release-l05-sheets-product.live.cjs (RED on the pre-release build, GREEN on the rehearsal build).
 * move_sheet and copy_sheet are released by later units (L-05/2, L-05/3: see the allowlist comment in write-tools.cjs).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

const J = x => JSON.stringify(x);
const SRC_CELLS = () => ({ A1: { v: 'h1', f: '' }, B1: { v: 'h2', f: '', bold: true }, C2: { v: '3', f: '', nf: '0.00' }, A3: { v: '4', f: '=B2+1' }, D1: { v: 'mrg', f: '', merge: 'D1:E1' } });
const book = () => [
  { name: 'Keep', cells: { A1: { v: 'k' }, A2: { v: 2 } } },
  { name: 'Empty', cells: {} },
  { name: 'Data', cells: { A1: { v: 'h' }, A2: { v: 7 }, B2: { v: 8 } } },
  { name: 'Named', cells: { A1: { v: 'n' } } },
  { name: 'User', cells: { A1: { v: 7, f: '=Named!A1' } } },
];
const del = sheet => ({ intent: 'delete_sheet', id: 's', sheet });
const names = h => h.store.durable.map(x => x.name);
const statuses = r => (r.operations || []).map(o => o.status);
const validate = step => W.validateWriteRequest(req([step]));

test('L-05: delete_sheet is a released intent, schema + constraints wired, a sole step', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('delete_sheet'));
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.ok(W.SOLE_STEP.includes('delete_sheet'));
  assert.equal(validate(del('Data')).ok, true);
});

test('L-05: the schema refuses a malformed step before any dispatch', () => {
  const bad = [
    [{ intent: 'delete_sheet', id: 's' }, 'sheet missing'],
    [{ ...del('Data'), name: 'x' }, 'an unknown field (name belongs to rename_sheet)'],
    [{ ...del('Data'), before: 'Keep' }, 'an unknown field (before belongs to move_sheet)'],
    [del(''), 'an empty sheet name'],
    [del('a/b'), 'a sheet name with a forbidden character'],
    [del(3), 'a sheet name that is not a string'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-05: delete an empty sheet and a sheet with data -- APPLIED, the others untouched, a retry is SATISFIED with 0 dispatch', async () => {
  for (const target of ['Empty', 'Data']) {
    const h = harness({ store: createStore(book()) });
    const r = await h.tool.handler(req([del(target)]));
    assert.equal(r.ok, true, `${target}: ${J(r).slice(0, 500)}`);
    assert.deepEqual(statuses(r), ['APPLIED'], target);
    assert.equal(names(h).includes(target), false, target);
    assert.deepEqual(h.store.durable.find(x => x.name === 'Keep').cells, { A1: { v: 'k' }, A2: { v: 2 } }, `${target}: Keep untouched`);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([del(target)]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `${target}: retry`);
  }
});

test('L-05: an absent sheet is SATISFIED with 0 dispatch (nothing to delete)', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([del('Nope')]));
  assert.deepEqual([r.ok, statuses(r), h.calls().dispatch], [true, ['SATISFIED'], 0], J(r).slice(0, 400));
});

test('L-05 NAMED limit: a sheet a formula elsewhere names is refused -- 0 dispatch, the document untouched, the family reason passes through', async () => {
  const h = harness({ store: createStore(book()) });
  const before = J(h.store.durable);
  const r = await h.tool.handler(req([del('Named')]));
  assert.equal(r.ok, false);
  assert.ok(J(r).includes('xlsx-delete-sheet-referenced-by-formulas'), J(r).slice(0, 600));
  assert.equal(h.calls().dispatch, 0);
  assert.equal(J(h.store.durable), before);
});


test('L-05 live harness verdict: a synthetic GREEN run passes, the pre-release shape (delete_sheet not allowlisted) and each broken gate fail', { skip: needsPrivate('test/live/release-l05-sheets-product.live.cjs') }, () => {
  const h = require('../live/release-l05-sheets-product.live.cjs');
  const tag = 'abc123';
  const sheetOf = k => `L05${k}_${tag}`;
  const ap = { ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }] };
  const sat = { ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] };
  const start = [sheetOf('MV1'), sheetOf('MV0'), sheetOf('MV2')];
  const plan = h.movePlan(start);
  const green = () => ({
    MOVE: { start, plan, runs: { MOV_AFTER: { first: ap, second: sat }, MOV_BEFORE: { first: ap, second: sat } } },
    SEEDS: { a: ap, b: ap },
    PAIRS: { DEL_EMPTY: { first: ap, second: sat }, DEL_DATA: { first: ap, second: sat }, CPY_BASIC: { first: ap, second: sat } },
    ABSENT: sat,
    NEG_NAMED: { DEL_REF: { ok: false, code: null, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.NEG_REASON } }] }, MOV_NOREF: { ok: false, code: null, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.MOVE_REASON } }] }, CPY_EXISTS: { ok: false, code: null, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.CPY_REASON } }] } },
    RUN_ID: 'run-abc123',
    NEG_SHAPE: { empty_name: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, unknown_field: { ok: false, code: 'UNKNOWN_FIELD', persistence: null, ops: [] }, mov_self: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, mov_both: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, cpy_same_name: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] } },
    READ: { ok: true, sheets: [sheetOf('DR'), sheetOf('U'), sheetOf('KEEP'), `L05CP_abc123`, ...plan.end], copy: SRC_CELLS(), source: SRC_CELLS(), crossRef: { A1: { v: '3', f: `=${sheetOf('CPS')}!B2` } } },
    CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO',
  });
  assert.deepEqual(h.verdict(green(), sheetOf), { PASS: true, failed: [], checks: h.verdict(green(), sheetOf).checks });
  // RED: the pre-release build refuses the first step before dispatch
  const red = green();
  red.PAIRS.DEL_EMPTY = { first: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] } };
  const rv = h.verdict(red, sheetOf);
  assert.equal(rv.PASS, false);
  assert.ok(rv.failed.includes('DEL_EMPTY_applied') && rv.failed.includes('DEL_EMPTY_retry'), J(rv.failed));
  // the sheet still in the saved file / the referenced sheet gone / the named refusal missing
  const still = green(); still.READ.sheets.push(sheetOf('DEL_DATA'));
  assert.ok(h.verdict(still, sheetOf).failed.includes('DEL_DATA_gone_in_saved_file'));
  const gone = green(); gone.READ.sheets = gone.READ.sheets.filter(n => n !== sheetOf('DR'));
  assert.ok(h.verdict(gone, sheetOf).failed.includes('neg_DEL_REF_both_sheets_kept'));
  const unnamed = green(); unnamed.NEG_NAMED.DEL_REF.ops[0].failure = { reason: 'something-else' };
  assert.ok(h.verdict(unnamed, sheetOf).failed.includes('neg_DEL_REF'));
});

test('L-05/2 live harness: the move plan always CHANGES the order (end != start, end != after step 1), and the verdict catches a missing release / a wrong end order / a moved refusal', { skip: needsPrivate('test/live/release-l05-sheets-product.live.cjs') }, () => {
  const h = require('../live/release-l05-sheets-product.live.cjs');
  const sheetOf = k => `L05${k}_abc123`;
  const ap = { ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }] };
  const sat = { ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] };
  // for EVERY starting order of the trio
  for (const start of [['a', 'b', 'c'], ['a', 'c', 'b'], ['b', 'a', 'c'], ['b', 'c', 'a'], ['c', 'a', 'b'], ['c', 'b', 'a']]) {
    const mp = h.movePlan(start);
    assert.notDeepEqual(mp.end, start, `end == start for ${start}`);
    assert.notDeepEqual(mp.end, mp.steps[0].order, `end == after step 1 for ${start}`);
    assert.notDeepEqual(mp.steps[0].order, start, `step 1 is a no-op for ${start}`);
    assert.deepEqual(mp.steps[0].step, { intent: 'move_sheet', sheet: start[0], after: start[1] });
    assert.deepEqual(mp.steps[1].step, { intent: 'move_sheet', sheet: start[2], before: start[1] });
  }
  const start = [sheetOf('MV1'), sheetOf('MV0'), sheetOf('MV2')];
  const plan = h.movePlan(start);
  const base = () => ({
    MOVE: { start, plan, runs: { MOV_AFTER: { first: ap, second: sat }, MOV_BEFORE: { first: ap, second: sat } } },
    SEEDS: { a: ap }, PAIRS: { DEL_EMPTY: { first: ap, second: sat }, DEL_DATA: { first: ap, second: sat }, CPY_BASIC: { first: ap, second: sat } }, ABSENT: sat,
    NEG_NAMED: { DEL_REF: { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.NEG_REASON } }] }, MOV_NOREF: { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.MOVE_REASON } }] }, CPY_EXISTS: { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.CPY_REASON } }] } },
    RUN_ID: 'run-abc123',
    NEG_SHAPE: { empty_name: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, unknown_field: { ok: false, code: 'UNKNOWN_FIELD', persistence: null, ops: [] }, mov_self: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, mov_both: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, cpy_same_name: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] } },
    READ: { ok: true, sheets: [sheetOf('DR'), sheetOf('U'), sheetOf('KEEP'), `L05CP_abc123`, ...plan.end], copy: SRC_CELLS(), source: SRC_CELLS(), crossRef: { A1: { v: '3', f: `=${sheetOf('CPS')}!B2` } } }, CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO',
  });
  assert.equal(h.verdict(base(), sheetOf).PASS, true, J(h.verdict(base(), sheetOf).failed));
  const red = base(); red.MOVE.runs.MOV_AFTER = { first: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] } };
  assert.ok(h.verdict(red, sheetOf).failed.includes('MOV_AFTER_applied'), 'the pre-release build (move_sheet not allowlisted) is RED');
  const wrong = base(); wrong.READ.sheets = [sheetOf('DR'), sheetOf('U'), sheetOf('KEEP'), ...plan.steps[0].order];
  assert.ok(h.verdict(wrong, sheetOf).failed.includes('move_end_order_in_saved_file'), 'only the first step took effect');
  const unmoved = base(); unmoved.READ.sheets = [sheetOf('DR'), sheetOf('U'), sheetOf('KEEP'), ...start];
  assert.ok(h.verdict(unmoved, sheetOf).failed.includes('move_end_order_in_saved_file'), 'nothing moved');
  const noref = base(); noref.NEG_NAMED.MOV_NOREF.ops[0].failure = { reason: 'other' };
  assert.ok(h.verdict(noref, sheetOf).failed.includes('neg_MOV_NOREF'));
  const dispatched = base(); dispatched.NEG_SHAPE.mov_self = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [{ status: 'FAILED' }] };
  assert.ok(h.verdict(dispatched, sheetOf).failed.includes('neg_shape_mov_self'));
  const nostart = base(); nostart.MOVE = { start: [] };
  assert.ok(h.verdict(nostart, sheetOf).failed.includes('move_start_observed'));
});

test('L-05/3 live harness: the copy gates -- the pre-release build (copy_sheet refused), a copy missing from the saved file, a copy that differs from its source, an unnamed or missing refusal', { skip: needsPrivate('test/live/release-l05-sheets-product.live.cjs') }, () => {
  const h = require('../live/release-l05-sheets-product.live.cjs');
  const sheetOf = k => `L05${k}_abc123`;
  const ap = { ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }] };
  const sat = { ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] };
  const start = [sheetOf('MV1'), sheetOf('MV0'), sheetOf('MV2')];
  const plan = h.movePlan(start);
  const base = () => ({
    RUN_ID: 'run-abc123', MOVE: { start, plan, runs: { MOV_AFTER: { first: ap, second: sat }, MOV_BEFORE: { first: ap, second: sat } } },
    SEEDS: { a: ap }, PAIRS: { DEL_EMPTY: { first: ap, second: sat }, DEL_DATA: { first: ap, second: sat }, CPY_BASIC: { first: ap, second: sat } }, ABSENT: sat,
    NEG_NAMED: { DEL_REF: { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.NEG_REASON } }] }, MOV_NOREF: { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.MOVE_REASON } }] }, CPY_EXISTS: { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: h.CPY_REASON } }] } },
    NEG_SHAPE: { empty_name: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, unknown_field: { ok: false, code: 'UNKNOWN_FIELD', persistence: null, ops: [] }, mov_self: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, mov_both: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] }, cpy_same_name: { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] } },
    READ: { ok: true, sheets: [sheetOf('DR'), sheetOf('U'), sheetOf('KEEP'), 'L05CP_abc123', ...plan.end], copy: SRC_CELLS(), source: SRC_CELLS(), crossRef: { A1: { v: '3', f: `=${sheetOf('CPS')}!B2` } } }, CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO',
  });
  assert.equal(h.verdict(base(), sheetOf).PASS, true, J(h.verdict(base(), sheetOf).failed));
  const red = base(); red.PAIRS.CPY_BASIC = { first: { ok: false, code: 'REFUSED_TEMPORARY_BOUNDARY', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_TEMPORARY_BOUNDARY', persistence: null, ops: [] } };
  assert.ok(h.verdict(red, sheetOf).failed.includes('CPY_BASIC_applied'), 'the pre-release build (the B03 boundary) is RED');
  const missing = base(); missing.READ.sheets = missing.READ.sheets.filter(n => !n.startsWith('L05CP_'));
  assert.ok(h.verdict(missing, sheetOf).failed.includes('CPY_BASIC_in_saved_file'));
  const differs = base(); differs.READ.copy.C2 = { v: '999', f: '', nf: '0.00' };
  assert.ok(h.verdict(differs, sheetOf).failed.includes('CPY_BASIC_cells_equal_source'), 'a changed value');
  for (const [name, mut] of [['the bold', c => { c.B1.bold = false; }], ['the number format', c => { delete c.C2.nf; }], ['the merge', c => { delete c.D1.merge; }], ['the formula text', c => { c.A3.f = ''; }]]) {
    const lost = base(); mut(lost.READ.copy);
    assert.ok(h.verdict(lost, sheetOf).failed.includes('CPY_BASIC_cells_equal_source'), `the copy lost ${name}`);
  }
  const vacuous = base(); delete vacuous.READ.copy.B1.bold; delete vacuous.READ.source.B1.bold;
  assert.ok(h.verdict(vacuous, sheetOf).failed.includes('CPY_BASIC_cells_equal_source'), 'an equality over a source that never had the format is vacuous');
  const repointed = base(); repointed.READ.crossRef.A1.f = '=L05CP_abc123!B2';
  assert.ok(h.verdict(repointed, sheetOf).failed.includes('CPY_BASIC_other_sheet_not_repointed'));
  const empty = base(); empty.READ.copy = {}; empty.READ.source = {};
  assert.ok(h.verdict(empty, sheetOf).failed.includes('CPY_BASIC_cells_equal_source'), 'two empty reads are not "equal"');
  const unnamed = base(); unnamed.NEG_NAMED.CPY_EXISTS.ops[0].failure = { reason: 'other' };
  assert.ok(h.verdict(unnamed, sheetOf).failed.includes('neg_CPY_EXISTS'));
  const dispatched = base(); dispatched.NEG_SHAPE.cpy_same_name = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [{ status: 'FAILED' }] };
  assert.ok(h.verdict(dispatched, sheetOf).failed.includes('neg_shape_cpy_same_name'));
});
