'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const range = require('../../src/adapters/xlsx/families/range.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 5: copy_range / move_range on the fake transport. Copy: the natural terminal-state retry, an occupied
 * target refused (the source overwrote silently, B-S10).
 *
 * Move -- L-06/b (owner decision, zero 2026-10-15), REPLACING the owner decision 2026-09-30 caller-
 * supplied-`expect` contract: the caller NEVER supplies `expect` -- classify() builds it itself,
 * from its OWN fresh read, directly before dispatch (moveExpectation, the SAME mechanism, now called
 * SERVER-side instead of by the caller). A blank source (any state that LOOKS already-moved,
 * including a genuine retry) therefore has nothing to build a fresh expectation FROM, and is refused
 * named (xlsx-move-already-applied-not-provable), 0/0, BEFORE even checking the target -- v1 never
 * attempts to verify a retry's target content. Never a second Cut.
 */

const book = () => [
  { name: 'P', cells: { A1: { v: 'a' }, B1: { v: 1 }, A2: { v: 2, f: '=B1*2' }, B2: { v: 1, f: '=$B$1' }, H1: { v: 'm' }, H2: { v: 3 }, E9: { v: 'busy' } } },
  { name: 'Q', cells: { A1: { v: 'keep' } } },
];

const copy = (target, extra = {}) => ({ id: 'm6-w5r', steps: [{ intent: 'copy_range', sheet: 'P', range: 'A1:B2', target, ...extra }] });
const MOVE = { intent: 'move_range', sheet: 'P', range: 'H1:H2' };
const H12 = Object.freeze([[{ t: 'text', v: 'm' }], [{ t: 'number', v: '3' }]]);
const exp = (target, cells = H12, over = {}) => ({ version: 1, fileId: 'disposable', source: { sheet: 'P', range: 'H1:H2' }, target: { sheet: 'P', range: `${target}:${target.replace(/\d+/, n => String(Number(n) + 1))}` }, cells, ...over });
const move = (target, extra = {}) => ({ id: 'm6-w5r', steps: [{ ...MOVE, target, ...extra }] });

/** The preparatory read: a fresh live read of both blocks, then the expectation from it (what a caller does). */
async function prepared(store, step, script = {}) {
  const d = createFakeXlsxDriver(store, script);
  await d.open({ fileId: 'disposable' });
  const r = await d.readLive(null, range.moveReadScope(step));
  return range.moveExpectation(normalizeInventory(r.payload), step);
}

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const S = (store, n) => store.durable.find(s => s.name === n).cells;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('copy: APPLIED -- values, a relative formula translated, an absolute kept; retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, copy('D5'));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  assert.equal(S(store, 'P').D5.v, 'a');
  assert.equal(S(store, 'P').D6.f, '=E5*2');
  assert.equal(S(store, 'P').E6.f, '=$B$1');
  const again = await run(store, copy('D5'));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('copy to ANOTHER sheet: APPLIED, the other sheet keeps the rest', async () => {
  const store = createStore(book());
  const { result } = await run(store, copy('C3', { targetSheet: 'Q' }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification));
  assert.equal(S(store, 'Q').C3.v, 'a');
  assert.equal(S(store, 'Q').A1.v, 'keep');
});

test('move: classify() builds the expectation itself (no `expect` on the step), APPLIED; the reopened retry is REFUSED named, 0/0, NEVER SATISFIED (L-06/b, owner decision 2026-10-15)', async () => {
  const store = createStore(book());
  const t = move('J5');
  assert.equal(Object.keys(t.steps[0]).includes('expect'), false, 'the caller never supplies expect');
  const { result } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(S(store, 'P').J5.v, 'm');
  assert.equal(S(store, 'P').H1, undefined);
  const again = await run(store, t);
  assert.equal(again.result.ok, false, JSON.stringify(again.result.execution.appliedOperations[0].evidence));
  assert.equal(reason(again.result), 'xlsx-move-already-applied-not-provable');
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
  assert.equal(again.driver.calls.dispatch, 0);
  // the saved state is untouched by the refused retry
  assert.equal(S(store, 'P').J5.v, 'm');
  assert.equal(S(store, 'P').H1, undefined);
});

test('a blank source is ALWAYS refused xlsx-move-already-applied-not-provable, regardless of what the target holds (v1: never attempts to verify a retry)', async () => {
  for (const target of [{ J5: { v: 'm' }, J6: { v: 3 } }, { J5: { v: 'm' }, J6: { v: 'WRONG' } }, {}]) {
    const b = book();
    delete b[0].cells.H1; delete b[0].cells.H2;
    Object.assign(b[0].cells, target);
    const { result, driver, m } = await run(createStore(b), move('J5'));
    assert.equal(result.ok, false, JSON.stringify(target));
    assert.equal(m.commands, 0);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), 'xlsx-move-already-applied-not-provable');
  }
});

for (const [name, why, script] of [
  ['the document identity is not observable', 'xlsx-move-document-identity-unobservable', { fileIdUnobservable: true }],
  ['a cell type is unobservable (COUNT unread)', 'xlsx-move-expect-type-unobservable', { countUnreadable: true }],
]) {
  test(`RED: refused BEFORE dispatch -- ${name}`, async () => {
    const store = createStore(book());
    const { result, driver } = await run(store, move('J5'), script);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
    assert.deepEqual(store.durable, book());
  });
}

test('L-06/b: the whole-task check (L2) only confirms the source is left blank -- it can no longer re-prove the target content (that is L1\'s job, on the SAME op classify() mutated); the narrowed scope is NAMED via notProven', () => {
  const op = planOperations(move('J5')).operations[0];
  // op.expect stays null: wholeTaskChecks receives a FRESH re-plan (adapter.verifyWholeTask), and
  // finalObservation is POST-write (source already blank) -- there is no pre-write snapshot left
  // to build one from at this layer. See the function's own header comment.
  const obs = cells => ({ primary: normalizeInventory({ sheets: [{ name: 'P', index: 0, used: 'A1:J6' }], cells: Object.entries(cells).map(([a, c]) => ({ sheet: 'P', a, v: c.v, f: c.f || '', count: typeof c.v === 'number' ? 1 : 0, merge: a })), unsupported: [], runtime: { identity: 'b', fileId: 'disposable' } }) });
  const blank = { v: '' };
  const r1 = range.wholeTaskChecks(op, obs({ H1: blank, H2: blank, J5: { v: 'm' }, J6: { v: 3 } }))[0];
  assert.equal(r1.ok, true);
  assert.deepEqual(r1.notProven, ['move-target-content-equality-not-reprovable-at-l2']);
  // the target content being WRONG is no longer this layer's concern (named above) -- it still
  // passes, since the source IS blank; L1 is what caught a wrong target, in the SAME dispatch.
  assert.equal(range.wholeTaskChecks(op, obs({ H1: blank, H2: blank, J5: { v: 'm' }, J6: { v: 'WRONG' } }))[0].ok, true, 'target content is no longer checked here (L1\'s job)');
  assert.equal(range.wholeTaskChecks(op, obs({ H1: { v: 'm' }, H2: blank, J5: { v: 'm' }, J6: { v: 3 } }))[0].ok, false, 'the source must be blank');
  assert.equal(range.wholeTaskChecks(op, { primary: null })[0].ok, false, 'no final observation -> no whole-task pass');
});

test('the expectation builder refuses what a fresh read cannot prove', async () => {
  const bad = async (mutate, step, script, re) => { const b = book(); if (mutate) mutate(b); const e = await prepared(createStore(b), step, script); assert.equal(e.ok, false, JSON.stringify(e)); assert.match(e.reason, re); };
  await bad(null, { ...MOVE, target: 'J5' }, { countUnreadable: true }, /xlsx-move-expect-type-unobservable/);
  await bad(b => { b[0].cells.H2 = { v: 'TRUE' }; }, { ...MOVE, target: 'J5' }, {}, /xlsx-move-expect-type-unobservable/);
  await bad(b => { delete b[0].cells.H1; delete b[0].cells.H2; }, { ...MOVE, target: 'J5' }, {}, /xlsx-move-expect-nothing-to-prove/);
  await bad(null, { ...MOVE, target: 'J5' }, { fileIdUnobservable: true }, /xlsx-move-document-identity-unobservable/);
  await bad(null, { ...MOVE, sheet: 'NOPE', target: 'J5' }, {}, /xlsx-sheet-not-found/);
});

test('RED (owner review 8527132): a correct copy with a corrupted SOURCE -> L3 FAIL (same sheet AND another target sheet)', async () => {
  for (const t of [copy('D5'), copy('C3', { targetSheet: 'Q' })]) {
    const store = createStore(book());
    const { result, driver } = await run(store, t, { fault: 'copy-corrupts-source' });
    assert.equal(result.ok, false);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS', 'L1 sees a correct target');
    assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', 'L3 protects the source');
    assert.equal(driver.calls.forceSave, 0);
    assert.deepEqual(store.durable, book());
  }
});

for (const [name, t, why, mutate] of [
  ['copy onto an occupied, different target (no silent overwrite)', copy('D8'), 'xlsx-copy-target-occupied'],
  ['copy whose formula would leave the sheet', { id: 'x', steps: [{ intent: 'copy_range', sheet: 'P', range: 'A2', target: 'A1' }] }, 'xlsx-copy-formula-leaves-the-sheet'],
  ['move onto an occupied target', move('E9'), 'xlsx-move-target-occupied'],
  // M7 E R7: a reference ENTIRELY inside the moved block is re-pointed (xlsx-m7-e-move.test.cjs); a partial one refuses
  ['move of a block another formula references PARTLY (it would not follow)', move('J5'), 'xlsx-move-reference-shape-not-measured', b => { b[1].cells.B1 = { v: 3, f: '=SUM(P!H2:H3)' }; }],
  ['a merged source cell', copy('D5'), 'xlsx-range-merged-cell-refused', b => { b[0].merges = ['A1:B1']; delete b[0].cells.B1; }],
]) {
  test(`refused BEFORE dispatch: ${name}`, async () => {
    const b = book();
    if (mutate) mutate(b);
    const store = createStore(b);
    const { result, driver } = await run(store, t);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('RED: ignored / partial / not translated / source kept -> L1 FAIL; a side effect -> L3 FAIL; nothing persisted', async () => {
  for (const [t, fault] of [[copy('D5'), 'ignore'], [copy('D5'), 'partial'], [move('J5'), 'move-keeps-source'], [move('J5'), 'ignore']]) {
    const store = createStore(book());
    const { result, driver } = await run(store, t, { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
    assert.deepEqual(store.durable, book(), fault);
  }
  const side = await run(createStore(book()), copy('D5'), { fault: 'side-effect' });
  assert.equal(side.result.ok, false);
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('L-06 r6: a Copy that did NOT translate the formula (fault copy-no-translate) is REPAIRED by the planner-translated explicit write -- APPLIED, the saved target is the translated text, a retry SATISFIED', async () => {
  const store = createStore(book());
  const first = await run(store, copy('D5'), { fault: 'copy-no-translate' });
  assert.equal(first.result.ok, true, JSON.stringify(first.result.execution.appliedOperations[0] && first.result.execution.appliedOperations[0].evidence).slice(0, 400));
  const again = await run(store, copy('D5'), { fault: 'copy-no-translate' });
  assert.equal(again.result.execution.appliedOperations[0].status, 'SATISFIED');
});

test('planner: target shape, overlap, the only step of its task -- move_range no longer accepts `expect` at all (L-06/b)', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  refuse(copy('D5:E6'), /target must be ONE cell/);
  refuse(copy('B2'), /overlap on the same sheet/);
  refuse(copy('XFD1'), /leaves the sheet/);
  refuse({ id: 'x', steps: [{ ...MOVE, target: 'J5', anchor: { value: 'm' } }] }, /does not support field\(s\): anchor/);
  // L-06/b: `expect` itself is now an UNKNOWN field for move_range (the caller never supplies it --
  // classify() builds it, see the move_range tests above) -- same rejection path copy_range's own
  // `expect: {}` already used, below.
  refuse({ id: 'x', steps: [{ ...MOVE, target: 'J5', expect: exp('J5') }] }, /does not support field\(s\): expect/);
  refuse({ id: 'x', steps: [{ intent: 'copy_range', sheet: 'P', range: 'A1', target: 'D5', expect: {} }] }, /does not support field/);
  refuse({ id: 'x', steps: [copy('D5').steps[0], { intent: 'write_range', sheet: 'Q', range: 'Z9', values: [['v']] }] }, /only step of its task/);
  assert.equal(planOperations(copy('B2', { targetSheet: 'Q' })).ok, true, 'the same coordinates on ANOTHER sheet do not overlap');
});

test('L-06 r6: the Copy body writes the PLANNER\'s translated formulas (target.formulas) and reads each back; a target that took no formula fails the command, named', () => {
  const op = planOperations(copy('D5', { targetSheet: 'Q' })).operations[0];
  const cmd = range.lower(op, { from: {}, to: op.dest, formulas: { E6: '=E5+X9' } });
  assert.deepEqual(cmd.meta.formulaWrites, { E6: '=E5+X9' });
  assert.equal(cmd.meta.explicitFormulaWrite, true);
  const run = (formulaOf) => {
    const sets = [];
    const ws = n => ({ GetName: () => n, GetRange: a => ({ a, Copy: () => {}, GetFormula: () => formulaOf(n, a), SetValue: v => sets.push([n, a, v]) }) });
    return { sets, out: new Function('Api', cmd.body)({ GetSheet: n => ws(n) }) };
  };
  const good = run((n, a) => (n === 'Q' && a === 'E6' ? '=E5+X9' : ''));
  assert.deepEqual(good.out, { ok: true });
  assert.deepEqual(good.sets, [['Q', 'E6', '=E5+X9']], 'ONLY the planner-translated cell, on the TARGET sheet, with the planner text');
  assert.deepEqual(run(() => '').out, { ok: false, stage: 'xlsx-copy-formula-not-written', cell: 'E6' });
  // no formulas in the source block -> nothing is rewritten
  const none = range.lower(op, { from: {}, to: op.dest, formulas: {} });
  assert.doesNotMatch(none.body, /SetValue\("/);
  assert.deepEqual(none.meta.formulaWrites, {});
});

test('the command bodies: exact-name sheets, the public Copy / Cut onto the target top-left', () => {
  for (const [t, call] of [[copy('D5', { targetSheet: 'Q' }), 'Copy'], [move('J5'), 'Cut']]) {
    const cmd = range.lower(planOperations(t).operations[0]);
    const log = [];
    const ws = n => ({ GetName: () => n, GetRange: a => ({ a, GetFormula: () => '', [call]: dest => log.push([n, a, call, dest.a]) }) }); // L-06 r5: no formula cell -> nothing rewritten
    assert.deepEqual(new Function('Api', cmd.body)({ GetSheet: n => ws(n) }), { ok: true });
    assert.deepEqual(log, call === 'Copy' ? [['P', 'A1:B2', 'Copy', 'D5']] : [['P', 'H1:H2', 'Cut', 'J5']]);
  }
});
