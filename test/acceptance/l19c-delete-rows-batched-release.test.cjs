'use strict';

/**
 * L-19 C (owner-decision C, owner channel, valasz a 37193-ra: "C, ha kell tobb lepesben toroljunk";
 * kiosztó: zero, kártya an internal branch): több `delete_rows`
 * lépés egy taskban, kötegelve -- a 100-soros per-lépés korlát (L-19) változatlan.
 *
 * TERVEZÉSI DÖNTÉS (a kártya #ref-es kommentje): minden delete_rows lépés `at`/`anchor`-je a task
 * SAJÁT, korábbi lépései UTÁN kialakult állapotra vonatkozik -- ezt a motor (executor.cjs,
 * plan.cjs NOT_PROVEN szabálya, structure.classify live-anchor olvasása) már MA is így hajtja végre,
 * ez a fájl a PLANNER-SZINTŰ kompozíciós tiltás három pontját pinneli:
 *   1. write-tools.cjs SOLE_STEP -- delete_rows kikerült, BATCHED_REPEATABLE nézi a homogenitást
 *   2. adapter.cjs planOperations -- a "structural step alone" tiltás alól delete_rows-only taskok kivételt kapnak
 *   3. adapter.cjs pairOf -- két structure-op plan-idejű literal-range overlap-vizsgálata nem fut (más időpillanat, más szám)
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { validateWriteRequest } = require('../../src/server/write-tools.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const FILE = '1303626';
const req = steps => ({ fileId: FILE, task: { id: 't', steps } });
const delAt = (at, count, value, extra = {}) => ({ intent: 'delete_rows', id: `s${at}`, sheet: 'S', at, count, anchor: { column: 'A', value }, ...extra });

test('two delete_rows in one task, second step in post-first-step coordinates: plans, both operations', () => {
  const r = validateWriteRequest(req([delAt(10, 5, 'a'), delAt(12, 5, 'b')]));
  assert.equal(r.ok, true, J(r));
  assert.equal(r.plannedOperations, 2);
});

test('the card scenario: 3 x 100-row delete_rows chained in one task plans cleanly', () => {
  const r = validateWriteRequest(req([delAt(10, 100, 'a'), delAt(110, 100, 'b'), delAt(210, 100, 'c')]));
  assert.equal(r.ok, true, J(r));
  assert.equal(r.plannedOperations, 3);
});

test('each step of a batched task still obeys the 100-row per-step ceiling (L-19 unchanged)', () => {
  const r = validateWriteRequest(req([delAt(10, 100, 'a'), delAt(110, 101, 'b')]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_SHAPE');
  assert.match(r.reason, /count/);
});

test('delete_rows mixed with insert_rows: still refused (insert_rows is still in SOLE_STEP, so that rule fires first)', () => {
  const insert = { intent: 'insert_rows', id: 'i', sheet: 'S', at: 2, count: 1, anchor: { column: 'A', value: 'z' } };
  const r = validateWriteRequest(req([delAt(10, 5, 'a'), insert]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_COMPOSITION');
  assert.match(r.reason, /insert_rows must be the only step of its task/);
});

test('delete_rows mixed with a NON-sole-step, non-structure intent still hits the BATCHED_REPEATABLE homogeneity message', () => {
  const other = { intent: 'format_range', id: 'f', sheet: 'S', range: 'A1', format: { numberFormat: '0' } };
  const r = validateWriteRequest(req([delAt(10, 5, 'a'), other]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_COMPOSITION');
  assert.match(r.reason, /delete_rows may only share a task with other delete_rows steps/);
});

test('delete_rows mixed with delete_columns: still refused (different intent, not just different family)', () => {
  const delCol = { intent: 'delete_columns', id: 'c', sheet: 'S', at: 'B', count: 1, anchor: { row: 1, value: 'z' } };
  const r = validateWriteRequest(req([delAt(10, 5, 'a'), delCol]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_COMPOSITION');
});

test('delete_rows mixed with write_range: still refused', () => {
  const other = { intent: 'write_range', id: 'w', sheet: 'S', range: 'A1', values: [['x']] };
  const r = validateWriteRequest(req([delAt(10, 5, 'a'), other]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_COMPOSITION');
});

test('insert_rows still may not repeat with itself: BATCHED_REPEATABLE is delete_rows-only, SOLE_STEP still catches insert_rows', () => {
  const a = { intent: 'insert_rows', id: 'a', sheet: 'S', at: 1, count: 1, anchor: { column: 'A', value: 'a' } };
  const b = { intent: 'insert_rows', id: 'b', sheet: 'S', at: 5, count: 1, anchor: { column: 'A', value: 'b' } };
  const r = validateWriteRequest(req([a, b]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_COMPOSITION');
  assert.match(r.reason, /insert_rows must be the only step of its task/);
});

test('adapter.cjs planOperations directly: homogeneous delete_rows batch plans; a lone delete_rows is unaffected', () => {
  const single = planOperations({ steps: [delAt(10, 5, 'a')] });
  assert.equal(single.ok, true, J(single));
  assert.equal(single.operations.length, 1);

  const batched = planOperations({ steps: [delAt(10, 5, 'a'), delAt(12, 5, 'b'), delAt(14, 5, 'c')] });
  assert.equal(batched.ok, true, J(batched));
  assert.equal(batched.operations.length, 3);
});

test('adapter.cjs planOperations: delete_rows batch still refuses when any step is a different structure intent', () => {
  const mixed = planOperations({ steps: [
    { id: 's1', intent: 'delete_rows', sheet: 'S', at: 10, count: 5, anchor: { column: 'A', value: 'a' } },
    { id: 's2', intent: 'insert_rows', sheet: 'S', at: 2, count: 1, anchor: { column: 'A', value: 'b' } },
  ] });
  assert.equal(mixed.ok, false);
  assert.match(mixed.detail, /unless every step of the task is delete_rows/);
});

test('plan-time literal-range overlap between two delete_rows steps is NOT refused -- it compares two different points in the task\'s own timeline (the pairOf exemption); correctness is the live per-step anchor read, not a plan-time number comparison', () => {
  // step 1 deletes rows 10-14; step 2's `at`=12 is a post-shift position (it names rows that, before
  // step 1 ran, sat at 17-21) -- literally 12-16 overlaps 10-14, but they are not the same rows.
  const r = planOperations({ steps: [
    { id: 's1', intent: 'delete_rows', sheet: 'S', at: 10, count: 5, anchor: { column: 'A', value: 'a' } },
    { id: 's2', intent: 'delete_rows', sheet: 'S', at: 12, count: 5, anchor: { column: 'A', value: 'b' } },
  ] });
  assert.equal(r.ok, true, J(r));
  assert.equal(r.operations.length, 2);
  assert.equal(r.operations.some(o => o.dependsOn), false, 'no dependsOn is manufactured between them either -- each is classified/resolved live, on its own turn');
});

test('a non-structure overlap pair is still refused as before (the exemption is structure-vs-structure only)', () => {
  const r = planOperations({ steps: [
    { id: 'w1', intent: 'write_range', sheet: 'S', range: 'A1:A5', cells: [{ a: 'A1', v: 1 }], values: [[1]] },
    { id: 'w2', intent: 'write_range', sheet: 'S', range: 'A3:A6', cells: [{ a: 'A3', v: 2 }], values: [[2]] },
  ] });
  // this pair's exact refusal shape depends on cells.cjs's own selector/footprint; the assertion that
  // matters here is narrower: it must NOT be let through as if it were a structure pair.
  assert.notEqual(r.ok, true, 'two overlapping write_range steps must not be silently accepted by the structure-only exemption');
});

/**
 * EXECUTION-level tests (a maintainer's review, card comment #ref, four numbered items): the plan-time
 * tests above pin the composition refusal; these four pin the actual chained-delete BEHAVIOUR
 * through the full classify -> resolve -> dispatch -> verify loop, on the fake transport (M6 wave 5's
 * own harness, test/fakes/fake-xlsx-driver.cjs -- the same one xlsx-w5-structure.test.cjs uses).
 */

const execBook = () => [
  { name: 'S', cells: Object.fromEntries(Array.from({ length: 9 }, (_, i) => [`A${i + 1}`, { v: `r${i + 1}` }])) },
];

const execTask = steps => ({ id: 'l19c', steps: steps.map((s, i) => ({ id: `s${i}`, sheet: 'S', ...s })) });

async function runExec(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver };
}

const S = store => store.durable.find(s => s.name === 'S').cells;
const statusOf = (result, id) => result.execution.appliedOperations.find(e => e.operationId === id).status;

// original rows 1..9 = r1..r9. Step 1 deletes rows 2-3 (original r2, r3): survivors become
// r1 r4 r5 r6 r7 r8 r9 at rows 1..7. Step 2's `at` is POST-shift: it targets what were originally
// r7/r8, now sitting at rows 5-6 (after step 1's -2 shift) -- the anchor that must be there BEFORE
// step 2 runs is `anchor.now` = at + count = 7, which (after step 1) holds the original r9.
const chainedSteps = (step2Anchor = 'r9') => [
  { intent: 'delete_rows', at: 2, count: 2, anchor: { column: 'A', value: 'r4' } },
  { intent: 'delete_rows', at: 5, count: 2, anchor: { column: 'A', value: step2Anchor } },
];

test('1. two delete_rows chained, the 2nd `at` in post-first-step coordinates: both APPLIED, final state = r1,r4,r5,r6,r9', async () => {
  const store = createStore(execBook());
  const t = execTask(chainedSteps());
  const { result, driver } = await runExec(store, t);
  assert.equal(result.ok, true, J([result.execution.appliedOperations, result.execution.wholeTaskVerification]));
  assert.equal(statusOf(result, 's0'), OPERATION_STATUS.APPLIED);
  assert.equal(statusOf(result, 's1'), OPERATION_STATUS.APPLIED);
  assert.equal(driver.calls.dispatch, 2);
  const s = S(store);
  assert.equal(s.A1.v, 'r1');
  assert.equal(s.A2.v, 'r4');
  assert.equal(s.A3.v, 'r5');
  assert.equal(s.A4.v, 'r6');
  assert.equal(s.A5.v, 'r9');
  assert.equal(s.A6, undefined, 'the book shrank from 9 to 5 rows');
});

test('2. the 2nd step\'s anchor wrong for the state the 1st step produced: FAIL_CLOSED, step 1 stays APPLIED, no data loss', async () => {
  const store = createStore(execBook());
  // step 2 names 'r8' -- a value that DOES exist in column A, but after step 1's shift it sits at
  // row 6, which is neither this step's expected-before position (anchor.now = at+count = 7) nor
  // its expected-after position (anchor.terminal = at = 5). A caller who forgot to re-derive the
  // anchor for the state step 1 already produced lands exactly here -- the mistake this whole
  // design decision exists to make loud, not silent.
  const t = execTask(chainedSteps('r8'));
  const { result, driver } = await runExec(store, t);
  assert.equal(result.ok, false, J(result.execution.appliedOperations));
  assert.equal(statusOf(result, 's0'), OPERATION_STATUS.APPLIED, 'the 1st step is not rolled back by the 2nd step\'s refusal');
  assert.equal(statusOf(result, 's1'), OPERATION_STATUS.FAILED);
  assert.equal(result.execution.appliedOperations.find(e => e.operationId === 's1').evidence.evidence.reason, 'xlsx-structure-anchor-state-conflict');
  assert.equal(driver.calls.dispatch, 1, 'the 2nd step never reached the transport');
  const s = S(store);
  assert.equal(s.A1.v, 'r1');
  assert.equal(s.A2.v, 'r4', 'step 1\'s own effect is durable -- no data loss from the 2nd step\'s refusal');
  assert.equal(s.A5.v, 'r7', 'step 2 never ran: the original r7 is still at its step-1-shifted position');
});

test('3. partial-task retry: the already-applied step is SATISFIED (0 writes), the not-yet-applied one is APPLY and runs', async () => {
  const store = createStore(execBook());
  // reproduce test 2's partial state first (step 1 applied, step 2 refused) ...
  await runExec(store, execTask(chainedSteps('r8')));
  assert.equal(S(store).A2.v, 'r4');
  assert.equal(S(store).A5.v, 'r7', 'still unshifted: step 2 has not run yet');
  // ... then retry the SAME task, this time with step 2's anchor corrected to match the state step 1
  // already produced (exactly what a caller does after reading the FAIL_CLOSED reason).
  const t2 = execTask(chainedSteps('r9'));
  const { result, driver } = await runExec(store, t2);
  assert.equal(result.ok, true, J(result.execution.appliedOperations));
  assert.equal(statusOf(result, 's0'), OPERATION_STATUS.SATISFIED, 'step 1: already done, 0 re-write');
  assert.equal(statusOf(result, 's1'), OPERATION_STATUS.APPLIED, 'step 2: now runs');
  assert.equal(driver.calls.dispatch, 1, 'only the retried (2nd) step dispatches; the satisfied 1st does not');
  const s = S(store);
  assert.equal(s.A2.v, 'r4');
  assert.equal(s.A5.v, 'r9');
});

test('4. the two compositions zero named as still-refused: delete_rows + write_range, and insert_rows repeated', () => {
  const dr = { intent: 'delete_rows', id: 'd', sheet: 'S', at: 2, count: 1, anchor: { column: 'A', value: 'r2' } };
  const wr = { intent: 'write_range', id: 'w', sheet: 'S', range: 'B1', values: [['x']] };
  const r1 = validateWriteRequest(req([dr, wr]));
  assert.equal(r1.ok, false);
  assert.equal(r1.code, 'REFUSED_COMPOSITION');

  const ir1 = { intent: 'insert_rows', id: 'i1', sheet: 'S', at: 1, count: 1, anchor: { column: 'A', value: 'r1' } };
  const ir2 = { intent: 'insert_rows', id: 'i2', sheet: 'S', at: 3, count: 1, anchor: { column: 'A', value: 'r3' } };
  const r2 = validateWriteRequest(req([ir1, ir2]));
  assert.equal(r2.ok, false);
  assert.equal(r2.code, 'REFUSED_COMPOSITION');
  assert.match(r2.reason, /insert_rows must be the only step of its task/);
});

/**
 * a maintainer's product-defect finding (card comment #ref, code-reading of structure.cjs classify()
 * line 354): a later delete_rows step's anchor can coincidentally already sit at ITS OWN terminal
 * position while an earlier step of the SAME run still needed a fresh APPLY -- the classic trigger
 * is a caller who (wrongly) gives every step in TASK-START coordinates. That reads as SATISFIED
 * (0 writes, ok=true) instead of the loud failure this whole design decision exists to produce.
 * The fix: SATISFIED operations of one run must be an unbroken PREFIX (a retry can only ever find
 * an EARLIER run's prefix already done) -- adapter.cjs classifyOperation now refuses a SATISFIED
 * verdict that follows an operation THIS run already had to APPLY, named
 * `xlsx-structure-satisfied-after-applied-in-batch`.
 */

const bigBook = () => [
  { name: 'S', cells: Object.fromEntries(Array.from({ length: 30 }, (_, i) => [`A${i + 1}`, { v: `R${i + 1}` }])) },
];

test('5. zero\'s exact example: a 2nd step given in TASK-START coordinates lands exactly on its own terminal -- refused, not silently SATISFIED', async () => {
  const store = createStore(bigBook());
  const t = execTask([
    { intent: 'delete_rows', at: 10, count: 5, anchor: { column: 'A', value: 'R15' } },
    // task-START coordinates (the natural mistake): at=17/anchor=R22 is correct for the ORIGINAL
    // book, but after step 1's -5 shift, R22 now sits at row 17 -- exactly this step's own
    // `terminal` (17), so the unguarded classify would read it as already-done.
    { intent: 'delete_rows', at: 17, count: 5, anchor: { column: 'A', value: 'R22' } },
  ]);
  const { result, driver } = await runExec(store, t);
  assert.equal(result.ok, false, J(result.execution.appliedOperations));
  assert.equal(statusOf(result, 's0'), OPERATION_STATUS.APPLIED);
  assert.equal(statusOf(result, 's1'), OPERATION_STATUS.FAILED);
  assert.equal(result.execution.appliedOperations.find(e => e.operationId === 's1').evidence.evidence.reason, 'xlsx-structure-satisfied-after-applied-in-batch');
  assert.equal(driver.calls.dispatch, 1, 'the 2nd step is refused BEFORE dispatch, not just after a no-op');
  const s = S(store);
  assert.equal(s.A17.v, 'R22', 'rows 17-21 (current numbering) were never deleted -- no silent false success');
});

test('6. legit retry (1st and 2nd already done from an earlier run, 3rd is new) is NOT caught by the invariant: SATISFIED, SATISFIED, APPLIED', async () => {
  const store = createStore(execBook());
  // each step's own anchor value must SURVIVE the later steps (never itself be the thing a later
  // step deletes), or a retry could not re-verify it by that anchor at all -- step 0 deletes r2/r3
  // (survivor anchor r4), step 1 deletes the (post-shift) current rows 3-4 = r5/r6 (survivor r7),
  // step 2 deletes the (post-shift) current row 4 = r8 (survivor r9). None of r4/r7/r9 is ever removed.
  const threeSteps = [
    { intent: 'delete_rows', at: 2, count: 2, anchor: { column: 'A', value: 'r4' } },
    { intent: 'delete_rows', at: 3, count: 2, anchor: { column: 'A', value: 'r7' } },
  ];
  // run only the first two first, to reproduce "1st and 2nd already done" as an EARLIER run
  await runExec(store, execTask(threeSteps));
  assert.equal(S(store).A1.v, 'r1');
  assert.equal(S(store).A2.v, 'r4');
  assert.equal(S(store).A3.v, 'r7', 'r5/r6 deleted by step 1; r7 is the survivor at row 3');
  // retry with all three -- 1 and 2 must come back SATISFIED (0 writes each), 3 is genuinely new
  const t3 = execTask([...threeSteps, { intent: 'delete_rows', at: 4, count: 1, anchor: { column: 'A', value: 'r9' } }]);
  const { result, driver } = await runExec(store, t3);
  assert.equal(result.ok, true, J(result.execution.appliedOperations));
  assert.equal(statusOf(result, 's0'), OPERATION_STATUS.SATISFIED);
  assert.equal(statusOf(result, 's1'), OPERATION_STATUS.SATISFIED);
  assert.equal(statusOf(result, 's2'), OPERATION_STATUS.APPLIED);
  assert.equal(driver.calls.dispatch, 1, 'only the 3rd (genuinely new) step dispatches');
  assert.equal(S(store).A4.v, 'r9');
});
