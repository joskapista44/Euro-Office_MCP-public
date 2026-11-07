'use strict';

/**
 * GH #326 spinoff (origin: 9feecff3 design v2 review, maintainer's own NEM-MERT): does the sheet-level
 * protection-gate (protection-gate.cjs) falsely refuse a [unprotect_sheet, format_range]-shaped task on a
 * protected sheet, because adapter.cjs's preflightDecisive() treats EVERY protection refusal as decisive
 * "REGARDLESS of `earlier`" (a premise its OLD comment justified with "no intent in this product sets or
 * clears sheet protection" -- false since GH #338-#342 added protect_sheet / unprotect_sheet)?
 *
 * MEASURED ANSWER (the composition itself): no -- that composition cannot be built at all, on ANY path
 * through this codebase, not just the product's write-tools.cjs entry point. adapter.cjs's own
 * planOperations() (lines ~171-174) refuses EVERY multi-step task containing a 'protection'-family op
 * ('protect_sheet / unprotect_sheet / set_cell_protection must be the only step of its task')
 * unconditionally, and executeTask() (executor.cjs line ~85) calls planOperations() FIRST, before the
 * X-M01 preflightDecisive loop (~line 125) ever runs. So `operations` can never contain a protection op
 * beside another op, which means `earlier` can never contain a protection-CHANGING op either, when
 * preflightDecisive evaluates a LATER step's protection refusal. Part E measures this directly against the
 * raw executor (bypassing write-tools.cjs entirely); Part F measures the SAME refusal at the product
 * surface (write-tools.cjs has its OWN, redundant SOLE_STEP pre-check, line ~1353, that fires first there
 * and gives the nicer REFUSED_COMPOSITION code -- but even without it, planOperations() at write-tools.cjs
 * line ~1362 would catch it anyway).
 *
 * CORRECTED (a maintainer's review of v1): the conclusion above does NOT make the `return true` branch itself dead
 * code. "`earlier` can never contain a protection op" is a narrower guarantee than "the branch never fires
 * usefully" -- `earlier` routinely contains NON-protection ops (write_range, clear_range, ...), and when a
 * LATER step in that same task gets a protection refusal, the branch is exactly what stops the EARLIER,
 * non-protection step from having already dispatched (Part H, C5b): without it, a task like
 * [write_range(unlocked cell), format_range(forbidden)] would let write_range run, then fail on
 * format_range, leaving the task half-applied -- the SAME partial-write hazard preflightDecisive exists to
 * prevent for write_range/unmerge_range/set_chart_title below (lines ~441+). The branch is load-bearing;
 * only its OLD justification (no such intent exists) was stale. protectionRacePrelude (the card's guess for
 * the "reverse direction") is still never reached by [protect_sheet, format_range]: that composition is
 * refused at the SAME planOperations() step, for the SAME reason, regardless of the sheet's starting
 * protection state (Part G) -- measured, not assumed.
 *
 * Verdict: NO product-code change to behaviour. adapter.cjs's comment at the `return true` line is
 * corrected to state the TRUE premise (planOperations() keeps protection ops alone in their task, so no
 * `earlier` step -- none of which can be a protection change -- could have cured a protection refusal
 * already true at planning) instead of the false one ("no intent... sets or clears protection").
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const RT22 = '9.3.5|euro-office-bridge-v1/9.3.5.0/22';
const cell = (v, f = '') => ({ v, f });
const book = (ws = {}) => [{ name: 'P', cells: { A1: cell('a'), B2: cell(1), C1: cell(2, '=B2*2') }, objects: { charts: [], pivots: [], names: [] }, ...ws }];
const unprotect = (extra = {}) => ({ intent: 'unprotect_sheet', sheet: 'P', ...extra });
const protect = (extra = {}) => ({ intent: 'protect_sheet', sheet: 'P', forbids: { formatCells: false, insertRows: false }, ...extra });
const fmt = (extra = {}) => ({ intent: 'format_range', sheet: 'P', range: 'A1', format: { fill: '#FFC000' }, ...extra });

// ---- Part E: direct executor call, bypassing write-tools.cjs's validateWriteRequest entirely (the gh338- --
// ---- gh342 test file's own run() pattern) -- proves the refusal is NOT merely a write-tools.cjs front-end --

async function run(store, task, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT22, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(task, { observation: { live: liveScope(task) } }));
  return { result, driver };
}
const sheetOf = store => store.durable[0];

test('E1 (MEASURED, bypassing write-tools.cjs): [unprotect_sheet, format_range] on a protected sheet never reaches classification or dispatch -- adapter.cjs\'s planOperations() itself refuses the composition', async () => {
  const store = createStore(book({ protection: { forbids: { formatCells: true } } }));
  const task = { id: 'e1', steps: [unprotect(), fmt()] };
  const { result, driver } = await run(store, task);
  assert.equal(result.ok, false);
  assert.equal(result.outcome, 'plan-failed', 'planOperations(), not the preflightDecisive (X-M01) loop -- it never gets that far');
  assert.match(result.execution.finalEvidence ? JSON.stringify(result.execution.finalEvidence) : JSON.stringify(result), /unprotect_sheet must be the only step of its task/);
  assert.equal(driver.calls.bridgeWrites || 0, 0, 'ZERO dispatch');
  assert.deepEqual(sheetOf(store).protection.forbids, { formatCells: true }, 'left protected, unchanged');
});

test('E2 (MEASURED): the mirror order [format_range, unprotect_sheet] -- same refusal, same reason, order-independent', async () => {
  const store = createStore(book({ protection: { forbids: { formatCells: true } } }));
  const task = { id: 'e2', steps: [fmt(), unprotect()] };
  const { result, driver } = await run(store, task);
  assert.equal(result.ok, false);
  assert.equal(result.outcome, 'plan-failed');
  assert.equal(driver.calls.bridgeWrites || 0, 0);
});

// ---- Part F: the product entry point (write-tools.cjs) -- the SAME refusal, reached a different way --------

test('F1: [unprotect_sheet, format_range] is REFUSED_COMPOSITION at write-tools.cjs request validation, before planning or classification', () => {
  const v = W.validateWriteRequest(req([unprotect(), fmt()]));
  assert.equal(v.ok, false);
  assert.equal(v.code, 'REFUSED_COMPOSITION');
  assert.match(v.reason, /unprotect_sheet must be the only step of its task/);
});

test('F2: the mirror composition [format_range, unprotect_sheet] -- SOLE_STEP catches it at whatever index it sits', () => {
  const v = W.validateWriteRequest(req([fmt(), unprotect()]));
  assert.equal(v.ok, false);
  assert.equal(v.code, 'REFUSED_COMPOSITION');
  assert.match(v.reason, /step 1: unprotect_sheet must be the only step of its task/);
});

test('F3: protect_sheet is caught the same way as unprotect_sheet', () => {
  const v = W.validateWriteRequest(req([protect(), fmt()]));
  assert.equal(v.ok, false);
  assert.equal(v.code, 'REFUSED_COMPOSITION');
  assert.match(v.reason, /protect_sheet must be the only step of its task/);
});

test('F4: write-tools.cjs\'s own SOLE_STEP pre-check is redundant with (not the sole reason for) the refusal -- planOperations() alone already proves it (Part E); this just confirms the front-end also names it, with a 3rd operation present', () => {
  for (const step of [protect(), unprotect()]) {
    const v = W.validateWriteRequest(req([fmt(), step, fmt({ range: 'B1' })]));
    assert.equal(v.ok, false, step.intent);
    assert.equal(v.code, 'REFUSED_COMPOSITION', step.intent);
  }
});

// ---- Part G: the card's "reverse direction" ([protect_sheet, format_range], starting unprotected) ----------
// Hypothesis in the card: this direction hangs on protectionRacePrelude instead. MEASURED: it does not --
// it is refused at the SAME planOperations() step as Part E, for the SAME reason, regardless of the
// sheet's starting protection state. protectionRacePrelude is never reached; nothing dispatches at all.

test('G1 (product surface): also REFUSED_COMPOSITION, same mechanism as F3', () => {
  const v = W.validateWriteRequest(req([protect({ forbids: { formatCells: true } }), fmt()]));
  assert.equal(v.ok, false);
  assert.equal(v.code, 'REFUSED_COMPOSITION');
});

test('G2 (MEASURED, bypassing write-tools.cjs, sheet starts UNPROTECTED): [protect_sheet, format_range] is refused at planning too -- NOT at protectionRacePrelude, NOT after a dispatch; protect_sheet itself never dispatches either', async () => {
  const store = createStore(book({})); // no `protection` key -- starts unprotected
  const task = { id: 'g2', steps: [protect({ forbids: { formatCells: true } }), fmt()] };
  const { result, driver } = await run(store, task);
  assert.equal(result.ok, false);
  assert.equal(result.outcome, 'plan-failed', 'the SAME composition rule, independent of the starting protection state');
  assert.equal(driver.calls.bridgeWrites || 0, 0, 'protect_sheet never dispatches either -- the whole task is refused before planning completes');
  assert.equal(sheetOf(store).protection, undefined, 'the sheet is untouched');
});

// ---- Part H (a maintainer's review of v1, B1/B2): the `return true` branch is LOAD-BEARING, not dead -----------
// `earlier` can never contain a protection-CHANGING op (Parts E-G), but it routinely contains OTHER ops --
// and when a LATER step in the SAME task draws a protection refusal, the branch is what stops those
// earlier, unrelated ops from having already dispatched. Without it, this exact composition would let
// write_range apply, then fail on format_range, leaving the task half-written.

test('H1 (CONTROL): write_range into an UNLOCKED, NON-EMPTY cell, alone -- applies, no protection refusal of its own (B2 is pre-existing, non-empty: book()\'s own rangeCells() never lists an EMPTY cell as unlocked, Excel\'s fail-closed default -- protection-gate.cjs\'s own doc comment)', async () => {
  const store = createStore(book({ protection: { forbids: { formatCells: true }, unlocked: ['B2'] } }));
  const task = { id: 'h1-control', steps: [{ intent: 'write_range', sheet: 'P', range: 'B2', values: [['x']] }] };
  const { result, driver } = await run(store, task);
  assert.equal(result.ok, true, JSON.stringify(result).slice(0, 300));
  assert.equal(driver.calls.dispatch, 1, 'write_range dispatches through the plain cells family, not a bridgeWrites-counted one');
  assert.equal(sheetOf(store).cells.B2.v, 'x');
});

test('H2 (C5b -- the case the "dead code" claim missed): [write_range(B2, unlocked), format_range(A1, forbidden)] on the SAME protected sheet -- the earlier write_range is refused BEFORE it dispatches, by the SAME preflightDecisive branch Parts E-G proved can never see a protection op in `earlier`', async () => {
  const store = createStore(book({ protection: { forbids: { formatCells: true }, unlocked: ['B2'] } }));
  const task = {
    id: 'h2',
    steps: [
      { intent: 'write_range', sheet: 'P', range: 'B2', values: [['x']] },
      fmt(), // A1, forbidden by forbids.formatCells
    ],
  };
  const { result, driver } = await run(store, task);
  assert.equal(result.ok, false, JSON.stringify(result).slice(0, 300));
  assert.equal(result.outcome, 'preflight-fail-closed', 'the X-M01 loop -- THIS is the branch zero\'s review named, not plan-failed (H1/H2 do not touch any protection-family op, so planOperations() lets the composition through)');
  assert.equal(driver.calls.dispatch || 0, 0, 'ZERO dispatch -- write_range never ran, even though it carried no refusal of its own');
  assert.equal(sheetOf(store).cells.B2.v, 1, 'B2 untouched: a partial write was prevented, not just a protection change');
});

// Local-only NEG. CONTROL (a maintainer's B2, not committed): commenting out the `if (evidence && PROTECTION_
// REFUSAL.test(...)) return true;` line at the top of preflightDecisive and re-running H2 by hand gives
// result.outcome 'operation-fail-closed' and driver.calls.dispatch === 1 -- write_range DID dispatch
// before format_range's refusal stopped the task, i.e. exactly the partial-write hazard the branch exists
// to prevent. Confirmed locally; not kept as a standing test (it would require live-patching the adapter).
