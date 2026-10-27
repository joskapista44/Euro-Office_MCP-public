'use strict';

/**
 * alice-1 (owner: CLASSIFICATION=BUG, backlog card, owner channel; LIVE repro test/live/xlsx-mona1-merged-anchor.live.cjs):
 *   (a) write_range into the TOP-LEFT (anchor) cell of a merged range works with normal worksheet semantics, the merge kept;
 *   (b) a non-anchor cell of a merged range is refused with a stable, non-empty reason -- never a silent FAILED;
 *   (c) that refusal comes BEFORE any mutation when every earlier step is a value/style write (no applied prefix);
 *   (d) a legitimate multistep sequence that CURES the refusal (unmerge_range, then the write) still runs;
 *   (e) every failure stage the canonical executor can record maps to a non-empty, specific reason.
 * Offline: the product WRITE tool with the fake XLSX editor behind the createDriver seam (test/fakes/write-wiring.cjs;
 * book(): Budget C1:D1 merged, C1 = 'merged').
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const W = require('../../src/server/write-tools.cjs');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const wr = (id, range, value) => ({ intent: 'write_range', id, sheet: 'Budget', range, values: [[value]] });
const val = (h, a) => { const s = h.store.durable.find(x => x.name === 'Budget'); return s.cells[a] ? s.cells[a].v : undefined; };
const merges = h => h.store.durable.find(x => x.name === 'Budget').merges;

test('alice-1a: the anchor (top-left) of a merged range is writable; the merge is kept; a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([wr('anchor', 'C1', 'Zöldterület kezelés')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.id, o.status, o.l1, o.l3]), [['anchor', 'APPLIED', 'PASS', 'PASS']]);
  assert.equal(val(h, 'C1'), 'Zöldterület kezelés');
  assert.deepEqual([...merges(h)], ['C1:D1']);
  const before = h.calls().dispatch;
  const again = await h.tool.handler(req([wr('anchor', 'C1', 'Zöldterület kezelés')]));
  assert.equal(again.ok, true);
  assert.deepEqual(again.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch - before, 0, 'retry: 0 dispatch');
});

test('alice-1b/c: a non-anchor merged cell after value writes -> REFUSED_PRE_MUTATION, a named reason, dispatch 0, no applied prefix', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([wr('pre', 'F1', 'pre'), wr('inside', 'D1', 'x'), wr('post', 'F2', 'post')]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_PRE_MUTATION', JSON.stringify(r));
  assert.match(r.reason, /^operation inside: xlsx-write-into-merged-area-refused$/);
  const inside = r.operations.find(o => o.id === 'inside');
  assert.deepEqual([inside.status, inside.failure.stage, inside.failure.reason, inside.failure.cell, inside.failure.mergeArea],
    ['FAILED', 'PREFLIGHT', 'xlsx-write-into-merged-area-refused', 'D1', 'C1:D1']);
  for (const id of ['pre', 'post']) {
    const o = r.operations.find(x => x.id === id);
    assert.deepEqual([o.status, o.failure.reason], ['SKIPPED_DUE_TO_PRIOR_FAILURE', 'skipped-due-to-prior-failure']);
  }
  assert.equal(h.calls().dispatch, 0, 'nothing dispatched');
  assert.equal(val(h, 'F1'), undefined, 'no applied prefix');
  assert.equal(r.persistence.state === 'FORCESAVE_CONFIRMED', false, 'nothing to persist');
});

test('alice-1d: a non-intersecting unmerge before the refused write does not cure it -> refused pre-mutation; an INTERSECTING one is a composition refusal', async () => {
  const two = book();
  two[0].merges = ['C1:D1', 'C3:D3'];
  const h = harness({ store: createStore(two) });
  const r = await h.tool.handler(req([{ intent: 'unmerge_range', id: 'unmerge', sheet: 'Budget', range: 'C3:D3' }, wr('inside', 'D1', 'x')]));
  assert.equal(r.code, 'REFUSED_PRE_MUTATION', JSON.stringify(r));
  assert.equal(h.calls().dispatch, 0);
  assert.deepEqual([...merges(h)], ['C1:D1', 'C3:D3'], 'the earlier unmerge did not run');
  // the curing order (unmerge the area, then write into it) cannot be one task: the canonical plan refuses overlapping writes
  const h2 = harness({ store: createStore(book()) });
  const r2 = await h2.tool.handler(req([{ intent: 'unmerge_range', id: 'unmerge', sheet: 'Budget', range: 'C1:D1' }, wr('inside', 'D1', 'x')]));
  assert.deepEqual([r2.ok, r2.stage, r2.code, h2.calls().dispatch], [false, 'PRE_DISPATCH', 'REFUSED_COMPOSITION', 0]);
});

test('GH79b: a non-intersecting earlier merge_range no longer forces a partial application -- REFUSED_PRE_MUTATION, 0 dispatch', async () => {
  // merge_range is not a product intent, but it cannot create/move a merge INTO a region it never touches: F5:G5 is
  // disjoint from the C1:D1 merge area that refuses the later write, so the whole task refuses before ANY dispatch --
  // before this fix, ANY earlier merge_range made the check non-decisive and this merge got APPLIED first (the
  // document was left in an intermediate state: a merge the task never meant to keep, the write still refused).
  const { withSession } = require('../../src/core/session/lifecycle.cjs');
  const { createExecutor } = require('../../src/core/execution/executor.cjs');
  const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
  const { createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
  const t = { id: 't', steps: [{ intent: 'merge_range', id: 'm', sheet: 'Budget', range: 'F5:G5' }, wr('inside', 'D1', 'x')] };
  let n = 0;
  const driver = createFakeXlsxDriver(createStore(book()), {});
  const final = await withSession({ driver, fileId: 'f', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  const r = W.mapExecutionResult(final, { fileId: 'f', task: t });
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_PRE_MUTATION', JSON.stringify(r));
  assert.equal(driver.calls.dispatch, 0, 'nothing dispatched, including the earlier merge');
  assert.deepEqual(r.operations.map(o => o.status), ['SKIPPED_DUE_TO_PRIOR_FAILURE', 'FAILED']);
  const inside = r.operations[1];
  assert.deepEqual([inside.failure.stage, inside.failure.reason, inside.failure.mergeArea], ['PREFLIGHT', 'xlsx-write-into-merged-area-refused', 'C1:D1']);
  assert.equal(r.reason, 'operation inside: xlsx-write-into-merged-area-refused');
});

test('GH79b: an earlier merge_range that DOES reach into the refused merge area keeps today\'s behavior (not decisive)', () => {
  // direct unit coverage of preflightDecisive: a geometric check, independent of whether this specific merge_range
  // would itself plan/classify (composition and the merge family's own content guard are separate concerns, already
  // covered by xlsx-w9-composition.test.cjs and xlsx-w4-merge.test.cjs)
  const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');
  const op = { intent: 'write_range', selector: { sheet: 'Budget', range: 'D1' } };
  const evidence = { reason: 'xlsx-write-into-merged-area-refused', cell: 'D1', mergeArea: 'C1:D1' };
  const touching = [{ intent: 'merge_range', selector: { sheet: 'Budget', range: 'B1:C1' } }];
  const apart = [{ intent: 'merge_range', selector: { sheet: 'Budget', range: 'F5:G5' } }];
  assert.equal(xlsxAdapter.preflightDecisive(touching, op, evidence), false, 'touches the merge area -> not decisive, as before');
  assert.equal(xlsxAdapter.preflightDecisive(apart, op, evidence), true, 'disjoint from the merge area -> decisive');
});

test('GH79b: the same geometric rule for a formula refusal, and for an earlier sort_range', () => {
  const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');
  const op = { intent: 'write_range', selector: { sheet: 'Budget', range: 'C2' } };
  const evidence = { reason: 'xlsx-write-over-formula-refused', cell: 'C2', formula: '=B2*2' };
  const touchingMerge = [{ intent: 'merge_range', selector: { sheet: 'Budget', range: 'B2:C2' } }];
  const apartMerge = [{ intent: 'merge_range', selector: { sheet: 'Budget', range: 'F5:G5' } }];
  const touchingSort = [{ intent: 'sort_range', selector: { sheet: 'Budget', range: 'A1:C3' } }];
  const apartSort = [{ intent: 'sort_range', selector: { sheet: 'Budget', range: 'A10:C12' } }];
  assert.equal(xlsxAdapter.preflightDecisive(touchingMerge, op, evidence), false);
  assert.equal(xlsxAdapter.preflightDecisive(apartMerge, op, evidence), true);
  assert.equal(xlsxAdapter.preflightDecisive(touchingSort, op, evidence), false);
  assert.equal(xlsxAdapter.preflightDecisive(apartSort, op, evidence), true);
});

test('GH79b: for any OTHER refusal reason, merge_range / sort_range still disqualify, unchanged', () => {
  const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');
  const op = { intent: 'write_range', selector: { sheet: 'Budget', range: 'A1' } };
  const evidence = { reason: 'xlsx-sheet-not-found' };
  const earlier = [{ intent: 'merge_range', selector: { sheet: 'Budget', range: 'F5:G5' } }];
  assert.equal(xlsxAdapter.preflightDecisive(earlier, op, evidence), false);
});

test('alice-1e: every failure stage the canonical executor records maps to a specific, non-empty reason', () => {
  const src = fs.readFileSync(path.join(__dirname, '../../src/core/execution/executor.cjs'), 'utf8');
  const stages = [...new Set([...src.matchAll(/stage: '([a-z-]+)'/g)].map(m => m[1]))];
  assert.ok(stages.length >= 6, `the executor stage literals were found (${stages.join(',')})`);
  for (const stage of stages) {
    const f = W.failureOf({ status: 'FAILED', evidence: { stage } }, { l1: null, l3: null });
    assert.equal(typeof f.reason, 'string');
    assert.notEqual(f.reason, '');
    assert.notEqual(f.stage, 'DISPATCH', `executor stage '${stage}' is mapped explicitly, not by the generic fallback`);
  }
  // degenerate evidence still yields a reason
  for (const e of [{ status: 'FAILED', evidence: null }, { status: 'FAILED' }, { status: 'DISPATCHED_OUTCOME_UNKNOWN' }, { status: 'SKIPPED_DUE_TO_PRIOR_FAILURE', evidence: {} }]) {
    const f = W.failureOf(e, { l1: null, l3: null });
    assert.ok(typeof f.reason === 'string' && f.reason.length > 0, JSON.stringify(e));
  }
  assert.equal(W.failureOf({ status: 'FAILED', evidence: { ok: true } }, { l1: 'FAIL', l3: null }).reason, 'l1-fail');
});
