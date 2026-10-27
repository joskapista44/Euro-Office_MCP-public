'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * L-16/2 (class 1.1 of the format-remove refusals plan; LIVE-measured test/live/xlsx-l16-2-merged-ring-probe.live.cjs, prod DS, session + reopen):
 * the composite writes only removeCells() -- the target cell(s) and the neighbour that faces each removed side. A ring cell that is a merge
 * part and is NOT written (a corner, a neighbour on a side that is not removed) is read back unchanged (corner R1B, left R1C_L, below R1C_B,
 * plain control PASS); a merged neighbour that IS written (R1A) gets its merge dissolved (MEASURED_FAIL). So `xlsx-format-remove-merged-ring-refused`
 * stays for the written cells and goes for the untouched ones. Every other refusal class is unchanged.
 * L-16/3 then released the WRITTEN merged cells too (C5: UnMerge -> composite -> Merge, xlsx-l16-3-c5-remove-merge.test.cjs): the merged-ring refusal is now only for an
 * area that cannot be re-issued, and what C5 was not measured for is refused under its own names. The cases below that pinned the L-16/2 refusals now pin the C5 contract.
 */

const thin = { s: 7, c: 255 };
const blank = { v: 0 };
const task = (range, sides) => ({ id: 'l16-2-merged-ring', steps: [{ intent: 'format_range', sheet: 'B', range, format: { border: { mode: 'remove', sides } } }] });
const ws = store => store.durable.find(s => s.name === 'B');

/** the target block (D3 / D3:D4) plus its one-cell ring, every cell readable; D2 carries the shared bottom, D3 the removed top. */
function seed(rows = [2, 3, 4]) {
  const cells = {};
  for (const c of ['C', 'D', 'E']) for (const r of rows) cells[`${c}${r}`] = { ...blank };
  cells.D2 = { ...blank, border: { b: thin } };
  cells.D3 = { v: 'x', border: { t: thin } };
  return cells;
}

async function run(cells, merges, range = 'D3:D3', sides = ['top']) {
  const store = createStore([{ name: 'B', cells, merges }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  let n = 0;
  const t = task(range, sides);
  const result = await withSession({ driver: createFakeXlsxDriver(store, {}), fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, store, evidence: result.execution.appliedOperations[0].evidence };
}
const passed = (r, merges) => {
  assert.equal(r.result.ok, true, JSON.stringify([r.evidence, r.result.execution.primitiveVerification, r.result.execution.preservationVerification]));
  assert.deepEqual([r.m.writes, r.m.commands], [1, 1]);
  assert.deepEqual(ws(r.store).cells.D3.border.t, { s: 0, c: 0 }, 'the removed top is gone');
  assert.deepEqual(ws(r.store).cells.D2.border.b, { s: 0, c: 0 }, 'and so is the facing bottom of the neighbour');
  assert.deepEqual(ws(r.store).merges, merges, 'the untouched merge survives');
};

test('R1A: the neighbour that FACES the removed side is a merge (the composite writes it) -> released through C5 (L-16/3): applied, the merge survives', async () => {
  const r = await run({ ...seed(), D1: { v: 'm' } }, ['D1:D2']);
  passed(r, ['D1:D2']);
});

test('R1B: a merged ring CORNER (never written) passes; the merge survives', async () => {
  passed(await run({ ...seed(), C1: { v: 'm' } }, ['C1:C2']), ['C1:C2']);
});

test('R1C: a merged neighbour on a side that is NOT removed -- left, right, below -- passes; the merge survives', async () => {
  passed(await run({ ...seed(), B3: { v: 'm' } }, ['B3:C3']), ['B3:C3']);
  passed(await run({ ...seed(), E3: { v: 'm' } }, ['E3:F3']), ['E3:F3']);
  passed(await run({ ...seed([2, 3, 4, 5, 6]), D4: { v: 'm' } }, ['D4:D5']), ['D4:D5']);
});

test('a merged target that CROSSES the requested range stays refused (named); a corner-anchored merge that reaches the facing neighbour is a C5 area and passes', async () => {
  const t = await run({ ...seed(), E3: { v: 'm' } }, ['D3:E3']);
  assert.deepEqual([t.m.writes, t.m.commands], [0, 0]);
  assert.equal(t.result.ok, false);
  assert.equal(t.evidence.evidence.reason, 'xlsx-format-merge-crosses-range-refused');
  passed(await run({ ...seed(), C1: { v: 'm' } }, ['C1:D2']), ['C1:D2']);
});

test('a range: bottom removal on D3:D4 writes the merge on D5 (C5, applied), top removal never touches it (applied, merge untouched); both keep the merge', async () => {
  const rows = [2, 3, 4, 5, 6];
  const cells = () => { const c = seed(rows); c.D4 = { v: 'y', border: { b: thin } }; c.D5 = { ...blank, border: { t: thin } }; return c; };
  const bottom = await run({ ...cells(), D6: { v: 'm' } }, ['D5:D6'], 'D3:D4', ['bottom']);
  assert.equal(bottom.result.ok, true, JSON.stringify(bottom.evidence));
  assert.deepEqual(ws(bottom.store).merges, ['D5:D6']);
  assert.deepEqual(bottom.result.execution.appliedOperations[0].status, 'APPLIED');
  const top = await run({ ...cells(), D6: { v: 'm' } }, ['D5:D6'], 'D3:D4', ['top']);
  assert.equal(top.result.ok, true, JSON.stringify(top.evidence));
  assert.deepEqual(ws(top.store).merges, ['D5:D6']);
});

test('every other refusal class is unchanged next to a passing merge: an unrestorable diagonal on the target', async () => {
  // L-16/1 (on main) restores a drawn diagonal WITH a direction flag; a drawn diagonal with NO direction flag is the remaining refusal on this axis
  const diag = { ...seed(), C1: { v: 'm' }, D3: { v: 'x', border: { t: thin, d: { s: 7, c: 0 } } } };
  const d = await run(diag, ['C1:C2']);
  assert.deepEqual([d.m.writes, d.m.commands], [0, 0]);
  assert.deepEqual([d.evidence.evidence.reason, d.evidence.evidence.why], ['xlsx-format-remove-unrestorable-value', 'no-direction']);
});
