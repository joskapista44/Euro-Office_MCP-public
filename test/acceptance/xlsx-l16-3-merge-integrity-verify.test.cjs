'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * L-16/3 (a): the product's verification must notice a MERGE changing on a cell the write did not target -- a ring cell, the facing neighbour, any other read cell
 * of the sheet. Before this check L1 compared `merge` only for `op.cells` and L3 compared STYLE_KEYS (no `merge`), so a ring cell's merge dissolving after a border-remove write
 * read ok=true, L1 PASS, L3 PASS (docs/evidence/l16-3-ring-merge-dissolve-undetected-measure.cjs; the C5 path needs this before it can ship).
 * The scenario is the L-16/2 corner merge (C1:C2, a ring corner the composite never writes). The mutation is applied to the READS after the write (the fake's reader shape = the
 * DS reader's `merge` field), and a POSITIVE control (a ring fill change) proves the mutation channel reaches L3.
 */

const thin = { s: 7, c: 255 };
const blank = { v: 0 };

function seed() {
  const cells = {};
  for (const c of ['C', 'D', 'E']) for (const r of [2, 3, 4]) cells[`${c}${r}`] = { ...blank };
  cells.D2 = { ...blank, border: { b: thin } };
  cells.D3 = { v: 'x', border: { t: thin } };
  cells.C1 = { v: 'm' };
  return cells;
}

/** `mutate(cell)` rewrites a cell of the reads taken AFTER the write; returns the number of cells it touched. */
async function run({ merges = ['C1:C2'], mutate = null, cells = seed() }) {
  const store = createStore([{ name: 'B', cells, merges }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  const inner = createFakeXlsxDriver(store, {});
  let wrote = false; let touched = 0;
  const driver = new Proxy(inner, {
    get: (t, k) => {
      if (k === 'dispatch') return async (...a) => { wrote = true; return t.dispatch(...a); };
      if (k === 'readLive') return async (...a) => { const r = await t.readLive(...a); if (mutate && wrote && r && r.payload && Array.isArray(r.payload.cells)) for (const c of r.payload.cells) if (mutate(c)) touched += 1; return r; };
      return typeof t[k] === 'function' ? t[k].bind(t) : t[k];
    },
  });
  let n = 0; const t = { id: 'x', steps: [{ intent: 'format_range', sheet: 'B', range: 'D3:D3', format: { border: { mode: 'remove', sides: ['top'] } } }] };
  const result = await withSession({ driver, fileId: 'd', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  const l3 = (result.execution.preservationVerification || [])[0];
  return { result, touched, l3: l3 && l3.verdict, ev: l3 && l3.evidence && l3.evidence.evidence, l1: (result.execution.primitiveVerification || []).map(v => v.verdict) };
}

test('NEG control (no false alarm): a write that touches no merge -> ok, L1 PASS, L3 PASS, no mergeChanged', async () => {
  const r = await run({});
  assert.equal(r.result.ok, true, JSON.stringify(r.ev));
  assert.deepEqual([r.l1, r.l3], [['PASS'], 'PASS']);
  assert.deepEqual(r.ev.mergeChanged, []);
  assert.equal(r.ev.mergeChangedCount, 0);
});

test('RED->GREEN: a ring cell\'s merge DISSOLVING after the write fails L3 with the cell, was and now named', async () => {
  const r = await run({ mutate: c => (c.a === 'C1' || c.a === 'C2') && ((c.merge = c.a), true) });
  assert.equal(r.touched, 2, 'the mutation reached the post-write reads');
  assert.equal(r.l3, 'FAIL');
  assert.equal(r.result.ok, false);
  assert.deepEqual(r.ev.mergeChanged, [{ cell: 'C1', was: 'C1:C2', now: 'C1' }, { cell: 'C2', was: 'C1:C2', now: 'C2' }]);
  assert.equal(r.ev.mergeChangedCount, 2);
});

test('the merge EXTENT or ANCHOR changing (not only dissolving) fails too; so does a merge APPEARING on a cell that was not one', async () => {
  const ext = await run({ mutate: c => (c.a === 'C1' || c.a === 'C2') && ((c.merge = 'C1:C3'), true) });
  assert.equal(ext.l3, 'FAIL');
  assert.deepEqual(ext.ev.mergeChanged.map(x => [x.cell, x.was, x.now]), [['C1', 'C1:C2', 'C1:C3'], ['C2', 'C1:C2', 'C1:C3']]);
  const moved = await run({ mutate: c => (c.a === 'C1' || c.a === 'C2') && ((c.merge = 'B1:B2'), true) });
  assert.equal(moved.l3, 'FAIL');
  const appeared = await run({ merges: [], mutate: c => c.a === 'E2' && ((c.merge = 'E2:E3'), true) });
  assert.equal(appeared.touched, 1);
  assert.equal(appeared.l3, 'FAIL');
  assert.deepEqual(appeared.ev.mergeChanged, [{ cell: 'E2', was: 'E2', now: 'E2:E3' }]);
});

test('POSITIVE control (unchanged): a ring cell\'s FILL changing still fails L3 through styleChanged; mergeChanged stays empty', async () => {
  const r = await run({ mutate: c => (c.a === 'C2') && ((c.fill = 123456), true) });
  assert.equal(r.touched, 1);
  assert.equal(r.l3, 'FAIL');
  assert.deepEqual(r.ev.styleChanged.map(x => x.cell), ['C2']);
  assert.deepEqual(r.ev.mergeChanged, []);
});

test('an unread or missing cell after the write is not read as a merge change (no false alarm from absence), and an intact merge stays PASS', async () => {
  const r = await run({ mutate: c => c.a === 'C1' && ((delete c.merge), true) });
  assert.equal(r.l3 === 'FAIL', r.ev.mergeChanged.length > 0, 'a verdict FAIL here can only come from a named merge change, never from the absence of the field');
  assert.deepEqual(r.ev.mergeChanged, []);
  const intact = await run({ merges: ['C1:C2', 'F1:G1'] });
  assert.equal(intact.result.ok, true);
  assert.deepEqual(intact.ev.mergeChanged, []);
});
