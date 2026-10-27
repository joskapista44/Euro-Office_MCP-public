'use strict';

/**
 * GitHub #88 (owner: release the cell background color first): format_range format.fill '#RRGGBB' on the PRODUCT
 * WRITE surface -- the canonical PUBLIC SetFillColor write + the public fill read (X-F08). fill 'none' (owner
 * request: colour REMOVAL must also work; the ClearFormats + explicit-restore composite, format-fill-clear.cjs,
 * X-F08) is released too, same key, same schema pattern. One format key per step is unchanged. GH #90 (design
 * #ref, accepted #ref): fill on a range covering the WHOLE merge (C1:D1) is now APPLIED, every member
 * carries the colour; a range CROSSING the merge's boundary stays refused, named. Offline: the product WRITE
 * tool with the fake XLSX editor (test/fakes/write-wiring.cjs; book(): Budget, C1:D1 merged, every other cell
 * border-free so the fill-clear shared-edge refusal has nothing to trip on here).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const fill = (id, range, value) => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { fill: value } });

test('GH #88: fill #RRGGBB is APPLIED on the product surface, a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([fill('f', 'A1:B2', '#FFC000')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.id, o.status, o.l1]), [['f', 'APPLIED', 'PASS']]);
  const before = h.calls().dispatch;
  const again = await h.tool.handler(req([fill('f', 'A1:B2', '#FFC000')]));
  assert.equal(again.ok, true);
  assert.deepEqual(again.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch - before, 0);
});

test('GH #88: fill "none" clears the fill, every kept axis unaffected, retry SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const colored = await h.tool.handler(req([fill('f', 'A1:B2', '#FFC000')]));
  assert.equal(colored.ok, true, JSON.stringify(colored));
  const cleared = await h.tool.handler(req([fill('fn', 'A1:B2', 'none')]));
  assert.equal(cleared.ok, true, JSON.stringify(cleared));
  assert.deepEqual(cleared.operations.map(o => [o.id, o.status, o.l1]), [['fn', 'APPLIED', 'PASS']]);
  assert.equal(h.store.durable.find(x => x.name === 'Budget').cells.A1.fill, 'S:No Fill');
  const before = h.calls().dispatch;
  const again = await h.tool.handler(req([fill('fn', 'A1:B2', 'none')]));
  assert.equal(again.ok, true);
  assert.deepEqual(again.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch - before, 0);
});

test('GH #88: a malformed colour/keyword, fill + border, and fill "none" + another key are refused before dispatch', async () => {
  for (const [format, rx] of [[{ fill: 'red' }, /pattern|fill/], [{ fill: 'None' }, /pattern|fill/],
    [{ fill: '#FFC000', border: { mode: 'cells', sides: ['top'], style: 'thin', color: '#000000' } }, /exactly one of numberFormat \/ border \/ fill/],
    [{ fill: 'none', border: { mode: 'cells', sides: ['top'], style: 'thin', color: '#000000' } }, /exactly one of numberFormat \/ border \/ fill/]]) {
    const h = harness({ store: createStore(book()) });
    const r = await h.tool.handler(req([{ intent: 'format_range', id: 'x', sheet: 'Budget', range: 'A1', format }]));
    assert.equal(r.ok, false, JSON.stringify(format));
    assert.match(JSON.stringify(r), rx, JSON.stringify(format));
    assert.equal(h.calls().dispatch, 0);
  }
});

// GH #90 (design #ref, accepted #ref): fill is merge-aware -- a range covering the WHOLE merge is APPLIED,
// every member carries the colour; a range that CROSSES the merge's boundary stays refused, named.
test('GH #90: fill over the WHOLE merged area is APPLIED, every member carries the colour, retry SATISFIED', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([fill('m', 'C1:D1', '#FFC000')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.id, o.status, o.l1]), [['m', 'APPLIED', 'PASS']]);
  const cells = h.store.durable.find(x => x.name === 'Budget').cells;
  assert.equal(cells.C1.fill, 0xFFC000);
  assert.equal(cells.D1.fill, 0xFFC000);
  const before = h.calls().dispatch;
  const again = await h.tool.handler(req([fill('m', 'C1:D1', '#FFC000')]));
  assert.equal(again.ok, true);
  assert.deepEqual(again.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch - before, 0);
});

test('GH #90: fill over a range that CROSSES the merge boundary is refused, named (xlsx-format-merge-crosses-range-refused)', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([fill('m', 'C1', '#FFC000')])); // C1 alone -- the merge is C1:D1
  assert.equal(r.ok, false);
  const f = r.operations.find(o => o.id === 'm').failure;
  assert.equal(f.reason, 'xlsx-format-merge-crosses-range-refused', JSON.stringify(f));
  assert.equal(f.mergeArea, 'C1:D1');
});
