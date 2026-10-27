'use strict';

/**
 * GitHub #93 (owner request via the owner channel, alice's NAD F80:F84 centering case): format_range
 * format.align { horizontal?, vertical? } on the PRODUCT WRITE surface -- the canonical PUBLIC
 * SetAlignHorizontal/SetAlignVertical writes, read back on the public alignment reader (format.cjs, the
 * SAME range.appearance.read/v1 bridge capability border already uses, B04 cutover). Offline: the product
 * WRITE tool with the fake XLSX editor (test/fakes/write-wiring.cjs).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const al = (id, range, align) => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { align } });

test("GH #93: align.horizontal 'center' (alice's NAD F80:F84 case) is APPLIED; a retry is SATISFIED with 0 dispatch", async () => {
  const h = harness({ store: createStore(book()) });
  const step = al('a', 'A2:A5', { horizontal: 'center' });
  const r = await h.tool.handler(req([step]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([step]));
  assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
});

for (const horizontal of ['left', 'center', 'right', 'justify']) {
  test(`GH #93: align.horizontal '${horizontal}' is APPLIED; a retry is SATISFIED with 0 dispatch`, async () => {
    const h = harness({ store: createStore(book()) });
    const step = al('a', 'A2', { horizontal });
    const r = await h.tool.handler(req([step]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([step]));
    assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
  });
}

// format.cjs's own documented limit: vertical 'bottom' equals the untouched default -- a bottom request on an
// untouched cell is SATISFIED at once (0 write), unlike every other vertical value.
test("GH #93: align.vertical 'bottom' on an untouched cell is SATISFIED at once (the documented default-equals-bottom limit)", async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([al('a', 'A2', { vertical: 'bottom' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch, 0);
});

for (const vertical of ['center', 'distributed', 'justify', 'top']) {
  test(`GH #93: align.vertical '${vertical}' is APPLIED; a retry is SATISFIED with 0 dispatch`, async () => {
    const h = harness({ store: createStore(book()) });
    const step = al('a', 'A2', { vertical });
    const r = await h.tool.handler(req([step]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([step]));
    assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
  });
}

test('GH #93: horizontal + vertical together in ONE align object is APPLIED (not the one-key-per-step rule -- that applies to format, not to align)', async () => {
  const h = harness({ store: createStore(book()) });
  const step = al('a', 'A2', { horizontal: 'center', vertical: 'top' });
  const r = await h.tool.handler(req([step]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
});

test('GH #93: an empty align object, an unsupported align key, align + another format key, and a bad enum value are refused before dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  for (const format of [{ align: {} }, { align: { direction: 'rtl' } }, { align: { horizontal: 'center' }, numberFormat: '0%' }, { align: { horizontal: 'diagonal' } }]) {
    const r = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format }]));
    assert.equal(r.ok, false, JSON.stringify(format));
    assert.equal(h.calls().dispatch, 0);
  }
});

test('GH #93: align over a PART of a merged area is refused, named (xlsx-format-merge-crosses-range-refused) -- not silently partial (X-B10: the whole merge is allowed)', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([al('m', 'C1', { horizontal: 'center' })]));
  assert.equal(r.ok, false);
  const f = r.operations.find(o => o.id === 'm').failure;
  assert.equal(f.reason, 'xlsx-format-merge-crosses-range-refused', JSON.stringify(f));
  assert.equal(f.mergeArea, 'C1:D1');
});
