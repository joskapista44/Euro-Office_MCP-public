'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

/**
 * LIVE finding (prod, release-l11-3 harness): filter_range { range: A1:B5 } with data DIRECTLY below the table (same
 * columns) -- the engine applied the AutoFilter to the whole current region (A1:B1877), filtering / hiding rows
 * OUTSIDE the requested range, and the product answered DISPATCHED_OUTCOME_UNKNOWN only after the write. With no
 * AutoFilter on the sheet yet, a non-blank cell in the ring around the requested range (the row below, the row above,
 * the column on either side, the corners) is now refused BEFORE the write, named.
 */
const J = x => JSON.stringify(x);
const TABLE = { A1: { v: 'id' }, B1: { v: 'tag' }, A2: { v: 1 }, B2: { v: 'x' }, A3: { v: 2 }, B3: { v: 'y' }, A4: { v: 3 }, B4: { v: 'x' } };
const book = extra => [{ name: 'F', cells: { ...TABLE, ...extra } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const flt = { intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:B4', field: 2, value: 'x' };
const run = async extra => { const h = harness({ store: createStore(book(extra)) }); const r = await h.tool.handler(req([flt])); return { r, h }; };
const REASON = 'xlsx-filter-range-not-the-current-region';

for (const [label, extra] of [['the row below (A5)', { A5: { v: 9 } }], ['the row below, second column (B5)', { B5: { v: 'z' } }],
  ['the column to the right (C2)', { C2: { v: 'side' } }], ['the corner below-right (C5)', { C5: { v: 1 } }]]) {
  test(`filter_range: a non-blank cell in the ring -- ${label} -- refuses before the write, named, the sheet untouched`, async () => {
    const { r, h } = await run(extra);
    assert.equal(r.ok, false, J(r).slice(0, 400));
    assert.match(J(r), new RegExp(REASON));
    assert.equal(h.calls().dispatch, 0);
    assert.equal(h.store.durable.find(x => x.name === 'F').autoFilter, undefined, 'no AutoFilter written');
  });
}

test('filter_range: data separated by a blank row / column is not adjacent -> APPLIED on exactly the requested range', async () => {
  const { r, h } = await run({ A6: { v: 'far' }, D2: { v: 'far' } });
  assert.equal(r.ok, true, J(r).slice(0, 400));
  assert.equal(h.store.durable.find(x => x.name === 'F').autoFilter.range, 'A1:B4');
});

test('LIVE harness verdict: GREEN passes; the pre-fix shape (dispatched, UNCERTAIN, a widened filter saved) fails on its own checks', { skip: needsPrivate('test/live/release-filter-adjacent-data-product.live.cjs') }, () => {
  const h = require('../live/release-filter-adjacent-data-product.live.cjs');
  const green = () => ({ SEEDS: [{ ok: true, ops: [{ status: 'APPLIED' }] }],
    ADJ: { ok: false, persistence: 'NOT_REQUIRED', ops: [{ status: 'FAILED', failure: { reason: h.REASON } }] },
    SEP: { ok: true, ops: [{ status: 'APPLIED' }] }, READ_ADJ: { ok: true, result: { range: null } }, READ_SEP: { ok: true, result: { range: 'A1:B5' } },
    CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO' });
  assert.equal(h.verdict(green()).PASS, true, J(h.verdict(green())));
  const red = green(); red.ADJ = { ok: false, persistence: 'UNCERTAIN_MAY_PERSIST', ops: [{ status: 'DISPATCHED_OUTCOME_UNKNOWN', failure: { reason: 'VERIFY_CANNOT_PROVE_TERMINAL_STATE' } }] }; red.READ_ADJ.result.range = 'A1:B6';
  assert.ok(h.verdict(red).failed.includes('adj_refused_named_pre_dispatch'));
  assert.ok(h.verdict(red).failed.includes('adj_no_filter_saved'));
  const blanket = green(); blanket.SEP = { ok: false, ops: [{ status: 'FAILED', failure: { reason: h.REASON } }] };
  assert.ok(h.verdict(blanket).failed.includes('sep_applied'), 'the guard must not refuse a separated table');
});
