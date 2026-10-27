'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const layout = require('../../src/adapters/xlsx/families/layout.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

/**
 * LIVE finding (prod, release-l11-3 harness): a layout_range on a small sheet, in a workbook where ANOTHER sheet's
 * layout read was truncated (maxCol + maxRow > MAX_LAYOUT_ITEMS), and the workbook itself below the read bound (FULL
 * tier), was DISPATCHED and ended DISPATCHED_OUTCOME_UNKNOWN / INTEGRAL_PRESERVATION_NOT_PROVEN: the L3 could not
 * prove that other sheet's layout unchanged, but only AFTER the write. The gate now refuses it BEFORE the write, named;
 * under BOUNDED the existing named notProven ('layout-on-unobserved-sheets') stays the contract.
 */
// BOUNDED is a property of the tiered read (inventory.cjs); here it is set on the normalized observation directly
const withBounded = (b, p) => (b ? { ...p, bounded: true } : p);
const step = { intent: 'layout_range', id: 's', sheet: 'Data', rows: '3', hidden: true };
const op = () => planOperations({ id: 't', steps: [step] }).operations[0];
const inv = ({ bounded = false, big = 'UNOBSERVABLE:layout-truncated' } = {}) => withBounded(bounded, normalizeInventory({
  sheets: [{ name: 'Data', index: 0, used: 'A1:B5' }, { name: 'Big', index: 1, used: 'A1:B19000' }],
  cells: [{ sheet: 'Data', a: 'A1', v: 'x', f: 'x', count: 0, merge: 'A1' }],
  truncated: false, unsupported: [],
  layout: {
    Data: { cols: { A: { width: 8.43, hidden: false }, B: { width: 8.43, hidden: false } }, rows: Object.fromEntries([1, 2, 3, 4, 5].map(r => [String(r), { height: 15, hidden: false }])) },
    Big: big,
  },
}));

test('FULL tier: another sheet\'s truncated layout refuses the layout write BEFORE dispatch, named, with the sheet', () => {
  const c = layout.classify(inv(), op());
  assert.equal(c.state, 'FAIL_CLOSED', JSON.stringify(c));
  assert.equal(c.evidence.reason, 'xlsx-layout-other-sheet-unobservable');
  assert.deepEqual(c.evidence.sheets, ['Big']);
});

test('the same write with every sheet\'s layout observed is not refused by this guard', () => {
  const p = inv({ big: { cols: { A: { width: 8.43, hidden: false } }, rows: { 1: { height: 15, hidden: false } } } });
  const c = layout.classify(p, op());
  assert.notEqual(c.evidence && c.evidence.reason, 'xlsx-layout-other-sheet-unobservable', JSON.stringify(c));
  assert.equal(c.state, 'APPLY', JSON.stringify(c));
});

test('BOUNDED tier: the other sheet\'s truncation stays the existing named notProven, not this pre-dispatch refusal', () => {
  const p = inv({ bounded: true });
  assert.equal(p.bounded, true, 'the fixture really is BOUNDED');
  const c = layout.classify(p, op());
  assert.notEqual(c.evidence && c.evidence.reason, 'xlsx-layout-other-sheet-unobservable', JSON.stringify(c));
});

test('LIVE harness verdict: GREEN passes; the pre-fix DISPATCHED_OUTCOME_UNKNOWN shape and a hidden row 3 fail on their own checks', { skip: needsPrivate('test/live/release-layout-other-sheet-truncated-product.live.cjs') }, () => {
  const h = require('../live/release-layout-other-sheet-truncated-product.live.cjs');
  const green = () => ({ SEEDS: [{ ok: true, ops: [{ status: 'APPLIED' }] }], BOUND: { tier: 'full' },
    CASE: { ok: false, persistence: 'NOT_REQUIRED', ops: [{ status: 'FAILED', failure: { reason: h.REASON } }] },
    READ_SMALL: { ok: true, rows: { 3: { hidden: false } } }, CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO' });
  assert.equal(h.verdict(green()).PASS, true, JSON.stringify(h.verdict(green())));
  const red = green(); red.CASE = { ok: false, persistence: 'UNCERTAIN_MAY_PERSIST', ops: [{ status: 'DISPATCHED_OUTCOME_UNKNOWN', failure: { reason: 'INTEGRAL_PRESERVATION_NOT_PROVEN' } }] };
  assert.ok(h.verdict(red).failed.includes('refused_named_pre_dispatch'));
  const hidden = green(); hidden.READ_SMALL.rows['3'].hidden = true;
  assert.ok(h.verdict(hidden).failed.includes('small_row3_not_hidden'));
  const bounded = green(); bounded.BOUND.tier = 'tiered';
  assert.ok(h.verdict(bounded).failed.includes('full_tier'));
});
