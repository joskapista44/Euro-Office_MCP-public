'use strict';

/**
 * L-01 (the owner GO 2026-10-14, leltar L-01): format_range format.wrap on the PRODUCT surface. The family already planned,
 * wrote (PUBLIC SetWrap) and verified (PUBLIC GetWrapText) it (format.cjs, W4a); only the product schema held it back. Offline, the
 * product WRITE tool with the fake XLSX editor (test/fakes/write-wiring.cjs). The LIVE proof is the harness
 * test/live/release-l01-wrap-product.live.cjs (RED on the pre-release build, GREEN on the rehearsal build).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const wrap = (range, value, id = 'w') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { wrap: value } });
const durable = (h, a) => h.store.durable.find(x => x.name === 'Budget').cells[a];
const statuses = r => r.operations.map(o => o.status);

test('L-01: wrap is a released format key (and the only change in the release list)', () => {
  assert.deepEqual([...W.FORMAT_RANGE_RELEASED_KEYS], ['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']);
});

test('L-01: wrap true, then false, then true -- each APPLIED + read back through the public reader; a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  assert.notEqual(durable(h, 'A2').wrap, true);
  for (const value of [true, false, true]) {
    const r = await h.tool.handler(req([wrap('A2', value)]));
    assert.equal(r.ok, true, `${value}: ${JSON.stringify(r).slice(0, 600)}`);
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']], String(value));
    assert.equal(durable(h, 'A2').wrap, value, String(value));
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([wrap('A2', value)]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${value}`);
  }
});

test('L-01: the neighbouring format and the value survive a wrap change (numberFormat, bold, align, the value; the untouched neighbour)', async () => {
  const store = createStore(book());
  const h = harness({ store });
  for (const step of [
    { intent: 'format_range', id: 'nf', sheet: 'Budget', range: 'B2', format: { numberFormat: '0.00' } },
    { intent: 'format_range', id: 'b', sheet: 'Budget', range: 'B2', format: { font: { bold: true } } },
    { intent: 'format_range', id: 'al', sheet: 'Budget', range: 'B2', format: { align: { horizontal: 'center' } } },
  ]) assert.equal((await h.tool.handler(req([step]))).ok, true, step.id);
  const before = { ...durable(h, 'B2') };
  const neighbour = { ...durable(h, 'B3') };
  const r = await h.tool.handler(req([wrap('B2', true)]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  const after = durable(h, 'B2');
  assert.equal(after.wrap, true);
  assert.deepEqual([after.v, after.nf, after.bold, after.align], [before.v, before.nf, before.bold, before.align]);
  assert.deepEqual(durable(h, 'B3'), neighbour, 'an untouched neighbour cell is byte-equal');
  const off = await h.tool.handler(req([wrap('B2', false)]));
  assert.equal(off.ok, true);
  assert.deepEqual([durable(h, 'B2').wrap, durable(h, 'B2').nf, durable(h, 'B2').bold], [false, before.nf, before.bold]);
});

test('L-01: a wrap over a range changes every cell of it and nothing outside', async () => {
  const h = harness({ store: createStore(book()) });
  const outside = { ...durable(h, 'A5') };
  const r = await h.tool.handler(req([wrap('A2:A4', true)]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  for (const a of ['A2', 'A3', 'A4']) assert.equal(durable(h, a).wrap, true, a);
  assert.deepEqual(durable(h, 'A5'), outside);
});

test('L-01: a non-boolean wrap, a wrap next to another key and an empty object are refused BEFORE dispatch (0 dispatch)', async () => {
  const h = harness({ store: createStore(book()) });
  const bad = [
    ['string "true"', { wrap: 'true' }], ['number 1', { wrap: 1 }], ['null', { wrap: null }], ['object', { wrap: {} }], ['array', { wrap: [true] }],
    ['wrap + numberFormat', { wrap: true, numberFormat: '0.00' }], ['wrap + align', { wrap: true, align: { horizontal: 'left' } }],
  ];
  for (const [label, format] of bad) {
    const r = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A2', format }]));
    assert.deepEqual([r.ok, r.dispatched], [false, false], label);
    assert.equal(r.code, 'REFUSED_SHAPE', `${label}: ${r.reason}`);
    assert.match(r.reason, label.startsWith('wrap +') ? /exactly one of numberFormat \/ border \/ fill \/ font \/ align \/ wrap/ : /format\.wrap must be a boolean/, label);
  }
  assert.equal(h.calls().dispatch, 0);
});

test('L-01: an unreleased key beside wrap is still refused by name (the boundary is intact)', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A2', format: { wrap: true, textDirection: 'rtl' } }]));
  assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_TEMPORARY_BOUNDARY', false]);
  assert.match(r.reason, /format\.textDirection/);
  assert.match(r.reason, /format\.align and format\.wrap only/);
});

test('L-01 RED control: a write that does not take effect (fault "ignore") is NOT reported as APPLIED', async () => {
  const h = harness({ store: createStore(book()), script: { fault: 'ignore' } });
  const r = await h.tool.handler(req([wrap('A2', true)]));
  assert.notEqual(statuses(r)[0], 'APPLIED', JSON.stringify(r).slice(0, 500));
  assert.notEqual(durable(h, 'A2').wrap, true);
});

test('L-01 RED control: a wrap step that ALSO changes another axis (fault "format-extra-key": the number format) fails the preservation proof', async () => {
  const h = harness({ store: createStore(book()), script: { fault: 'format-extra-key' } });
  const r = await h.tool.handler(req([wrap('A2', true)]));
  assert.equal(r.ok, false, JSON.stringify(r).slice(0, 500));
  assert.notEqual(statuses(r)[0], 'APPLIED');
});

test('L-01 RED control: a wrap that leaks to a neighbour (fault "format-leaks") fails the proof', async () => {
  const h = harness({ store: createStore(book()), script: { fault: 'format-leaks', leakCell: 'A5' } });
  const r = await h.tool.handler(req([wrap('A2', true)]));
  assert.equal(r.ok, false, JSON.stringify(r).slice(0, 500));
  assert.notEqual(statuses(r)[0], 'APPLIED');
});

test('L-01 RED control: a write that only reaches the first cell of a range (fault "partial") is not APPLIED', async () => {
  const h = harness({ store: createStore(book()), script: { fault: 'partial' } });
  const r = await h.tool.handler(req([wrap('A2:A4', true)]));
  assert.notEqual(statuses(r)[0], 'APPLIED', JSON.stringify(r).slice(0, 500));
});
