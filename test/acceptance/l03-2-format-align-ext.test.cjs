'use strict';

/**
 * L-03/2 (owner GO 2026-10-14, inventory row L-03): format_range format.align.rotation / readingOrder / indent / shrink on the PRODUCT surface.
 * The LIVE measurement on the prod DS (/11) found rotation and readingOrder WRITABLE and indent / shrink FAIL_CLOSED -- on a stale runtime list
 * (the B05 Bridge WRITE list was /6 /7 /8 while the handler is byte-identical through /11; test/static/b05-write-handler-identity.test.cjs is
 * the warrant). Offline, the product WRITE tool with the fake XLSX editor; the LIVE proof is test/live/release-l03-align-product.live.cjs
 * (RED on the pre-release build, GREEN on the rehearsal build) and the L-03 measurement re-run (all four axes WRITABLE).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the historical 9.3.4.60 chain (including
// /11, the LIVE-measured prod runtime this file's tests were pinned on) is retired -- the single ACCEPTED
// runtime is now 9.3.5.0/16, so `RT(n)` collapses to that one identity regardless of `n`. `RT_RETIRED(n)`
// keeps naming an actual, now-retired 9.3.4.60 revision, for the OFF-list negative test below.
const RT = () => '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const RT_RETIRED = n => `9.3.4|euro-office-bridge-v1/9.3.4.60/${n}`;
const align = (range, a, id = 'a') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { align: a } });
const wrapStep = (range, v, id = 'w') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { wrap: v } });
const durable = (h, a) => h.store.durable.find(x => x.name === 'Budget').cells[a];
const statuses = r => r.operations.map(o => o.status);
const J = x => JSON.stringify(x);
const validate = a => W.validateWriteRequest(req([align('A2', a)]));
const H = { right: 0, left: 1, center: 2, justify: 3 };

test('L-03/2: the four axes are released align keys (the format key set is unchanged)', () => {
  for (const a of [{ rotation: 45 }, { readingOrder: 'rtl' }, { indent: 3, horizontal: 'left' }, { shrink: true }, { rotation: 255 }, { rotation: -90 }]) assert.equal(validate(a).ok, true, J(a));
  assert.deepEqual([...W.FORMAT_RANGE_RELEASED_KEYS], ['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']);
});

test('L-03/2: rotation -- the MEASURED values APPLIED + read back, a retry SATISFIED with 0 dispatch (public path: 0 Bridge writes)', async () => {
  const h = harness({ store: createStore(book()) });
  for (const rotation of [45, -90, 90, 255, 0]) {
    const r = await h.tool.handler(req([align('A2', { rotation })]));
    assert.equal(r.ok, true, `${rotation}: ${J(r).slice(0, 500)}`);
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']], String(rotation));
    assert.equal(durable(h, 'A2').rot, rotation, String(rotation));
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([align('A2', { rotation })]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${rotation}`);
  }
});

test('L-03/2: readingOrder context / ltr / rtl -- APPLIED + read back (0 / 1 / 2), retry SATISFIED, back to context', async () => {
  const h = harness({ store: createStore(book()) });
  for (const [name, code] of [['rtl', 2], ['ltr', 1], ['rtl', 2], ['context', 0]]) {
    const r = await h.tool.handler(req([align('A3', { readingOrder: name })]));
    assert.equal(r.ok, true, `${name}: ${J(r).slice(0, 500)}`);
    assert.equal(durable(h, 'A3').readingOrder, code, name);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([align('A3', { readingOrder: name })]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${name}`);
  }
});

const bridgeWrites = h => h.drivers.reduce((n, d) => n + d.calls.bridgeWrites, 0);
for (const runtime of [6, 7, 8, 9, 10, 11].map(RT)) {
  test(`L-03/2: indent + shrink are written through the Bridge on ${runtime.slice(runtime.lastIndexOf('/'))} -- APPLIED + read back, retry SATISFIED with 0 dispatch, one Bridge write per step`, async () => {
    const h = harness({ store: createStore(book()), script: { runtime } });
    const cases = [
      ['A4', { horizontal: 'left', indent: 3 }, c => c.indent === 3 && c.align.h === H.left],
      ['A5', { horizontal: 'right', indent: 15 }, c => c.indent === 15 && c.align.h === H.right],
      ['B2', { shrink: true }, c => c.shrink === true],
    ];
    for (const [range, a, holds] of cases) {
      const w0 = bridgeWrites(h);
      const r = await h.tool.handler(req([align(range, a)]));
      assert.equal(r.ok, true, `${J(a)}: ${J(r).slice(0, 500)}`);
      assert.deepEqual(statuses(r), ['APPLIED'], J(a));
      assert.equal(holds(durable(h, range)), true, `${range} read back ${J(durable(h, range))}`);
      assert.equal(bridgeWrites(h) - w0, 1, `${J(a)}: one Bridge write`);
      const d0 = h.calls().dispatch;
      const again = await h.tool.handler(req([align(range, a)]));
      assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${J(a)}`);
    }
  });
}

test('L-03/2: on a runtime OFF the B05 write list (the retired 9.3.4.60 chain, /5 and /17) indent and shrink still FAIL CLOSED -- 0 writes, 0 dispatch, the document untouched', async () => {
  // DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the ENTIRE historical 9.3.4.60 chain
  // (/5 through /17, widened progressively by X-O07/3 / X-O08 / X-O05/5-B / GH139 B as listed in the prior
  // revision of this comment) is now off the accepted list -- the single accepted runtime is 9.3.5.0/16.
  for (const runtime of [RT_RETIRED(5), RT_RETIRED(17)]) {
    const h = harness({ store: createStore(book()), script: { runtime } });
    const before = J(h.store.durable);
    for (const a of [{ horizontal: 'left', indent: 3 }, { shrink: true }]) {
      const r = await h.tool.handler(req([align('A4', a)]));
      assert.equal(r.ok, false, `${runtime.slice(-3)} ${J(a)}`);
      assert.match(J(r.operations || r), /xlsx-align-ext-bridge-write-runtime-not-proven/, `${runtime.slice(-3)} ${J(a)}`);
    }
    assert.equal(J(h.store.durable), before, 'DOCUMENT_MUTATION=0');
  }
});

test('L-03/2: the neighbouring axes survive (horizontal, vertical, wrap, value) -- rotation, readingOrder and shrink each change ONE axis', async () => {
  const h = harness({ store: createStore(book()), script: { runtime: RT(11) } });
  assert.equal((await h.tool.handler(req([align('B3', { horizontal: 'center', vertical: 'top' })]))).ok, true);
  assert.equal((await h.tool.handler(req([wrapStep('B3', true)]))).ok, true);
  const before = { ...durable(h, 'B3'), align: { ...durable(h, 'B3').align } };
  const neighbour = { ...durable(h, 'B4') };
  for (const a of [{ rotation: 45 }, { readingOrder: 'rtl' }, { shrink: true }]) {
    const r = await h.tool.handler(req([align('B3', a)]));
    assert.equal(r.ok, true, `${J(a)}: ${J(r).slice(0, 500)}`);
    const now = durable(h, 'B3');
    assert.deepEqual([now.align.h, now.align.v, now.wrap, now.v], [before.align.h, before.align.v, before.wrap, before.v], J(a));
  }
  assert.deepEqual(durable(h, 'B4'), neighbour);
});

test('L-03/2: out-of-domain values are refused BEFORE dispatch (REFUSED_SHAPE, 0 dispatch) -- the engine would reset a rotation to 0 silently', async () => {
  const h = harness({ store: createStore(book()), script: { runtime: RT(11) } });
  for (const a of [{ rotation: 91 }, { rotation: -91 }, { rotation: 180 }, { rotation: 254 }, { rotation: 256 }, { rotation: 1.5 }, { rotation: '45' }, { rotation: null },
    { readingOrder: 'RTL' }, { readingOrder: 2 }, { readingOrder: '' }, { indent: 16, horizontal: 'left' }, { indent: -1, horizontal: 'left' }, { indent: 1.5, horizontal: 'left' },
    { indent: '3', horizontal: 'left' }, { shrink: 'yes' }, { shrink: 1 }, { shrink: null }]) {
    const r = await h.tool.handler(req([align('A2', a)]));
    assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_SHAPE', false], J(a));
  }
  assert.equal(h.calls().dispatch, 0);
});

test('L-03/2: indent WITHOUT horizontal left / right in the same step is refused by name (the engine silently zeroes it) -- 0 dispatch', async () => {
  const h = harness({ store: createStore(book()), script: { runtime: RT(11) } });
  for (const a of [{ indent: 2 }, { indent: 2, horizontal: 'center' }, { indent: 2, horizontal: 'justify' }, { indent: 2, vertical: 'top' }]) {
    const r = await h.tool.handler(req([align('A2', a)]));
    assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_SHAPE', false], J(a));
    assert.match(r.reason, /align\.indent requires align\.horizontal 'left' or 'right'/, J(a));
  }
  assert.equal(h.calls().dispatch, 0);
});

test('L-03/2 RED controls: a write that does not take effect / only reaches the first cell / leaks / changes the number format too is NOT APPLIED', async () => {
  for (const [fault, extra, range] of [['ignore', {}, 'A2'], ['partial', {}, 'A2:A4'], ['format-leaks', { leakCell: 'A5' }, 'A2'], ['format-extra-key', {}, 'A2']]) {
    for (const a of [{ rotation: 45 }, { readingOrder: 'rtl' }]) {
      const h = harness({ store: createStore(book()), script: { fault, runtime: RT(11), ...extra } });
      const r = await h.tool.handler(req([align(range, a)]));
      assert.notEqual(statuses(r)[0], 'APPLIED', `${fault} ${J(a)}: ${J(r).slice(0, 300)}`);
    }
  }
});
