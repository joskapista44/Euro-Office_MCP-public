'use strict';

/**
 * L-05/3 (owner decision, the owner owner channel 2026-10-15 "A"): copy_sheet on the PRODUCT surface -- the B03 TEMPORARY_RC_BOUNDARY is lifted. The write path is the DS Bridge WRITE
 * worksheet.copy/v1 (the B03 direct internals were removed, M11.5c.4/5), served on /8 .. /12; the LIVE measurement (docs/evidence/l05-sheets-fail-closed-20261015) found a sheet with
 * values + formulas copied to a new name WRITABLE on /12 (retry SATISFIED / 0) and a target name that exists and differs REFUSED, named. Offline, the product WRITE tool with the fake
 * XLSX editor; the LIVE proof is test/live/release-l05-sheets-product.live.cjs (RED on the pre-release build, GREEN on the rehearsal build).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [
  { name: 'Src', cells: { A1: { v: 'h1' }, B1: { v: 'h2' }, A2: { v: 1 }, B2: { v: 2 }, A3: { v: 3, f: '=B2+1' } } },
  { name: 'Other', cells: { A1: { v: 'o' } } },
];
// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): `rt(n)` keeps naming an actual, now-retired
// 9.3.4.60 revision (used below as an OFF-list example). `mk()`'s default runtime is the single ACCEPTED
// 9.3.5.0/16 identity.
const rt = n => `9.3.4|euro-office-bridge-v1/9.3.4.60/${n}`;
const RT935 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const mk = (runtime = RT935) => harness({ store: createStore(book()), script: { runtime } });
const cp = extra => ({ intent: 'copy_sheet', id: 's', sheet: 'Src', name: 'Cp', ...extra });
const statuses = r => (r.operations || []).map(o => o.status);
const validate = step => W.validateWriteRequest(req([step]));
const names = h => h.store.durable.map(x => x.name);

test('L-05/3: copy_sheet is a released intent (the B03 boundary lifted), schema + constraints wired, a sole step', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('copy_sheet'));
  assert.equal(Object.prototype.hasOwnProperty.call(W.TEMPORARY_BOUNDARY, 'copy_sheet'), false, 'no longer a boundary');
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.ok(W.SOLE_STEP.includes('copy_sheet'));
  assert.equal(validate(cp()).ok, true);
});

test('L-05/3: the schema + constraints refuse a malformed step before any dispatch', () => {
  const bad = [
    [cp({ name: 'src' }), 'the copy name equals the source name, case-insensitively'],
    [cp({ name: '' }), 'an empty copy name'],
    [cp({ name: 'a/b' }), 'a copy name with a forbidden character'],
    [cp({ sheet: '' }), 'an empty source'],
    [cp({ name: 3 }), 'a copy name that is not a string'],
    [{ intent: 'copy_sheet', id: 's', sheet: 'Src' }, 'name missing'],
    [cp({ before: 'Other' }), 'an unknown field (before belongs to move_sheet)'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-05/3: copy a sheet to a new name -- APPLIED, the copy equals the source (values + formula), the source and the others untouched, a retry is SATISFIED with 0 dispatch', async () => {
  const h = mk();
  const r = await h.tool.handler(req([cp()]));
  assert.equal(r.ok, true, J(r).slice(0, 600));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.ok(names(h).includes('Cp') && names(h).includes('Src') && names(h).includes('Other'));
  const copy = h.store.durable.find(x => x.name === 'Cp').cells;
  assert.deepEqual([copy.A1.v, copy.B2.v, copy.A3.f], ['h1', 2, '=B2+1']);
  assert.deepEqual(h.store.durable.find(x => x.name === 'Src').cells, book()[0].cells, 'the source is untouched');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([cp()]));
  assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0]);
});

test('L-05/3 NAMED limit: a target name that exists and DIFFERS is refused -- 0 dispatch, the document untouched, the family reason passes through', async () => {
  const h = mk();
  const before = J(h.store.durable);
  const r = await h.tool.handler(req([cp({ name: 'Other' })]));
  assert.equal(r.ok, false);
  assert.ok(J(r).includes('xlsx-copy-target-exists-and-differs'), J(r).slice(0, 500));
  assert.equal(h.calls().dispatch, 0);
  assert.equal(J(h.store.durable), before);
  const missing = mk();
  const nr = await missing.tool.handler(req([cp({ sheet: 'Nope' })]));
  assert.equal(nr.ok, false);
  assert.equal(missing.calls().dispatch, 0);
});

test('L-05/3: the Bridge WRITE runtime list is the gate -- the accepted 9.3.5.0/16 identity writes, EVERY historical 9.3.4.60 revision is refused BEFORE any dispatch, named (xlsx-copy-bridge-write-runtime-not-proven)', async () => {
  {
    const h = mk();
    const r = await h.tool.handler(req([cp()]));
    assert.equal(r.ok, true, `accepted: ${J(r).slice(0, 300)}`);
    assert.deepEqual(statuses(r), ['APPLIED'], 'accepted');
  }
  // DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): the ENTIRE historical 9.3.4.60 chain (/4..17,
  // progressively widened by M11.5c.4 / X-O07/3 / X-O08 / X-O05/5-B / GH139 B as listed in the prior revision
  // of this comment) is now off the accepted list -- the single accepted runtime is 9.3.5.0/16.
  for (const n of [4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]) {
    const h = mk(rt(n));
    const before = J(h.store.durable);
    const r = await h.tool.handler(req([cp()]));
    assert.equal(r.ok, false, `/${n}`);
    assert.ok(J(r).includes('xlsx-copy-bridge-write-runtime-not-proven'), `/${n}: ${J(r).slice(0, 300)}`);
    assert.equal(h.calls().dispatch, 0, `/${n}`);
    assert.equal(J(h.store.durable), before, `/${n}`);
  }
});
