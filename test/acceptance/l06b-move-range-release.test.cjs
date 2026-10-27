'use strict';

/**
 * L-06/b (owner decision, zero 2026-10-15): move_range on the PRODUCT surface (the WRITE tool with
 * the fake XLSX editor), REPLACING the owner decision 2026-09-30 caller-supplied-`expect` contract.
 * The caller NEVER supplies `expect` -- classify() builds it itself, from its OWN fresh read,
 * directly before dispatch, the SAME session the write uses (range.cjs's own classify()/
 * moveExpectation, no new executor-level mechanism). A blank source (any state that looks already-
 * moved, including a genuine retry) has nothing to build a fresh expectation FROM, and is refused
 * named (xlsx-move-already-applied-not-provable), 0/0, BEFORE even checking the target -- v1 never
 * attempts to verify a retry's target content, and never dispatches a second Cut.
 *
 * The family-level behaviour (reference re-pointing, the dependent guard, the lower()/verify
 * mechanics) is in xlsx-w5-range.test.cjs and xlsx-m7-e-move.test.cjs; this file is the
 * product-tool level, mirroring l06-copy-range-release.test.cjs's own structure. The LIVE proof is
 * test/live/release-l06b-move-range-product.live.cjs (RED on the pre-release build, GREEN on the
 * rehearsal build) -- a maintainer's own run, per the kiosztás.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

const J = x => JSON.stringify(x);
const book = () => [
  { name: 'S', cells: { A1: { v: 'h1' }, B1: { v: 'h2' }, A2: { v: 1 }, B2: { v: 2.5 }, A3: { v: 101, f: '=A2+H1' }, H1: { v: 100 } } },
  { name: 'T', cells: {} },
];
const mv = (extra = {}) => ({ intent: 'move_range', id: 's', sheet: 'S', range: 'A1:A3', target: 'D1', ...extra });
const statuses = r => (r.operations || []).map(o => o.status);
const validate = step => W.validateWriteRequest(req([step]));
const cells = (h, n) => h.store.durable.find(x => x.name === n).cells;

test('L-06/b: move_range is a released intent, schema + constraints wired, a sole step; the caller never supplies `expect`', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('move_range'));
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.ok(W.SOLE_STEP.includes('move_range'));
  assert.equal(validate(mv()).ok, true);
  assert.equal(validate(mv({ targetSheet: 'T' })).ok, true);
});

test('L-06/b: the schema + constraints refuse a malformed step before any dispatch', () => {
  const bad = [
    [mv({ range: 'A1' + ':' }), 'a range that is not A1'],
    [mv({ target: 'D1:E2' }), 'the target is ONE cell'],
    [mv({ target: 'd1' }), 'a lower-case cell'],
    [mv({ target: 3 }), 'a target that is not a string'],
    [mv({ sheet: '' }), 'an empty sheet'],
    [mv({ sheet: 'a/b' }), 'a sheet name with a forbidden character'],
    [mv({ targetSheet: '' }), 'an empty targetSheet'],
    [mv({ expect: { version: 1, fileId: 'x', source: {}, target: {}, cells: [] } }), 'expect is no longer accepted at all (L-06/b)'],
    [{ intent: 'move_range', id: 's', sheet: 'S', range: 'A1:A3' }, 'target missing'],
    [mv({ range: 'A1:ZZ9999' }), 'a block over the cell ceiling'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-06/b: same-sheet move -- APPLIED, the source blank, the target holds the moved content; the reopened retry is REFUSED named, 0/0 (never SATISFIED, never a second Cut)', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([mv()]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  const S = cells(h, 'S');
  assert.deepEqual([S.D1.v, S.A1, S.D3.f], ['h1', undefined, '=D2+H1'], 'A2 (inside the block) shifts to D2; H1 (outside it) stays H1');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([mv()]));
  assert.equal(again.ok, false, J(again));
  assert.match(J(again), /xlsx-move-already-applied-not-provable/);
  assert.equal(h.calls().dispatch - d0, 0, 'the retry never dispatches');
  // the saved state is untouched by the refused retry
  assert.deepEqual([cells(h, 'S').D1.v, cells(h, 'S').A1], ['h1', undefined]);
});

test('L-06/b: cross-sheet move -- APPLIED, the source blank, the OTHER sheet holds the moved content; the reopened retry is REFUSED named, 0/0', async () => {
  const h = harness({ store: createStore(book()) });
  // B1:B2 (plain values, no formula) -- a formula inside a block moved to ANOTHER sheet is its own,
  // separately-tested family concern (xlsx-m7-e-move.test.cjs's R7 "inner reference" cases); this
  // test is about the PRODUCT release mechanics, not the reference-shift edge cases.
  const r = await h.tool.handler(req([mv({ range: 'B1:B2', target: 'C2', targetSheet: 'T' })]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  const T = cells(h, 'T');
  assert.deepEqual([T.C2.v, T.C3.v], ['h2', 2.5]);
  assert.equal(cells(h, 'S').B1, undefined, 'the source is blank');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([mv({ range: 'B1:B2', target: 'C2', targetSheet: 'T' })]));
  assert.equal(again.ok, false);
  assert.match(J(again), /xlsx-move-already-applied-not-provable/);
  assert.equal(h.calls().dispatch - d0, 0);
});

test('L-06/b NAMED limits: an occupied target, a merged cell, a reference-shape-not-measured case are refused -- 0 dispatch, the document untouched', async () => {
  const merged = () => { const b = book(); b[0].merges = ['A1:B1']; return b; };
  const occupied = () => { const b = book(); b[1].cells.C2 = { v: 'occ' }; return b; };
  const referenced = () => { const b = book(); b[1].cells.B1 = { v: 3, f: '=SUM(S!A1:A4)' }; return b; };
  const cases = [
    ['an occupied target', occupied, mv({ targetSheet: 'T', target: 'C2' }), /xlsx-move-target-occupied/],
    ['a merged cell in the source', merged, mv(), /merged/],
    ['a formula that partly overlaps the moved block (not measured)', referenced, mv(), /xlsx-move-reference-shape-not-measured/],
  ];
  for (const [name, shape, step, re] of cases) {
    const h = harness({ store: createStore(shape()) });
    const before = J(h.store.durable);
    const r = await h.tool.handler(req([step]));
    assert.equal(r.ok, false, name);
    assert.match(J(r), re, `${name}: ${J(r).slice(0, 400)}`);
    assert.equal(h.calls().dispatch, 0, `${name}: 0 dispatch`);
    assert.equal(J(h.store.durable), before, `${name}: untouched`);
  }
});

test('L-06/b LIVE harness verdict: a synthetic GREEN run passes; the pre-release shape (move_range not allowlisted) and each broken gate fail', { skip: needsPrivate('test/live/release-l06b-move-range-product.live.cjs') }, () => {
  const h = require('../live/release-l06b-move-range-product.live.cjs');
  const ap = () => ({ ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED' }] });
  const refused = { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: 'xlsx-move-already-applied-not-provable' } }] };
  const readOf = want => { const o = {}; for (const [sh, cells] of Object.entries(want || {})) { o[sh] = {}; for (const [a, w] of Object.entries(cells)) o[sh][a] = { v: w.v, f: w.f || '' }; } return o; };
  const namedFail = reason => ({ ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason } }] });
  const green = () => {
    const o = { SEEDS: { a: ap() }, PAIRS: {}, NEG_NAMED: {}, NEG_SHAPE: {}, READ: {}, CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO' };
    for (const p of h.POS) { o.PAIRS[p.key] = { first: ap(), second: refused }; o.READ[p.key] = readOf(p.want); }
    for (const n of h.NEG_NAMED) { o.NEG_NAMED[n.key] = namedFail(n.reason); o.READ[n.key] = readOf(n.keep); }
    for (const k of Object.keys(h.NEG_SHAPE)) o.NEG_SHAPE[k] = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] };
    return o;
  };
  assert.equal(h.verdict(green()).PASS, true, J(h.verdict(green()).failed));
  const red = green(); red.PAIRS.MV_VAL = { first: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] } };
  assert.equal(h.verdict(red).PASS, false, 'the pre-release build is RED');
  assert.equal(h.verdict(red).outcome, 'DS_NOT_RELEASED');
  const stillSatisfied = green(); stillSatisfied.PAIRS.MV_VAL.second = { ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] };
  assert.ok(h.verdict(stillSatisfied).failed.includes('MV_VAL_retry_refused_not_satisfied'), 'a SATISFIED retry is now WRONG, not right (L-06/b)');
  const wrongTarget = green(); wrongTarget.READ.MV_VAL.S.E1 = { v: 'WRONG', f: '' };
  assert.ok(h.verdict(wrongTarget).failed.includes('MV_VAL_saved'));
  const sourceNotBlank = green(); sourceNotBlank.READ.MV_VAL.S.A1 = { v: 'h1', f: '' };
  assert.ok(h.verdict(sourceNotBlank).failed.includes('MV_VAL_saved'), 'the source must be left blank');
  const modified = green(); modified.READ.NG_OCCUPIED.S.F1 = { v: 'changed', f: '' };
  assert.ok(h.verdict(modified).failed.includes('neg_NG_OCCUPIED_untouched'));
  const unnamed = green(); unnamed.NEG_NAMED.NG_REF.ops[0].failure = { reason: 'other' };
  assert.ok(h.verdict(unnamed).failed.includes('neg_NG_REF'));
  const dispatched = green(); dispatched.NEG_SHAPE.overlap = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [{ status: 'FAILED' }] };
  assert.ok(h.verdict(dispatched).failed.includes('neg_shape_overlap'));
  const notCleanedUp = green(); notCleanedUp.CLEANUP = { gone: false };
  assert.ok(h.verdict(notCleanedUp).failed.includes('cleanup'));
});

test('L-06/b: xlsx-move-source-differs-from-expectation is now STRUCTURALLY UNREACHABLE (named, not silently assumed) -- op.expect and the source read are built from the SAME fresh observation inside one classify() call, so a caller can no longer dispatch a move against a stale, separately-captured expectation', async () => {
  // The OLD (owner decision 2026-09-30) caller-supplied-`expect` contract could race: the caller
  // read the source, built `expect`, then dispatched later -- if the source changed meanwhile,
  // classifyMove's srcOff check caught the mismatch. L-06/b removes that race by construction: there
  // is no separate caller-side read anymore, so there is nothing LEFT to go stale. A task simply
  // dispatches against whatever the source holds AT CLASSIFY TIME, every time.
  const h = harness({ store: createStore(book()) });
  h.store.durable.find(s => s.name === 'S').cells.A2 = { v: 'CHANGED' };
  const r = await h.tool.handler(req([mv({ sheet: 'S', range: 'A2', target: 'F1' })]));
  assert.equal(r.ok, true, J(r).slice(0, 300));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.equal(cells(h, 'S').F1.v, 'CHANGED', 'the CURRENT content moved, not a stale snapshot');
});
