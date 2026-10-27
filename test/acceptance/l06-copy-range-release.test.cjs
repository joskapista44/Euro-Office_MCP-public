'use strict';

/**
 * L-06 (owner decision L-06): copy_range on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The LIVE measurement on the prod DS (DS /12,
 * docs/evidence/l06-range-copy-fail-closed-20261015) found copy_range WRITABLE for values, formulas and formats on the same sheet, and for values + formulas onto another
 * sheet (the format does not travel there: named notProven in the receipt). The family-level behaviour (the explicit formula write, the delayed-Copy fault) is in
 * l06-copy-range-explicit-write.test.cjs; this file is the product-tool level. The LIVE proof is test/live/release-l06-copy-range-product.live.cjs
 * (RED on the pre-release build, GREEN on the rehearsal build). move_range's OWN product release
 * (L-06/b, owner decision, zero 2026-10-15) is test/acceptance/l06b-move-range-release.test.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

const J = x => JSON.stringify(x);
const book = () => [
  { name: 'S', cells: { A1: { v: 'h1' }, B1: { v: 'h2', bold: true }, A2: { v: 1 }, B2: { v: 2.5, nf: '0.00' }, A3: { v: 101, f: '=A2+H1' }, H1: { v: 100 } } },
  { name: 'T', cells: {} },
];
const cp = (extra = {}) => ({ intent: 'copy_range', id: 's', sheet: 'S', range: 'A1:B3', target: 'D1', ...extra });
const statuses = r => (r.operations || []).map(o => o.status);
const validate = step => W.validateWriteRequest(req([step]));
const cells = (h, n) => h.store.durable.find(x => x.name === n).cells;

test('L-06: copy_range is a released intent, schema + constraints wired, a sole step', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('copy_range'));
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.ok(W.SOLE_STEP.includes('copy_range'));
  assert.equal(validate(cp()).ok, true);
  assert.equal(validate(cp({ targetSheet: 'T' })).ok, true);
});

test('L-06: the schema + constraints refuse a malformed step before any dispatch', () => {
  const bad = [
    [cp({ range: 'A1' + ':' }), 'a range that is not A1'],
    [cp({ target: 'D1:E2' }), 'the target is ONE cell'],
    [cp({ target: 'd1' }), 'a lower-case cell'],
    [cp({ target: 3 }), 'a target that is not a string'],
    [cp({ sheet: '' }), 'an empty sheet'],
    [cp({ sheet: 'a/b' }), 'a sheet name with a forbidden character'],
    [cp({ targetSheet: '' }), 'an empty targetSheet'],
    [cp({ expect: {} }), 'an unknown field (expect belongs to move_range)'],
    [{ intent: 'copy_range', id: 's', sheet: 'S', range: 'A1:B3' }, 'target missing'],
    [cp({ range: 'A1:ZZ9999' }), 'a block over the cell ceiling'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-06: same-sheet copy -- APPLIED with values, the translated formula and the FORMAT; a retry is SATISFIED with 0 dispatch (also under the editor\'s save defect)', async () => {
  for (const script of [{}, { fault: 'copy-saves-literal' }]) {
    const h = harness({ store: createStore(book()), script });
    const r = await h.tool.handler(req([cp()]));
    assert.equal(r.ok, true, J(r).slice(0, 500));
    assert.deepEqual(statuses(r), ['APPLIED']);
    assert.equal(r.operations[0].semanticsNotProven, undefined, 'a same-sheet copy carries the format: nothing is left unproven');
    const S = cells(h, 'S');
    assert.deepEqual([S.D1.v, S.D2.v, S.E1.bold, S.E2.nf], ['h1', 1, true, '0.00'], 'values and format arrived');
    assert.equal(S.D3.f, '=D2+K1', `${J(script)}: the SAVED formula is the planner translation`);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([cp()]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `${J(script)}: retry`);
  }
});

test('L-06: cross-sheet copy -- APPLIED with the typed values and the translated formula, NO format, the receipt NAMES the limit, a retry SATISFIED / 0 dispatch (also under the delayed-Copy defect)', async () => {
  for (const script of [{}, { fault: 'cross-copy-delayed' }]) {
    const h = harness({ store: createStore(book()), script });
    const r = await h.tool.handler(req([cp({ targetSheet: 'T', target: 'C2' })]));
    assert.equal(r.ok, true, J(r).slice(0, 500));
    assert.deepEqual(statuses(r), ['APPLIED']);
    assert.deepEqual(r.operations[0].semanticsNotProven, ['cross-sheet-copy-format-not-carried']);
    const T = cells(h, 'T');
    assert.deepEqual([T.C2.v, T.D2.v, T.C3.v, T.D3.v], ['h1', 'h2', 1, 2.5], 'a number stays a number');
    assert.equal(T.C4.f, '=C3+J2', `${J(script)}: the SAVED formula is the planner translation`);
    assert.equal(T.D2.bold, undefined, 'the format did not travel');
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([cp({ targetSheet: 'T', target: 'C2' })]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `${J(script)}: retry`);
  }
});

test('L-06 NAMED limits: an overlap, an occupied target, a merged cell, a formula that leaves the sheet are refused -- 0 dispatch, the document untouched, the family reason passes through', async () => {
  const merged = () => { const b = book(); b[0].merges = ['A1:B1']; return b; };
  const broken = () => [{ name: 'S', cells: { B2: { v: 1, f: '=A1' } } }, { name: 'T', cells: {} }];
  const cases = [
    ['an overlap on the same sheet', book, cp({ target: 'B2' }), /overlap/],
    ['an occupied target', () => { const b = book(); b[1].cells.C2 = { v: 'occ' }; return b; }, cp({ targetSheet: 'T', target: 'C2' }), /xlsx-copy-target-occupied/],
    ['a merged cell in the source', merged, cp(), /merged/],
    ['a formula whose translation leaves the sheet', broken, { intent: 'copy_range', id: 's', sheet: 'S', range: 'B2', target: 'D1' }, /xlsx-copy-formula-leaves-the-sheet/],
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

test('L-06: a cross-sheet copy whose source holds a value the reader cannot type (TRUE) is refused before any write, named', async () => {
  const h = harness({ store: createStore([{ name: 'S', cells: { A1: { v: 'TRUE' } } }, { name: 'T', cells: {} }]) });
  const r = await h.tool.handler(req([{ intent: 'copy_range', id: 's', sheet: 'S', range: 'A1', target: 'C2', targetSheet: 'T' }]));
  assert.equal(r.ok, false);
  assert.match(J(r), /xlsx-copy-cross-sheet-value-type-not-observable/);
  assert.equal(h.calls().dispatch, 0);
});

test('L-06 live harness verdict: a synthetic GREEN run passes; the pre-release shape (copy_range not allowlisted) and each broken gate fail', { skip: needsPrivate('test/live/release-l06-copy-range-product.live.cjs') }, () => {
  const h = require('../live/release-l06-copy-range-product.live.cjs');
  const ap = (np) => ({ ok: true, persistence: 'FORCESAVE_CONFIRMED', ops: [{ status: 'APPLIED', ...(np ? { notProven: np } : {}) }] });
  const sat = { ok: true, persistence: null, ops: [{ status: 'SATISFIED' }] };
  const readOf = (want, extra = {}) => { const o = { S: { A1: { v: 'h1' } } }; for (const [sh, cells] of Object.entries(want || {})) { o[sh] = { ...(o[sh] || {}) }; for (const [a, w] of Object.entries(cells)) o[sh][a] = { v: w.v, f: w.f || '', ...(w.bold !== undefined ? { bold: w.bold } : {}), ...(w.nf ? { nf: w.nf } : {}) }; } return { ...o, ...extra }; };
  const green = () => {
    const o = { SEEDS: { a: ap() }, PAIRS: {}, NEG_NAMED: {}, NEG_SHAPE: {}, READ: {}, CLEANUP: { gone: true }, FIXTURE_MUTATED: 'NO' };
    for (const p of h.POS) { o.PAIRS[p.key] = { first: ap(p.cross ? [h.NOT_CARRIED] : null), second: sat }; o.READ[p.key] = readOf(p.want); }
    for (const n of h.NEG_NAMED) { o.NEG_NAMED[n.key] = { ok: false, persistence: null, ops: [{ status: 'FAILED', failure: { reason: n.reason } }] }; o.READ[n.key] = readOf(n.keep); }
    for (const k of Object.keys(h.NEG_SHAPE)) o.NEG_SHAPE[k] = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [] };
    return o;
  };
  assert.equal(h.verdict(green()).PASS, true, J(h.verdict(green()).failed));
  assert.deepEqual(h.POS.filter(p => p.cross).map(p => p.key), ['XS_VAL', 'XS_FORM', 'XS_FMT']);
  const red = green(); red.PAIRS.CP_VAL = { first: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] }, second: { ok: false, code: 'REFUSED_NOT_ALLOWLISTED', persistence: null, ops: [] } };
  assert.ok(h.verdict(red).failed.includes('CP_VAL_applied'), 'the pre-release build is RED');
  const literal = green(); literal.READ.CP_FORM.S.G3 = { v: '101', f: '=A2+H1' };
  assert.ok(h.verdict(literal).failed.includes('CP_FORM_saved'), 'a literal copy (value 101, source text) is caught by the VALUE-discriminating seed');
  const silentFormat = green(); silentFormat.PAIRS.XS_FMT.first = ap();
  assert.ok(h.verdict(silentFormat).failed.includes('XS_FMT_format_limit_named'), 'the cross-sheet format limit must be NAMED in the receipt');
  const wrongNamed = green(); wrongNamed.PAIRS.CP_FMT.first = ap([h.NOT_CARRIED]);
  assert.ok(h.verdict(wrongNamed).failed.includes('CP_FMT_format_limit_absent'), 'a same-sheet copy carries the format: no limit to name');
  const boldLeaked = green(); boldLeaked.READ.XS_FMT.T.C3.bold = true;
  assert.ok(h.verdict(boldLeaked).failed.includes('XS_FMT_saved'), 'the documented limit: the bold does not travel across sheets');
  const modified = green(); modified.READ.NG_OCCUPIED.S.F1 = { v: 'changed', f: '' };
  assert.ok(h.verdict(modified).failed.includes('neg_NG_OCCUPIED_untouched'));
  const unnamed = green(); unnamed.NEG_NAMED.NG_BROKEN.ops[0].failure = { reason: 'other' };
  assert.ok(h.verdict(unnamed).failed.includes('neg_NG_BROKEN'));
  const dispatched = green(); dispatched.NEG_SHAPE.overlap = { ok: false, code: 'REFUSED_SHAPE', persistence: null, ops: [{ status: 'FAILED' }] };
  assert.ok(h.verdict(dispatched).failed.includes('neg_shape_overlap'));
});
