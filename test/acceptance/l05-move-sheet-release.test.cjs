'use strict';

/**
 * L-05/2 (owner decision L-05): move_sheet on the PRODUCT surface, on top of the released delete_sheet. The LIVE measurement on the prod DS (DS /12,
 * docs/evidence/l05-sheets-fail-closed-20261015) found a move WRITABLE three ways (before / after a reference, a formula naming the moved sheet keeps its
 * text; retry SATISFIED / 0 commands) and REFUSED, named, 0 writes / 0 commands: a sheet relative to itself (refused before dispatch) and an absent reference
 * sheet (xlsx-move-reference-not-found). Offline, the product WRITE tool with the fake XLSX editor; the LIVE proof is
 * test/live/release-l05-sheets-product.live.cjs (RED on the pre-release build, GREEN on the rehearsal build).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [
  { name: 'Alpha', cells: { A1: { v: 1 } } },
  { name: 'Beta', cells: { A1: { v: 2 } } },
  { name: 'Gamma', cells: { A1: { v: 3, f: '=Beta!A1+1' } } },
];
const mv = extra => ({ intent: 'move_sheet', id: 's', ...extra });
const names = h => h.store.durable.map(x => x.name);
const statuses = r => (r.operations || []).map(o => o.status);
const validate = step => W.validateWriteRequest(req([step]));

test('L-05/2: move_sheet is a released intent, schema + constraints wired, a sole step', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('move_sheet'));
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.ok(W.SOLE_STEP.includes('move_sheet'));
  assert.equal(validate(mv({ sheet: 'Gamma', before: 'Alpha' })).ok, true);
  assert.equal(validate(mv({ sheet: 'Alpha', after: 'Gamma' })).ok, true);
});

test('L-05/2: the schema + constraints refuse a malformed step before any dispatch', () => {
  const bad = [
    [mv({ before: 'Alpha' }), 'sheet missing'],
    [mv({ sheet: 'Gamma' }), 'neither before nor after'],
    [mv({ sheet: 'Gamma', before: 'Alpha', after: 'Beta' }), 'before AND after'],
    [mv({ sheet: 'Gamma', before: 'Gamma' }), 'relative to itself'],
    [mv({ sheet: 'Gamma', after: 'gamma' }), 'relative to itself, case-insensitively'],
    [mv({ sheet: 'Gamma', before: '' }), 'an empty reference name'],
    [mv({ sheet: 'Gamma', before: 'a/b' }), 'a reference name with a forbidden character'],
    [mv({ sheet: 'Gamma', before: 'Alpha', index: 0 }), 'an unknown field'],
    [mv({ sheet: 3, before: 'Alpha' }), 'a sheet name that is not a string'],
  ];
  for (const [s, why] of bad) assert.equal(validate(s).ok, false, why);
});

test('L-05/2: move before / after a reference -- APPLIED, the order changes, a formula naming the moved sheet keeps its text, a retry is SATISFIED with 0 dispatch', async () => {
  for (const [step, want] of [[mv({ sheet: 'Gamma', before: 'Alpha' }), ['Gamma', 'Alpha', 'Beta']], [mv({ sheet: 'Alpha', after: 'Beta' }), ['Beta', 'Alpha', 'Gamma']], [mv({ sheet: 'Alpha', after: 'Gamma' }), ['Beta', 'Gamma', 'Alpha']]]) {
    const h = harness({ store: createStore(book()) });
    const r = await h.tool.handler(req([step]));
    assert.equal(r.ok, true, `${J(step)}: ${J(r).slice(0, 500)}`);
    assert.deepEqual(statuses(r), ['APPLIED'], J(step));
    assert.deepEqual(names(h), want, J(step));
    assert.equal(h.store.durable.find(x => x.name === 'Gamma').cells.A1.f, '=Beta!A1+1', 'the formula text is kept');
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([step]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `${J(step)}: retry`);
  }
});

test('L-05/2 NAMED limit: a reference sheet that is not in the workbook is refused -- 0 dispatch, the order untouched, the family reason passes through', async () => {
  const h = harness({ store: createStore(book()) });
  const before = J(h.store.durable);
  const r = await h.tool.handler(req([mv({ sheet: 'Gamma', before: 'Nope' })]));
  assert.equal(r.ok, false);
  assert.ok(J(r).includes('xlsx-move-reference-not-found'), J(r).slice(0, 600));
  assert.equal(h.calls().dispatch, 0);
  assert.equal(J(h.store.durable), before);
});
