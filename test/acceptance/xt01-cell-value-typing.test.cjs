'use strict';

/**
 * X-T01 CELL_VALUE_TYPING_AND_PRESERVATION (owner: an explicit string typing contract, the native cell typing, no
 * apostrophe workaround, edge whitespace kept; LIVE measurement test/live/xlsx-xt01-cell-value-typing-probe.live.cjs).
 * Offline, the product WRITE tool with the fake XLSX editor (test/fakes/write-wiring.cjs), whose getters follow the
 * measured shapes (GetFormula returns a non-formula cell's value text).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const OWNER_VALUES = ['1.', '001', '1E10', '01/02', '=1+1', '+1', '-1', ' 1 ', '1 ', ' 1', '@x'];
const col = (vals, c = 'Z') => ({ intent: 'write_range', id: 'typed', sheet: 'Budget', range: `${c}1:${c}${vals.length}`, values: vals.map(v => [v]) });
const durable = (h, a) => h.store.durable.find(x => x.name === 'Budget').cells[a];

test('X-T01: { text } stores every owner-named value EXACTLY as Text (nf @), L1/L3 PASS; a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([col(OWNER_VALUES.map(text => ({ text })))]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  OWNER_VALUES.forEach((v, i) => {
    const c = durable(h, `Z${i + 1}`);
    assert.deepEqual([c.v, c.nf, c.f === undefined || c.f === null || c.f === ''], [v, '@', true], `Z${i + 1}`);
  });
  const before = h.calls().dispatch;
  const again = await h.tool.handler(req([col(OWNER_VALUES.map(text => ({ text })))]));
  assert.equal(again.ok, true, JSON.stringify(again));
  assert.deepEqual(again.operations.map(o => o.status), ['SATISFIED'], 'the =1+1 TEXT reads as a text literal, not a formula');
  assert.equal(h.calls().dispatch - before, 0);
});

test('X-T01: a Text "=1+1" can be overwritten by a later plain write (it is not a formula)', async () => {
  const h = harness({ store: createStore(book()) });
  assert.equal((await h.tool.handler(req([col([{ text: '=1+1' }])]))).ok, true);
  const r = await h.tool.handler(req([col(['replaced'])]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(durable(h, 'Z1').v, 'replaced');
});

test('X-T01: an untyped (plain) coercible string is refused before dispatch with the exact reason AND the typed alternative', async () => {
  const h = harness({ store: createStore(book()) });
  for (const [v, why] of [['1.', /looks like a number/], [' 1 ', /leading or trailing whitespace/], ['01/02', /looks like a date/], ['=1+1', /starts with = \+ - @/]]) {
    const r = await h.tool.handler(req([col([v])]));
    assert.deepEqual([r.ok, r.stage, r.code, r.dispatched], [false, 'PRE_DISPATCH', 'REFUSED_SHAPE', false], `${JSON.stringify(v)}: ${JSON.stringify(r)}`);
    assert.match(r.reason, why);
    assert.match(r.reason, /\{ text: '\.\.\.' \}/, 'the refusal names the explicit typing');
  }
  assert.equal(h.calls().dispatch, 0);
});

test('X-T01: the typed value shape is strict (only { text }, a non-empty string)', async () => {
  const h = harness({ store: createStore(book()) });
  for (const bad of [{ text: '' }, { text: 1 }, { text: 'a', extra: 1 }, { str: 'a' }]) {
    const r = await h.tool.handler(req([col([bad])]));
    assert.equal(r.ok, false, JSON.stringify(bad));
    assert.equal(r.dispatched, false);
  }
});
