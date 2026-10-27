'use strict';

/**
 * L-22c: set_validation type 'list' -- the real ceiling is the Excel-compatible combined list-source LITERAL
 * LENGTH (<=255 chars, the documented Excel "List Source" field limit), not the item count. The item-count
 * ceiling (64) that stood here before this unit was an unmeasured proxy: an earlier LIVE probe (an internal
 * tracker; the measurement cited in this release's own leadership comment) accepted 100 short items
 * (well over 255 combined characters) through the editor motor without error -- so item-count alone was never
 * the real wall. This file is the RED-first offline proof of the switch; the LIVE round that would confirm (or
 * correct) the EXACT 255-character edge on the installed motor is test/live/xlsx-l22c-list-limit-probe.live.cjs
 * (a PLAN, not run here -- "LIVE-ot NE futtass" per the card).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const validation = require('../../src/adapters/xlsx/families/validation.cjs');
const W = require('../../src/server/write-tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const step = (over = {}) => ({ intent: 'set_validation', id: 's', sheet: 'S', range: 'A1', type: 'list', alertStyle: 'stop', ...over });
const plan = s => validation.plan(s, 0);

test('L-22c: MAX_LIST_LITERAL_CHARS is 255 (the documented Excel "List Source" field limit), MAX_LIST_ITEMS is a loose 128-item upper bound, not the real ceiling', () => {
  assert.equal(validation.MAX_LIST_LITERAL_CHARS, 255);
  assert.equal(validation.MAX_LIST_ITEMS, 128);
});

test('RED (the shape this unit CHANGED): 65 short items -- refused on the OLD item-count-only rule, accepted now (combined literal is 129 chars, well under 255)', () => {
  const items = Array.from({ length: 65 }, () => 'a');
  const combined = items.join(',');
  assert.equal(combined.length, 129, 'sanity: a 65-item short list is nowhere near 255 chars');
  const r = plan(step({ items }));
  assert.equal(r.ok, true, J(r));
  assert.equal(r.operation.want.formula1, `"${combined}"`);
});

test('a single item of exactly 255 characters (the literal-length edge) is accepted', () => {
  const r = plan(step({ items: ['x'.repeat(255)] }));
  assert.equal(r.ok, true, J(r));
  assert.equal(r.operation.want.formula1, `"${'x'.repeat(255)}"`);
});

test('a single item of exactly 254 characters is accepted', () => {
  const r = plan(step({ items: ['x'.repeat(254)] }));
  assert.equal(r.ok, true, J(r));
});

test('a single item of 256 characters is refused, named, over the Excel list-source limit', () => {
  const r = plan(step({ items: ['x'.repeat(256)] }));
  assert.equal(r.ok, false);
  assert.match(r.detail, /combined list-source literal is 256 characters \(the Excel list-source limit is 255\)/);
});

test('128 one-char items (the MAX_LIST_ITEMS edge) combine to exactly 255 chars and are accepted', () => {
  const items = Array.from({ length: 128 }, () => 'x');
  const r = plan(step({ items }));
  assert.equal(r.ok, true, J(r));
  assert.equal(r.operation.want.formula1.length, 255 + 2, '+2 for the surrounding quotes');
});

test('129 one-char items are refused on the item-count ceiling BEFORE the literal-length check ever runs', () => {
  const items = Array.from({ length: 129 }, () => 'x');
  const r = plan(step({ items }));
  assert.equal(r.ok, false);
  assert.match(r.detail, /items must be 1\.\.128 non-empty strings/);
});

test('two items whose combined literal is 256 characters (each individually well under 255) is refused -- proves the check sums, not per-item', () => {
  const items = ['a'.repeat(200), 'b'.repeat(55)];
  assert.equal(items.join(',').length, 256);
  const r = plan(step({ items }));
  assert.equal(r.ok, false);
  assert.match(r.detail, /combined list-source literal is 256 characters/);
});

test('the empty-items / shape refusals (count, type, comma/quote content) are unchanged', () => {
  assert.equal(plan(step({ items: [] })).ok, false);
  assert.match(plan(step({ items: [] })).detail, /items must be 1\.\.128/);
  assert.match(plan(step({ items: ['a,b'] })).detail, /without a comma/);
  assert.match(plan(step({ items: ['"a'] })).detail, /without a comma/);
});

// ---- RED CONTROL: proves the checks above discriminate, not pass-by-construction ----

test('RED CONTROL: a 255-char item combined with a second item (256 total) is NOT silently accepted', () => {
  const r = plan(step({ items: ['x'.repeat(255), 'y'] }));
  assert.equal(r.ok, false, 'a fabricated pass would wrongly accept a 257-char literal');
});

// ---- PRODUCT SURFACE (write-tools.cjs schema + CONSTRAINTS + the family's plan(), end to end) ----

const setV = (over = {}) => ({ intent: 'set_validation', sheet: 'Budget', range: 'A1', type: 'list', alertStyle: 'stop', ...over });

test('L-22c: the schema maxItems (128) is the loose upper bound, never smaller than the family ceiling -- 65 short items pass the FULL product pipeline now', () => {
  const items = Array.from({ length: 65 }, () => 'a');
  const v = W.validateWriteRequest(req([setV({ items })]));
  assert.equal(v.ok, true, J(v));
});

test('a single item over the per-item schema cap (255 chars) is refused at the SCHEMA layer (REFUSED_SHAPE), before the family ever sees it', () => {
  const v = W.validateWriteRequest(req([setV({ items: ['x'.repeat(256)] })]));
  assert.equal(v.ok, false);
  assert.equal(v.code, 'REFUSED_SHAPE', J(v));
});

test('two items that each pass the schema (<=255 chars each) but combine to 256 are refused at the FAMILY layer (plan/REFUSED_SHAPE from validation.cjs), not the schema', () => {
  const items = ['a'.repeat(200), 'b'.repeat(55)];
  assert.equal(items.join(',').length, 256);
  const v = W.validateWriteRequest(req([setV({ items })]));
  assert.equal(v.ok, false, J(v));
  assert.match(v.reason || J(v), /combined list-source literal is 256 characters/);
});

test('APPLIED through the fake driver with 65 short items, retry SATISFIED 0 dispatch', async () => {
  const items = Array.from({ length: 65 }, () => 'a');
  const h = harness({ store: createStore([{ name: 'Budget', cells: { A1: { v: 'x' } }, objects: { charts: [], pivots: [], names: [] } }]) });
  const first = await h.tool.handler(req([setV({ items })]));
  assert.deepEqual((first.operations || []).map(o => o.status), ['APPLIED'], J(first).slice(0, 400));
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([setV({ items })]));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});
