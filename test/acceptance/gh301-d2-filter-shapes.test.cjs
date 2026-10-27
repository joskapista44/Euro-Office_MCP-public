'use strict';

/**
 * GH #301 D2: filter_range's further criteria shapes, each LIVE-measured in GH #301 AF-02 (SetAutoFilter on 9.3.5,
 * session = reopen = stored) and read back by the public getters as below:
 *   values: 3..10 different plain values  SetAutoFilter(field, [..], 'xlFilterValues') -> operator xlFilterValues,
 *                                         criteria1 the ARRAY (the 2-value shape is unchanged: xlOr =a / =b)
 *   top: n (1..500)                        SetAutoFilter(field, n, 'xlTop10Items')        -> xlTop10Items, criteria1 n
 *   dynamic: 'aboveAverage' | 'belowAverage'  SetAutoFilter(field, 'xlFilterAboveAverage', 'xlFilterDynamic')
 *                                         -> xlFilterDynamic, criteria1 'xlFilterAboveAverage'
 * The cell-COLOUR filter is NOT added: GetCriteria1 reads it back null (AF-02), so a write could never be verified.
 * read_filter mirrors the same shapes (filter-read.cjs), so a filter the writer made reads back as its own criteria.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const filter = require('../../src/adapters/xlsx/families/filter.cjs');
const { projectFilterRead } = require('../../src/adapters/xlsx/families/filter-read.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'F', cells: {
  A1: { v: 'id' }, B1: { v: 'tag' },
  A2: { v: 1 }, B2: { v: 'w' }, A3: { v: 2 }, B3: { v: 'x' }, A4: { v: 3 }, B4: { v: 'y' }, A5: { v: 4 }, B5: { v: 'z' }, A6: { v: 5 }, B6: { v: 'w' },
} }];
const flt = extra => ({ intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:B6', ...extra });
const statuses = r => (r.operations || []).map(o => o.status);
const v = s => W.validateWriteRequest(req([s]));

test('D2 schema: values 3..10, top 1..500, dynamic above/below average accepted; exactly one criteria; colour and bad shapes refused', () => {
  assert.equal(v(flt({ field: 2, values: ['w', 'x', 'y'] })).ok, true);
  assert.equal(v(flt({ field: 2, values: ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j'] })).ok, true, 'ten values');
  assert.equal(v(flt({ field: 1, top: 3 })).ok, true);
  assert.equal(v(flt({ field: 1, dynamic: 'aboveAverage' })).ok, true);
  assert.equal(v(flt({ field: 1, dynamic: 'belowAverage' })).ok, true);
  assert.equal(v(flt({ field: 2, values: ['x', 'y'] })).ok, true, 'the two-value shape is unchanged');
  for (const [step, why] of [
    [flt({ field: 2, values: ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k'] }), 'eleven values'],
    [flt({ field: 2, values: ['w', 'x', 'w'] }), 'a repeated value'],
    [flt({ field: 1, top: 0 }), 'top 0'],
    [flt({ field: 1, top: 2.5 }), 'a fractional top'],
    [flt({ field: 1, dynamic: 'today' }), 'an unmeasured dynamic type'],
    [flt({ field: 1, top: 3, value: 'x' }), 'two criteria'],
    [flt({ field: 1, color: '#FF0000' }), 'a colour filter (not verifiable)'],
  ]) assert.equal(v(step).ok, false, why);
});

test('D2: the lowered SetAutoFilter calls are exactly the measured AF-02 shapes', () => {
  const body = step => filter.lower(filter.plan(flt(step), 0).operation).body;
  assert.match(body({ field: 2, values: ['w', 'x', 'y'] }), /SetAutoFilter\(2, \["w","x","y"\], 'xlFilterValues'\)/);
  assert.match(body({ field: 1, top: 3 }), /SetAutoFilter\(1, 3, 'xlTop10Items'\)/);
  assert.match(body({ field: 1, dynamic: 'aboveAverage' }), /SetAutoFilter\(1, 'xlFilterAboveAverage', 'xlFilterDynamic'\)/);
  assert.match(body({ field: 1, dynamic: 'belowAverage' }), /SetAutoFilter\(1, 'xlFilterBelowAverage', 'xlFilterDynamic'\)/);
});

for (const [name, step] of [
  ['three values', { field: 2, values: ['w', 'x', 'y'] }],
  ['top 3 items', { field: 1, top: 3 }],
  ['above average', { field: 1, dynamic: 'aboveAverage' }],
  ['below average', { field: 1, dynamic: 'belowAverage' }],
]) {
  test(`D2: ${name} -- APPLIED, then retry SATISFIED, 0 dispatch`, async () => {
    const h = harness({ store: createStore(book()) });
    const first = await h.tool.handler(req([flt(step)]));
    assert.equal(first.ok, true, J(first).slice(0, 600));
    assert.deepEqual(statuses(first), ['APPLIED']);
    const before = h.calls().dispatch;
    const again = await h.tool.handler(req([flt(step)]));
    assert.deepEqual(statuses(again), ['SATISFIED']);
    assert.equal(h.calls().dispatch, before, 'the retry dispatches nothing');
  });
}

test('D2: a fake that stores a DIFFERENT value list / top count / dynamic type fails the verify (the read-back is the proof)', async () => {
  for (const step of [{ field: 2, values: ['w', 'x', 'y'] }, { field: 1, top: 3 }, { field: 1, dynamic: 'aboveAverage' }]) {
    const h = harness({ store: createStore(book()), script: { fault: 'filter-wrong-value' } });
    const r = await h.tool.handler(req([flt(step)]));
    assert.ok(h.calls().dispatch >= 1, `${J(step)}: the write was dispatched -- this is the VERIFY failing, not a refusal`);
    assert.equal(r.ok, false, `${J(step)} must not pass with a wrong stored criterion`);
  }
});

test('D2: read_filter mirrors the new shapes as the writer would write them', () => {
  const inv = filters => ({ sheets: [{ name: 'F' }], filterReads: { F: { range: 'A1:B6', filters } } });
  const read = filters => projectFilterRead(inv(filters), { sheet: 'F' }).filters;
  assert.deepEqual(read([{ field: 2, on: true, operator: 'xlFilterValues', criteria1: ['w', 'x', 'y'], criteria2: null }]), [{ field: 2, values: ['w', 'x', 'y'] }]);
  assert.deepEqual(read([{ field: 1, on: true, operator: 'xlTop10Items', criteria1: 3, criteria2: null }]), [{ field: 1, top: 3 }]);
  assert.deepEqual(read([{ field: 1, on: true, operator: 'xlFilterDynamic', criteria1: 'xlFilterAboveAverage', criteria2: null }]), [{ field: 1, dynamic: 'aboveAverage' }]);
  // the colour filter stays a named gap (GetCriteria1 reads it back null)
  assert.equal(read([{ field: 1, on: true, operator: 'xlFilterCellColor', criteria1: null, criteria2: null }])[0].unsupported, true);
  // an unmeasured dynamic type stays a named gap too
  assert.equal(read([{ field: 1, on: true, operator: 'xlFilterDynamic', criteria1: 'xlFilterToday', criteria2: null }])[0].unsupported, true);
});
