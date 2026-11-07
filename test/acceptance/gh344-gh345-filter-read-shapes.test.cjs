'use strict';

/**
 * GH #344 (colour filter) + #345 (empty value-list): read_filter against the getter shapes MEASURED LIVE
 * (docs/evidence/gh344-gh345-filter-probe-live-r2-20261031.json, in-session AND after a save + fresh reopen):
 *   cell-colour filter  GetOperator xlFilterCellColor, GetCriteria1 null, GetCriteria2 null   (colour only in the saved <dxf>)
 *   font-colour filter  GetOperator xlFilterFontColor, GetCriteria1 null, GetCriteria2 null
 *   empty value-list    in session: on true, xlFilterValues, GetCriteria1 []; saved as an EMPTY <filterColumn/>; after reopen on:false
 * No colour is ever invented, and `[]` is never called "the empty list": it is three engine states to the public getters.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { projectFilterRead } = require('../../src/adapters/xlsx/families/filter-read.cjs');

const SHEETS = [{ name: 'S', index: 0, used: 'A1:E12' }];
const read = filters => projectFilterRead({ sheets: SHEETS, filterReads: { S: { range: 'A1:E12', filters } } }, { sheet: 'S' });
const eng = (field, operator, criteria1, criteria2 = null, on = true) => ({ field, on, operator, criteria1, criteria2 });

test('#344: a MEASURED cell-colour / font-colour filter is named by kind, colour null, reason COLOR_NOT_PUBLIC -- not the misleading NO_WRITE_PATH', () => {
  const cell = read([eng(2, 'xlFilterCellColor', null)]).filters[0];
  assert.equal(cell.kind, 'cellColor');
  assert.equal(cell.color, null);
  assert.match(cell.reason, /^COLOR_NOT_PUBLIC:/);
  assert.doesNotMatch(cell.reason, /NO_WRITE_PATH/);
  assert.equal(cell.unsupported, true, 'still not a filter_range shape: existing consumers keep skipping it');
  assert.deepEqual(cell.engine, { operator: 'xlFilterCellColor', criteria1: null, criteria2: null });
  const font = read([eng(3, 'xlFilterFontColor', null)]).filters[0];
  assert.equal(font.kind, 'fontColor');
  assert.equal(font.color, null);
  assert.match(font.reason, /^COLOR_NOT_PUBLIC:/);
});

test('#344: only the MEASURED shape (criteria1 AND criteria2 empty) is the colour state; any other criteria under a colour operator stays the old named gap', () => {
  const odd = read([eng(1, 'xlFilterCellColor', 'red', 'blue')]).filters[0];
  assert.equal(odd.unsupported, true);
  assert.match(odd.reason, /^NO_WRITE_PATH:/, 'an unmeasured colour shape is not promoted');
  assert.equal(odd.kind, undefined);
  assert.equal(read([eng(1, 'xlFilterCellColor', null, null, false)]).filters[0].on, false, 'an inactive filter is still just off');
});

test('#345: [] under xlFilterValues is a NAMED ambiguous state listing the three engine states -- not unsupported/NO_WRITE_PATH, not "the empty list"', () => {
  const e = read([eng(4, 'xlFilterValues', [])]).filters[0];
  assert.deepEqual(e.values, []);
  assert.deepEqual(e.ambiguous, ['empty-list', 'blanks-only', 'dates-only']);
  assert.match(e.reason, /^AMBIGUOUS_EMPTY_LIST:/);
  assert.doesNotMatch(e.reason, /NO_WRITE_PATH/);
  assert.equal(e.unsupported, true, 'not a writable filter_range shape: existing consumers keep skipping it');
  assert.deepEqual(e.engine, { operator: 'xlFilterValues', criteria1: [], criteria2: null });
});

test('#345: the empty list is only the MEASURED operator (xlFilterValues, no second criterion); [] elsewhere keeps the old gap', () => {
  for (const f of [eng(1, 'xlOr', []), eng(1, '', []), eng(1, 'xlFilterValues', [], '=5')]) {
    const out = read([f]).filters[0];
    assert.equal(out.unsupported, true, JSON.stringify(f));
    assert.match(out.reason, /^NO_WRITE_PATH:/, JSON.stringify(f));
    assert.equal(out.ambiguous, undefined, JSON.stringify(f));
  }
});

test('#345: after a save + reopen the engine reads the persisted empty list as on:false (measured) -- the projection says off, as before', () => {
  assert.deepEqual(read([eng(4, null, null, null, false)]).filters[0], { field: 4, on: false });
});

test('the measured r2 read as a whole: the 3-value control is still a values list, the two colours and the empty list are the new named states', () => {
  const out = read([eng(5, 'xlFilterValues', ['a', 'b', 'c']), eng(2, 'xlFilterCellColor', null), eng(3, 'xlFilterFontColor', null), eng(4, 'xlFilterValues', [])]);
  assert.deepEqual(out.filters[0], { field: 5, values: ['a', 'b', 'c'] }, 'POZ control: unchanged');
  assert.deepEqual(out.filters.map(f => f.kind || (f.ambiguous ? 'ambiguous' : 'values')), ['values', 'cellColor', 'fontColor', 'ambiguous']);
  assert.deepEqual(out.engineShapes.map(s => s.criteria1), ['array', 'null', 'null', 'array'], 'engineShapes (diagnostic) unchanged');
  assert.equal(out.engineShapes[3].criteria1Length, 0);
});
