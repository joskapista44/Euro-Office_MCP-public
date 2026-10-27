'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * #REF! (C), per writing family: a broken reference is no dependency, so a #REF! on ANOTHER sheet no longer refuses the
 * whole workbook (the pair -- each of these was refused before, xlsx-*-dependency-observation-incomplete:
 * broken-reference); a #REF! hosted on the sheet the write changes is refused by the ONE named guard
 * (xlsx-ref-on-target-unmeasured), 0 dispatch.
 */
const J = x => JSON.stringify(x);
const rows = () => { const c = {}; for (let i = 1; i <= 8; i += 1) c[`A${i}`] = { v: `r${i}` }; return c; };
const BROKEN = { v: '#REF!', f: '= #REF!*2' };
const book = brokenOn => [
  { name: 'Data', cells: { ...rows(), ...(brokenOn === 'Data' ? { F9: BROKEN } : {}) } },
  { name: 'Other', cells: { A1: { v: 'keep' }, ...(brokenOn === 'Other' ? { F9: BROKEN } : {}) } },
];
const CASES = [
  ['delete_rows', { intent: 'delete_rows', id: 's', sheet: 'Data', at: 2, count: 1, anchor: { column: 'A', value: 'r3' } }],
  ['move_range', { intent: 'move_range', id: 's', sheet: 'Data', range: 'A1:A3', target: 'D1' }],
  ['rename_sheet', { intent: 'rename_sheet', id: 's', sheet: 'Data', name: 'Data2' }],
  ['set_defined_name (sheet scope)', { intent: 'set_defined_name', id: 's', name: 'Rate', refersTo: '=Data!$A$1', scope: 'Data' }],
];
const runOn = async (brokenOn, step) => {
  const h = harness({ store: createStore(book(brokenOn)) });
  const r = await h.tool.handler(req([step]));
  return { r, dispatch: h.calls().dispatch };
};

for (const [label, step] of CASES) {
  test(`#REF! (C) ${label}: a broken formula on ANOTHER sheet does not block -> APPLIED`, async () => {
    const { r } = await runOn('Other', step);
    assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED'], J(r));
  });
  test(`#REF! (C) ${label}: a broken formula ON the changed sheet -> refused, xlsx-ref-on-target-unmeasured, 0 dispatch`, async () => {
    const { r, dispatch } = await runOn('Data', step);
    assert.equal(dispatch, 0, J(r));
    assert.match(J(r), /xlsx-ref-on-target-unmeasured/);
    assert.doesNotMatch(J(r), /dependency-observation-incomplete/);
  });
}
