'use strict';

/**
 * GH #301 D1: set_validation's four remaining one-value operators -- equal | notEqual | less | greaterEqual --
 * on the ONE axis (validation.cjs's ONE table, next to the already-measured greater | lessEqual), same types
 * and same `value` field greater/lessEqual already take. The family-level behaviour (the ONE table itself, the
 * lower() Add() call shape) is in test/acceptance/xlsx-m7-d-rules.test.cjs; this file is the product-tool level,
 * mirroring l22d-validation-date-release.test.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { FILE, harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'Data', cells: { A1: { v: 1 } }, objects: { charts: [], pivots: [], names: [] } }];
const base = { intent: 'set_validation', id: 's', sheet: 'Data', range: 'A1:A3', type: 'wholeNumber', alertStyle: 'stop' };
const step = (operator, value, over = {}) => ({ ...base, operator, value, ...over });
const statuses = r => (r.operations || []).map(o => o.status);

test('GH #301 D1: the schema operator enum gains equal, notEqual, less, greaterEqual', () => {
  const enumNow = W.STEP_SCHEMAS.set_validation.properties.operator.enum;
  for (const op of ['equal', 'notEqual', 'less', 'greaterEqual']) assert.ok(enumNow.includes(op), `${op} missing from the operator enum`);
  // the already-measured operators are still there -- this is an addition, not a replacement
  for (const op of ['between', 'notBetween', 'greater', 'lessEqual']) assert.ok(enumNow.includes(op), `${op} lost from the operator enum`);
});

test('GH #301 D1: the schema accepts the four new operators (value, not min/max) and still refuses an unknown one', () => {
  const v = s => W.validateWriteRequest(req([s]));
  for (const op of ['equal', 'notEqual', 'less', 'greaterEqual']) assert.equal(v(step(op, 7)).ok, true, op);
  assert.equal(v(step('equal', 7, { min: 1 })).ok, false, 'min/max do not belong with a one-value operator');
  assert.equal(v({ ...base, operator: 'xlNotAnOperator', value: 7 }).ok, false, 'an unknown operator string is refused by the schema itself');
});

for (const [op, value] of [['equal', 7], ['notEqual', 7], ['less', 7], ['greaterEqual', 7]]) {
  test(`GH #301 D1: ${op} -- APPLIED, then retry SATISFIED, 0 dispatch`, async () => {
    const h = harness({ store: createStore(book()) });
    const first = await h.tool.handler(req([step(op, value)]));
    assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
    const d0 = h.calls().dispatch;
    const retry = await h.tool.handler(req([step(op, value)]));
    assert.deepEqual(statuses(retry), ['SATISFIED']);
    assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
  });
}

test('GH #301 D1: read-symmetry -- read_validations reports the new operator verbatim (xlEqual), no translation gap', async () => {
  // write the rule, then open a SEPARATE read session on the SAME store -- read_validations is a plain
  // field-select over inv.validationReads (read-tools.cjs XLSX_PROJECTIONS.read_validations); no operator
  // allowlist of its own, so the new operator needs no read-side change -- this proves that, not re-derives it.
  const store = createStore(book());
  const w = harness({ store });
  const written = await w.tool.handler(req([step('equal', 7)]));
  assert.deepEqual(statuses(written), ['APPLIED'], J(written).slice(0, 400));

  const { createReadTool } = require('../../src/server/read-tools.cjs');
  const { createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
  const readTool = createReadTool('xlsx', {
    configSource: { kind: 'BROKER' },
    runtime: { check: () => ({ ok: true }) },
    createDriver: () => createFakeXlsxDriver(store, {}),
  });
  const r = await readTool.handler({ fileId: FILE, operation: 'read_validations', sheet: 'Data', range: 'A1:A3' });
  assert.equal(r.ok, true, J(r));
  assert.deepEqual(r.result.validation, { type: 'xlValidateWholeNumber', alertStyle: 'xlValidAlertStop', operator: 'xlEqual', formula1: '7', formula2: '' });
});
