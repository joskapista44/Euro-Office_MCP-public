'use strict';

/**
 * GH #87: a format_range refusal like xlsx-format-border-unobservable names the cell, but the classifier's own `why`
 * (what exactly was not observable on that cell: 'not-read' or the reader's UNOBSERVABLE:<reason>) never reached
 * the product answer -- the user saw the code and the cell, not the cause. `why` is a scalar on the evidence; it is
 * now carried on failure.why (classification and preflight stages), like cell / mergeArea / formula / detail.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const border = { intent: 'format_range', id: 'b', sheet: 'Budget', range: 'A1', format: { border: { mode: 'cells', sides: ['top'], style: 'thin', color: '#000000' } } };

test('GH #87: a border refusal on an unobservable cell carries failure.why to the product answer', async () => {
  const h = harness({ store: createStore(book()), script: { appearanceBuild: false } });
  const r = await h.tool.handler(req([border]));
  assert.equal(r.ok, false, JSON.stringify(r));
  const f = r.operations.find(o => o.id === 'b').failure;
  assert.equal(f.reason, 'xlsx-format-border-unobservable', JSON.stringify(f));
  assert.equal(f.cell, 'A1');
  assert.equal(typeof f.why, 'string', JSON.stringify(f));
  assert.notEqual(f.why, '');
});

test('GH #87: failureOf carries a scalar why and drops a non-scalar one (classification + preflight)', () => {
  for (const stage of ['classification', 'preflight']) {
    const ok = W.failureOf({ status: 'FAILED', evidence: { stage, evidence: { reason: 'x', cell: 'A1', why: 'UNOBSERVABLE:engine' } } }, { l1: null, l3: null });
    assert.equal(ok.why, 'UNOBSERVABLE:engine', stage);
    const obj = W.failureOf({ status: 'FAILED', evidence: { stage, evidence: { reason: 'x', cell: 'A1', why: { nested: 1 } } } }, { l1: null, l3: null });
    assert.equal('why' in obj, false, `${stage}: a non-scalar why is not carried`);
  }
});
