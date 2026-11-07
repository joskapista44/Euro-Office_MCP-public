'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const DC = require('../../src/adapters/xlsx/families/delete-cells.cjs');

/**
 * Owner decision A (2026-11-02, owner channel): delete_cells carries the same required second witness as delete_rows /
 * delete_columns (decision B, docs/DELETE-ROWS-COLUMNS-SECOND-WITNESS.md). The anchor names the first SURVIVING cell
 * (just past the block, in a lane within the band); the witness names the first REMOVED cell (the block's first row for
 * 'up', first column for 'left'). The defect: an anchor given as the first REMOVED cell's value already sat at the
 * block start, read as "already done" -> SATISFIED, 0 cells deleted, reported as success. Unit level (plan / classify /
 * verifyPrimitive on hand-built observations), the l18-1-delete-cells-offline.test.cjs convention.
 */

const key = (sheet, addr) => `${sheet}\u0000${addr}`;
function obs(cells) {
  const c = {};
  for (const [addr, v] of Object.entries(cells)) c[key('P', addr)] = typeof v === 'object' ? v : { v, f: '' };
  return {
    format: 'xlsx', complete: true, truncated: false, sheets: [{ name: 'P' }], cells: c,
    dependencies: { complete: true, unobservable: [], refs: [], broken: [] },
    rules: { complete: true, unobservable: [], bySheet: {} }, objects: {}, filterReads: { P: { range: null } },
  };
}
// the block C2:D3, shift up: C2 'c2' / D2 'd2' are the first REMOVED row, C4 'c4' / D4 'd4' the first SURVIVING one
const BEFORE = { C1: 'c1', D1: 'd1', C2: 'c2', D2: 'd2', C3: 'c3', D3: 'd3', C4: 'c4', D4: 'd4', C5: 'c5', D5: 'd5', E2: 'e2' };
// the same sheet after the up-delete of C2:D3: rows 4-5 of the band moved to rows 2-3
const AFTER = { C1: 'c1', D1: 'd1', C2: 'c4', D2: 'd4', C3: 'c5', D3: 'd5', E2: 'e2' };
const step = extra => ({ intent: 'delete_cells', id: 's', sheet: 'P', range: 'C2:D3', shift: 'up', ...extra });
const plan = extra => DC.plan(step(extra), 0);
const op = extra => { const p = plan(extra); assert.equal(p.ok, true, p.detail); return p.operation; };
const classify = (cells, extra) => DC.classify(obs(cells), op(extra));

test('the reported case, RED before decision A: the anchor given as the first REMOVED cell (c2) is refused, not SATISFIED', () => {
  const r = classify(BEFORE, { anchor: { column: 'C', value: 'c2' }, witness: { column: 'D', value: 'd2' } });
  assert.notEqual(r.state, 'SATISFIED');
  assert.equal(r.state, 'FAIL_CLOSED');
  assert.equal(r.evidence.reason, 'xlsx-delete-cells-witnesses-disagree');
});

test('both witnesses where the step says -> APPLY', () => {
  const r = classify(BEFORE, { anchor: { column: 'C', value: 'c4' }, witness: { column: 'D', value: 'd2' } });
  assert.equal(r.state, 'APPLY', JSON.stringify(r));
  assert.deepEqual(r.evidence.witnessAt, [2]);
});

test('the after-shape of a genuinely executed delete is refused as unprovable -- 0 writes, nothing deleted twice', () => {
  const r = classify(AFTER, { anchor: { column: 'C', value: 'c4' }, witness: { column: 'D', value: 'd2' } });
  assert.equal(r.state, 'FAIL_CLOSED');
  assert.equal(r.evidence.reason, 'xlsx-delete-cells-prior-execution-unprovable');
  assert.equal(r.evidence.readsAsAlreadyApplied, true);
  assert.equal(r.evidence.written, 0);
});

test('the after-shape with a witness that never existed is not success either (the shape the old rule accepted)', () => {
  const r = classify(BEFORE, { anchor: { column: 'C', value: 'c2' }, witness: { column: 'D', value: 'never-there' } });
  assert.equal(r.evidence.reason, 'xlsx-delete-cells-prior-execution-unprovable');
});

test('pre-dispatch refusals: the witness one row off, the witness twice, the anchor at neither position', () => {
  assert.equal(classify(BEFORE, { anchor: { column: 'C', value: 'c4' }, witness: { column: 'D', value: 'd3' } }).evidence.reason, 'xlsx-delete-cells-witnesses-disagree');
  assert.equal(classify({ ...BEFORE, D5: 'd2' }, { anchor: { column: 'C', value: 'c4' }, witness: { column: 'D', value: 'd2' } }).evidence.reason, 'xlsx-delete-cells-witness-not-unique');
  assert.equal(classify(BEFORE, { anchor: { column: 'C', value: 'c5' }, witness: { column: 'D', value: 'd2' } }).evidence.reason, 'xlsx-delete-cells-anchor-state-conflict');
});

test('shift left follows the same contract: APPLY before, unprovable after', () => {
  // block C2:D3 shifted left: the first REMOVED column C, the first SURVIVING E (row 2 lane)
  const left = { anchor: { row: 2, value: 'e2' }, witness: { row: 3, value: 'c3' }, shift: 'left' };
  const before = { C2: 'c2', D2: 'd2', E2: 'e2', C3: 'c3', D3: 'd3', E3: 'e3' };
  const after = { C2: 'e2', C3: 'e3' };
  assert.equal(classify(before, left).state, 'APPLY');
  assert.equal(classify(after, left).evidence.reason, 'xlsx-delete-cells-prior-execution-unprovable');
});

test('plan / schema: the witness is required, in the band, distinct from the anchor, on the shift\'s own axis', () => {
  const refuse = (extra, re) => { const p = plan(extra); assert.equal(p.ok, false, JSON.stringify(extra)); assert.match(p.detail, re); };
  refuse({ anchor: { column: 'C', value: 'c4' } }, /witness must be \{ column, value \}.*FIRST cell the delete removes/);
  refuse({ anchor: { column: 'C', value: 'c4' }, witness: { column: 'E', value: 'e2' } }, /witness.column must be within the block's own column band/);
  refuse({ anchor: { column: 'C', value: 'c4' }, witness: { column: 'C', value: 'c4' } }, /same value in the same column/);
  refuse({ anchor: { column: 'C', value: 'c4' }, witness: { row: 2, value: 'c2' } }, /witness must be \{ column, value \}/);
  refuse({ anchor: { column: 'C', value: 'c4' }, witness: { column: 'D', value: '' } }, /witness.value must be/);
  assert.deepEqual(W.STEP_SCHEMAS.delete_cells.required.slice().sort(), ['anchor', 'intent', 'range', 'sheet', 'shift', 'witness']);
  assert.equal(W.validateWriteRequest({ fileId: '123', task: { id: 't', steps: [step({ anchor: { column: 'C', value: 'c4' } })] } }).ok, false, 'the old anchor-only call is refused by the product schema');
});

test('L1: the witness must be gone after the write -- a delete that left it in its lane does not verify', () => {
  const o = op({ anchor: { column: 'C', value: 'c4' }, witness: { column: 'D', value: 'd2' } });
  const good = DC.verifyPrimitive(o, { primary: obs(BEFORE) }, { primary: obs(AFTER) });
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  assert.equal(good.evidence.witnessGone, true);
  const bad = DC.verifyPrimitive(o, { primary: obs(BEFORE) }, { primary: obs({ ...AFTER, D9: 'd2' }) });
  assert.equal(bad.ok, false);
  assert.equal(bad.evidence.witnessGone, false);
});
