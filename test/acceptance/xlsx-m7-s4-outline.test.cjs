'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const PROBE = require('../../docs/evidence/m7-s4-probe2-20261003.json').SCENARIOS.BORDER;

/**
 * M7 S4 S01 (owner GO 2026-10-03, owner channel): format_range border mode 'outline'. The fake is CALIBRATED on the
 * LIVE probe (run m7-s4-probe2-20261003): its seed is the probe's measured OWN borders (R0, after a reopen), and after
 * the product's outline the fake's full read must equal the probe's measured full read after the reopen (R2), cell
 * by cell and side by side over the range and its one-cell ring -- for the 3x3, the 1-row and the 1-column shape.
 */

const SHAPES = { G3: 'C3:E5', G1R: 'C9:E9', G1C: 'C13:C15' };
const SPEC = { border: { mode: 'outline', style: 'thin', color: '#0000FF' } };
const SIDES = ['t', 'b', 'l', 'r'];
const norm = x => (x && x.s > 0 ? { s: x.s, c: x.c === null ? 0 : x.c } : { s: 0, c: 0 });

/** The probe's measured own borders (R0) as the fake's seed, every shape on one sheet. */
function seededBook() {
  const cells = {};
  for (const g of Object.keys(SHAPES)) {
    for (const [a, x] of Object.entries(PROBE.R0.CELL[g].cells)) {
      const own = x.own;
      const border = Object.fromEntries(SIDES.filter(k => own[k].s > 0).map(k => [k, { s: own[k].s, c: own[k].c }]));
      if (own.d.s > 0) Object.assign(border, { d: own.d, dd: own.dd, du: own.du });
      if (Object.keys(border).length) cells[a] = { v: null, border };
    }
  }
  cells.C3 = { ...(cells.C3 || {}), v: 'v' };
  cells.D4 = { ...(cells.D4 || {}), v: 4 };
  return [{ name: 'B', cells }, { name: 'Other', cells: { A1: { v: 'keep' } } }];
}

/** The fake's full read of one cell (own side, else the neighbour's facing side -- the measured shared edge). */
function fullRead(ws, a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  const col = m[1].charCodeAt(0) - 64;
  const row = Number(m[2]);
  const at = (c, r) => ws.cells[`${String.fromCharCode(64 + c)}${r}`];
  const own = (at(col, row) || {}).border || {};
  const nb = { t: [col, row - 1, 'b'], b: [col, row + 1, 't'], l: [col - 1, row, 'r'], r: [col + 1, row, 'l'] };
  const out = {};
  for (const k of SIDES) {
    if (own[k] && own[k].s > 0) { out[k] = norm(own[k]); continue; }
    const [c, r, f] = nb[k];
    const n = c >= 1 && r >= 1 ? ((at(c, r) || {}).border || {})[f] : null;
    out[k] = norm(n);
  }
  out.d = norm(own.d);
  return out;
}

const task = (range, spec = SPEC) => ({ id: 'm7-s4-outline', steps: [{ intent: 'format_range', sheet: 'B', range, format: spec }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const op = result => result.execution.appliedOperations[0];
const ws = store => store.durable.find(s => s.name === 'B');

for (const [g, range] of Object.entries(SHAPES)) {
  test(`${g} ${range}: APPLIED in one command; the fake's end state equals the LIVE probe's reopened full read, every side of the range + ring; retry SATISFIED 0/0`, async () => {
    const store = createStore(seededBook());
    const { result, m } = await run(store, task(range));
    assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.primitiveVerification, result.execution.preservationVerification]));
    assert.deepEqual([m.writes, m.commands], [1, 1]);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    const diff = [];
    for (const [a, x] of Object.entries(PROBE.R2.CELL[g].cells)) {
      const want = Object.fromEntries([...SIDES, 'd'].map(k => [k, norm(x.full[k])]));
      const got = fullRead(ws(store), a);
      if (JSON.stringify(got) !== JSON.stringify(want)) diff.push({ a, got, want });
    }
    assert.deepEqual(diff, [], 'the calibrated end state');
    const again = await run(store, task(range));
    assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
    assert.equal(again.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  });
}

test('selected outline sides only (top + left): the other outer edges and every inner edge unchanged', async () => {
  const store = createStore(seededBook());
  const before = { e5: fullRead(ws(store), 'E5'), b4: fullRead(ws(store), 'B4'), e3: fullRead(ws(store), 'E3') };
  const { result } = await run(store, task('C3:E5', { border: { mode: 'outline', sides: ['top', 'left'], style: 'thin', color: '#0000FF' } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification));
  assert.deepEqual(fullRead(ws(store), 'E3').t, { s: 7, c: 255 });
  assert.deepEqual(fullRead(ws(store), 'C5').l, { s: 7, c: 255 });
  assert.deepEqual(fullRead(ws(store), 'E5').b, before.e5.b, 'bottom not requested: unchanged');
  assert.deepEqual(fullRead(ws(store), 'E3').r, before.e3.r, 'right not requested: unchanged');
  assert.deepEqual(fullRead(ws(store), 'B4').r, { s: 7, c: 255 }, 'the left ring neighbour shows the shared edge');
});

test('the command batches the per-edge-cell SetBorders in ONE callCommand (edge cells only, outer sides only)', () => {
  const o = planOperations(task('C3:E5')).operations[0];
  const cmd = require('../../src/adapters/xlsx/families/format.cjs').lower(o);
  assert.deepEqual(cmd.meta.edges, [['C3', ['Top', 'Left']], ['D3', ['Top']], ['E3', ['Top', 'Right']], ['C4', ['Left']], ['E4', ['Right']], ['C5', ['Bottom', 'Left']], ['D5', ['Bottom']], ['E5', ['Bottom', 'Right']]]);
  assert.doesNotThrow(() => new Function(cmd.body));
  assert.equal((cmd.body.match(/SetBorders\(/g) || []).length, 1, 'one call site in a loop');
});

for (const [fault, extra] of [
  ['outline-miss-side', {}],
  ['outline-inner-side', {}],
  ['outline-ring-not-cleared', {}],
  ['outline-ring-other-side', { faultCell: 'D2' }],
  ['outline-inner-edge-lost', { faultCell: 'D4' }],
  ['outline-diagonal-lost', { faultCell: 'D5' }],
]) {
  test(`RED (mutant '${fault}'): the verification fails, the task is not ok`, async () => {
    const store = createStore(seededBook());
    const { result } = await run(store, task('C3:E5'), { fault, ...extra });
    assert.equal(result.ok, false);
    // an L1 miss leaves the operation unproven (not APPLIED); an L3 miss fails the preservation layer
    const l3 = (result.execution.preservationVerification || []).map(v => v.verdict);
    assert.ok(op(result).status !== OPERATION_STATUS.APPLIED || l3.includes('FAIL'), JSON.stringify([op(result).status, l3]));
  });
}

test('refused before the write, 0/0: a merged cell in the ring; an unverified border reader', async () => {
  const b = seededBook();
  b[0].merges = ['F4:G4'];
  b[0].cells.F4 = { v: 'm' };
  const merged = await run(createStore(b), task('C3:E5'));
  assert.deepEqual([merged.m.writes, merged.m.commands], [0, 0]);
  assert.equal(op(merged.result).evidence.evidence.reason, 'xlsx-format-outline-merged-ring-refused');
  const blind = await run(createStore(seededBook()), task('C3:E5'), { appearanceBuild: false });
  assert.deepEqual([blind.m.writes, blind.m.commands], [0, 0]);
  assert.equal(op(blind.result).evidence.evidence.reason, 'xlsx-format-border-unobservable');
});

test('plan: mode values, default sides, the measured styles only; removal still refused', () => {
  const ok = planOperations(task('C3:E5', { border: { mode: 'outline', style: 'medium', color: '#FF0000' } })).operations[0];
  assert.deepEqual(ok.outline.sides, ['t', 'b', 'l', 'r']);
  assert.equal(ok.footprint.style, 'B2:F6', 'the ring is in the footprint: its facing sides change');
  const refuse = (spec, re) => assert.match(planOperations(task('C3', spec)).detail || '', re);
  refuse({ border: { mode: 'bogus', style: 'thin', color: '#000000' } }, /border.mode must be/);
  refuse({ border: { mode: 'outline', style: 'triple', color: '#000000' } }, /the measured styles/);
  refuse({ border: { mode: 'outline', style: 'none', color: '#000000' } }, /border\.style 'none' is not a style/);
  refuse({ border: { mode: 'outline', sides: [], style: 'thin', color: '#000000' } }, /border.sides/);
});

test('not SATISFIED while a ring neighbour shows a DIFFERENT shared edge (its own side): APPLY, and the edge is taken over', async () => {
  const thin = { s: 7, c: 255 };
  const cells = {};
  for (const a of ['C3', 'D3', 'E3']) cells[a] = { v: null, border: { t: thin } };
  for (const a of ['C5', 'D5', 'E5']) cells[a] = { v: null, border: { ...(cells[a] || {}).border, b: thin } };
  for (const a of ['C3', 'C4', 'C5']) cells[a] = { v: null, border: { ...((cells[a] || {}).border || {}), l: thin } };
  for (const a of ['E3', 'E4', 'E5']) cells[a] = { v: null, border: { ...((cells[a] || {}).border || {}), r: thin } };
  cells.D2 = { v: null, border: { b: { s: 13, c: 0 } } }; // the ring neighbour's own facing side, differing
  const store = createStore([{ name: 'B', cells }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  const { result, m } = await run(store, task('C3:E5'));
  assert.equal(op(result).status, OPERATION_STATUS.APPLIED);
  assert.equal(m.commands, 1);
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification));
  assert.deepEqual(fullRead(ws(store), 'D2').b, thin);
});
