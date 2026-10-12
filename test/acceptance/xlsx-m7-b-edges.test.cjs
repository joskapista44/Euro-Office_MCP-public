'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const EDGES = require('../../docs/evidence/m7-b-probe1-20261003.json').SCENARIOS.EDGES;

/**
 * M7 backlog X-B02 (inside) / X-B03 (all) -- owner mandate 2026-10-03 (owner channel), zero GO #ref.
 * The fake is CALIBRATED on the LIVE probe (run m7-b-probe1-20261003, `xlsx-m7-b-probe.live.cjs`'s EDGES
 * scenario): its seed is the probe's measured OWN borders (R0), and after the product's inside/all write the
 * fake's full read must equal the probe's measured full read after a reopen (R2), cell by cell and side by side,
 * over the range and its one-cell ring -- for the 3x3, the 1-row and the 1-column shape. INSIDE and ALL share the
 * identical R0 seed (verified byte-equal), so one seed serves both modes.
 *
 * The fake driver's `applyOutline`/fault injection is GENERIC over `cmd.meta.edges` (a [cell, [sides]] list) --
 * it was not written "for outline only", so the same fault names already used by the outline acceptance test
 * (`outline-miss-side`, `outline-inner-side`, `outline-ring-not-cleared`, `outline-ring-other-side`,
 * `outline-inner-edge-lost`, `outline-diagonal-lost`) are reused here verbatim, unmodified.
 */

const SHAPES = { G3: 'C3:E5', G1R: 'C9:E9', G1C: 'C13:C15' };
const SIDES = ['t', 'b', 'l', 'r'];
const norm = x => (x && x.s > 0 ? { s: x.s, c: x.c === null ? 0 : x.c } : { s: 0, c: 0 });
const task = (range, spec) => ({ id: 'm7-b-edges', steps: [{ intent: 'format_range', sheet: 'B', range, format: spec }] });
const specFor = mode => ({ border: { mode, style: 'thin', color: '#0000FF' } });

/** The probe's measured own borders (R0) as the fake's seed -- identical for INSIDE and ALL, every shape, one sheet. */
function seededBook() {
  const cells = {};
  for (const g of Object.keys(SHAPES)) {
    for (const [a, x] of Object.entries(EDGES.R0.INSIDE[g].cells)) {
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

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const op = result => result.execution.appliedOperations[0];
const ws = store => store.durable.find(s => s.name === 'B');

for (const mode of ['inside', 'all']) {
  for (const [g, range] of Object.entries(SHAPES)) {
    test(`${mode} ${g} ${range}: APPLIED in one command; the fake's end state equals the LIVE probe's reopened full read, every side of the range + ring; retry SATISFIED 0/0`, async () => {
      const store = createStore(seededBook());
      const { result, m } = await run(store, task(range, specFor(mode)));
      assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.primitiveVerification, result.execution.preservationVerification]));
      assert.deepEqual([m.writes, m.commands], [1, 1]);
      assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
      const diff = [];
      for (const [a, x] of Object.entries(EDGES.R2[mode.toUpperCase()][g].cells)) {
        const want = Object.fromEntries([...SIDES, 'd'].map(k => [k, norm(x.full[k])]));
        const got = fullRead(ws(store), a);
        if (JSON.stringify(got) !== JSON.stringify(want)) diff.push({ a, got, want });
      }
      assert.deepEqual(diff, [], 'the calibrated end state');
      const again = await run(store, task(range, specFor(mode)));
      assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
      assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
      assert.equal(again.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
    });
  }
}

test('inside: never touches the outer ring (only the grid-internal shared edges); all: DOES touch the ring, same mechanism as outline', async () => {
  const insideStore = createStore(seededBook());
  const ringBefore = fullRead(ws(insideStore), 'B3');
  await run(insideStore, task('C3:E5', specFor('inside')));
  assert.deepEqual(fullRead(ws(insideStore), 'B3'), ringBefore, 'inside: the ring neighbour is untouched');

  const allStore = createStore(seededBook());
  await run(allStore, task('C3:E5', specFor('all')));
  assert.deepEqual(fullRead(ws(allStore), 'B3').r, { s: 7, c: 255 }, 'all: the left ring neighbour shows the outline edge, same as outline mode');
});

test('the command batches the per-cell SetBorders in ONE callCommand (inside: 8 cells, 12 calls on C3:E5; matches the LIVE probe call count)', () => {
  const o = planOperations(task('C3:E5', specFor('inside'))).operations[0];
  const cmd = require('../../src/adapters/xlsx/families/format.cjs').lower(o);
  assert.deepEqual(cmd.meta.edges, [
    ['C3', ['Bottom', 'Right']], ['D3', ['Bottom', 'Right']], ['E3', ['Bottom']],
    ['C4', ['Bottom', 'Right']], ['D4', ['Bottom', 'Right']], ['E4', ['Bottom']],
    ['C5', ['Right']], ['D5', ['Right']],
  ]);
  assert.equal(cmd.meta.edges.reduce((n, [, s]) => n + s.length, 0), 12, 'matches the LIVE probe (m7-b-probe1): INSIDE G3 calls=12');
  assert.doesNotThrow(() => new Function(cmd.body));
  assert.equal((cmd.body.match(/SetBorders\(/g) || []).length, 1, 'one call site in a loop');
});

test('the command batches the per-cell SetBorders in ONE callCommand (all: 9 cells, 24 calls on C3:E5; matches the LIVE probe call count)', () => {
  const o = planOperations(task('C3:E5', specFor('all'))).operations[0];
  const cmd = require('../../src/adapters/xlsx/families/format.cjs').lower(o);
  assert.equal(cmd.meta.edges.length, 9, 'every cell of the 3x3 carries at least one side');
  assert.equal(cmd.meta.edges.reduce((n, [, s]) => n + s.length, 0), 24, 'matches the LIVE probe (m7-b-probe1): ALL G3 calls=24');
  assert.doesNotThrow(() => new Function(cmd.body));
});

for (const [fault, extra] of [
  ['outline-miss-side', {}],
  ['outline-ring-not-cleared', {}],
  ['outline-ring-other-side', { faultCell: 'D2' }],
  ['outline-inner-edge-lost', { faultCell: 'D4' }],
  ['outline-diagonal-lost', { faultCell: 'D5' }],
]) {
  for (const mode of ['inside', 'all']) {
    test(`RED (mutant '${fault}', mode '${mode}'): the verification fails, the task is not ok`, async () => {
      const store = createStore(seededBook());
      const { result } = await run(store, task('C3:E5', specFor(mode)), { fault, ...extra });
      assert.equal(result.ok, false);
      const l3 = (result.execution.preservationVerification || []).map(v => v.verdict);
      assert.ok(op(result).status !== OPERATION_STATUS.APPLIED || l3.includes('FAIL'), JSON.stringify([op(result).status, l3]));
    });
  }
}

test('refused before the write, 0/0: a merged cell in the window; an unverified border reader (inside mode)', async () => {
  const b = seededBook();
  b[0].merges = ['F4:G4'];
  b[0].cells.F4 = { v: 'm' };
  const merged = await run(createStore(b), task('C3:E5', specFor('inside')));
  assert.deepEqual([merged.m.writes, merged.m.commands], [0, 0]);
  assert.equal(op(merged.result).evidence.evidence.reason, 'xlsx-format-outline-merged-ring-refused');
  const blind = await run(createStore(seededBook()), task('C3:E5', specFor('all')), { appearanceBuild: false });
  assert.deepEqual([blind.m.writes, blind.m.commands], [0, 0]);
  assert.equal(op(blind.result).evidence.evidence.reason, 'xlsx-format-border-unobservable');
});

/**
 * review: classify() must not call this SATISFIED. Every inside-write's own WRITING side already
 * holds the request, but D4's own Top (the side that would read D3's Bottom through the shared edge, unless
 * D4 carries its own, stale value -- which it does here) still shows an untouched, differing style.
 *
 * Measured directly (not assumed): classify()'s SATISFIED line is plain `wantsAt(...).every(holds)` -- it does
 * NOT use the new `expectationsAt` helper, and does not need to. `ringMisses(op, primary)`, already ANDed into
 * the same condition, iterates the WHOLE window (range + ring, a superset of `op.cells`) via the now-generalized
 * `facingSides`, so it alone already catches D4's stale Top for inside/all, exactly as it already caught a
 * ring cell's stale facing side for outline. (Reverting `ringMisses`'s own facingSides generalization -- not
 * `expectationsAt` -- is what would make this test wrongly pass.) `expectationsAt` is used ONLY in
 * `verifyPrimitive` (L1), which has no window-wide check of its own; the genuine mutation-kill proof for THAT
 * fix is the existing `RED (mutant 'outline-ring-not-cleared', mode 'inside')` test below, confirmed by reverting
 * `expectationsAt` back to `wantsAt` there: that one test alone flips from catching the fault to missing it.
 */
test('classify never calls this SATISFIED: a stale facing side forces APPLY even though every writing side already holds the request', async () => {
  const thin = { s: 7, c: 255 }; // the requested style (thin, blue) -- already present on every WRITING side
  const cells = {
    C3: { v: null, border: { b: thin, r: thin } },
    D3: { v: null, border: { b: thin, r: thin } },
    E3: { v: null, border: { b: thin } },
    C4: { v: null, border: { b: thin, r: thin } },
    // D4's own Top is STALE (thick/black) -- D3's Bottom already holds the request, but D4 has its own,
    // untouched value for the shared edge, so the full read of D4.t is WRONG until a write clears it.
    D4: { v: 4, border: { b: thin, r: thin, t: { s: 13, c: 0 } } },
    E4: { v: null, border: { b: thin } },
    C5: { v: null, border: { r: thin } },
    D5: { v: null, border: { r: thin } },
  };
  const store = createStore([{ name: 'B', cells }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  const { result, m } = await run(store, task('C3:E5', specFor('inside')));
  assert.notEqual(op(result).status, OPERATION_STATUS.SATISFIED, 'every writing side already holds, but D4.t does not -- must be APPLY, not SATISFIED');
  assert.equal(op(result).status, OPERATION_STATUS.APPLIED);
  assert.deepEqual([m.writes, m.commands], [1, 1]);
  assert.equal(result.ok, true, JSON.stringify(result.execution.preservationVerification));
  assert.deepEqual(fullRead(ws(store), 'D4').t, thin, 'after the write, the stale own side is cleared and the shared edge resolves correctly');
});

test('plan: inside/all take no sides option, always both inner directions / all four outer sides; removal still refused', () => {
  const insideOp = planOperations(task('C3:E5', specFor('inside'))).operations[0];
  assert.deepEqual(insideOp.outline.sides, []);
  assert.equal(insideOp.inside, true);
  const allOp = planOperations(task('C3:E5', specFor('all'))).operations[0];
  assert.deepEqual(allOp.outline.sides, ['t', 'b', 'l', 'r']);
  assert.equal(allOp.inside, true);
  assert.equal(allOp.footprint.style, 'B2:F6', 'all: the ring is in the footprint, same as outline');
  const refuse = (spec, re) => assert.match(planOperations(task('C3', spec)).detail || '', re);
  refuse({ border: { mode: 'inside', sides: ['top'], style: 'thin', color: '#000000' } }, /border.sides is not supported with mode 'inside'/);
  refuse({ border: { mode: 'all', sides: ['top'], style: 'thin', color: '#000000' } }, /border.sides is not supported with mode 'all'/);
  refuse({ border: { mode: 'inside', style: 'none', color: '#000000' } }, /border\.style 'none' is not a style/);
  refuse({ border: { mode: 'inside', style: 'triple', color: '#000000' } }, /the measured styles/);
});

test('an outline-mode op is unaffected by the inside generalization: op.inside is false, edgeSides/facingSides unchanged', () => {
  const o = planOperations(task('C3:E5', { border: { mode: 'outline', style: 'thin', color: '#0000FF' } })).operations[0];
  assert.equal(o.inside, false);
  const cmd = require('../../src/adapters/xlsx/families/format.cjs').lower(o);
  assert.deepEqual(cmd.meta.edges, [['C3', ['Top', 'Left']], ['D3', ['Top']], ['E3', ['Top', 'Right']], ['C4', ['Left']], ['E4', ['Right']], ['C5', ['Bottom', 'Left']], ['D5', ['Bottom']], ['E5', ['Bottom', 'Right']]], 'byte-identical to the outline-only acceptance test');
});
