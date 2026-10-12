'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 backlog X-B04 (border removal) -- owner mandate 2026-10-03 (owner channel); zero GO #ref
 * (points 1/2/4) and decision #ref (point 3: BOTH cells of every removed edge, candidate C over an
 * unmeasured internal 'own' read (A) or a write-then-verify-fails attempt (B)).
 *
 * UNLIKE the outline / inside / all acceptance tests, this file is NOT calibrated against a byte-exact LIVE
 * probe of the FULL R0/R2 shape: P2.0 (`xlsx-m7-b4-probe.live.cjs`, decision #ref) specifically probes the
 * both-cells approach's correctness. P2.0 HAS since reported back (m7-b4-probe1-20261004) and confirmed it:
 * the shared edge disappears from both cells, every kept side/fill/bold/italic/numFmt round-trips byte-exact
 * -- but THIS file's seeds are still hand-built, not lifted from that probe's exact fixture, so the relation
 * stated below (what IS measured, via `docs/evidence/M7-B-PROBE.md` sections 3-5, P0.3/P0.5) still holds:
 * `SetBorders(side, 'None')` removes nothing, self-owned or shared (P0.3/a-b); the composite (read every
 * observable axis, `ClearFormats()`, restore every KEPT axis with an EXPLICIT value) is byte-exact R0=R2 on a
 * non-border-only cell (P0.5/a) and removes a shared edge from the OWNER side (P0.5/b); the restore MUST use
 * the explicit, colour-bearing write call -- omitting the colour argument never matches the original (P0.5/c).
 * The fake driver's `applyRemove` simulates exactly that net effect (clear the named sides, leave every other
 * axis as it was) -- not a byte-level replay of the generated script, same relationship the outline/inside
 * fakes have to their own `applyOutline`.
 */

const task = (range, sides, sheet = 'B') => ({ id: 'm7-b04-remove', steps: [{ intent: 'format_range', sheet, range, format: { border: { mode: 'remove', sides } } }] });
const thin = { s: 7, c: 255 };
const thick = { s: 13, c: 65280 };
const blank = { v: 0 };

/** Every window cell (the range, D3:D3, plus its one-cell ring) style-readable, none merged, none diagonal --
 * the baseline `removeWindowGap` needs before any removal can proceed. */
function windowSeed() {
  const win = ['C2', 'D2', 'E2', 'C3', 'D3', 'E3', 'C4', 'D4', 'E4'];
  const cells = {};
  for (const a of win) cells[a] = { ...blank };
  return cells;
}

async function run(cells, script = {}) {
  const store = createStore([{ name: 'B', cells }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const t = task('D3:D3', ['top']);
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, store };
}

const op = result => result.execution.appliedOperations[0];
const ws = store => store.durable.find(s => s.name === 'B');

test('a thin top (the shared edge with D2) removed: D3.top and D2.bottom both go to REMOVED; every other axis unchanged; retry SATISFIED 0/0', async () => {
  const cells = windowSeed();
  cells.D2 = { ...blank, border: { b: thin } };
  cells.D3 = { v: 'x', border: { t: thin, b: thick }, fill: 0x00ff00, font: { bold: true }, align: { h: 1, v: 4 } }; // h:1='left', v:4='top' -- the fake's internal codes, not the spec's word strings
  const { result, m, store } = await run(cells);
  assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  const d3 = ws(store).cells.D3;
  const d2 = ws(store).cells.D2;
  assert.deepEqual(d3.border.t, { s: 0, c: 0 }, 'D3 own top: removed');
  assert.deepEqual(d2.border.b, { s: 0, c: 0 }, 'D2 own bottom (the same physical edge): removed too -- both cells, a maintainer decision #ref');
  assert.deepEqual(d3.border.b, thick, 'D3 own bottom: kept, byte-exact');
  assert.equal(d3.v, 'x', 'content untouched');
  assert.equal(d3.fill, 0x00ff00, 'fill: kept, byte-exact');
  assert.equal(d3.font.bold, true, 'font.bold: kept');
  assert.deepEqual(d3.align, { h: 1, v: 4 }, 'align h+v: kept');
  const again = await run(ws(store).cells);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
});

test('four sides removed on a 2-cell range: the inner shared edge (D3/D4) and all four outer ring facings clear on both cells; one command', async () => {
  const win = ['C2', 'D2', 'E2', 'C3', 'D3', 'E3', 'C4', 'D4', 'E4', 'C5', 'D5', 'E5'];
  const cells = {};
  for (const a of win) cells[a] = { ...blank };
  cells.D3 = { v: 1, border: { t: thin, l: thin, r: thin, b: thin } };
  cells.D4 = { v: 2, border: { t: thin, b: thin, l: thin, r: thin } };
  cells.D2 = { ...blank, border: { b: thin } };
  cells.D5 = { ...blank, border: { t: thin } };
  cells.C3 = { ...blank, border: { r: thin } };
  cells.E3 = { ...blank, border: { l: thin } };
  cells.C4 = { ...blank, border: { r: thin } };
  cells.E4 = { ...blank, border: { l: thin } };
  const store = createStore([{ name: 'B', cells }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  let n = 0;
  const driver = createFakeXlsxDriver(store, {});
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const t = task('D3:D4', ['top', 'bottom', 'left', 'right']);
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.deepEqual([result.metrics.writes, result.metrics.commands], [1, 1]);
  const after = ws(store);
  for (const [a, side] of [['D3', 't'], ['D3', 'l'], ['D3', 'r'], ['D4', 'b'], ['D4', 'l'], ['D4', 'r'], ['D2', 'b'], ['D5', 't'], ['C3', 'r'], ['E3', 'l'], ['C4', 'r'], ['E4', 'l']]) {
    assert.deepEqual(after.cells[a].border[side], { s: 0, c: 0 }, `${a}.${side} removed`);
  }
  // the inner shared edge between the two targets themselves: D3.bottom / D4.top, both removed, neither restored
  assert.deepEqual(after.cells.D3.border.b, { s: 0, c: 0 }, 'D3.bottom (inner edge with D4): removed');
  assert.deepEqual(after.cells.D4.border.t, { s: 0, c: 0 }, 'D4.top (inner edge with D3): removed');
});

for (const [fault, extra, cells] of [
  ['remove-one-cell-only', {}, () => { const c = windowSeed(); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin } }; return c; }],
  ['remove-kept-axis-lost', {}, () => { const c = windowSeed(); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin }, fill: '#00FF00' }; return c; }],
  ['remove-ring-other-side', { faultCell: 'C2' }, () => { const c = windowSeed(); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin } }; return c; }],
]) {
  test(`RED (mutant '${fault}'): the verification fails, the task is not ok`, async () => {
    const { result } = await run(cells(), { fault, ...extra });
    assert.equal(result.ok, false);
    const l1 = (result.execution.primitiveVerification || []).map(v => v.verdict);
    const l3 = (result.execution.preservationVerification || []).map(v => v.verdict);
    assert.ok(l1.includes('FAIL') || l3.includes('FAIL'), JSON.stringify({ status: op(result).status, l1, l3 }));
  });
}

test('refused before the write, 0/0: a diagonal on the target cell (no public restore path); a merge in the window; an unmeasured axis anywhere in the window', async () => {
  const diag = windowSeed();
  diag.D2 = { ...blank, border: { b: thin } };
  diag.D3 = { v: 'x', border: { t: thin, d: { s: 7, c: 0 }, dd: true } };
  const r1 = await run(diag);
  assert.deepEqual([r1.m.writes, r1.m.commands], [0, 0]);
  assert.equal(op(r1.result).evidence.evidence.reason, 'xlsx-format-remove-diagonal-unsupported');

  const merged = windowSeed();
  merged.D2 = { ...blank, border: { b: thin } };
  merged.D3 = { v: 'x', border: { t: thin } };
  const store2 = createStore([{ name: 'B', cells: { ...merged, E3: { v: 'm' } }, merges: ['E3:E4'] }]);
  const driver2 = createFakeXlsxDriver(store2, {});
  let n2 = 0;
  const result2 = await withSession({ driver: driver2, fileId: 'disposable', now: () => (n2 += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(task('D3:D3', ['top']), { observation: { live: liveScope(task('D3:D3', ['top'])) } }));
  assert.deepEqual([result2.metrics.writes, result2.metrics.commands], [0, 0]);
  assert.equal(op(result2).evidence.evidence.reason, 'xlsx-format-remove-merged-ring-refused');

  // X-F02: an empty cell's font is now observable by default -- forcing the reader to
  // throw (emptyFontUnobservable) is what still leaves this removeCells() cell's kept axis unmeasured.
  const empty = {};
  for (const a of ['C2', 'D2', 'E2', 'C3', 'E3', 'C4', 'D4', 'E4']) empty[a] = { ...blank };
  empty.D3 = { v: '' }; // D3 IS a removeCells() cell, strong check applies
  const r3 = await run(empty, { emptyFontUnobservable: ['D3'] });
  assert.deepEqual([r3.m.writes, r3.m.commands], [0, 0]);
  assert.equal(op(r3.result).evidence.evidence.reason, 'xlsx-format-remove-unmeasured-axis');
});

/**
 * review (X-B04 usability, acc1): removeWindowGap used to require full-axis readability on
 * EVERY window cell, even ones the composite never clears or restores (a ring corner, a neighbour of a
 * side that is NOT being removed). A merge-NEG test's own ring corner (blank, font UNOBSERVABLE:empty-
 * cell) refused the whole removal even though it is irrelevant to the write. Fail-closed throughout --
 * this was a usability limitation, not a security gap. removeWindowGap now applies the STRONG check
 * (cellGap) only to removeCells() cells; every other window cell gets the WEAK check (weakGap: border +
 * merge only, the same contract outline/inside/all already use for their own ring).
 */
test('a non-removeCells() window cell only needs border + merge readability -- an unreadable font or a diagonal there does not block the removal; a merge there still does', async () => {
  const base = () => { const c = windowSeed(); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin } }; return c; };

  // C2 is a ring CORNER -- not in removeCells() for a D3-top removal (only D3 and D2 are). Genuinely blank
  // (no `v` key at all), so the fake's font reader answers UNOBSERVABLE:empty-cell -- which would have
  // refused under the old, uniform cellGap-everywhere check.
  const blankCorner = base();
  delete blankCorner.C2.v;
  blankCorner.C2 = {};
  const r1 = await run(blankCorner);
  assert.equal(r1.result.ok, true, JSON.stringify(op(r1.result).evidence));
  assert.deepEqual([r1.m.writes, r1.m.commands], [1, 1]);

  // a diagonal on that SAME non-target corner must not refuse either -- cellGap's diagonal check only
  // applies through the strong path, which a ring corner never reaches.
  const diagCorner = base();
  diagCorner.C2 = { ...blank, border: { d: { s: 7, c: 0 }, dd: true } };
  const r2 = await run(diagCorner);
  assert.equal(r2.result.ok, true, JSON.stringify(op(r2.result).evidence));

  // a MERGE on that same non-target corner must still refuse -- weakGap keeps the merge check.
  const mergedCorner = base();
  const store3 = createStore([{ name: 'B', cells: { ...mergedCorner, C1: { v: 'm' } }, merges: ['C1:C2'] }]);
  const driver3 = createFakeXlsxDriver(store3, {});
  let n3 = 0;
  const result3 = await withSession({ driver: driver3, fileId: 'disposable', now: () => (n3 += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(task('D3:D3', ['top']), { observation: { live: liveScope(task('D3:D3', ['top'])) } }));
  assert.deepEqual([result3.metrics.writes, result3.metrics.commands], [0, 0]);
  assert.equal(op(result3).evidence.evidence.reason, 'xlsx-format-remove-merged-ring-refused');
});

/**
 * review (BLOCKER): observable is not the same as restorable. The fake can seed a border style
 * code or an alignment code that ISN'T in STYLE_REV/H_REV/V_REV -- not UNOBSERVABLE, just a value our write
 * vocabulary cannot turn back into a real call. Without this refusal, `removeRestoreData` would silently
 * emit `undefined`/`null` into the generated script's `SetBorders`/`SetAlignHorizontal` call, AFTER
 * `ClearFormats()` already ran -- a real, unrecoverable format loss, not merely an unproven write. This must
 * be caught BEFORE the dispatch (0 writes, 0 commands), same as the diagonal/merge/unmeasured-axis refusals.
 */
test('refused before the write, 0/0: a KEPT axis reads a value outside the measured write vocabulary (unrestorable, not merely unobservable)', async () => {
  const base = () => { const c = windowSeed(); c.D2 = { ...blank, border: { b: thin } }; return c; };

  const unmappedStyle = base();
  unmappedStyle.D3 = { v: 'x', border: { t: thin, l: { s: 99, c: 0 } } }; // s=99 is not a measured STYLE value
  const r1 = await run(unmappedStyle);
  assert.deepEqual([r1.m.writes, r1.m.commands], [0, 0]);
  assert.deepEqual({ reason: op(r1.result).evidence.evidence.reason, key: op(r1.result).evidence.evidence.key }, { reason: 'xlsx-format-remove-unrestorable-value', key: 'border.l' });

  const unmappedAlignH = base();
  unmappedAlignH.D3 = { v: 'x', border: { t: thin }, align: { h: 4, v: 0 } }; // h:4 is not in H (left/center/right/justify)
  const r2 = await run(unmappedAlignH);
  assert.deepEqual([r2.m.writes, r2.m.commands], [0, 0]);
  assert.deepEqual({ reason: op(r2.result).evidence.evidence.reason, key: op(r2.result).evidence.evidence.key }, { reason: 'xlsx-format-remove-unrestorable-value', key: 'align.h' });

  const unmappedAlignV = base();
  unmappedAlignV.D3 = { v: 'x', border: { t: thin }, align: { h: null, v: 9 } }; // v:9 is not in V
  const r3 = await run(unmappedAlignV);
  assert.deepEqual([r3.m.writes, r3.m.commands], [0, 0]);
  assert.deepEqual({ reason: op(r3.result).evidence.evidence.reason, key: op(r3.result).evidence.evidence.key }, { reason: 'xlsx-format-remove-unrestorable-value', key: 'align.v' });

  // the SAME unmapped style on a side that IS being removed must NOT refuse -- it is cleared, never restored
  const removedSideItself = base();
  removedSideItself.D3 = { v: 'x', border: { t: { s: 99, c: 0 } } };
  const r4 = await run(removedSideItself);
  assert.equal(r4.result.ok, true, JSON.stringify(op(r4.result).evidence));
});

test('plan: mode "remove" rejects style/color, empty/missing/duplicate/unknown sides, the removed sides land in `want` as REMOVED', () => {
  const o = planOperations(task('D3:D3', ['top', 'bottom'])).operations[0];
  assert.deepEqual(o.remove.sides, ['t', 'b']);
  assert.deepEqual(o.want, [{ key: 'border.t', expect: { s: 0, c: 0 } }, { key: 'border.b', expect: { s: 0, c: 0 } }]);
  assert.equal(o.outline, null);
  assert.equal(o.footprint.style, 'C2:E4', 'remove: the ring is in the footprint, same as outline/inside/all');
  const refuse = (sides, re) => assert.match(planOperations(task('D3', sides)).detail || '', re);
  const refuseSpec = (spec, re) => assert.match(planOperations({ id: 't', steps: [{ intent: 'format_range', sheet: 'B', range: 'D3', format: spec }] }).detail || '', re);
  refuse([], /border\.sides must be a non-empty list/);
  refuse(['top', 'top'], /border\.sides must be a non-empty list/);
  refuse(['diagonal'], /border\.sides must be a non-empty list/);
  refuseSpec({ border: { mode: 'remove' } }, /border\.sides must be a non-empty list/);
  refuseSpec({ border: { mode: 'remove', sides: ['top'], style: 'thin', color: '#000000' } }, /border key\(s\) not supported with mode 'remove': style, color/);
});

test('the command is one callCommand, data-driven, one call site each for SetBorders / ClearFormats, and compiles', () => {
  const o = planOperations(task('D3:D3', ['top'])).operations[0];
  const before = {
    D3: { wrap: false, nf: 'General', fill: 'S:none', font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 }, align: { h: null, v: 0 }, border: { t: { s: 7, c: 255 }, b: { s: 0, c: 0 }, l: { s: 0, c: 0 }, r: { s: 0, c: 0 }, d: { s: 0, c: 0 }, dd: false, du: false } },
    D2: { wrap: false, nf: 'General', fill: 'S:none', font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 }, align: { h: null, v: 0 }, border: { t: { s: 0, c: 0 }, b: { s: 7, c: 255 }, l: { s: 0, c: 0 }, r: { s: 0, c: 0 }, d: { s: 0, c: 0 }, dd: false, du: false } },
  };
  const cmd = require('../../src/adapters/xlsx/families/format.cjs').lower(o, { before });
  assert.deepEqual(cmd.meta.removeCells, ['D3', 'D2']);
  assert.deepEqual(cmd.meta.removeClear, { D3: ['t'], D2: ['b'] });
  assert.doesNotThrow(() => new Function(cmd.body));
  assert.equal((cmd.body.match(/\.ClearFormats\(\)/g) || []).length, 1, 'one call site in a loop');
  assert.equal((cmd.body.match(/\.SetBorders\(/g) || []).length, 1, 'one call site in a loop (remove\'s own loop; no outline/plain-cells border write coexists)');

  // a maintainer review v3 #ref (point 3): the fake applies the NET effect, so a setter dropped from the generated
  // script would pass offline -- every restored axis must have exactly one setter call in the restore loop
  // (the data-to-setter wiring). Deleting any one of these lines from removeLine() makes it 0, not 1.
  const loop = cmd.body.slice(cmd.body.indexOf('c.ClearFormats();'));
  for (const call of ['c.SetWrap(rs.wrap)', 'c.SetNumberFormat(rs.nf)', 'c.SetFillColor(rgb(rs.fill))', 'c.SetBold(true)', 'c.SetItalic(true)', 'c.SetFontName(rs.name)', 'c.SetFontSize(rs.size)', 'c.SetFontColor(rgb(rs.color))', 'c.SetAlignHorizontal(rs.h)', 'c.SetAlignVertical(rs.v)', 'c.SetBorders(side, style, rgb(col))']) {
    assert.equal(loop.split(call).length - 1, 1, call);
  }
});

/**
 * a maintainer review v2 #ref: two mutants SURVIVED the offline suite (`cell.font.color === null` disabled;
 * `removeRestoreData`'s fill forced to `null`) because the fake driver's `applyRemove` simulates the NET
 * EFFECT (every other axis stays as it was) rather than consuming the actual generated restore data -- a bug
 * that drops a kept axis from `removeRestoreData`'s OUTPUT is invisible to every test that only checks the
 * resulting cell state through the fake. This test closes that gap directly: it inspects `lower()`'s own
 * `cmd.meta.restoreData` (the exact per-cell data the generated script restores from) against `target.before`,
 * axis by axis, for every cell `removeCells()` names -- target AND neighbour -- and asserts every KEPT axis
 * (the removed sides excluded) is carried forward unchanged. A `fill: null` / `bold: false` / a dropped kept
 * border side would fail here even though the fake-driven pair tests above stay green.
 */
test('lower(op, target).meta.restoreData matches target.before, axis by axis, for every removeCells() cell -- the removed sides excluded', () => {
  const { H, V, STYLE, lower } = require('../../src/adapters/xlsx/families/format.cjs');
  const H_OF = word => H[word]; // word -> int, the forward map; rs.h/rs.v are words (SetAlignHorizontal/Vertical vocabulary)
  const API_STYLE_INV = { Thin: STYLE.thin, Medium: STYLE.medium, Thick: STYLE.thick, Dashed: STYLE.dashed };
  const before = {
    D3: { wrap: true, nf: '0.00', fill: 0x112233, font: { bold: true, italic: true, name: 'Arial', size: 14, color: 0x445566 }, align: { h: 2, v: 4 }, border: { t: { s: 7, c: 255 }, b: { s: 13, c: 65280 }, l: { s: 12, c: 0 }, r: { s: 0, c: 0 }, d: { s: 0, c: 0 }, dd: false, du: false } },
    D2: { wrap: false, nf: 'General', fill: 'S:none', font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 }, align: { h: null, v: 0 }, border: { t: { s: 0, c: 0 }, b: { s: 7, c: 255 }, l: { s: 0, c: 0 }, r: { s: 6, c: 16711680 }, d: { s: 0, c: 0 }, dd: false, du: false } },
  };
  const o = planOperations(task('D3:D3', ['top'])).operations[0];
  const cmd = lower(o, { before });
  assert.deepEqual(cmd.meta.removeCells.slice().sort(), ['D2', 'D3']);
  for (const [a, rs] of cmd.meta.restoreData) {
    const b = before[a];
    const clear = new Set(cmd.meta.removeClear[a]);
    assert.equal(rs.wrap, b.wrap, `${a}.wrap`);
    assert.equal(rs.nf, b.nf, `${a}.nf`);
    assert.equal(rs.fill, typeof b.fill === 'number' ? b.fill : null, `${a}.fill`);
    assert.equal(rs.bold, b.font.bold === true, `${a}.bold`);
    assert.equal(rs.italic, b.font.italic === true, `${a}.italic`);
    assert.equal(rs.name, b.font.name, `${a}.font.name`);
    assert.equal(rs.size, b.font.size, `${a}.font.size`);
    assert.equal(rs.color, b.font.color, `${a}.font.color`);
    assert.equal(rs.h === null ? null : H_OF(rs.h), b.align.h, `${a}.align.h`);
    assert.equal(rs.v === null ? null : V[rs.v], b.align.v, `${a}.align.v`);
    const wantBorder = ['t', 'b', 'l', 'r'].filter(x => !clear.has(x) && b.border[x] && b.border[x].s > 0)
      .map(x => [{ t: 'Top', b: 'Bottom', l: 'Left', r: 'Right' }[x], b.border[x].s, b.border[x].c]);
    const gotBorder = rs.border.map(([side, styleName, col]) => [side, API_STYLE_INV[styleName], col]);
    assert.deepEqual(gotBorder.sort(), wantBorder.sort(), `${a}.border (kept sides)`);
    for (const x of clear) assert.ok(!rs.border.some(([side]) => side === { t: 'Top', b: 'Bottom', l: 'Left', r: 'Right' }[x]), `${a}: removed side ${x} must not be in the restore border list`);
  }
});

test('refused before the write, 0/0: font.name / font.size / font.color read as null on a KEPT axis (restorable would require guessing a default, P0.5/c forbids it)', async () => {
  for (const [key, override] of [['font.name', { name: null }], ['font.size', { size: null }], ['font.color', { color: null }]]) {
    const cells = windowSeed();
    cells.D2 = { ...blank, border: { b: thin } };
    cells.D3 = { v: 'x', border: { t: thin }, font: override };
    const { result, m } = await run(cells);
    assert.deepEqual([m.writes, m.commands], [0, 0], key);
    assert.deepEqual({ reason: op(result).evidence.evidence.reason, key: op(result).evidence.evidence.key }, { reason: 'xlsx-format-remove-unrestorable-value', key }, key);
  }
});
