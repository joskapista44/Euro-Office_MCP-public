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
 * M7 backlog X-F08 (fill removal, `format: { fill: 'none' }`) --, built on the X-B04
 * remove machinery's own pattern (format-fill-clear.cjs). UNLIKE X-B04, this is single-cell: no
 * shared-edge write of its own (a maintainer decision #ref does not apply), so this file's fake-driven
 * tests only model the net effect of `applyFillClear` (fill reads the measured 'S:No Fill', every
 * other axis exactly as it was) -- the byte-exact restore sequence itself is LIVE-measured
 * (m7-s15-probe4 accepted). Mirrors xlsx-m7-b04-remove.test.cjs's own structure and
 * fault-injection convention.
 *
 * The shared-edge refusal (format-fill-clear.cjs's `sharedEdge`) fires whenever the target has a
 * non-zero border side facing a READABLE neighbour at all -- the probe's own accepted run never
 * checked a neighbour's state, so most of this file's fixtures stay border-free on D3 to reach
 * the restore/verification path; the border axis itself is proven separately, on a sheet corner
 * where the relevant sides have no neighbour to disturb.
 */

const task = (range, sheet = 'B') => ({ id: 'm7-f08-fill-clear', steps: [{ intent: 'format_range', sheet, range, format: { fill: 'none' } }] });
const thin = { s: 7, c: 255 };
const blank = { v: 0 };

/** Every window cell (the target, D3:D3, plus its one-cell ring) style-readable, none merged,
 * none diagonal, no border at all (so the shared-edge check has nothing to trip on by default). */
function windowSeed() {
  const win = ['C2', 'D2', 'E2', 'C3', 'D3', 'E3', 'C4', 'D4', 'E4'];
  const cells = {};
  for (const a of win) cells[a] = { ...blank };
  return cells;
}

async function run(cells, script = {}, range = 'D3:D3') {
  const store = createStore([{ name: 'B', cells }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const t = task(range);
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, store };
}

const op = result => result.execution.appliedOperations[0];
const ws = store => store.durable.find(s => s.name === 'B');

test('fill cleared, every kept axis byte-exact (wrap, nf, bold, italic, name/size/color, align); content untouched; retry SATISFIED 0/0', async () => {
  const cells = windowSeed();
  cells.D3 = { v: 'x', fill: 0x00ff00, wrap: true, nf: '0.00', font: { bold: true, italic: true, name: 'Arial', size: 14, color: 0x112233 }, align: { h: 1, v: 4 }, bold: true, italic: true };
  const { result, m, store } = await run(cells);
  assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  const d3 = ws(store).cells.D3;
  assert.equal(d3.fill, 'S:No Fill', 'fill cleared -- the measured no-fill read');
  assert.equal(d3.v, 'x', 'content untouched');
  assert.equal(d3.wrap, true, 'wrap: kept');
  assert.equal(d3.nf, '0.00', 'numberFormat: kept');
  assert.deepEqual(d3.font, { bold: true, italic: true, name: 'Arial', size: 14, color: 0x112233 }, 'font: kept');
  assert.deepEqual(d3.align, { h: 1, v: 4 }, 'align h+v: kept');
  const again = await run(ws(store).cells);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
});

test('fill already \'none\' on an untouched cell is SATISFIED, 0 writes -- the same shape as any already-true request', async () => {
  const cells = windowSeed();
  cells.D3 = { v: 'x' };
  const { result, m } = await run(cells);
  assert.equal(op(result).status, OPERATION_STATUS.SATISFIED, JSON.stringify(op(result).evidence));
  assert.deepEqual([m.writes, m.commands], [0, 0]);
});

/** A border side at the SHEET'S OUTER EDGE (A1's top / left) has no neighbour to disturb -- the
 * one case where a kept border can be proven restored without tripping the shared-edge refusal. */
test('a kept border on a sheet-edge side (no neighbour at all) is also restored byte-exact', async () => {
  const cells = { A1: { v: 'x', fill: 0x00ff00, border: { t: thin, l: thin } }, A2: { ...blank }, B1: { ...blank }, B2: { ...blank } };
  const { result, m, store } = await run(cells, {}, 'A1:A1');
  assert.equal(result.ok, true, JSON.stringify([op(result).evidence, result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.deepEqual([m.writes, m.commands], [1, 1]);
  const a1 = ws(store).cells.A1;
  assert.equal(a1.fill, 'S:No Fill', 'fill cleared');
  assert.deepEqual(a1.border.t, thin, 'border.t: kept, byte-exact (no neighbour above row 1)');
  assert.deepEqual(a1.border.l, thin, 'border.l: kept, byte-exact (no neighbour left of column A)');
});

for (const fault of ['fillclear-fill-not-cleared', 'fillclear-kept-axis-lost']) {
  test(`RED (mutant '${fault}'): the verification fails, the task is not ok`, async () => {
    const cells = windowSeed();
    cells.D3 = { v: 'x', fill: 0x00ff00, wrap: true, font: { name: 'Arial', size: 11, color: 0 } };
    const { result } = await run(cells, { fault });
    assert.equal(result.ok, false);
    const l1 = (result.execution.primitiveVerification || []).map(v => v.verdict);
    const l3 = (result.execution.preservationVerification || []).map(v => v.verdict);
    assert.ok(l1.includes('FAIL') || l3.includes('FAIL'), JSON.stringify({ status: op(result).status, l1, l3 }));
  });
}

test('refused before the write, 0/0: a diagonal on the target cell (no public restore path); a merge on the target; an unmeasured kept axis', async () => {
  const diag = windowSeed();
  diag.D3 = { v: 'x', fill: 0x00ff00, border: { d: { s: 7, c: 0 }, dd: true } };
  const r1 = await run(diag);
  assert.deepEqual([r1.m.writes, r1.m.commands], [0, 0]);
  assert.equal(op(r1.result).evidence.evidence.reason, 'xlsx-format-fill-clear-diagonal-unsupported');

  const merged = windowSeed();
  merged.D3 = { v: 'x', fill: 0x00ff00 };
  const store2 = createStore([{ name: 'B', cells: { ...merged, D4: { v: 'm' } }, merges: ['D3:D4'] }]);
  const driver2 = createFakeXlsxDriver(store2, {});
  let n2 = 0;
  const t2 = task('D3:D3');
  const result2 = await withSession({ driver: driver2, fileId: 'disposable', now: () => (n2 += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(t2, { observation: { live: liveScope(t2) } }));
  assert.deepEqual([result2.metrics.writes, result2.metrics.commands], [0, 0]);
  assert.equal(op(result2).evidence.evidence.reason, 'xlsx-format-in-merged-area-refused');

  // X-F02: an empty cell's font is now observable by default -- only a reader that
  // throws for that cell (emptyFontUnobservable) still leaves the kept axis unmeasured.
  const empty = windowSeed();
  empty.D3 = { v: '' };
  const r3 = await run(empty, { emptyFontUnobservable: ['D3'] });
  assert.deepEqual([r3.m.writes, r3.m.commands], [0, 0]);
  assert.equal(op(r3.result).evidence.evidence.reason, 'xlsx-format-fill-clear-unmeasured-axis');
});

test('refused before the write, 0/0: a KEPT axis reads a value outside the measured write vocabulary (unrestorable, not merely unobservable)', async () => {
  const base = () => { const c = windowSeed(); c.D3 = { v: 'x', fill: 0x00ff00 }; return c; };

  const unmappedStyle = base();
  unmappedStyle.D3 = { ...unmappedStyle.D3, border: { l: { s: 99, c: 0 } } }; // s=99 is not a measured STYLE value
  const r1 = await run(unmappedStyle);
  assert.deepEqual([r1.m.writes, r1.m.commands], [0, 0]);
  assert.deepEqual({ reason: op(r1.result).evidence.evidence.reason, key: op(r1.result).evidence.evidence.key }, { reason: 'xlsx-format-fill-clear-unrestorable-value', key: 'border.l' });

  const unmappedAlignH = base();
  unmappedAlignH.D3 = { ...unmappedAlignH.D3, align: { h: 4, v: 0 } }; // h:4 is not in H (left/center/right/justify)
  const r2 = await run(unmappedAlignH);
  assert.deepEqual([r2.m.writes, r2.m.commands], [0, 0]);
  assert.deepEqual({ reason: op(r2.result).evidence.evidence.reason, key: op(r2.result).evidence.evidence.key }, { reason: 'xlsx-format-fill-clear-unrestorable-value', key: 'align.h' });
});

test('refused before the write, 0/0: the target carries a border side facing a readable neighbour -- the neighbour interaction the probe never measured, regardless of what the neighbour itself carries', async () => {
  const cells = windowSeed(); // D2 carries NO border of its own
  cells.D3 = { v: 'x', fill: 0x00ff00, border: { t: thin } };
  const { result, m } = await run(cells);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.deepEqual({ reason: op(result).evidence.evidence.reason, side: op(result).evidence.evidence.side, neighbour: op(result).evidence.evidence.neighbour }, { reason: 'xlsx-format-fill-clear-shared-edge-unsupported', side: 't', neighbour: 'D2' });
});

test('plan: "none" rejects combination with any other format key, rejects a non-hex/non-"none" string; the want shape', () => {
  const refuseSpec = spec => assert.match(planOperations({ id: 't', steps: [{ intent: 'format_range', sheet: 'B', range: 'D3', format: spec }] }).detail || '', /fill 'none' must be the only format key|fill must be '#RRGGBB' or 'none'/);
  refuseSpec({ fill: 'none', wrap: true });
  refuseSpec({ fill: 'none', border: { sides: ['top'], style: 'thin', color: '#000000' } });
  refuseSpec({ fill: 'notahex' });
  const o = planOperations(task('D3:D3')).operations[0];
  assert.deepEqual(o.want, [{ key: 'fill', expect: null }]);
  assert.equal(o.fillClear, true);
  assert.equal(o.outline, null);
  assert.equal(o.remove, null);
});

test('the command is one callCommand, data-driven, one call site each for ClearFormats / the kept setters, NO fill-restore call; compiles', () => {
  const o = planOperations(task('D3:D3')).operations[0];
  const before = { D3: { wrap: true, nf: '0.00', fill: 0x00ff00, font: { bold: true, italic: true, name: 'Arial', size: 14, color: 0x112233 }, align: { h: 2, v: 4 }, border: { t: { s: 7, c: 255 }, b: { s: 0, c: 0 }, l: { s: 0, c: 0 }, r: { s: 0, c: 0 }, d: { s: 0, c: 0 }, dd: false, du: false } } };
  const cmd = require('../../src/adapters/xlsx/families/format.cjs').lower(o, { before });
  assert.deepEqual(cmd.meta.fillClearCells, ['D3']);
  assert.doesNotThrow(() => new Function(cmd.body));
  assert.equal((cmd.body.match(/\.ClearFormats\(\)/g) || []).length, 1, 'one call site in a loop');
  assert.equal((cmd.body.match(/\.SetFillColor\(/g) || []).length, 0, 'fill is never re-applied -- the one axis deliberately left cleared');
  const loop = cmd.body.slice(cmd.body.indexOf('c.ClearFormats();'));
  for (const call of ['c.SetWrap(rs.wrap)', 'c.SetNumberFormat(rs.nf)', 'c.SetBold(true)', 'c.SetItalic(true)', 'c.SetFontName(rs.name)', 'c.SetFontSize(rs.size)', 'c.SetFontColor(rgbFC(rs.color))', 'c.SetAlignHorizontal(rs.h)', 'c.SetAlignVertical(rs.v)', 'c.SetBorders(side, style, rgbFC(col))']) {
    assert.equal(loop.split(call).length - 1, 1, call);
  }
});

/**
 * review's lesson, applied here: the fake's applyFillClear simulates the NET EFFECT
 * (every other axis stays as it was), so a bug that drops a kept axis from
 * `fillClearRestoreData`'s OUTPUT is invisible to every fake-driven test above. This inspects
 * `lower()`'s own `cmd.meta.restoreData` against `target.before`, axis by axis.
 */
test('lower(op, target).meta.restoreData matches target.before, axis by axis -- fill excluded', () => {
  const { H, V, STYLE, lower } = require('../../src/adapters/xlsx/families/format.cjs');
  const H_OF = word => H[word];
  const API_STYLE_INV = { Thin: STYLE.thin, Medium: STYLE.medium, Thick: STYLE.thick, Dashed: STYLE.dashed };
  const before = { D3: { wrap: true, nf: '0.00', fill: 0x112233, font: { bold: true, italic: true, name: 'Arial', size: 14, color: 0x445566 }, align: { h: 2, v: 4 }, border: { t: { s: 7, c: 255 }, b: { s: 13, c: 65280 }, l: { s: 12, c: 0 }, r: { s: 0, c: 0 }, d: { s: 0, c: 0 }, dd: false, du: false } } };
  const o = planOperations(task('D3:D3')).operations[0];
  const cmd = lower(o, { before });
  assert.deepEqual(cmd.meta.fillClearCells, ['D3']);
  const [[a, rs]] = cmd.meta.restoreData;
  assert.equal(a, 'D3');
  const b = before.D3;
  assert.equal(rs.wrap, b.wrap);
  assert.equal(rs.nf, b.nf);
  assert.equal('fill' in rs, false, 'fill is not part of the restore data at all');
  assert.equal(rs.bold, b.font.bold === true);
  assert.equal(rs.italic, b.font.italic === true);
  assert.equal(rs.name, b.font.name);
  assert.equal(rs.size, b.font.size);
  assert.equal(rs.color, b.font.color);
  assert.equal(H_OF(rs.h), b.align.h);
  assert.equal(V[rs.v], b.align.v);
  const wantBorder = ['t', 'b', 'l', 'r'].filter(x => b.border[x] && b.border[x].s > 0).map(x => [{ t: 'Top', b: 'Bottom', l: 'Left', r: 'Right' }[x], b.border[x].s, b.border[x].c]);
  const gotBorder = rs.border.map(([side, styleName, col]) => [side, API_STYLE_INV[styleName], col]);
  assert.deepEqual(gotBorder.sort(), wantBorder.sort());
});
