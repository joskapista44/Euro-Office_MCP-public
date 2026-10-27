'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * L-16/3: border remove over a MERGE -- C5 = UnMerge -> the composite -> Merge, in ONE body (src/adapters/xlsx/families/format-remove.cjs c5Wrap).
 * LIVE-measured (test/live/xlsx-l16-3-merged-side-remove-probe.live.cjs, xlsx-l16-3-c5-risk-probe.live.cjs, prod DS, session + reopen): the SetBorders('None') family removes nothing
 * (plain cells too), C5 is the only path that keeps the merge, and it WORKS on the realistic merged header and on 2x2 / 1x3 / 3x1 x four sides. The fake driver models the NET
 * effect (sides cleared, merge intact); the generated body itself is executed here against a recording model of the editor API (the UnMerge / ClearFormats / Merge order, the
 * `finally` re-merge, the member-value guard), and the verification is exercised by mutating the reads taken AFTER the write.
 */

const thin = { s: 7, c: 255 };
const blank = { v: 0 };
const grid = (cols, rows) => { const c = {}; for (const x of cols) for (const r of rows) c[`${x}${r}`] = { ...blank }; return c; };
const ws = store => store.durable.find(s => s.name === 'B');
const task = (range, sides) => ({ id: 'l16-3-c5', steps: [{ intent: 'format_range', sheet: 'B', range, format: { border: { mode: 'remove', sides } } }] });

/** `mutate(cell)` rewrites a cell of the reads taken AFTER the write. Captures the dispatched body. */
async function run(cells, merges, range = 'D3:D3', sides = ['top'], mutate = null) {
  const store = createStore([{ name: 'B', cells, merges }, { name: 'Other', cells: { A1: { v: 'keep' } } }]);
  const inner = createFakeXlsxDriver(store, {});
  let wrote = false; let touched = 0; let body = null;
  const driver = new Proxy(inner, {
    get: (t, k) => {
      if (k === 'dispatch') return async (h, cmd) => { wrote = true; body = cmd.body; return t.dispatch(h, cmd); };
      if (k === 'readLive') return async (...a) => { const r = await t.readLive(...a); if (mutate && wrote && r && r.payload && Array.isArray(r.payload.cells)) for (const c of r.payload.cells) if (mutate(c)) touched += 1; return r; };
      return typeof t[k] === 'function' ? t[k].bind(t) : t[k];
    },
  });
  let n = 0; const t = task(range, sides);
  const result = await withSession({ driver, fileId: 'd', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  const op = result.execution.appliedOperations[0];
  return { result, m: result.metrics, store, body, touched, evidence: op && op.evidence && op.evidence.evidence, status: op && op.status, l1: result.execution.primitiveVerification || [], l3: result.execution.preservationVerification || [] };
}

// ---- 1. the released shapes: applied, the merge survives, retry SATISFIED ------------------------------------------------------------------------------

const SHAPES = { '2x2': { area: 'D4:E5', anchor: 'D4' }, '1x3': { area: 'D4:F4', anchor: 'D4' }, '3x1': { area: 'D4:D6', anchor: 'D4' } };
const edgeCells = (shape, side) => {
  const m = { '2x2': { top: ['D4', 'E4'], bottom: ['D5', 'E5'], left: ['D4', 'D5'], right: ['E4', 'E5'] }, '1x3': { top: ['D4', 'E4', 'F4'], bottom: ['D4', 'E4', 'F4'], left: ['D4'], right: ['F4'] }, '3x1': { top: ['D4'], bottom: ['D6'], left: ['D4', 'D5', 'D6'], right: ['D4', 'D5', 'D6'] } };
  return m[shape][side];
};
const KEY = { top: 't', bottom: 'b', left: 'l', right: 'r' };
function shapeCells(shape, side) {
  const cells = grid(['B', 'C', 'D', 'E', 'F', 'G', 'H'], [2, 3, 4, 5, 6, 7, 8]);
  cells[SHAPES[shape].anchor] = { v: 'm', border: {} };
  for (const a of edgeCells(shape, side)) cells[a] = { ...cells[a], border: { [KEY[side]]: thin } };
  return cells;
}

for (const shape of Object.keys(SHAPES)) {
  test(`C5 target: a merged ${shape} header, each of the four sides -> APPLIED 1/1, the merge survives, the edge is gone, the retry is SATISFIED 0/0`, async () => {
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const r = await run(shapeCells(shape, side), [SHAPES[shape].area], SHAPES[shape].area, [side]);
      assert.equal(r.result.ok, true, `${shape} ${side} ${JSON.stringify(r.evidence)}`);
      assert.deepEqual([r.m.writes, r.m.commands], [1, 1], `${shape} ${side}`);
      assert.deepEqual(ws(r.store).merges, [SHAPES[shape].area], `${shape} ${side}: merge`);
      for (const a of edgeCells(shape, side)) assert.deepEqual(ws(r.store).cells[a].border[KEY[side]], { s: 0, c: 0 }, `${shape} ${side} ${a}`);
      assert.ok(r.body.includes('UnMerge') && r.body.includes('finally'), `${shape} ${side}: the body is the C5 body`);
      const again = await run(ws(r.store).cells, [SHAPES[shape].area], SHAPES[shape].area, [side]);
      assert.deepEqual([again.m.writes, again.m.commands, again.status], [0, 0, 'SATISFIED'], `${shape} ${side}: retry`);
    }
  });
}

test('C5 neighbour: the merge that FACES the removed side (above / left / below / right of a plain target) is a C5 area -> applied, merge survives', async () => {
  const base = () => grid(['B', 'C', 'D', 'E', 'F'], [1, 2, 3, 4, 5]);
  const cases = [
    { side: 'top', merges: ['D1:D2'], edge: { D3: { t: thin }, D2: { b: thin } } },
    { side: 'bottom', merges: ['D4:D5'], edge: { D3: { b: thin }, D4: { t: thin } } },
    { side: 'left', merges: ['B3:C3'], edge: { D3: { l: thin }, C3: { r: thin } } },
    { side: 'right', merges: ['E3:F3'], edge: { D3: { r: thin }, E3: { l: thin } } },
  ];
  for (const c of cases) {
    const cells = base();
    for (const [a, b] of Object.entries(c.edge)) cells[a] = { ...cells[a], v: a === 'D3' ? 'x' : cells[a].v, border: b };
    const r = await run(cells, c.merges, 'D3:D3', [c.side]);
    assert.equal(r.result.ok, true, `${c.side} ${JSON.stringify(r.evidence)}`);
    assert.deepEqual([r.m.writes, r.m.commands], [1, 1], c.side);
    assert.deepEqual(ws(r.store).merges, c.merges, c.side);
    assert.ok(r.body.includes('UnMerge'), `${c.side}: C5 body`);
  }
});

// ---- 2. what C5 was NOT measured for stays refused, named ------------------------------------------------------------------------------------------------

test('REFUSED, named, 0/0: two merge areas in one write (adjacent merges, M5) | an area too large for the member guard | a merged target that crosses the range', async () => {
  const cells = grid(['C', 'D', 'E', 'F'], [1, 2, 3, 4]);
  const two = await run({ ...cells, D3: { v: 'm' }, D2: { v: 'n' } }, ['D3:E3', 'D2:E2'], 'D3:E3', ['top']);
  assert.deepEqual([two.m.writes, two.m.commands], [0, 0]);
  assert.equal(two.evidence.reason, 'xlsx-format-remove-multi-merge-unmeasured');
  assert.equal(two.evidence.count, 2);
  const big = await run({ ...cells, D2: { v: 'n' } }, ['D2:QZ2'], 'D3:D3', ['top']);
  assert.deepEqual([big.m.writes, big.m.commands], [0, 0]);
  assert.deepEqual([big.evidence.reason, big.evidence.mergeArea, big.evidence.max], ['xlsx-format-remove-merge-too-large', 'D2:QZ2', 400]);
  const cross = await run({ ...cells, D3: { v: 'm' } }, ['D3:E3'], 'D3:D3', ['top']);
  assert.deepEqual([cross.m.writes, cross.m.commands, cross.evidence.reason], [0, 0, 'xlsx-format-merge-crosses-range-refused']);
});

// ---- 3. the generated body, executed against a recording model of the editor API -------------------------------------------------------------------------

function modelApi({ values = {}, merges = [], throwOn = null } = {}) {
  const log = []; const live = new Set(merges);
  const range = a => new Proxy({}, {
    get: (t, k) => {
      if (k === 'GetValue') return () => { if (throwOn === 'GetValue') throw new Error('boom GetValue'); log.push(['GetValue', a]); return values[a]; };
      if (k === 'UnMerge') return () => { if (throwOn === 'UnMerge') throw new Error('boom UnMerge'); log.push(['UnMerge', a]); live.delete(a); };
      if (k === 'Merge') return flag => { if (throwOn === 'Merge') throw new Error('boom Merge'); log.push(['Merge', a, flag]); live.add(a); };
      if (k === 'ClearFormats') return () => { if (throwOn === 'ClearFormats') throw new Error('boom ClearFormats'); log.push(['ClearFormats', a]); };
      return () => undefined;
    },
  });
  const sheet = { GetName: () => 'B', GetRange: range };
  return { Api: { GetSheet: () => sheet, CreateColorFromRGB: () => 0 }, log, live };
}
const exec = (body, model) => new Function('Api', body)(model.Api);
const names = log => log.map(x => x[0]);

async function bodyOf(shape) {
  if (shape === 'A') { const c = grid(['B', 'C', 'D', 'E'], [1, 2, 3, 4]); c.D2 = { ...c.D2, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin } }; return (await run(c, ['D1:D2'], 'D3:D3', ['top'])).body; }
  const c = shapeCells('1x3', 'top'); return (await run(c, ['D4:F4'], 'D4:F4', ['top'])).body;
}

test('the body: guard reads, then UnMerge(area), the composite, then Merge(area, false) -- exact order, one area', async () => {
  const body = await bodyOf('A');
  const model = modelApi({ merges: ['D1:D2'] });
  assert.deepEqual(exec(body, model), { ok: true });
  assert.deepEqual(model.log.filter(x => x[0] === 'ClearFormats').map(x => x[1]).sort(), ['D2', 'D3'], 'the composite runs on exactly the removeCells (the target and the facing neighbour)');
  assert.deepEqual(model.log.filter(x => x[0] === 'GetValue').map(x => x[1]), ['D2'], 'the guard reads exactly the non-anchor member(s) of the area');
  const order = names(model.log);
  assert.ok(order.lastIndexOf('GetValue') < order.indexOf('UnMerge') && order.indexOf('UnMerge') < order.indexOf('ClearFormats') && order.lastIndexOf('ClearFormats') < order.indexOf('Merge'));
  assert.deepEqual(model.log.find(x => x[0] === 'Merge'), ['Merge', 'D1:D2', false], 'the same area, across = false');
  assert.deepEqual([...model.live], ['D1:D2']);
});

test('FAILURE after the UnMerge: the finally puts the area back and the original error is reported; a failing re-Merge is not masked as success by the verification', async () => {
  const body = await bodyOf('A');
  for (const throwOn of ['ClearFormats']) {
    const model = modelApi({ merges: ['D1:D2'], throwOn });
    const r = exec(body, model);
    assert.equal(r.ok, false);
    assert.equal(r.stage, 'threw');
    assert.match(r.error, /boom ClearFormats/);
    assert.deepEqual(model.log.filter(x => x[0] === 'Merge' || x[0] === 'UnMerge').map(x => x[0]), ['UnMerge', 'Merge'], 'UnMerge happened, then the finally re-merged');
    assert.deepEqual([...model.live], ['D1:D2'], 'the merge is back');
  }
  const noReMerge = modelApi({ merges: ['D1:D2'], throwOn: 'Merge' });
  assert.deepEqual(exec(body, noReMerge), { ok: true }, 'a re-Merge that throws is swallowed by the finally -- the VERIFICATION (below) is what names the merge that did not come back');
  assert.deepEqual([...noReMerge.live], [], 'and the area is indeed gone');
  const unmergeFails = modelApi({ merges: ['D1:D2'], throwOn: 'UnMerge' });
  const rr = exec(body, unmergeFails);
  assert.equal(rr.stage, 'threw');
  assert.deepEqual(unmergeFails.log.filter(x => x[0] === 'ClearFormats' || x[0] === 'Merge'), [], 'nothing was written and nothing needs re-merging when the UnMerge itself threw');
});

test('GUARD: a value in a non-anchor member aborts BEFORE the UnMerge -- nothing is touched', async () => {
  for (const which of ['A', 'B']) {
    const body = await bodyOf(which);
    const members = which === 'A' ? ['D2'] : ['E4', 'F4'];
    for (const m of members) {
      const model = modelApi({ merges: [which === 'A' ? 'D1:D2' : 'D4:F4'], values: { [m]: 'hidden' } });
      const r = exec(body, model);
      assert.deepEqual([r.ok, r.stage, r.cell], [false, 'merge-member-value', m]);
      assert.deepEqual(model.log.filter(x => ['UnMerge', 'ClearFormats', 'Merge'].includes(x[0])), [], `${which} ${m}: no write at all`);
    }
    const empty = modelApi({ merges: [which === 'A' ? 'D1:D2' : 'D4:F4'], values: { [members[0]]: '' } });
    assert.equal(exec(body, empty).ok, true, 'an empty member does not abort');
  }
});

test('GUARD is best-effort: a member the raw GetValue() cannot answer for (it throws) does not abort the write -- the C5 sequence still runs and the area comes back', async () => {
  const body = await bodyOf('A');
  const model = modelApi({ merges: ['D1:D2'], throwOn: 'GetValue' });
  assert.deepEqual(exec(body, model), { ok: true });
  assert.deepEqual(names(model.log).filter(x => x === 'UnMerge' || x === 'Merge'), ['UnMerge', 'Merge']);
  assert.deepEqual([...model.live], ['D1:D2']);
});

// ---- 4. the verification sees a merge that did not come back -------------------------------------------------------------------------------------------------

test('VERIFY: a merge that does not come back fails the write -- the merged TARGET through L1 (mergeChanged), the merged NEIGHBOUR through L3 (mergeChanged); an intact C5 passes', async () => {
  const cellsB = shapeCells('1x3', 'top');
  const okB = await run(cellsB, ['D4:F4'], 'D4:F4', ['top']);
  assert.equal(okB.result.ok, true);
  const dissolveB = await run(cellsB, ['D4:F4'], 'D4:F4', ['top'], c => ['D4', 'E4', 'F4'].includes(c.a) && ((c.merge = c.a), true));
  assert.equal(dissolveB.touched, 3);
  assert.equal(dissolveB.result.ok, false);
  assert.ok(dissolveB.l1.some(v => v.verdict === 'FAIL'), 'L1 fails the merged target');
  assert.deepEqual(dissolveB.l1[0].evidence.evidence.mergeChanged.map(x => [x.cell, x.was, x.now]), [['D4', 'D4:F4', 'D4'], ['E4', 'D4:F4', 'E4'], ['F4', 'D4:F4', 'F4']]);

  const cellsA = grid(['B', 'C', 'D', 'E'], [1, 2, 3, 4]); cellsA.D2 = { ...cellsA.D2, border: { b: thin } }; cellsA.D3 = { v: 'x', border: { t: thin } };
  const okA = await run(cellsA, ['D1:D2'], 'D3:D3', ['top']);
  assert.equal(okA.result.ok, true);
  const dissolveA = await run(cellsA, ['D1:D2'], 'D3:D3', ['top'], c => (c.a === 'D1' || c.a === 'D2') && ((c.merge = c.a), true));
  assert.equal(dissolveA.touched, 2);
  assert.equal(dissolveA.result.ok, false);
  assert.ok(dissolveA.l3.some(v => v.verdict === 'FAIL'), 'L3 fails the merged neighbour');
  assert.deepEqual(dissolveA.l3[0].evidence.evidence.mergeChanged.map(x => x.cell), ['D1', 'D2']);
});

// ---- 5. NEG control: a write with no merge in reach has the EXACT body it had before this change -------------------------------------------------------------

const GOLDEN = {
  d3top: '320d2c92a90668c001141c7b5e0801a2eea3484c45b245d706d4fe922059280d',
  four: '6d091ca9d57b4b5993b0e55ef8ec738db988896d33b49c62e568cbcddef52529',
  diag: '372c5f7ee9ea5458c22f56f23f29305c1873e66612fa36a541d11dc16bf0d0e4',
  ringcorner: 'f163c68a1cf649d7a5bb8ca28425d2150c271ab0e253ebcc84ef1a4a0943668d',
  leftbottom: 'f58c233bc8e3a3be0f7666e5e1517b0a8142a78f4e45a1a2c65201152fc637c9',
};
const thick = { s: 13, c: 65280 };
const SCEN = {
  d3top: () => { const c = grid(['C', 'D', 'E'], [2, 3, 4]); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin, b: thick }, fill: 0x00ff00, font: { bold: true }, align: { h: 1, v: 4 } }; return [c, [], 'D3:D3', ['top']]; },
  four: () => { const c = grid(['C', 'D', 'E'], [2, 3, 4, 5]); c.D3 = { v: 1, border: { t: thin, l: thin, r: thin, b: thin } }; c.D4 = { v: 2, border: { t: thin, b: thin, l: thin, r: thin } }; return [c, [], 'D3:D4', ['top', 'bottom', 'left', 'right']]; },
  diag: () => { const c = grid(['C', 'D', 'E'], [2, 3, 4]); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin, d: { s: 7, c: 0 }, dd: true } }; return [c, [], 'D3:D3', ['top']]; },
  ringcorner: () => { const c = grid(['C', 'D', 'E'], [2, 3, 4]); c.D2 = { ...blank, border: { b: thin } }; c.D3 = { v: 'x', border: { t: thin } }; c.C1 = { v: 'm' }; return [c, ['C1:C2'], 'D3:D3', ['top']]; },
  leftbottom: () => { const c = grid(['C', 'D', 'E'], [2, 3, 4, 5]); c.D3 = { v: 'x', border: { l: thin, b: thin } }; c.C3 = { ...blank, border: { r: thin } }; return [c, [], 'D3:D3', ['left', 'bottom']]; },
};

test('NEG control (no behaviour change): without a merge in reach the lowered body is BYTE-IDENTICAL to the pre-change body (sha256 of five scenarios, incl. a ring merge the composite never writes)', async () => {
  for (const [k, f] of Object.entries(SCEN)) {
    const r = await run(...f());
    assert.equal(r.result.ok, true, k);
    assert.ok(!r.body.includes('UnMerge') && !r.body.includes('finally'), `${k}: no C5 in the body`);
    assert.equal(crypto.createHash('sha256').update(r.body).digest('hex'), GOLDEN[k], `${k}: the body changed`);
  }
});
