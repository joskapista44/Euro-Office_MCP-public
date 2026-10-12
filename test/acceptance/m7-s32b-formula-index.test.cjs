'use strict';

// M7 S32 slice 2 (owner GO 2026-10-03, owner channel; plan docs/M7-S32B-FORMULA-INDEX-PLAN.md): the T3 formula index
// above 5000 used cells, the dependency analysis on it, and copy_range / set_defined_name / delete_defined_name on the
// tiered observation. The editor bodies (the probe, the T1 window read, the T2 digest + T3 index) run against an
// in-memory spreadsheet stand-in; the dependency READERS' output (names, validations, conditional formats, objects,
// AutoFilter) is supplied in its measured shape. The LIVE harness proves the editor behaviour.

const test = require('node:test');
const assert = require('node:assert/strict');
const { readQuery, usedProbeQuery, digestQuery, tierOf, tieredPayload, withLimitNote } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { parseRange, cellsOf, addr } = require('../../src/adapters/xlsx/a1.cjs');
const { translateFormula } = require('../../src/adapters/xlsx/shift.cjs');

// ---- the in-memory workbook (the slice-1 stand-in, plus defined names) -------------------------------------------

const empty = raw => raw === undefined || raw === '';
function book(spec, names = []) {
  return { names: names.map(n => ({ ...n })), sheets: Object.entries(spec).map(([name, entries]) => ({ name, cells: new Map(Object.entries(entries).map(([a, raw]) => [a, { raw, bold: false, nf: 'General', fill: null }])) })) };
}
const set = (b, sheet, a, raw) => { const s = b.sheets.find(x => x.name === sheet); s.cells.set(a, { ...(s.cells.get(a) || { bold: false, nf: 'General', fill: null }), raw }); return b; };
function usedOf(s) {
  const gs = [...s.cells].filter(([, c]) => !empty(c.raw)).map(([a]) => parseRange(a));
  if (!gs.length) return null;
  return `${addr(Math.min(...gs.map(g => g.c1)), Math.min(...gs.map(g => g.r1)))}:${addr(Math.max(...gs.map(g => g.c1)), Math.max(...gs.map(g => g.r1)))}`;
}
function apiOf(b) {
  const cellOf = (s, a) => {
    const c = s.cells.get(a) || { raw: undefined, bold: false, nf: 'General', fill: null };
    const f = typeof c.raw === 'string' && c.raw.startsWith('=') ? c.raw : '';
    const v = f ? '7' : empty(c.raw) ? '' : String(c.raw);
    return { GetAddress: () => a, GetValue: () => v, GetFormula: () => f || v, MergeArea: { GetAddress: () => a }, _num: typeof c.raw === 'number' || !!f,
      GetNumberFormat: () => c.nf, GetFillColor: () => null, range: { getFont: () => ({ getBold: () => c.bold }) } };
  };
  const rangeOf = (s, a1) => ({ GetAddress: () => a1, ForEach: fn => { for (const a of cellsOf(parseRange(a1))) fn(cellOf(s, a)); } });
  const ws = s => ({ GetName: () => s.name, GetIndex: () => b.sheets.indexOf(s), GetUsedRange: () => (usedOf(s) ? rangeOf(s, usedOf(s)) : null), GetRange: a1 => rangeOf(s, a1) });
  return { GetSheets: () => b.sheets.map(ws), GetSheet: n => { const s = b.sheets.find(x => x.name === n); return s ? ws(s) : null; }, WorksheetFunction: { COUNT: cell => (cell._num ? 1 : 0) } };
}
const run = (body, b) => new Function('Api', body)(apiOf(b));

/** The dependency readers' output in the measured raw shapes (W3d/W3e): names, validations, conditional, objects, filters. */
function readers(b, request) {
  const sheets = b.sheets.map(s => s.name);
  return {
    dependents: { names: b.names.map(n => ({ name: n.name, ref: n.ref, scopeSheet: n.scopeSheet || null })), validations: Object.fromEntries(sheets.map(n => [n, []])) },
    conditional: Object.fromEntries(sheets.map(n => [n, []])),
    objects: Object.fromEntries(sheets.map(n => [n, { charts: [], pivots: [], names: [] }])),
    filterReads: Object.fromEntries((request.filterSheets || []).map(n => [n, { range: null }])),
  };
}

/** What readLive returns: the probe, the tier decision, the matching read; T3 asked exactly when the readers are. */
function observe(b, request) {
  const probe = run(usedProbeQuery(), b);
  const tier = tierOf(probe.sheets, request);
  const deps = request.dependents ? readers(b, request) : {};
  if (tier.tier === 'tiered') {
    const p = run(readQuery(request.windows, { usedRanges: false }), b);
    const payload = tieredPayload(p, run(digestQuery(request.windows, { formulaIndex: !!request.dependents }), b), tier, { build: 'test', fileId: null });
    return { tier, payload: { ...payload, ...deps, objects: deps.objects || payload.objects } };
  }
  const p = run(readQuery(request.windows, {}), b);
  const payload = { ...p, ...deps, objects: deps.objects || p.objects };
  if (tier.tier === 'refused' && p.truncated) payload.unsupported = withLimitNote(p.unsupported, tier);
  return { tier, payload };
}

/** The driver stand-in: reads through observe(); dispatch applies the command's structured meta (as the fake driver does). */
function standIn(b, { inject = null } = {}) {
  const calls = { dispatch: 0 };
  const tiers = [];
  let confirmedAt = null;
  const apply = m => {
    if (m.family === 'range') {
      const src = b.sheets.find(s => s.name === m.sheet);
      const dst = b.sheets.find(s => s.name === m.targetSheet);
      for (const { from, to } of m.pairs) {
        const c = src.cells.get(from);
        const raw = c && typeof c.raw === 'string' && c.raw.startsWith('=') ? translateFormula(c.raw, m.delta.dr, m.delta.dc).formula : (c ? c.raw : undefined);
        set(b, dst.name, to, raw);
      }
    }
    if (m.family === 'names') {
      if (m.action === 'set_defined_name') b.names.push({ name: m.name, ref: m.refersTo });
      else b.names = b.names.filter(n => n.name.toLowerCase() !== m.name.toLowerCase());
    }
  };
  return {
    calls, tiers,
    async open() { return { ok: true, id: 'stand-in' }; },
    async close() { return { ok: true, confirmation: confirmedAt === null ? null : (calls.dispatch > confirmedAt ? { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' } : { state: 'confirmed', reason: 'forcesave-confirmed' }) }; },
    async forceSave() { confirmedAt = calls.dispatch; return { ok: true, confirmed: true, forceSave: { state: 'confirmed', reason: 'forcesave-confirmed' } }; },
    async canSave() { return { ok: true, canSave: false }; },
    async readPackage() { return { ok: true, parts: {} }; },
    async readLive(_h, request = {}) { const o = observe(b, request); tiers.push(o.tier.tier); return { ok: true, authority: 'LIVE_READ', payload: o.payload }; },
    async dispatch(_h, command) { calls.dispatch += 1; apply(command.meta); if (inject) inject(calls.dispatch, b); return { ok: true, dispatched: true }; },
  };
}

async function session(b, task, opts) {
  const driver = standIn(b, opts);
  const options = { driver, fileId: 'stand-in', ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(task, { observation: { live: liveScope(task) } }));
  return { result, driver };
}

/** 1000 rows x 10 columns = 10 000 cells; column E holds 1000 formulas (=B+C). */
function bigBook({ extra = {}, names = [] } = {}) {
  const spec = {};
  for (let r = 1; r <= 1000; r += 1) Object.assign(spec, { [`A${r}`]: `T${r}`, [`B${r}`]: r, [`C${r}`]: r + 0.5, [`D${r}`]: `D${r}`, [`E${r}`]: `=B${r}+C${r}`, [`F${r}`]: `F${r}`, [`G${r}`]: 7 * r, [`H${r}`]: `H${r}`, [`I${r}`]: 3 * r, [`J${r}`]: `J${r}` });
  return book({ S: { ...spec, ...extra } }, names);
}

const copyTask = (range, target) => ({ id: 't', steps: [{ intent: 'copy_range', id: 'c', sheet: 'S', range, target }] });
const nameTask = (intent, name, refersTo) => ({ id: 't', steps: [{ intent, id: 'n', name, ...(refersTo ? { refersTo } : {}) }] });
const reasonOf = result => JSON.stringify(result.execution.appliedOperations[0].evidence);
const op = result => result.execution.appliedOperations[0];

// ---- T3 itself -----------------------------------------------------------------------------------------------------

test('T3: every formula cell OUTSIDE the windows, with its text; window formulas and non-formulas are not in it', () => {
  const b = book({ S: { A1: 1, B1: '=A1*2', C1: 'x', B2: '=SUM(A1:A3)', D9: '=B1' } });
  const g = run(digestQuery([{ sheet: 'S', range: 'D9' }], { formulaIndex: true }), b);
  assert.deepEqual(g.formulas, [['S', 'B1', '=A1*2'], ['S', 'B2', '=SUM(A1:A3)']]);
  assert.equal(g.formulaIndexOver, false);
  assert.equal(run(digestQuery([{ sheet: 'S', range: 'D9' }]), b).formulas, undefined, 'not asked: no index, the slice-1 body');
});

test('T3 ceiling: over 5000 formulas -> the index is cut AND marked over; the observation is incomplete (IMPLEMENTATION_LIMIT)', () => {
  const spec = {};
  for (let r = 1; r <= 1000; r += 1) for (const c of ['A', 'B', 'C', 'D', 'E', 'F']) spec[`${c}${r}`] = `=${r}`;
  const b = book({ S: spec });
  const o = observe(b, { windows: [{ sheet: 'S', range: 'Z1' }], tiered: true, dependents: true });
  assert.equal(o.tier.tier, 'tiered');
  const p = normalizeInventory(o.payload);
  assert.equal(p.complete, false);
  assert.match(JSON.stringify(p.unsupported), /IMPLEMENTATION_LIMIT:formula-index-over-5000-formulas/);
});

test('the inventory: T3 -> `formulas`; the dependency analysis sees a formula that lives ONLY in T3 (an INDIRECT -> incomplete)', () => {
  const clean = normalizeInventory(observe(bigBook(), { windows: [{ sheet: 'S', range: 'L1:L2' }], tiered: true, dependents: true }).payload);
  assert.equal(clean.tier, 'tiered');
  assert.equal(Object.keys(clean.formulas).length, 1000);
  assert.equal(clean.formulas['S\u0000E7'], '=B7+C7');
  assert.equal(clean.dependencies.complete, true);
  const dirty = normalizeInventory(observe(bigBook({ extra: { E500: '=INDIRECT("A1")' } }), { windows: [{ sheet: 'S', range: 'L1:L2' }], tiered: true, dependents: true }).payload);
  assert.equal(dirty.dependencies.complete, false, 'an unresolvable formula outside the windows is a gap, never "not affected"');
  assert.match(JSON.stringify(dirty.dependencies.unobservable), /indirect/);
});

// ---- copy_range above 5000 -------------------------------------------------------------------------------------------

test('E2E copy 10k: tiered; APPLIED, digest L3 PASS, FORCESAVE_CONFIRMED; the reopened retry SATISFIED 0 / 0', async () => {
  const b = bigBook();
  const t = copyTask('A1:B5', 'L1');
  const { result, driver } = await session(b, t);
  assert.equal(result.ok, true, reasonOf(result));
  assert.deepEqual([...new Set(driver.tiers)], ['tiered']);
  assert.equal(result.persistenceState, 'FORCESAVE_CONFIRMED');
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(b.sheets[0].cells.get('M3').raw, 3);
  const again = await session(b, t);
  assert.deepEqual([again.result.metrics.writes, again.result.metrics.commands], [0, 0]);
  assert.equal(op(again.result).status, 'SATISFIED');
});

test('copy 10k refused BEFORE the write, 0 dispatch: a defined name reaching the source block (the R8 guard on the tiered read)', async () => {
  const b = bigBook({ names: [{ name: 'Src', ref: 'S!$A$3' }] });
  const { result, driver } = await session(b, copyTask('A1:B5', 'L1'));
  assert.equal(driver.calls.dispatch, 0);
  assert.match(reasonOf(result), /xlsx-range-dependent-affected/);
});

test('copy 10k refused, 0 dispatch: an unresolvable formula OUTSIDE the windows (visible only through T3)', async () => {
  const b = bigBook({ extra: { E500: '=INDIRECT("A1")' } });
  const { result, driver } = await session(b, copyTask('A1:B5', 'L1'));
  assert.equal(driver.calls.dispatch, 0);
  assert.match(reasonOf(result), /xlsx-range-dependency-observation-incomplete/);
});

test('copy 10k L3 negative: a foreign change OUTSIDE the windows after the dispatch -> not APPLIED (the digest)', async () => {
  const b = bigBook();
  const { result } = await session(b, copyTask('A1:B5', 'L1'), { inject: (n, bk) => { if (n === 1) set(bk, 'S', 'G700', -1); } });
  assert.equal(result.ok, false);
  assert.notEqual(op(result).status, 'APPLIED');
});

test('copy 10k: a formula count over the T3 ceiling -> refused before the write, the IMPLEMENTATION_LIMIT named', async () => {
  const extra = {};
  for (let r = 1; r <= 1000; r += 1) for (const c of ['K', 'M', 'N', 'O', 'P']) extra[`${c}${r}`] = `=B${r}`;
  const b = bigBook({ extra });
  const { result, driver } = await session(b, copyTask('A1:B5', 'Z1'));
  assert.equal(driver.calls.dispatch, 0);
  assert.match(reasonOf(result), /IMPLEMENTATION_LIMIT:formula-index-over-5000-formulas/);
});

test('move_range above 5000 stays refused before the write (not migrated: its Cut re-points formulas outside the windows)', async () => {
  const b = bigBook();
  const task = { id: 't', steps: [{ intent: 'move_range', id: 'm', sheet: 'S', range: 'A1:A2', target: 'L1', expect: { version: 1 } }] };
  const { result, driver } = await session(b, task);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.ok, false);
});

// ---- names above 5000 ------------------------------------------------------------------------------------------------

test('E2E set_defined_name 10k: tiered; APPLIED; the reopened retry SATISFIED 0 / 0', async () => {
  const b = bigBook();
  const t = nameTask('set_defined_name', 'Total', '=S!$B$1');
  const { result, driver } = await session(b, t);
  assert.equal(result.ok, true, reasonOf(result));
  assert.deepEqual([...new Set(driver.tiers)], ['tiered']);
  assert.deepEqual(b.names, [{ name: 'Total', ref: 'S!$B$1' }]);
  const again = await session(b, t);
  assert.deepEqual([again.result.metrics.writes, again.result.metrics.commands], [0, 0]);
});

test('delete_defined_name 10k refused, 0 dispatch: a formula that uses the name lives ONLY in T3 (outside every window)', async () => {
  const b = bigBook({ extra: { E900: '=Total*2' }, names: [{ name: 'Total', ref: 'S!$B$1' }] });
  const { result, driver } = await session(b, nameTask('delete_defined_name', 'Total'));
  assert.equal(driver.calls.dispatch, 0);
  assert.match(reasonOf(result), /S!E900/);
  assert.equal(b.names.length, 1, 'kept');
});

test('delete_defined_name 10k of an unused name: APPLIED; retry SATISFIED 0 / 0', async () => {
  const b = bigBook({ names: [{ name: 'Unused', ref: 'S!$B$1' }] });
  const t = nameTask('delete_defined_name', 'Unused');
  const { result } = await session(b, t);
  assert.equal(result.ok, true, reasonOf(result));
  assert.equal(b.names.length, 0);
  const again = await session(b, t);
  assert.deepEqual([again.result.metrics.writes, again.result.metrics.commands], [0, 0]);
});

test('below 5000 the same tasks keep the FULL tier (no index, no digest) -- unchanged behaviour', async () => {
  const spec = {};
  for (let r = 1; r <= 100; r += 1) Object.assign(spec, { [`A${r}`]: `T${r}`, [`B${r}`]: r, [`E${r}`]: `=B${r}*2` });
  const b = book({ S: spec });
  const { result, driver } = await session(b, copyTask('A1:B5', 'L1'));
  assert.equal(result.ok, true, reasonOf(result));
  assert.deepEqual([...new Set(driver.tiers)], ['full']);
});

test('a malformed T3 (a row without formula text, a non-list) is a gap: the observation is incomplete, never trusted', () => {
  const o = observe(bigBook(), { windows: [{ sheet: 'S', range: 'L1' }], tiered: true, dependents: true });
  for (const bad of [[['S', 'E1']], [['S', 'E1', 'B1+C1']], 'UNOBSERVABLE:x']) {
    const p = normalizeInventory({ ...o.payload, formulaIndex: bad });
    assert.equal(p.complete, false, JSON.stringify(bad));
    assert.match(JSON.stringify(p.unsupported), /formula-index.*malformed/);
  }
  assert.equal(normalizeInventory(o.payload).complete, true, 'positive control: the well-formed index');
});
