'use strict';

/**
 * GH #96 follow-up (the reporter's LIVE file, 1.0.1-dbc780f6bae9): set_validation replace:true on a workbook with over
 * 5000 formulas refused xlsx-observation-unreadable-getters, failure.unsupported naming
 * formula-index / IMPLEMENTATION_LIMIT:formula-index-over-5000-formulas. The validation family asks for the dependency
 * readers only for the workbook's RULES (rules.bySheet, its L3); the tiered read also turned on the T3 formula index,
 * which a validation step never reads -- and its ceiling made the whole observation incomplete. A validation-only task
 * now opts out of the index (noFormulaIndex); every other dependents task (copy_range, names, ...) keeps it.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { readQuery, usedProbeQuery, digestQuery, tierOf, tieredPayload, formulaIndexWanted } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { parseRange, cellsOf, addr } = require('../../src/adapters/xlsx/a1.cjs');

// the in-memory workbook stand-in (the m7-s32b shape): 1000 rows x 10 columns, SIX formula columns = 6000 formulas
const empty = raw => raw === undefined || raw === '';
function bigFormulaBook() {
  const cells = new Map();
  for (let r = 1; r <= 1000; r += 1) {
    for (const c of ['A', 'B', 'C', 'D']) cells.set(`${c}${r}`, { raw: `${c}${r}` });
    for (const c of ['E', 'F', 'G', 'H', 'I', 'J']) cells.set(`${c}${r}`, { raw: `=${r}` });
  }
  cells.set('M3800', { raw: 'seed' }); // one far seed cell: used A1:M3800 = 49 400 cells, in the tiered band (45 000 .. 50 000]
  return { sheets: [{ name: 'S', cells }] };
}
function usedOf(s) {
  const gs = [...s.cells].filter(([, c]) => !empty(c.raw)).map(([a]) => parseRange(a));
  return `${addr(Math.min(...gs.map(g => g.c1)), Math.min(...gs.map(g => g.r1)))}:${addr(Math.max(...gs.map(g => g.c1)), Math.max(...gs.map(g => g.r1)))}`;
}
function apiOf(b) {
  const cellOf = (s, a) => {
    const c = s.cells.get(a) || { raw: undefined };
    const f = typeof c.raw === 'string' && c.raw.startsWith('=') ? c.raw : '';
    const v = f ? '7' : empty(c.raw) ? '' : String(c.raw);
    return { GetAddress: () => a, GetValue: () => v, GetFormula: () => f || v, MergeArea: { GetAddress: () => a }, _num: !!f,
      GetNumberFormat: () => 'General', GetFillColor: () => null, range: { getFont: () => ({ getBold: () => false }) } };
  };
  const rangeOf = (s, a1) => ({ GetAddress: () => a1, ForEach: fn => { for (const a of cellsOf(parseRange(a1))) fn(cellOf(s, a)); } });
  const ws = s => ({ GetName: () => s.name, GetIndex: () => b.sheets.indexOf(s), GetUsedRange: () => rangeOf(s, usedOf(s)), GetRange: a1 => rangeOf(s, a1) });
  return { GetSheets: () => b.sheets.map(ws), GetSheet: n => { const s = b.sheets.find(x => x.name === n); return s ? ws(s) : null; }, WorksheetFunction: { COUNT: cell => (cell._num ? 1 : 0) } };
}
const run = (body, b) => new Function('Api', body)(apiOf(b));

/** readLive's tiered branch, with the driver's OWN index decision (formulaIndexWanted) and the readers' measured shapes. */
function observe(b, request) {
  const tier = tierOf(run(usedProbeQuery(), b).sheets, request);
  assert.equal(tier.tier, 'tiered', 'the stand-in is above the 45000-cell bound');
  const p = run(readQuery(request.windows, { usedRanges: false }), b);
  const payload = tieredPayload(p, run(digestQuery(request.windows, { formulaIndex: formulaIndexWanted(request) }), b), tier, { build: 'test', fileId: null });
  const deps = request.dependents ? { dependents: { names: [], validations: { S: [] } }, conditional: { S: [] }, objects: { S: { charts: [], pivots: [], names: [] } } } : {};
  return normalizeInventory({ ...payload, ...deps, objects: deps.objects || payload.objects });
}

const validationTask = { id: 't', steps: [{ intent: 'set_validation', sheet: 'S', range: 'K18:K20', type: 'list', items: ['a', 'b'], alertStyle: 'stop', replace: true }] };
const copyTask = { id: 't', steps: [{ intent: 'copy_range', id: 'c', sheet: 'S', range: 'A1:B5', target: 'L1' }] };
const nameTask = { id: 't', steps: [{ intent: 'set_defined_name', id: 'n', name: 'Zz', refersTo: '=S!$A$1' }] };
// copy_range must be its task's only step, so the mixed case pairs a names step (also dependents) with the validation
const both = { id: 't', steps: [...nameTask.steps, { ...validationTask.steps[0], id: 'v' }] };

test('GH #96 f/u: a validation-only task asks for the dependency readers WITHOUT the formula index', () => {
  const r = liveScope(validationTask);
  assert.equal(r.dependents, true, 'the rules are still read (the L3)');
  assert.equal(r.noFormulaIndex, true);
  assert.equal(formulaIndexWanted(r), false);
});

test('GH #96 f/u: every other dependents task keeps the index -- alone, and mixed with a validation step', () => {
  for (const t of [copyTask, nameTask, both]) {
    const r = liveScope(t);
    assert.equal('noFormulaIndex' in r, false, JSON.stringify(t.steps.map(s => s.intent)));
    assert.equal(formulaIndexWanted(r), true);
  }
  assert.equal(formulaIndexWanted({}), false, 'no dependents: no index (unchanged)');
});

test('GH #96 f/u: over 5000 formulas a validation observation is COMPLETE; a copy observation still names the limit', () => {
  const b = bigFormulaBook();
  const v = observe(b, liveScope(validationTask));
  assert.equal(v.complete, true, JSON.stringify(v.unsupported));
  assert.equal(v.formulas, null, 'no index was read');
  assert.equal(v.rules.complete, true);
  const c = observe(b, liveScope(copyTask));
  assert.equal(c.complete, false);
  assert.match(JSON.stringify(c.unsupported), /IMPLEMENTATION_LIMIT:formula-index-over-5000-formulas/);
});
