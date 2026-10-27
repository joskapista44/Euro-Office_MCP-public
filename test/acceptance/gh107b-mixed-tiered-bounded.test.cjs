'use strict';

/**
 * GH #107 (b) (internal tracker #107b/2; design docs/GH107B-MIXED-TIERED-BOUNDED-DESIGN-20261014.md, accepted by a maintainer):
 * a task mixing a TIERED_INTENTS step with a step that only has a bounded variant (format_range, layout_range) is
 * observed above 5000 used cells through the merged boundedVariant. The two-direction condition per step: it reports the
 * scope its own family reports in its single-family task (format_range BOUNDED, the tiered families FULL / their own),
 * and the variant carries every extra the step's own proof rests on.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { tierOf, TIERED_EXTRAS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { liveScope, xlsxAdapter, planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const formatFam = require('../../src/adapters/xlsx/families/format.cjs');
const validationFam = require('../../src/adapters/xlsx/families/validation.cjs');

// 48 000 used cells: above the 45 000 full-read bound, below the 50 000 tiered ceiling
const sheets = [{ name: 'S', used: 'A1:AG1500' }];
const META = ['windows', 'tiered', 'windowStyle', 'boundedVariant', 'noFormulaIndex', 'mixedTask', 'readOnly'];

const fmt = { intent: 'format_range', id: 'f1', sheet: 'S', range: 'C18:C67', format: { fill: '#FFFF00' } };
const val = { intent: 'set_validation', id: 'v1', sheet: 'S', range: 'D18:D67', type: 'list', items: ['a', 'b'], alertStyle: 'stop', replace: true };
const wr = { intent: 'write_range', id: 'w1', sheet: 'S', range: 'L1', values: [['x']] };
const fr = { intent: 'filter_range', id: 'fr1', sheet: 'S', range: 'A1:B2', field: 1, value: 'x' };

// one valid step per TIERED_INTENTS member (the ids are unique per task, ranges away from fmt's C18:C67 ring)
const TIERED_STEPS = {
  write_range: wr,
  clear_range: { intent: 'clear_range', id: 'c1', sheet: 'S', range: 'A991:J1000' },
  copy_range: { intent: 'copy_range', id: 'cp1', sheet: 'S', range: 'A1:B2', target: 'N1' },
  set_defined_name: { intent: 'set_defined_name', id: 'n1', name: 'X_NAME', refersTo: '=S!$A$1' },
  delete_defined_name: { intent: 'delete_defined_name', id: 'dn1', name: 'X_NAME' },
  merge_range: { intent: 'merge_range', id: 'm1', sheet: 'S', range: 'P1:Q2' },
  unmerge_range: { intent: 'unmerge_range', id: 'um1', sheet: 'S', range: 'P1:Q2' },
  sort_range: { intent: 'sort_range', id: 's1', sheet: 'S', range: 'A1:B9', keyColumn: 'A', order: 'asc', hasHeaders: false },
  clear_filter: { intent: 'clear_filter', id: 'cf1', sheet: 'S', range: 'A1:B2' },
  set_print_area: { intent: 'set_print_area', id: 'pa1', sheet: 'S', range: 'A1:B9' },
  clear_print_area: { intent: 'clear_print_area', id: 'cpa1', sheet: 'S' },
  format_run: { intent: 'format_run', id: 'r1', sheet: 'S', cell: 'H1', runs: [{ start: 1, length: 1, font: { bold: true } }] },
  freeze_panes: { intent: 'freeze_panes', id: 'fz1', sheet: 'S', rows: 1 },
  unfreeze_panes: { intent: 'unfreeze_panes', id: 'ufz1', sheet: 'S' },
  add_conditional_format: { intent: 'add_conditional_format', id: 'cfa1', sheet: 'S', range: 'F1:F9', rule: { type: 'cellValue', operator: 'greater', formula1: '5' } },
  delete_conditional_format: { intent: 'delete_conditional_format', id: 'cfd1', sheet: 'S', range: 'F1:F9', rule: { type: 'cellValue', operator: 'greater', formula1: '5' } },
  set_validation: val,
  clear_validation: { intent: 'clear_validation', id: 'cv1', sheet: 'S', range: 'D18:D67' },
};

const scopeOf = steps => {
  assert.equal(xlsxAdapter.planOperations({ id: 't', steps }).ok, true, `the task plans: ${JSON.stringify(xlsxAdapter.planOperations({ id: 't', steps }))}`);
  return liveScope({ id: 't', steps });
};
const tier = steps => tierOf(sheets, scopeOf(steps));
const variantOf = steps => scopeOf(steps).boundedVariant;

test('T2 (guard): EVERY TIERED_INTENTS member has a fixture here, and its liveExtras keys are all in the driver TIERED_EXTRAS (the precondition of the merge)', () => {
  const adapterSrc = require('node:fs').readFileSync(require('node:path').join(__dirname, '../../src/adapters/xlsx/adapter.cjs'), 'utf8');
  const listed = /const TIERED_INTENTS = Object\.freeze\(\[([^\]]*)\]\)/.exec(adapterSrc)[1].split(',').map(x => x.trim().replace(/'/g, '')).filter(Boolean);
  assert.deepEqual(Object.keys(TIERED_STEPS).sort(), [...listed].sort(), 'a new tiered intent needs a fixture here');
  for (const [intent, step] of Object.entries(TIERED_STEPS)) {
    const planned = planOperations({ id: 't', steps: [step] });
    assert.equal(planned.ok, true, `${intent} plans`);
  }
});

test('T1: format_range + each tiered step with extras -> tiered + bounded; the variant is a subset of TIERED_EXTRAS and carries the step\'s own extras', () => {
  const expect = {
    set_validation: ['validationRanges', 'dependents', 'noFormulaIndex'],
    clear_validation: ['validationRanges', 'dependents'],
    add_conditional_format: ['cfRanges'],
    delete_conditional_format: ['cfRanges'],
    format_run: ['runCells'],
    freeze_panes: ['panes', 'sheetVisibility'],
    unfreeze_panes: ['panes', 'sheetVisibility'],
    set_defined_name: ['dependents'],
    delete_defined_name: ['dependents'],
    set_print_area: ['dependents'],
    sort_range: [], merge_range: [], unmerge_range: [], clear_filter: [], clear_range: [], write_range: [], clear_print_area: ['dependents'],
    // copy_range is not here: it must be the ONLY step of its task (planOperations), so it can never be part of a mix
  };
  for (const [intent, keys] of Object.entries(expect)) {
    const steps = [fmt, TIERED_STEPS[intent]];
    const t = tier(steps);
    assert.deepEqual([t.tier, t.bounded], ['tiered', true], `${intent}: ${t.reason}`);
    const v = variantOf(steps);
    assert.equal(v.tiered, true);
    assert.equal(v.windowStyle, true);
    assert.deepEqual(v.formatSheets.includes('S'), true, `${intent}: format_range's number-format digest sheet is carried`);
    const extraKeys = Object.keys(v).filter(k => !META.includes(k));
    assert.deepEqual(extraKeys.filter(k => !TIERED_EXTRAS.includes(k)), [], `${intent}: only TIERED_EXTRAS keys`);
    for (const k of keys) assert.ok(v[k], `${intent}: the variant carries ${k}`);
    assert.equal(v.styleSheets, undefined, `${intent}: the sheet-wide style read is NOT asked`);
  }
});

test('T5: noFormulaIndex in the merged variant is computed exactly as in the pure tiered task (same helper)', () => {
  assert.equal(variantOf([fmt, val]).noFormulaIndex, true);
  assert.equal(scopeOf([val]).noFormulaIndex, true);
  assert.equal(variantOf([fmt, TIERED_STEPS.set_defined_name]).noFormulaIndex, undefined);
  assert.equal(scopeOf([TIERED_STEPS.set_defined_name]).noFormulaIndex, undefined);
});

test('pure tasks are UNCHANGED: a pure tiered task has no variant; a format-only task keeps its old variant', () => {
  assert.equal(scopeOf([val]).boundedVariant, undefined);
  const v = variantOf([fmt]);
  // GH #326 (U4b): format_range's own protection check is sheet-level only (no cell-level unlocked list),
  // so boundedVariant now always carries protectionSheets too -- the SAME cross-cutting axis liveScope adds.
  assert.deepEqual(v, { windows: [{ sheet: 'S', range: 'B17:D68', style: true }], formatSheets: ['S'], tiered: true, windowStyle: true, protectionSheets: ['S'] });
});

test('T7 NEG: the windows cap applies to the SUM -- a merged task whose windows together exceed 45 000 cells is refused', () => {
  const f = (id, range) => ({ intent: 'format_range', id, sheet: 'S', range, format: { fill: '#FFFF00' } });
  const v2 = { ...val, range: 'AG300:AG310' };
  // format_range has a 2000-cell step ceiling, so the sum is built from many steps: each B..AE x 63 rows (1890 cells) has a
  // 32 x 65 = 2080-cell ring; 22 of them = 45 760 > 45 000, while the 4-row gaps keep the rings apart
  const ring = n => Array.from({ length: n }, (_, k) => f(`f${k}`, `B${k * 67 + 2}:AE${k * 67 + 64}`));
  const t = tier([...ring(22), v2]);
  assert.equal(t.tier, 'refused');
  assert.match(t.reason, /task-windows-over-45000-cells/);
  assert.deepEqual(tier([...ring(10), v2]).tier, 'tiered', 'POS control: ten of them fit');
});

test('NEG: a step with NEITHER path (filter_range) still refuses the mix, with the (a) note', () => {
  const t = tier([fmt, fr]);
  assert.equal(t.tier, 'refused');
  assert.match(t.reason, /not tiered: f1 \(format_range\), fr1 \(filter_range\); no bounded variant: fr1 \(filter_range\)/);
  assert.equal(tier([wr, fr]).tier, 'refused');
  assert.equal(tier([val, fr]).tier, 'refused');
});

// ---- T3: per-step scope in a MERGED observation (hand-built, same technique as gh86b / m7-s32) --------------------
const opOf = (step) => planOperations({ id: 't', steps: [step] }).operations[0];
const mergedObs = rules => ({ primary: { format: 'xlsx', complete: true, sheets: [{ name: 'S', used: 'A1:J1000' }, { name: 'Other', used: 'A1' }], tier: 'tiered', bounded: true, digests: [],
  cells: {}, ...(rules ? { rules: { bySheet: rules } } : {}) } });

test('T3: in a merged (tiered + bounded) observation the validation step reports FULL and the format step BOUNDED -- no step claims more than alone', () => {
  const rules = o => ({ S: [], Other: o });
  const keep = [{ kind: 'validation', range: 'B1', formula1: '1', formula2: '2' }];
  const rv = validationFam.verifyPreservation(opOf(val), mergedObs(rules(keep)), mergedObs(rules(keep)));
  assert.equal(rv.ok, true, JSON.stringify(rv.evidence));
  assert.equal(rv.evidence.preservationScope, 'FULL');
  const rf = formatFam.verifyPreservation(opOf(fmt), mergedObs(null), mergedObs(null));
  assert.equal(rf.evidence.preservationScope, 'BOUNDED');
  assert.deepEqual(rf.evidence.notProven, ['style-outside-target-and-ring']);
});

test('T3 NEG (drift kept): an out-of-window validation rule that changes -> the validation step FAILS even in the merged observation', () => {
  const rules = o => ({ S: [], Other: o });
  const rv = validationFam.verifyPreservation(opOf(val), mergedObs(rules([{ kind: 'validation', range: 'B1', formula1: '1', formula2: '2' }])), mergedObs(rules([{ kind: 'validation', range: 'B1', formula1: '1', formula2: '99' }])));
  assert.equal(rv.ok, false);
});

test('T4 key mutant: WITHOUT the dependents read (rules.bySheet absent) the validation step FAILS closed -- it never reports FULL on missing data', () => {
  const rv = validationFam.verifyPreservation(opOf(val), mergedObs(null), mergedObs(null));
  assert.equal(rv.ok, false, JSON.stringify(rv.evidence));
});

test('T4 (variant side): dropping `dependents` from the merged variant would be caught -- the variant of [format_range, set_validation] asks for it', () => {
  assert.equal(variantOf([fmt, val]).dependents, true);
  assert.ok(variantOf([fmt, val]).validationRanges.length >= 1);
});

// ---- T6: the window-union property (design 2.3, accepted by a maintainer): tiered steps' L3 is NOT widened -----------------
test('T6: a format_range window (style:true) next to a tiered step: the tiered step does not claim style preservation (its scope fields are unchanged), the format step keeps its own ring style read', () => {
  const v = variantOf([fmt, val]);
  assert.deepEqual(v.windows, [{ sheet: 'S', range: 'B17:D68', style: true }]);
  const rv = validationFam.verifyPreservation(opOf(val), mergedObs({ S: [], Other: [] }), mergedObs({ S: [], Other: [] }));
  assert.deepEqual(Object.keys(rv.evidence).filter(k => /style/i.test(k)), [], 'no style claim from the tiered family');
});
