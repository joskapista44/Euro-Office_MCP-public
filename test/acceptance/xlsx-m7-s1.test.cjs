'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const range = require('../../src/adapters/xlsx/families/range.cjs');
const { referencesBlock } = require('../../src/adapters/xlsx/shift.cjs');
const { analyze } = require('../../src/adapters/xlsx/dependencies.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 S1 (owner GO 2026-09-30, owner channel): copy_range / move_range refuse before the write when a defined name, a
 * validation, a conditional format, a chart series, a pivot or an AutoFilter reaches either block (directly, from
 * another sheet, through a name) or when any of those cannot be observed; L3 keeps names / rules / objects / the
 * AutoFilter unchanged. An xlCellValue conditional format is observable (type, operator, range, both formulas),
 * other types / operators / a partial read stay UNOBSERVABLE. Print_Area / Print_Titles / _xlnm.* are reserved.
 */

const book = () => [
  { name: 'P', cells: { A1: { v: 'a' }, B1: { v: 1 }, H1: { v: 'm' }, H2: { v: 3 } } },
  { name: 'Q', cells: { A1: { v: 'keep' } } },
];
const COPY = { id: 'm7s1', steps: [{ intent: 'copy_range', sheet: 'P', range: 'A1:B1', target: 'D5' }] };
// L-06/b: the caller no longer supplies `expect` -- classify() builds it from book()'s own H1:H2 content.
const MOVE = { id: 'm7s1', steps: [{ intent: 'move_range', sheet: 'P', range: 'H1:H2', target: 'J5' }] };

async function run(store, t, script = {}) {
  let n = 0;
  // M11.5c.4 / DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): copy_sheet writes only through an
  // ACCEPTED Bridge WRITE runtime -- a copy task runs on the accepted 9.3.5.0/16 identity unless the test names a runtime.
  const driver = createFakeXlsxDriver(store, t.steps.some(x => x.intent === 'copy_sheet') && script.runtime === undefined ? { ...script, runtime: '9.3.5|euro-office-bridge-v1/9.3.5.0/16' } : script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics, driver };
}
const op0 = result => result.execution.appliedOperations[0];
const reasonOf = result => { const o = op0(result); return o && o.evidence && o.evidence.evidence ? o.evidence.evidence.reason : JSON.stringify(result.execution); };
const refused = async (store, t, want, script = {}) => {
  const { result, m } = await run(store, t, script);
  assert.equal(result.ok, false);
  assert.equal(reasonOf(result), want, JSON.stringify(op0(result)));
  assert.equal(m.commands, 0, 'refused before the write: 0 command');
  assert.equal(result.metrics.writes || 0, 0);
  return result;
};

// ---- the read scope: copy / move ask for the dependency readers and both sheets' AutoFilter ----

test('liveScope: copy / move ask for the dependents and the AutoFilter of both sheets', () => {
  const s = liveScope({ id: 'x', steps: [{ intent: 'copy_range', sheet: 'P', range: 'A1', target: 'C3', targetSheet: 'Q' }] });
  assert.equal(s.dependents, true);
  assert.deepEqual(s.filterSheets, ['P', 'Q']);
});

// ---- affected dependents: refused 0 write / 0 command ----

const AFFECTED = [
  // M7 E R7: a name ENTIRELY inside the move source now moves with it (xlsx-m7-e-move.test.cjs); one PARTLY over it refuses
  ['a workbook name partly over the MOVE source', () => ({ names: [{ name: 'N1', ref: 'P!$H$1:$H$3' }] }), MOVE],
  ['a workbook name into the COPY target', () => ({ names: [{ name: 'N2', ref: 'P!$D$5:$E$6' }] }), COPY],
  ['a sheet-scoped name on ANOTHER sheet into the copy source', b => { b[1].objects = { charts: [], pivots: [], names: [{ name: 'QN', refersTo: '=P!$A$1' }] }; return {}; }, COPY],
  ['a validation over the copy target (applies-to)', b => { b[0].rules = [{ kind: 'validation', type: 1, range: 'E5', formula: '1', formula2: '10' }]; return {}; }, COPY],
  ['a conditional format on ANOTHER sheet whose formula reads the move source', b => { b[1].rules = [{ kind: 'conditional', range: 'A1', formula: '=P!$H$2>0' }]; return {}; }, MOVE],
  // J1 itself is in neither block; repeated over A1:A10 the relative J1 reads J1..J10, which reaches the target J5:J6
  ['a RELATIVE rule formula whose repetition over its area reaches the block', b => { b[1].rules = [{ kind: 'conditional', range: 'A1:A10', formula: '=P!J1>0' }]; return {}; }, MOVE],
  ['a chart series over the move source', b => { b[1].objects = { charts: [{ type: 'bar', series: [{ name: '', values: '=P!$H$1:$H$2', cat: '', x: '' }] }], pivots: [], names: [] }; return {}; }, MOVE],
  ['a pivot whose source reaches the copy source', b => { b[1].objects = { charts: [], pivots: [{ name: 'PV', source: 'P!$A$1:$B$3' }], names: [] }; return {}; }, COPY],
  ['a pivot HOSTED on the target sheet (its extent is not observable)', b => { b[0].objects = { charts: [], pivots: [{ name: 'PV', source: 'Q!$A$1:$A$2' }], names: [] }; return {}; }, COPY],
  ['an AutoFilter over the move source', b => { b[0].autoFilter = { range: 'H1:H9', filters: [] }; return {}; }, MOVE],
];

for (const [name, setup, t] of AFFECTED) {
  test(`affected, refused 0/0: ${name}`, async () => {
    const b = book();
    const extra = setup(b);
    const store = createStore(b);
    if (extra.names) store.names = extra.names;
    await refused(store, t, 'xlsx-range-dependent-affected');
    assert.deepEqual(store.durable, (() => { const x = book(); setup(x); return x; })(), 'the document is untouched');
  });
}

// ---- unobservable / unresolvable dependents: refused 0 write / 0 command (never "not affected") ----

const UNOBSERVABLE = [
  ['the names reader failed', {}, { namesFault: 'UNOBSERVABLE:threw' }, 'xlsx-range-dependency-observation-incomplete'],
  ['the dependency build is not proven', {}, { dependencyBuild: false }, 'xlsx-range-dependency-observation-incomplete'],
  ['a pivot reader that threw (the after-reopen shape)', { pivotThrows: true }, { objectsUnobservable: true }, 'xlsx-range-dependency-observation-incomplete'],
  ['an INDIRECT formula anywhere (text reference, unresolvable)', { indirect: true }, {}, 'xlsx-range-dependency-observation-incomplete'],
  ['an unmeasured conditional type anywhere', { colorScale: true }, {}, 'xlsx-range-dependency-observation-incomplete'],
  ['the AutoFilter field build not verified', {}, { filterBuild: false }, 'xlsx-range-dependent-unobservable'],
  ['a RELATIVE defined name (evaluates per user cell)', { relName: true }, {}, 'xlsx-range-dependent-unobservable'],
];

for (const [name, what, script, want] of UNOBSERVABLE) {
  test(`unobservable, refused 0/0: ${name}`, async () => {
    const b = book();
    if (what.indirect) b[1].cells.C3 = { v: 0, f: '=INDIRECT("P!H1")' };
    if (what.colorScale) b[1].rules = [{ kind: 'conditional', type: 'xlColorScale', range: 'C1', formula: '' }];
    const store = createStore(b);
    if (what.relName) store.names = [{ name: 'Rel', ref: 'Q!C9' }];
    await refused(store, MOVE, want, script);
  });
}

// ---- clean: unrelated dependents do not refuse; first run APPLIED, the reopened retry SATISFIED 0/0 ----

const unrelated = () => {
  const b = book();
  b[1].rules = [{ kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', range: 'A1:A4', formula: '5' }, { kind: 'validation', type: 1, range: 'C1', formula: '1', formula2: '10' }];
  b[1].objects = { charts: [{ type: 'bar', series: [{ name: '', values: '=Q!$A$1:$A$2', cat: '', x: '' }] }], pivots: [], names: [] };
  b[0].rules = [{ kind: 'conditional', range: 'Z1', formula: '=$C$9>0' }]; // on P, reads C9 only -- none of the blocks
  const store = createStore(b);
  store.names = [{ name: 'Far', ref: 'Q!$A$1' }];
  return store;
};

test('clean copy with unrelated names / rules / charts: APPLIED, reopened retry SATISFIED 0/0', async () => {
  const store = unrelated();
  const { result, m } = await run(store, COPY);
  assert.equal(result.ok, true, JSON.stringify([op0(result), result.execution.preservationVerification]));
  assert.equal(m.commands, 1);
  const again = await run(store, COPY);
  assert.equal(again.m.commands, 0);
  assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('clean move with unrelated dependents: APPLIED, the reopened retry is REFUSED named, 0/0 (L-06/b: never SATISFIED, never a second Cut)', async () => {
  const store = unrelated();
  const { result, m } = await run(store, MOVE);
  assert.equal(result.ok, true, JSON.stringify([op0(result), result.execution.preservationVerification]));
  assert.equal(m.commands, 1);
  const again = await run(store, MOVE);
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.ok, false);
  assert.equal(reasonOf(again.result), 'xlsx-move-already-applied-not-provable');
});

// ---- L3 negative controls: a dependent changed as a side effect is never APPLIED ----

for (const fault of ['range-side-name', 'range-side-rule', 'range-side-chart', 'range-side-filter']) {
  test(`L3 catches a dependent side effect: ${fault}`, async () => {
    const store = unrelated();
    const { result } = await run(store, COPY, { fault });
    assert.equal(result.ok, false);
    const pv = JSON.stringify(result.execution.preservationVerification);
    assert.match(pv, /dependentDiffs/, pv);
    assert.equal(op0(result).status === OPERATION_STATUS.APPLIED, false);
  });
}

// ---- the relative-reference spread ----

test('referencesBlock: a relative part repeats over the spread, an absolute part does not', () => {
  const block = { sheet: 'P', r1: 5, r2: 6, c1: 8, c2: 8 };
  assert.equal(referencesBlock('=P!H1>0', 'Q', block), false);
  assert.equal(referencesBlock('=P!H1>0', 'Q', block, { dr: [0, 9], dc: [0, 0] }), true);
  assert.equal(referencesBlock('=P!$H$1>0', 'Q', block, { dr: [0, 9], dc: [0, 0] }), false);
  assert.equal(referencesBlock('=P!$H1>0', 'Q', block, { dr: [0, 4], dc: [0, 0] }), true);
  assert.equal(referencesBlock('=P!$H1>0', 'Q', block, { dr: [0, 3], dc: [0, 0] }), false);
});

// ---- xlCellValue: observable with a measured operator and a full read; anything else UNOBSERVABLE ----

const cfList = r => analyze(['S'], {}, { names: [], validations: { S: [] }, conditional: { S: [r] }, objects: { S: { charts: [], pivots: [], names: [] } } });
const CV = { type: 'xlCellValue', operator: 'xlGreater', appliesTo: 'S!$A$1:$A$4', formula1: '5', formula2: '' };

test('xlCellValue read: a full read with a measured operator is complete; the rule keeps its operator', () => {
  const r = cfList(CV);
  assert.equal(r.rules.complete, true);
  assert.equal(r.dependencies.complete, true);
  assert.deepEqual(r.rules.bySheet.S, [{ kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', range: 'A1:A4', formula: '5', formula2: '' }]);
  assert.equal(cfList({ ...CV, operator: 'xlBetween', formula1: '2', formula2: '7' }).rules.complete, true);
});

// X-L13/7 (zero LIVE r1, cap-live/xl137-r1.out, internal tracker): the remaining six single-/two-formula
// xlCellValue operators measured round-trip-safe on BOTH readers this module consumes -- each must now
// be `rules.complete: true`, same as xlGreater/xlBetween already were, with its own formula arity.
for (const [op, f2] of [['xlLess', ''], ['xlLessEqual', ''], ['xlGreaterEqual', ''], ['xlEqual', ''], ['xlNotEqual', ''], ['xlNotBetween', '7']]) {
  test(`xlCellValue read: X-L13/7 newly-measured operator ${op} is complete`, () => {
    const r = cfList({ ...CV, operator: op, formula2: f2 });
    assert.equal(r.rules.complete, true, JSON.stringify(r.rules.unobservable));
    assert.equal(r.rules.bySheet.S[0].operator, op);
  });
}

// All eight of Excel's xlFormatConditionOperator values are measured as of X-L13/7 -- there is no
// REAL operator left to use as the "unmeasured" NEG. CONTROL below, so these two cases use a
// deliberately synthetic operator name that can never collide with a real engine value, to keep
// proving the "operator not in the whitelist" branch itself (not any specific Excel operator).
const UNMEASURED_OP = 'xlOperatorNotYetMeasured';

for (const [name, r, why] of [
  ['an unmeasured operator', { ...CV, operator: UNMEASURED_OP }, 'conditional-operator-not-measured'],
  ['the operator missing', { ...CV, operator: undefined }, 'conditional-operator-not-measured'],
  ['formula1 missing', { ...CV, formula1: '' }, 'conditional-partial-read'],
  ['between without formula2', { ...CV, operator: 'xlBetween', formula2: '' }, 'conditional-partial-read'],
  ['greater WITH a formula2', { ...CV, formula2: '9' }, 'conditional-partial-read'],
  ['applies-to missing', { ...CV, appliesTo: '' }, 'conditional-unexpected-shape'],
  ['another type (colour scale)', { type: 'xlColorScale' }, 'conditional-type-not-measured'],
  ['another type (data bar)', { type: 'xlDatabar' }, 'conditional-type-not-measured'],
  ['another type (icon set)', { type: 'xlIconSets' }, 'conditional-type-not-measured'],
]) {
  test(`xlCellValue read stays UNOBSERVABLE: ${name}`, () => {
    const a = cfList(r);
    assert.equal(a.rules.complete, false);
    assert.match(JSON.stringify(a.rules.unobservable), new RegExp(why));
  });
}

// ---- GH #141 (docs/GH141-CF-DEPENDENCY-OBSERVABILITY-20261018.md): aboveAverage / top10 / unique have no
// formula concept at all (the engine's own getters throw for them -- cf.cjs's doc comment), so only
// appliesTo is read; they join MEASURED_CF_TYPES. colorScale / databar / iconSet (regression-checked above)
// stay OUT: their CFVO stops can be formula-typed and no reader here has ever measured that.

for (const [name, r] of [
  ['aboveAverage', { type: 'xlAboveAverageCondition', appliesTo: 'S!$A$1:$A$4', formula1: null, formula2: null }],
  ['top10', { type: 'xlTop10', appliesTo: 'S!$A$1:$A$4', formula1: null, formula2: null }],
  ['unique', { type: 'xlUniqueValues', appliesTo: 'S!$A$1:$A$4', formula1: null, formula2: null }],
]) {
  test(`GH141: ${name} read (appliesTo only, no formula) is complete`, () => {
    const a = cfList(r);
    assert.equal(a.rules.complete, true, JSON.stringify(a.rules.unobservable));
    assert.equal(a.dependencies.complete, true, JSON.stringify(a.dependencies.unobservable));
    assert.deepEqual(a.rules.bySheet.S, [{ kind: 'conditional', type: r.type, range: 'A1:A4', formula: '', formula2: '' }]);
  });
}

const withNewCfType = type => createStore([
  { name: 'Data', cells: { A1: { v: 'r1' }, A2: { v: 'r2' } } },
  { name: 'Other', cells: { A1: { v: 7 } }, rules: [{ kind: 'conditional', type, range: 'A1:A4', formula: '' }] },
]);

for (const type of ['xlAboveAverageCondition', 'xlTop10', 'xlUniqueValues']) {
  test(`GH141: ${type} on ANOTHER sheet no longer blocks workbook-wide: set_defined_name APPLIED, reopened retry 0/0`, async () => {
    const store = withNewCfType(type);
    const t = { id: 'gh141', steps: [{ intent: 'set_defined_name', name: 'GH141N', refersTo: '=Data!$A$1' }] };
    const { result, m } = await run(store, t);
    assert.equal(result.ok, true, JSON.stringify(op0(result)));
    assert.equal(m.commands, 1);
    const again = await run(store, t);
    assert.equal(again.m.commands, 0);
    assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test('GH141 regression: xlColorScale on ANOTHER sheet still blocks workbook-wide (stays excluded)', async () => {
  const store = withNewCfType('xlColorScale');
  const { result } = await run(store, { id: 'gh141-cs', steps: [{ intent: 'set_defined_name', name: 'GH141N2', refersTo: '=Data!$A$1' }] });
  assert.equal(result.ok, false);
  assert.match(JSON.stringify(op0(result)), /conditional-type-not-measured/);
});

const withCellValueCf = () => {
  const b = [
    { name: 'Data', cells: { A1: { v: 'r1' }, A2: { v: 'r2' }, A3: { v: 'r3' } } },
    { name: 'Other', cells: { A1: { v: 7 } }, rules: [{ kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', range: 'A1:A4', formula: '5' }] },
  ];
  return createStore(b);
};

test('xlCellValue on ANOTHER sheet: set_defined_name APPLIED (was refused workbook-wide), reopened retry 0/0', async () => {
  const store = withCellValueCf();
  const t = { id: 'm7s1n', steps: [{ intent: 'set_defined_name', name: 'S1N', refersTo: '=Data!$A$1' }] };
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(op0(result)));
  assert.equal(m.commands, 1);
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
  assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('xlCellValue on ANOTHER sheet: insert_rows on Data APPLIED, reopened retry 0/0', async () => {
  const store = withCellValueCf();
  const t = { id: 'm7s1s', steps: [{ intent: 'insert_rows', sheet: 'Data', at: 2, count: 1, anchor: { column: 'A', value: 'r2' } }] };
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(op0(result)));
  assert.equal(m.commands, 1);
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
});

test('xlCellValue HOSTED on the shifted sheet: the structural step still refuses (rules on the sheet)', async () => {
  const b = [{ name: 'Data', cells: { A1: { v: 'r1' }, A2: { v: 'r2' } }, rules: [{ kind: 'conditional', type: 'xlCellValue', operator: 'xlGreater', range: 'B1:B4', formula: '5' }] }];
  await refused(createStore(b), { id: 'x', steps: [{ intent: 'insert_rows', sheet: 'Data', at: 2, count: 1, anchor: { column: 'A', value: 'r2' } }] }, 'xlsx-structure-dependent-not-measured');
});

test('xlCellValue with an unmeasured operator anywhere: set_defined_name still refused (incomplete)', async () => {
  const store = withCellValueCf();
  store.durable[1].rules[0].operator = UNMEASURED_OP;
  const { result } = await run(store, { id: 'x', steps: [{ intent: 'set_defined_name', name: 'S1N', refersTo: '=Data!$A$1' }] });
  assert.equal(result.ok, false);
  assert.match(JSON.stringify(op0(result)), /conditional-operator-not-measured/);
});

// X-L13/7: xlLess (and its five siblings) on ANOTHER sheet no longer blocks an unrelated set_defined_name --
// this is the exact whole-task shape of the e314d2b6 LIVE finding (an unrelated CF rule's operator fail-closing
// set_validation workbook-wide), replayed here for set_defined_name since that is what this file already covers.
test('xlCellValue with a newly-measured operator (xlLess) on ANOTHER sheet: set_defined_name APPLIED, reopened retry 0/0', async () => {
  const store = withCellValueCf();
  store.durable[1].rules[0].operator = 'xlLess';
  const t = { id: 'm7s1n-xlless', steps: [{ intent: 'set_defined_name', name: 'S1N2', refersTo: '=Data!$A$1' }] };
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(op0(result)));
  assert.equal(m.commands, 1);
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
  assert.equal(op0(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('copy_sheet of a sheet HOSTING an xlCellValue rule: refused (the copy of that type is not measured)', async () => {
  const store = withCellValueCf();
  await refused(store, { id: 'x', steps: [{ intent: 'copy_sheet', sheet: 'Other', name: 'Other2' }] }, 'xlsx-copy-conditional-type-not-measured-for-copy');
});

// ---- reserved names ----

for (const n of ['Print_Area', 'PRINT_AREA', 'print_titles', 'Print_Titles', '_xlnm.Print_Area', '_XLNM.Criteria']) {
  test(`reserved name refused at planning, 0 read: ${n}`, () => {
    for (const step of [{ intent: 'set_defined_name', name: n, refersTo: '=Data!$A$1' }, { intent: 'delete_defined_name', name: n }]) {
      const p = planOperations({ id: 'x', steps: [step] });
      assert.equal(p.ok, false);
      assert.match(p.detail, /reserved name/);
    }
  });
}

test('an ordinary name that merely contains the words is accepted', () => {
  for (const n of ['PrintArea2', 'My_Print_Area', 'xlnm_note']) assert.equal(planOperations({ id: 'x', steps: [{ intent: 'set_defined_name', name: n, refersTo: '=Data!$A$1' }] }).ok, true, n);
});

// ---- the guard itself, directly (a partial / wrong reader) ----

test('guard: dependencies not read at all -> refused', () => {
  const p = normalizeInventory({ sheets: [{ name: 'P', index: 0, used: 'A1:B1' }], cells: [{ sheet: 'P', a: 'A1', v: 'a', f: '' }, { sheet: 'P', a: 'B1', v: 1, f: '' }, { sheet: 'P', a: 'D5', v: null, f: '' }, { sheet: 'P', a: 'E5', v: null, f: '' }], objects: {}, runtime: { identity: 'x', fileId: 'f' } });
  const op = planOperations(COPY).operations[0];
  assert.equal(range.classify(p, op).evidence.reason, 'xlsx-range-dependents-not-read');
});
