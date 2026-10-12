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
 * M6 wave 3: create/rename/delete/move_sheet through the target path on the fake transport. The fake
 * follows the wave-3 LIVE probe (m6-w3p-1790584404615-f856e6): a rename rewrites the references, a
 * delete turns them into #REF!; its faults reproduce what the probe found silently wrong (a name
 * turned into 'Sheet2') and what a runtime could get wrong (references left behind, side effects).
 */

const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, B2: { v: 3 }, C2: { v: 6, f: '=B2*2' } } },
  { name: 'Other', cells: { A1: { v: 'keep me' }, B1: { v: 4, f: '=Data!B2+1' }, C1: { v: 'Data!B2', f: '="Data!"&"B2"' } } },
  { name: 'Notes', cells: { A1: { v: 'free text' } } },
];

const task = step => ({ id: 'm6-w3', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  // M11.5c.4: copy_sheet writes only through the /8 Bridge WRITE -- a copy task runs on /8 unless the test names a runtime
  const driver = createFakeXlsxDriver(store, t.steps.some(x => x.intent === 'copy_sheet') && script.runtime === undefined ? { ...script, runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/8' } : script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const names = store => store.durable.map(s => s.name);
const applied = async (store, t, script) => {
  const { result, m } = await run(store, t, script);
  assert.equal(result.ok, true, JSON.stringify(result.execution && result.execution.primitiveVerification));
  assert.equal(m.writes, 1);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
  return result;
};
const satisfied = async (store, t) => {
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true);
  assert.equal(result.execution.noOp, true);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
};
const refused = async (store, t, reason, script) => {
  const before = JSON.stringify(store.durable);
  const { result, driver, m } = await run(store, t, script);
  assert.equal(result.ok, false);
  assert.equal(m.writes, 0, 'refused before any write');
  assert.equal(driver.calls.dispatch, 0);
  assert.match(JSON.stringify(result), new RegExp(reason));
  assert.equal(JSON.stringify(store.durable), before, 'durable untouched');
};

test('create_sheet: FIRST applied (exact name, empty, other sheets untouched) -> reopened retry SATISFIED', async () => {
  const store = createStore(book());
  const t = task({ intent: 'create_sheet', name: 'M6W3_new' });
  await applied(store, t);
  assert.equal(names(store).filter(n => n === 'M6W3_new').length, 1);
  await satisfied(store, t);
});

test('create_sheet: names outside the measured class are refused at planning (the probe: bad/ accepted, 32 chars -> Sheet2)', () => {
  for (const name of ['bad/x', 'L'.padEnd(32, 'x'), 'A1', 'XFD7', 'R1C1', 'rc', 'History', 'TRUE', '1abc', '', 'Ünnep', ' lead', 'two  spaces', "it's"]) {
    const p = planOperations(task({ intent: 'create_sheet', name }));
    assert.equal(p.ok, false, name);
  }
  assert.equal(planOperations(task({ intent: 'create_sheet', name: 'L'.padEnd(31, 'x') })).ok, true, '31 characters is allowed');
  // M7 R26c: words joined by single spaces (accents too) are the measured quoted class
  for (const name of ['My Sheet', 'Bevétel 2026', 'Q 1']) assert.equal(planOperations(task({ intent: 'create_sheet', name })).ok, true, name);
});

test('create_sheet: a non-empty sheet of that name, or a case-only collision, is refused before dispatch', async () => {
  await refused(createStore(book()), task({ intent: 'create_sheet', name: 'Notes' }), 'xlsx-create-name-taken-by-a-non-empty-sheet');
  await refused(createStore(book()), task({ intent: 'create_sheet', name: 'notes' }), 'xlsx-sheet-name-case-collision');
});

test('RED: the measured silent default name (Sheet2) is NOT APPLIED and nothing is persisted', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ intent: 'create_sheet', name: 'M6W3_new' }), { fault: 'create-default-name' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
  assert.deepEqual(names(store), ['Data', 'Other', 'Notes']);
});

test('RED: a create that also reorders the other sheets passes L1 but fails L3 (sheet list)', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ intent: 'create_sheet', name: 'M6W3_new' }), { fault: 'create-reorders' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS');
  assert.equal(result.execution.preservationVerification[0].evidence.evidence.reason, 'xlsx-sheet-list-changed-beyond-the-operation');
  assert.equal(driver.calls.forceSave, 0);
});

test('rename_sheet: applied, the referencing formula follows, L3 PASS, same position -> retry SATISFIED', async () => {
  const store = createStore(book());
  const t = task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' });
  await applied(store, t);
  assert.deepEqual(names(store), ['Data2', 'Other', 'Notes']);
  assert.equal(store.durable[1].cells.B1.f, '=Data2!B2+1');
  assert.equal(store.durable[1].cells.C1.f, '="Data!"&"B2"', 'a string literal is not a reference');
  await satisfied(store, t);
});

test('RED: a rename whose references did NOT follow fails L3 (formula), not APPLIED', async () => {
  const store = createStore(book());
  const { result, driver } = await run(store, task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' }), { fault: 'rename-no-follow' });
  assert.equal(result.ok, false);
  const pv = result.execution.preservationVerification[0];
  assert.equal(pv.verdict, 'FAIL');
  assert.deepEqual(pv.evidence.evidence.changed.map(c => [c.cell, c.change]), [['Other!B1', 'formula']]);
  assert.equal(driver.calls.forceSave, 0);
});

test('rename_sheet refusals: target present, both present, case-only, source missing', async () => {
  await refused(createStore(book()), task({ intent: 'rename_sheet', sheet: 'Data', name: 'notes' }), 'xlsx-sheet-name-case-collision');
  await refused(createStore(book()), task({ intent: 'rename_sheet', sheet: 'Data', name: 'Notes' }), 'xlsx-rename-source-and-target-both-present');
  await refused(createStore(book()), task({ intent: 'rename_sheet', sheet: 'data', name: 'X1a' }), 'xlsx-sheet-case-mismatch');
  await refused(createStore(book()), task({ intent: 'rename_sheet', sheet: 'Nope', name: 'Yes' }), 'xlsx-sheet-not-found');
  assert.equal(planOperations(task({ intent: 'rename_sheet', sheet: 'Data', name: 'DATA' })).ok, false, 'case-only rename');
});

test('rename_sheet refused before any write: a conditional format naming the sheet -- its own, another sheet\'s, through a name (W3f run 2: a CF does not follow a rename)', async () => {
  const t = task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' });
  const cases = [
    ['own CF', b => { b[0].rules = [{ kind: 'conditional', range: 'F1', formula: '=Data!$B$2>0' }]; return {}; }, 'Data'],
    ['CF on another sheet', b => { b[1].rules = [{ kind: 'conditional', range: 'F1', formula: '=Data!$A$1="x"' }]; return {}; }, 'Other'],
    ['CF through a workbook name', b => { b[1].rules = [{ kind: 'conditional', range: 'F1', formula: '=WbN>0' }]; return { workbookNames: [{ name: 'WbN', ref: 'Data!$B$2' }] }; }, 'Other'],
    ['CF formula2, case-folded sheet name', b => { b[2].rules = [{ kind: 'conditional', range: 'F1', formula: '=Notes!$A$1>0', formula2: '=data!$B$2' }]; return {}; }, 'Notes'],
  ];
  for (const [label, setup, host] of cases) {
    const b = book();
    const script = setup(b);
    await refused(createStore(b), t, 'xlsx-rename-conditional-format-reference-not-followed', script);
    const { result } = await run(createStore(b), t, script);
    assert.match(JSON.stringify(result), new RegExp(`"kind":"conditional","host":"${host}"`), label);
  }
});

test('rename_sheet: a pivot over the sheet refuses before any write (W3g run 1: unobservable after the reopen); a chart over it does not (it follows)', async () => {
  const t = task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' });
  const b = book();
  b[1].objects = { charts: [], pivots: [{ name: 'P1', source: 'Data!$A$1:$B$3' }], names: [] };
  await refused(createStore(b), t, 'xlsx-rename-pivot-source-unobservable-after-reopen');
  const other = book();
  other[1].objects = { charts: [], pivots: [{ name: 'P1', source: 'Notes!$A$1:$B$3' }], names: [] };
  await applied(createStore(other), t, {});
  const c = book();
  c[1].objects = { charts: [{ type: 'bar', series: [{ name: '=Data!$B$1', values: '=Data!$B$2:$B$3', cat: '=Data!$A$2:$A$3', x: '' }] }], pivots: [], names: [] };
  const store = createStore(c);
  await applied(store, t, {});
  assert.equal(store.durable[1].objects.charts[0].series[0].values, '=Data2!$B$2:$B$3');
});

test('rename_sheet refused before any write when the dependency read is incomplete or unresolvable (build, readers, INDIRECT)', async () => {
  const t = task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' });
  const gap = 'xlsx-rename-dependency-observation-incomplete';
  for (const script of [{ dependencyBuild: false }, { namesFault: 'UNOBSERVABLE:asc_getDefinedNames-threw' }, { validationsFault: 'UNOBSERVABLE:dataValidations-threw' }, { conditionalFault: 'UNOBSERVABLE:threw' }]) {
    await refused(createStore(book()), t, gap, script);
  }
  const b = book();
  b[1].rules = [{ kind: 'conditional', range: 'F1', formula: '=NoSuchName>0' }];
  await refused(createStore(b), t, gap);
  const c = book();
  c[2].cells.D1 = { v: 1, f: '=INDIRECT("Data!B2")' };
  await refused(createStore(c), t, gap);
});

test('rename_sheet: conditional formats that do NOT name the sheet (own over another sheet, another sheet over itself) -> applied, retry SATISFIED', async () => {
  const b = book();
  b[0].rules = [{ kind: 'conditional', range: 'F1', formula: '=Other!$A$1="x"' }];
  b[1].rules = [{ kind: 'conditional', range: 'F1', formula: '=Other!$B$1>0' }, { kind: 'validation', range: 'E1', formula: '=Notes!$A$1' }];
  const store = createStore(b);
  const t = task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' });
  await applied(store, t, {});
  assert.deepEqual(names(store), ['Data2', 'Other', 'Notes']);
  await satisfied(store, t);
});

test('rename_sheet with dependents that follow (validation on another sheet, a sheet-level and a workbook name over the sheet): applied, L3 PASS, retry SATISFIED', async () => {
  const b = book();
  b[1].rules = [{ kind: 'validation', range: 'E1', formula: '=Data!$A$1:$A$2' }];
  b[1].objects = { charts: [], pivots: [], names: [{ name: 'LocN', refersTo: '=Data!$B$2' }] };
  const script = { workbookNames: [{ name: 'WbN', ref: 'Data!$B$2' }] };
  const store = createStore(b);
  const t = task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' });
  const r = await applied(store, t, script);
  assert.equal(r.execution.preservationVerification[0].evidence.evidence.workbook.ok, true);
  assert.equal(store.durable[1].rules[0].formula, '=Data2!$A$1:$A$2');
  await satisfied(store, t);
});

test('RED: a rename that leaves a validation or a defined name on the old sheet name fails L3 (workbook side); not APPLIED, nothing saved', async () => {
  for (const [fault, axis] of [['rename-dv-no-follow', 'rules'], ['rename-name-no-follow', 'defined-names']]) {
    const b = book();
    b[1].rules = [{ kind: 'validation', range: 'E1', formula: '=Data!$A$1:$A$2' }];
    const { result, driver } = await run(createStore(b), task({ intent: 'rename_sheet', sheet: 'Data', name: 'Data2' }), { fault, workbookNames: [{ name: 'WbN', ref: 'Data!$B$2' }] });
    assert.equal(result.ok, false, fault);
    const pv = result.execution.preservationVerification[0];
    assert.equal(pv.verdict, 'FAIL', fault);
    assert.match(JSON.stringify(pv.evidence.evidence.workbook.diffs), new RegExp(`"axis":"${axis}"`), fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('delete_sheet: with a COMPLETE dependency observation naming nothing, an unreferenced sheet is deleted -> retry SATISFIED', async () => {
  const store = createStore(book());
  const t = task({ intent: 'delete_sheet', sheet: 'Notes' });
  await applied(store, t);
  assert.deepEqual(names(store), ['Data', 'Other']);
  await satisfied(store, t);
});

test('delete_sheet with the dependency readers: each controlled dependent on another sheet refuses before any write, and names itself', async () => {
  const cases = [
    ['workbook-level name', () => ({ workbookNames: [{ name: 'WbN', ref: 'Notes!$A$1' }] }), 'name', null],
    ['sheet-level name on Other', b => { b[1].objects = { charts: [], pivots: [], names: [{ name: 'LocN', refersTo: '=Notes!$A$1' }] }; return {}; }, 'name', 'Other'],
    ['validation on Other over Notes', b => { b[1].rules = [{ kind: 'validation', range: 'E1', formula: '=Notes!$A$1:$A$2' }]; return {}; }, 'validation', 'Other'],
    ['conditional format on Other over Notes', b => { b[1].rules = [{ kind: 'conditional', range: 'F1', formula: '=Notes!$A$1="x"' }]; return {}; }, 'conditional', 'Other'],
    ['chart series on Data over Notes', b => { b[0].objects = { charts: [{ type: 'bar', series: [{ name: '', values: '=Notes!$A$1:$A$2', cat: '', x: '' }] }], pivots: [], names: [] }; return {}; }, 'chart', 'Data'],
    ['pivot on Data over Notes', b => { b[0].objects = { charts: [], pivots: [{ name: 'P1', source: 'Notes!$A$1:$B$3' }], names: [] }; return {}; }, 'pivot', 'Data'],
  ];
  for (const [label, setup, kind, host] of cases) {
    const b = book();
    const script = setup(b);
    const { result } = await (async () => { await refused(createStore(b), task({ intent: 'delete_sheet', sheet: 'Notes' }), 'xlsx-delete-sheet-referenced-by-dependents', script); return run(createStore(b), task({ intent: 'delete_sheet', sheet: 'Notes' }), script); })();
    const deps = JSON.parse(JSON.stringify(result)).execution;
    assert.match(JSON.stringify(deps), new RegExp(`"kind":"${kind}","host":${host ? `"${host}"` : 'null'}`), label);
  }
});

test('delete_sheet: a formula THROUGH a defined name reaches the name\'s sheet (resolved, not only the name text)', async () => {
  const b = book();
  b[1].cells.D1 = { v: 2, f: '=LocN*2' };
  b[1].objects = { charts: [], pivots: [], names: [{ name: 'LocN', refersTo: '=Notes!$A$1' }] };
  const { result } = await run(createStore(b), task({ intent: 'delete_sheet', sheet: 'Notes' }));
  assert.equal(result.ok, false);
  assert.match(JSON.stringify(result), /"kind":"formula","host":"Other","name":"D1","sheets":\["Notes"\]/);
});

test('delete_sheet: a dependent hosted ON the deleted sheet goes with it (not a refusal); a sheet-scoped name there still refuses', async () => {
  const store = createStore((() => { const b = book(); b[2].rules = [{ kind: 'validation', range: 'E1', formula: '=Notes!$A$1' }]; return b; })());
  await applied(store, task({ intent: 'delete_sheet', sheet: 'Notes' }));
  const b = book();
  b[2].objects = { charts: [], pivots: [], names: [{ name: 'Own', refersTo: '=Notes!$A$1' }] };
  await refused(createStore(b), task({ intent: 'delete_sheet', sheet: 'Notes' }), 'xlsx-delete-sheet-referenced-by-dependents');
});

test('delete_sheet UNOBSERVABLE -> refused before any write (0 command): a reader gap, an unproven build, broken / ambiguous / unknown references', async () => {
  const gap = 'xlsx-delete-dependency-observation-incomplete';
  const t = task({ intent: 'delete_sheet', sheet: 'Notes' });
  for (const script of [{ dependencyBuild: false }, { namesFault: 'UNOBSERVABLE:asc_getDefinedNames-threw' }, { validationsFault: 'UNOBSERVABLE:dataValidations-threw' },
    { conditionalFault: 'UNOBSERVABLE:threw' }, { objectsUnobservable: true }, { namesFault: [{ name: 'X', ref: 'A1' }] }, { validationsFault: { Data: [] } }]) {
    await refused(createStore(book()), t, gap, script);
  }
  const variants = [
    ['broken validation', b => { b[1].rules = [{ kind: 'validation', range: 'E1', formula: '=#REF!' }]; }, /broken-reference/],
    ['name on a missing sheet', b => { b[1].objects = { charts: [], pivots: [], names: [{ name: 'Gone', refersTo: '=Nowhere!$A$1' }] }; }, /missing-sheet:Nowhere/],
    ['unknown name in a formula', b => { b[1].cells.D1 = { v: 1, f: '=NoSuchName*2' }; }, /unknown-name:NoSuchName/],
    ['unmeasured conditional type', b => { b[1].rules = [{ kind: 'conditional', type: 'xlColorScale', range: 'F1', formula: '' }]; }, /conditional-type-not-measured/],
    ['a structured reference', b => { b[1].cells.D1 = { v: 1, f: '=SUM(Table1[Col])' }; }, /external-or-structured-reference/],
    ['a broken formula', b => { b[1].cells.D1 = { v: '#REF!', f: '= #REF!*2' }; }, /broken-reference/],
    ['a literal chart series name (unmeasured shape)', b => { b[0].objects = { charts: [{ type: 'bar', series: [{ name: 'Revenue', values: '=Data!$B$2', cat: '', x: '' }] }], pivots: [], names: [] }; }, /series-shape-not-measured/],
  ];
  for (const [label, edit, why] of variants) {
    const b = book();
    edit(b);
    const { result, driver } = await run(createStore(b), t);
    assert.equal(result.ok, false, label);
    assert.equal(driver.calls.dispatch, 0, label);
    assert.match(JSON.stringify(result), new RegExp(gap), label);
    assert.match(JSON.stringify(result), why, label);
  }
  // an ambiguous name (two workbook definitions) and a name written with '=' (W3e: never resolved)
  const amb = book();
  amb[1].cells.D1 = { v: 1, f: '=Dup*2' };
  const r1 = await run(createStore(amb), t, { workbookNames: [{ name: 'Dup', ref: 'Data!$B$2' }, { name: 'DUP', ref: 'Data!$B$2' }] });
  assert.match(JSON.stringify(r1.result), /ambiguous-name:Dup/);
  assert.equal(r1.driver.calls.dispatch, 0);
  const r2 = await run(createStore(book()), t, { workbookNames: [{ name: 'Eq', ref: '=Data!$B$2' }] });
  assert.match(JSON.stringify(r2.result), /name-text-starts-with-equals/);
  assert.equal(r2.driver.calls.dispatch, 0);
});

test('INDIRECT anywhere still refuses a delete (a text reference no parser resolves)', async () => {
  const b = book();
  b[1].cells.D1 = { v: 1, f: '=INDIRECT("Notes!A1")' };
  await refused(createStore(b), task({ intent: 'delete_sheet', sheet: 'Notes' }), 'xlsx-delete-indirect-reference-unresolvable');
});

test('RED: a delete that also adds a rule on another sheet or a workbook name fails L3 (workbook side); nothing persisted', async () => {
  for (const fault of ['side-rule', 'side-name']) {
    const store = createStore(book());
    const { result, driver } = await run(store, task({ intent: 'delete_sheet', sheet: 'Notes' }), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('delete_sheet refusals: referenced by a formula on another sheet (the measured silent #REF!), the last sheet', async () => {
  await refused(createStore(book()), task({ intent: 'delete_sheet', sheet: 'Data' }), 'xlsx-delete-sheet-referenced-by-formulas');
  await refused(createStore([{ name: 'Only', cells: { A1: { v: 1 } } }]), task({ intent: 'delete_sheet', sheet: 'Only' }), 'xlsx-delete-last-sheet-refused');
});

test('move_sheet: before and after, then retry SATISFIED', async () => {
  const store = createStore(book());
  const before = task({ intent: 'move_sheet', sheet: 'Notes', before: 'Data' });
  await applied(store, before);
  assert.deepEqual(names(store), ['Notes', 'Data', 'Other']);
  await satisfied(store, before);
  const after = task({ intent: 'move_sheet', sheet: 'Notes', after: 'Other' });
  await applied(store, after);
  assert.deepEqual(names(store), ['Data', 'Other', 'Notes']);
  await satisfied(store, after);
});

test('every sheet intent: an ignored command fails L1, a side effect fails L3; nothing persisted', async () => {
  const steps = [
    { intent: 'create_sheet', name: 'M6W3_new' },
    { intent: 'rename_sheet', sheet: 'Notes', name: 'Notes2' },
    { intent: 'delete_sheet', sheet: 'Notes' },
    { intent: 'move_sheet', sheet: 'Notes', before: 'Data' },
  ];
  for (const step of steps) {
    for (const [fault, layer] of [['ignore', 'primitiveVerification'], ['side-effect', 'preservationVerification']]) {
      const store = createStore(book());
      const { result, driver } = await run(store, task(step), { fault });
      assert.equal(result.ok, false, `${step.intent} ${fault}`);
      assert.equal(result.execution[layer][0].verdict, 'FAIL', `${step.intent} ${fault}`);
      assert.equal(driver.calls.forceSave, 0, `${step.intent} ${fault}`);
    }
  }
});

test('planning: a sheet operation is the only step of its task; a copy needs a distinct name', () => {
  const p = planOperations({ id: 'x', steps: [{ intent: 'create_sheet', name: 'A_b' }, { intent: 'write_range', sheet: 'A_b', range: 'A1', values: [[1]] }] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /only step/);
  assert.equal(planOperations(task({ intent: 'copy_sheet', sheet: 'Data', name: 'data' })).ok, false, 'a case-only copy name');
  assert.equal(planOperations(task({ intent: 'move_sheet', sheet: 'Data', before: 'Other', after: 'Notes' })).ok, false);
});

// ---- copy_sheet: EO_INTERNAL_DEPENDENCY WRITE (owner decision 2026-09-28), build-gated ------------------

const copyBook = () => [
  { name: 'Data', cells: { A1: { v: 'k' }, B2: { v: 5 }, C1: { v: 5, f: '=Data!B2' }, C2: { v: 5, f: '=B2' }, C3: { v: 'keep me', f: '=Other!A1' } },
    objects: { charts: [{ type: 'bar', series: [{ name: '=Data!$B$1', values: '=Data!$B$2:$B$3', cat: '=Other!$A$1:$A$2', x: '' }] }],
      pivots: [{ name: 'PivotTable1', source: 'Data!$A$1:$B$3' }], names: [{ name: 'LocalN', refersTo: '=Data!$B$2' }] },
    rules: [{ kind: 'validation', range: 'E1', formula: '=Other!$A$1:$A$2' }, { kind: 'conditional', range: 'F1', formula: '=Other!$A$1>0' }] },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];
const copyTask = (name = 'Data_copy') => task({ intent: 'copy_sheet', sheet: 'Data', name });
const RULES = {}; // the dependency readers answer by default; a test turns one off to refuse

test('copy_sheet FIRST: applied through the internal write; self references now name the copy, others unchanged; L3 PASS; last', async () => {
  const store = createStore(copyBook());
  await applied(store, copyTask(), RULES);
  assert.deepEqual(names(store), ['Data', 'Other', 'Data_copy']);
  const copy = store.durable[2];
  assert.equal(copy.cells.C1.f, '=Data_copy!B2');
  assert.equal(copy.cells.C3.f, '=Other!A1');
  assert.equal(copy.objects.charts[0].series[0].values, '=Data_copy!$B$2:$B$3');
  assert.equal(copy.objects.charts[0].series[0].cat, '=Other!$A$1:$A$2');
  assert.equal(copy.objects.names[0].refersTo, '=Data_copy!$B$2');
});

test('copy_sheet retry: the end state equals the desired copy -> SATISFIED, 0/0, provenance NOT proven (not this run\'s copy)', async () => {
  const store = createStore(copyBook());
  await applied(store, copyTask(), RULES);
  const { result, m } = await run(store, copyTask(), RULES);
  assert.equal(result.ok, true);
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0);
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.SATISFIED);
  assert.match(JSON.stringify(result.execution), /"provenance":"NOT_PROVEN"/);
});

test('copy_sheet refused before any write: an unobservable rule / name reader, an unproven build, unobservable objects, a broken name on the source', async () => {
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-rule-observation-incomplete', { validationsFault: 'UNOBSERVABLE:dataValidations-threw' });
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-rule-observation-incomplete', { conditionalFault: 'UNOBSERVABLE:threw' });
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-rule-observation-incomplete', { dependencyBuild: false });
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-source-dependency-unobservable', { namesFault: 'UNOBSERVABLE:asc_getDefinedNames-threw' });
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-bridge-write-runtime-not-proven', { ...RULES, runtime: '9.3.4-9475e9f1' });
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-bridge-write-runtime-not-proven', { ...RULES, runtime: null });
  await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-source-objects-unobservable', { ...RULES, objectsUnobservable: true });
  const b = copyBook();
  b[0].objects.names.push({ name: 'Broken', refersTo: '=Gone!$A$1' });
  await refused(createStore(b), copyTask(), 'xlsx-copy-source-dependency-unobservable', RULES);
  const c = copyBook();
  c[0].rules.push({ kind: 'validation', range: 'E3', formula: '=#REF!' });
  await refused(createStore(c), copyTask(), 'xlsx-copy-source-dependency-unobservable', RULES);
});

test('copy_sheet refused before any write when an L3 axis of ANOTHER sheet is unobservable (a pivot there: unreadable after a reopen, W3g); a resolution gap elsewhere is not a copy gap', async () => {
  const b = copyBook();
  b[1].objects = { charts: [], pivots: 'UNOBSERVABLE:TypeError: reading ws', names: [] };
  await refused(createStore(b), copyTask(), 'xlsx-copy-source-dependency-unobservable', RULES);
  const { result } = await run(createStore(b), copyTask(), RULES);
  assert.match(JSON.stringify(result), /"axis":"pivots:Other"/);
  // an unknown name in a formula on another sheet does not stop the copy (L3 compares that cell's text)
  const c = copyBook();
  c[1].cells.D1 = { v: 1, f: '=NoSuchName*2' };
  await applied(createStore(c), copyTask(), RULES);
});

test('copy_sheet refused before any write: a validation / conditional format on the source naming the source itself (W3f run 1: the engine leaves it on the source)', async () => {
  const cases = [
    { kind: 'conditional', range: 'F2', formula: '=Data!$B$2>0' },
    { kind: 'validation', range: 'E2', formula: '=Data!$A$1:$A$2' },
    { kind: 'validation', range: 'E3', formula: '=Other!$A$1', formula2: "='Data'!$A$2" },
    { kind: 'conditional', range: 'F3', formula: '=data!$B$2>0' },
  ];
  for (const rule of cases) {
    const b = copyBook();
    b[0].rules.push(rule);
    await refused(createStore(b), copyTask(), 'xlsx-copy-rule-self-reference-not-mapped', RULES);
  }
  // negative control: the same shapes naming ANOTHER sheet are not refused by this guard
  const ok = copyBook();
  ok[0].rules.push({ kind: 'conditional', range: 'F2', formula: '=Other!$B$2>0' }, { kind: 'validation', range: 'E3', formula: '=Other!$A$1', formula2: "='Other'!$A$2" });
  await applied(createStore(ok), copyTask(), RULES);
});

test('copy_sheet: the rules of the copy (validation + conditional over another sheet, unchanged) and the sheet-level name are verified; other sheets and workbook names unchanged', async () => {
  const store = createStore(copyBook());
  const r = await applied(store, copyTask(), { workbookNames: [{ name: 'WbN', ref: 'Other!$A$1' }] });
  const copy = store.durable[2];
  assert.deepEqual(copy.rules.map(x => [x.kind, x.range, x.formula]), [['validation', 'E1', '=Other!$A$1:$A$2'], ['conditional', 'F1', '=Other!$A$1>0']]);
  assert.equal(r.execution.preservationVerification[0].evidence.evidence.workbook.ok, true);
});

test('RED: a copy that also adds a rule on another sheet or a workbook name fails L3 (workbook side); nothing persisted', async () => {
  for (const fault of ['side-rule', 'side-name']) {
    const store = createStore(copyBook());
    const { result, driver } = await run(store, copyTask(), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('copy_sheet refused: the name taken (case-insensitively), or taken by a sheet that differs / is not last', async () => {
  const b = copyBook();
  b.push({ name: 'Data_copy', cells: { A1: { v: 'other content' } } });
  await refused(createStore(b), copyTask(), 'xlsx-copy-target-exists-and-differs', RULES);
  const c = copyBook();
  c.push({ name: 'data_COPY', cells: {} });
  await refused(createStore(c), copyTask(), 'xlsx-sheet-name-case-collision', RULES);
  // an identical sheet that is NOT last is not the desired end state
  const store = createStore(copyBook());
  await applied(store, copyTask(), RULES);
  store.durable.push({ name: 'Tail', cells: {} });
  await refused(store, copyTask(), 'xlsx-copy-target-exists-and-differs', RULES);
});

test('RED: a duplicate after the copy, or a timeout, is never APPLIED; nothing is saved and the copy is NOT dispatched again', async () => {
  for (const fault of ['copy-duplicate', 'copy-timeout']) {
    const store = createStore(copyBook());
    const { result, driver } = await run(store, copyTask(), { ...RULES, fault });
    assert.equal(result.ok, false, fault);
    assert.notEqual(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, fault);
    assert.equal(driver.calls.dispatch, 1, `${fault}: no blind second copy`);
    assert.equal(driver.calls.forceSave, 0, fault);
    assert.deepEqual(names(store), ['Data', 'Other'], `${fault}: durable untouched`);
  }
});

test('RED: a copy whose self references still point at the source, or that drops an object, a rule or a cell, fails L1; a side effect fails L3', async () => {
  for (const [fault, layer] of [['copy-keeps-self-reference', 'primitiveVerification'], ['copy-drops-object', 'primitiveVerification'], ['copy-drops-rule', 'primitiveVerification'], ['copy-drops-cell', 'primitiveVerification'], ['side-effect', 'preservationVerification']]) {
    const store = createStore(copyBook());
    const { result, driver } = await run(store, copyTask(), { ...RULES, fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution[layer][0].verdict, 'FAIL', fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

// M11.5c.4 cutover: B03 copy_sheet through the /8 Bridge WRITE ONLY -- /4 /5 /6 /7 /9 refused before the write (no fallback)
test('M11.5c.4 copy_sheet: APPLIED on /8 (one Bridge WRITE); refused on /4, /5, /6, /7 and /9 before the write', async () => {
  for (const rt of ['9.3.4|euro-office-bridge-v1/9.3.4.60/8']) await applied(createStore(copyBook()), copyTask(), { ...RULES, runtime: rt });
  for (const rt of ['4', '5', '6', '7', '9'].map(n => `9.3.4|euro-office-bridge-v1/9.3.4.60/${n}`)) await refused(createStore(copyBook()), copyTask(), 'xlsx-copy-bridge-write-runtime-not-proven', { ...RULES, runtime: rt });
});
