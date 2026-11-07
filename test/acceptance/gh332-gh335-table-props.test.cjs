'use strict';

/**
 * GH #332-#335 product side over the Bridge /19 sheet.table.write/v1: set_table_style / set_table_flag / set_table_totals_row / rename_table /
 * resize_table (table-props.cjs). The public API has no table setter (docs/evidence/gh332-gh335-public-probe-20261030.json), so the plan is ONE
 * Bridge WRITE served on /19 only; the family refuses, before the write and named, what the measurements (docs/DS-9.3.5-BRIDGE19-TABLE-WRITE-
 * DESIGN.md) show the engine gets wrong, and verifies from a fresh read_tables + cells. Offline: the fake driver models the handler's verdict;
 * this proves the PRODUCT logic -- NOT that the editor's table actions persist (the design's LIVE gate).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/table-props.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const RT19 = '9.3.5|euro-office-bridge-v1/9.3.5.0/19';
const RT18 = '9.3.5|euro-office-bridge-v1/9.3.5.0/18';
const RT24 = '9.3.5|euro-office-bridge-v1/9.3.5.0/24';
const RT25 = '9.3.5|euro-office-bridge-v1/9.3.5.0/25'; // W2 regression fix (workbook.protection.read/v1 null branch) -- sheet.table.write/v1 byte-identical to /24
const RT26 = '9.3.5|euro-office-bridge-v1/9.3.5.0/26'; // review fix (f0451f47, the two guard change) -- sheet.table.write/v1's totalsRow branch byte-identical to /25
const RT27 = '9.3.5|euro-office-bridge-v1/9.3.5.0/27'; // GH #346 AUTO-REPAIR 1. egyseg (chart.series.style.read/v1 dataPoint only) -- sheet.table.write/v1 byte-identical to /26
const RT28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28'; // GH #361/#363 pre-install readiness gate (internal tracker) -- byte-identical warrant, ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs
const RT29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker) -- byte-identical warrant, ds935-bridge29-veryhidden-write-handler-identity.test.cjs
const STYLE = { name: 'TableStyleLight9', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false };
const cell = (v, f = '') => ({ v, f });
const T1 = (extra = {}) => ({ name: 'Table1', ref: 'A1:C5', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: false, style: STYLE, autoFilter: true, ...extra });
const T1_TOTALS_ON = (extra = {}) => ({ name: 'Table1', ref: 'A1:C6', columns: ['Item', 'Qty', 'Note'], headerRow: true, totalsRow: true, style: STYLE, autoFilter: true, ...extra });
const book = (extra = {}) => [{
  name: 'A',
  cells: {
    A1: cell('Item'), B1: cell('Qty'), C1: cell('Note'),
    A2: cell('a'), B2: cell(1), C2: cell('x'), A3: cell('b'), B3: cell(2), C3: cell('y'),
    A4: cell('c'), B4: cell(3), C4: cell('z'), A5: cell('d'), B5: cell(4), C5: cell('w'),
    G1: cell('r1'), E8: cell(10, '=SUM(B2:B5)'),
    ...(extra.cells || {}),
  },
  tables: extra.tables || [T1()],
  merges: extra.merges || [],
  objects: { charts: [], pivots: extra.pivots || [], names: [] },
}, { name: 'Other', cells: { A1: cell('other') }, objects: { charts: [], pivots: [], names: [] } }];

const style = (extra = {}) => ({ intent: 'set_table_style', sheet: 'A', table: 'Table1', style: 'TableStyleMedium9', ...extra });
const flag = (extra = {}) => ({ intent: 'set_table_flag', sheet: 'A', table: 'Table1', flag: 'rowStripes', value: false, ...extra });
const totals = (extra = {}) => ({ intent: 'set_table_totals_row', sheet: 'A', table: 'Table1', value: true, ...extra });
const rename = (extra = {}) => ({ intent: 'rename_table', sheet: 'A', table: 'Table1', name: 'Sales', ...extra });
const resize = (extra = {}) => ({ intent: 'resize_table', sheet: 'A', table: 'Table1', ref: 'A1:C7', ...extra });
const task = step => ({ id: 'gh332', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT19, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const tableOf = store => store.durable[0].tables[0];
const snapshot = store => JSON.parse(JSON.stringify(store.durable));
const why = r => JSON.stringify([r.execution.primitiveVerification, r.execution.appliedOperations[0].evidence]);
// a classification refusal: appliedOperations[0].evidence = { stage: 'classification', state, evidence: <the family's own object> }
const evidenceOf = r => {
  const op = r.execution.appliedOperations && r.execution.appliedOperations[0];
  if (op) { const e = op.evidence; return e && e.evidence ? e.evidence : e; }
  return r.execution.resolution || r.execution; // a refusal at the RESOLVE step (an ambiguous table) has no applied operation
};
const deepReason = o => { if (!o || typeof o !== 'object') return undefined; if (typeof o.reason === 'string') return o.reason; if (typeof o.ambiguity === 'string') return o.ambiguity; for (const v of Object.values(o)) { const x = deepReason(v); if (x) return x; } return undefined; };
const reasonOf = r => deepReason(evidenceOf(r));

// ---- schema ----------------------------------------------------------------------------------------------

test('schema: each intent takes exactly its fields; a bad style / flag / name / ref and an unknown field are refused', () => {
  const v = s => W.validateWriteRequest(req([s]));
  for (const s of [style(), flag(), totals(), totals({ value: false }), rename(), resize()]) assert.equal(v(s).ok, true, JSON.stringify(v(s)));
  for (const [name, s] of [
    ['style free string', style({ style: 'Pretty' })], ['style Medium29', style({ style: 'TableStyleMedium29' })], ['style Light22', style({ style: 'TableStyleLight22' })],
    ['flag unknown', flag({ flag: 'headerRow' })], ['flag value string', flag({ value: 'false' })], ['flag no value', flag({ value: undefined })],
    // GH #335 remainder: value: false is now a real, accepted totals-row step -- the NEG control is a non-boolean value, not false itself.
    ['totals value string', totals({ value: 'false' })], ['totals value number', totals({ value: 1 })], ['totals no value', totals({ value: undefined })],
    ['rename space', rename({ name: 'My Table' })], ['rename no name', rename({ name: undefined })], ['resize unparsable', resize({ ref: 'A1' })], ['resize no ref', resize({ ref: undefined })],
    ['style step with a flag field', style({ flag: 'rowStripes' })], ['no table', style({ table: undefined })], ['bad table name', style({ table: '1bad name' })],
  ]) assert.equal(v(s).ok, false, name);
});

test('plan: totals OFF is now a real step (GH #335 remainder, PASS); only a non-boolean value is refused; a malformed name / ref is refused with the reason in the detail', () => {
  assert.equal(fam.plan({ intent: 'set_table_totals_row', sheet: 'A', table: 'T', value: false }, 0).ok, true);
  assert.equal(fam.plan({ intent: 'set_table_totals_row', sheet: 'A', table: 'T', value: false }, 0).operation.want.value, false);
  assert.match(fam.plan({ intent: 'set_table_totals_row', sheet: 'A', table: 'T', value: 'false' }, 0).detail, /requires value: a boolean/);
  assert.match(fam.plan({ intent: 'set_table_totals_row', sheet: 'A', table: 'T', value: undefined }, 0).detail, /requires value: a boolean/);
  assert.equal(fam.plan({ intent: 'rename_table', sheet: 'A', table: 'T', name: 'AB12' }, 0).ok, false);
  assert.equal(fam.plan({ intent: 'resize_table', sheet: 'A', table: 'T', ref: 'a1:c7' }, 0).operation.want.ref, 'A1:C7');
  assert.equal(fam.plan({ intent: 'resize_table', sheet: 'A', table: 'T', ref: '$A$1:$C$7' }, 0).operation.want.ref, 'A1:C7');
});

// ---- the five APPLY paths + retry ------------------------------------------------------------------------------

test('STYLE: one Bridge write, the style changes, nothing else does; retry SATISFIED with 0 commands and the named observation', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  const { result, m } = await run(store, task(style()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).style.name, 'TableStyleMedium9');
  assert.deepEqual(tableOf(store).style.rowStripes, true);
  assert.deepEqual(store.durable[0].cells, before[0].cells);
  assert.deepEqual(store.durable[1], before[1]);
  const again = await run(store, task(style()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.execution.appliedOperations[0].evidence.observation, 'table-property-already-at-the-target-value');
});

test('FLAG: each of the four flags flips only itself; retry SATISFIED', async () => {
  for (const [name, value] of [['rowStripes', false], ['columnStripes', true], ['firstColumn', true], ['lastColumn', true]]) {
    const store = createStore(book());
    const { result } = await run(store, task(flag({ flag: name, value })));
    assert.equal(result.ok, true, `${name}: ${why(result)}`);
    assert.equal(tableOf(store).style[name], value);
    for (const other of ['rowStripes', 'columnStripes', 'firstColumn', 'lastColumn'].filter(x => x !== name)) assert.equal(tableOf(store).style[other], STYLE[other], `${name} left ${other}`);
    assert.equal((await run(store, task(flag({ flag: name, value })))).m.commands, 0);
  }
});

test('TOTALS ROW ON: the ref grows by one row, totalsRow true; retry SATISFIED', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(totals()));
  assert.equal(result.ok, true, why(result));
  assert.equal(tableOf(store).totalsRow, true);
  assert.equal(tableOf(store).ref, 'A1:C6');
  const again = await run(store, task(totals()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('TOTALS ROW OFF (GH #335 remainder): the Bridge itself refuses value !== true on /19..23 -- FAIL_CLOSED, named, 0 dispatch', async () => {
  await NO_WRITE(createStore(book({ tables: [T1_TOTALS_ON()] })), totals({ value: false }), {}, 'xlsx-table-totals-off-bridge-runtime-not-proven', e => { assert.equal(e.capability, 'CAPABILITY_MISSING'); assert.deepEqual([...e.proven], [RT24, RT25, RT26, RT27, RT28, RT29]); });
});

// GH #338-#342 remainder: TOTALS_OFF_BRIDGE_RUNTIMES is live now -- /24 is the ONE Bridge revision proven (static
// byte-identity warrant) to accept value:false, so these two tests run on the real runtime, no monkeypatch.
test('TOTALS ROW OFF on /24: the ref shrinks back by one row, totalsRow false; retry SATISFIED; no header-row check applies', async () => {
  const store = createStore(book({ tables: [T1_TOTALS_ON()] }));
  const { result, m } = await run(store, task(totals({ value: false })), { runtime: RT24 });
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).totalsRow, false);
  assert.equal(tableOf(store).ref, 'A1:C5');
  const again = await run(store, task(totals({ value: false })), { runtime: RT24 });
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.execution.appliedOperations[0].evidence.observation, 'table-property-already-at-the-target-value');
  // OFF does not require a header row (only the ON path's product creates a totals row off a header)
  const noHeader = await run(createStore(book({ tables: [T1_TOTALS_ON({ headerRow: false })] })), task(totals({ value: false })), { runtime: RT24 });
  assert.equal(noHeader.result.ok, true, why(noHeader.result));
});

// W2 regression fix: /25 carries the SAME sheet.table.write/v1 handler byte-identical to /24 (the only /25 change
// is workbook.protection.read/v1's null branch, an unrelated capability) -- the totals-OFF path must work on /25
// exactly as it does on /24, not fall back to the FAIL_CLOSED branch a stale list would wrongly give it.
test('TOTALS ROW OFF on /25 (W2 regression fix): the same byte-identical write handler as /24 -- not FAIL_CLOSED', async () => {
  const store = createStore(book({ tables: [T1_TOTALS_ON()] }));
  const { result, m } = await run(store, task(totals({ value: false })), { runtime: RT25 });
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).totalsRow, false);
  assert.equal(tableOf(store).ref, 'A1:C5');
});

test('TOTALS ROW OFF on /24, B2 (a maintainer review): a structured reference anywhere, or a plain formula touching the totals row itself, refuses before the write; an unrelated formula does not', async () => {
  await NO_WRITE(createStore(book({ tables: [T1_TOTALS_ON()], cells: { E9: cell(10, '=SUM(Table1[Qty])') } })), totals({ value: false }), { runtime: RT24 }, 'xlsx-table-structured-reference');
  await NO_WRITE(createStore(book({ tables: [T1_TOTALS_ON()], cells: { E9: cell(10, '=A6') } })), totals({ value: false }), { runtime: RT24 }, 'xlsx-table-totals-row-referenced', e => assert.ok(e.formulas.some(x => x.includes('E9'))));
  await NO_WRITE(createStore(book({ tables: [T1_TOTALS_ON()], cells: { E9: cell(10, '=SUM(A6:C6)') } })), totals({ value: false }), { runtime: RT24 }, 'xlsx-table-totals-row-referenced');
  // NEG: a formula that reaches a DATA row of the table, not the totals row -- the OFF write proceeds
  const store = createStore(book({ tables: [T1_TOTALS_ON()], cells: { E9: cell(10, '=A5') } }));
  const { result } = await run(store, task(totals({ value: false })), { runtime: RT24 });
  assert.equal(result.ok, true, why(result));
});

// GH #335 remainder (measured on a disposable test server): on a table with a totals row, the engine's OWN totals-row formula (e.g. a SUBTOTAL over a
// structured column reference) is itself a structured reference to the table -- the B2 check above, read literally,
// refused it on EVERY totals-row table, making totals OFF unusable in practice. Fix: the totals row's own cells are
// excluded from both checks (they are removed WITH the row, so they never go stale) -- everything outside it is not.
test('TOTALS ROW OFF on /24 (G2 fix): the totals row\'s OWN formula does not refuse -- it is removed with the row; an outside structured/A1 reference still does', async () => {
  const ownRow = { A6: cell('Summary'), C6: cell(0, '=SUBTOTAL(103,Table1[Note])') };
  const store = createStore(book({ tables: [T1_TOTALS_ON()], cells: { ...ownRow } }));
  const { result, m } = await run(store, task(totals({ value: false })), { runtime: RT24 });
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).totalsRow, false);
  assert.equal(tableOf(store).ref, 'A1:C5');
  // an OUTSIDE structured reference is still caught, even with the own-row formula also present
  await NO_WRITE(createStore(book({ tables: [T1_TOTALS_ON()], cells: { ...ownRow, E9: cell(10, '=SUM(Table1[Qty])') } })), totals({ value: false }), { runtime: RT24 }, 'xlsx-table-structured-reference');
  // an OUTSIDE A1 reference into the totals row is still caught too
  await NO_WRITE(createStore(book({ tables: [T1_TOTALS_ON()], cells: { ...ownRow, E9: cell(10, '=A6') } })), totals({ value: false }), { runtime: RT24 }, 'xlsx-table-totals-row-referenced');
});

// the own-row exclusion must forgive ONLY the totals row's own unobservable gap -- it must not skip
// the dependency reader's REACH scan (a validation / CF / chart / defined name that actually depends on the
// table). Unit-level (direct classify(), the obs() pattern this file already uses) because p.dependencies is
// hand-built, not derived from cells.
test('TOTALS ROW OFF on /24: an own-row-only unobservable gap is forgiven, but a real dependent elsewhere is still caught, and an unrelated gap still blocks', () => {
  const ownRowGap = [{ axis: 'formula:A:C6', why: 'UNOBSERVABLE:external-or-structured-reference', host: 'A' }];
  const base = deps => ({
    format: 'xlsx', complete: true, tier: 'full', runtime: { identity: RT24 }, sheets: [{ name: 'A' }],
    tableReads: { A: { tables: [T1_TOTALS_ON()] } }, cells: {}, formulas: {}, objects: { A: { pivots: [] } },
    dependencies: deps,
  });
  const off = fam.plan(totals({ value: false }), 0).operation;
  // own-row-only gap, nothing else depends on the table -- forgiven, APPLY
  const ok = fam.classify(base({ complete: false, unobservable: ownRowGap, refs: [] }), off);
  assert.equal(ok.state, 'APPLY', JSON.stringify(ok));
  // the SAME own-row-only gap, but a validation elsewhere on the sheet also depends on the table -- the reach
  // scan still runs (it is never narrowed) and catches it, even though the gap itself was forgiven
  const blocked = fam.classify(base({ complete: false, unobservable: ownRowGap, refs: [{ kind: 'validation', host: 'A', sheets: ['A'] }] }), off);
  assert.equal(blocked.state, 'FAIL_CLOSED');
  assert.equal(blocked.evidence.reason, 'xlsx-table-props-dependent-not-measured');
  // a genuinely UNRELATED unobservable gap (not confined to the totals row) is NOT forgiven
  const unrelated = fam.classify(base({ complete: false, unobservable: [{ axis: 'formula:A:E9', why: 'UNOBSERVABLE:external-or-structured-reference', host: 'A' }], refs: [] }), off);
  assert.equal(unrelated.state, 'FAIL_CLOSED');
  assert.equal(unrelated.evidence.reason, 'xlsx-structure-dependency-observation-incomplete');
});

test('RENAME: renamed, the cells untouched; the retry (the old name gone, one table of the new name) is SATISFIED -- table-already-renamed', async () => {
  const store = createStore(book());
  const before = snapshot(store);
  const { result, m } = await run(store, task(rename()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).name, 'Sales');
  assert.deepEqual(store.durable[0].cells, before[0].cells);
  const again = await run(store, task(rename()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.result.execution.appliedOperations[0].evidence.observation, 'table-already-renamed');
});

test('RESIZE: grow by two rows (empty extension) and shrink by one; each retry SATISFIED', async () => {
  const store = createStore(book());
  const grown = await run(store, task(resize()));
  assert.equal(grown.result.ok, true, why(grown.result));
  assert.equal(tableOf(store).ref, 'A1:C7');
  assert.equal((await run(store, task(resize()))).m.commands, 0);
  const store2 = createStore(book());
  const shrunk = await run(store2, task(resize({ ref: 'A1:C4' })));
  assert.equal(shrunk.result.ok, true, why(shrunk.result));
  assert.equal(tableOf(store2).ref, 'A1:C4');
  assert.equal(store2.durable[0].cells.A5.v, 'd', 'a cell cut off by a shrink stays where it is');
});

// ---- the refusals BEFORE the write -------------------------------------------------------------------------

const NO_WRITE = async (store, step, script, reason, extra) => {
  const before = snapshot(store);
  const { result, m } = await run(store, task(step), script);
  assert.equal(result.ok, false, `${step.intent}: should have been refused`);
  assert.equal(m.commands, 0, `${step.intent}: no command may be dispatched`);
  assert.equal(reasonOf(result), reason, JSON.stringify(result));
  assert.deepEqual(store.durable, before, `${step.intent}: nothing may change`);
  if (extra) extra(evidenceOf(result));
};

test('RUNTIME: on /18 (no sheet.table.write/v1) every intent is FAIL_CLOSED, named CAPABILITY_MISSING, nothing dispatched', async () => {
  for (const s of [style(), flag(), totals(), rename(), resize()]) {
    await NO_WRITE(createStore(book()), s, { runtime: RT18 }, 'xlsx-table-props-bridge-write-runtime-not-proven', e => { assert.equal(e.capability, 'CAPABILITY_MISSING'); assert.equal(e.runtime, RT18); });
  }
});

test('REFUSALS: not found / ambiguous; flag without style info; totals without a header, with no room, with content under the table; named, 0 commands', async () => {
  await NO_WRITE(createStore(book()), style({ table: 'Nope' }), {}, 'xlsx-table-not-found');
  await NO_WRITE(createStore(book()), style({ sheet: 'Missing' }), {}, 'xlsx-sheet-not-found');
  await NO_WRITE(createStore(book({ tables: [T1(), T1({ ref: 'F1:H5' })] })), style(), {}, 'xlsx-table-name-ambiguous');
  await NO_WRITE(createStore(book({ tables: [T1({ style: null })] })), flag(), {}, 'xlsx-table-style-info-missing');
  await NO_WRITE(createStore(book({ tables: [T1({ headerRow: false })] })), totals(), {}, 'xlsx-table-no-header-row');
  await NO_WRITE(createStore(book({ cells: { A6: cell('under') } })), totals(), {}, 'xlsx-table-totals-row-below-occupied');
  await NO_WRITE(createStore(book({ merges: ['A6:B6'], cells: { A6: cell(''), B6: cell('') } })), totals(), {}, 'xlsx-table-band-has-merge');
  await NO_WRITE(createStore(book({ tables: [T1({ ref: 'A1:C1048575' })] })), totals(), {}, 'xlsx-table-totals-row-no-room');
});

test('GH #332/#333 (owner decision A): resize no longer refuses a structured reference -- APPLY, the dependent formula untouched (resize does not need a rewrite, measured LIVE 6d4606b2 #ref)', async () => {
  const withRef = book({ cells: { E9: cell(10, '=SUM(Table1[Qty])') } });
  const store = createStore(withRef);
  const before = snapshot(store);
  const { result, m } = await run(store, task(resize()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(tableOf(store).ref, 'A1:C7');
  // the dependent formula is the ONLY thing resize does not touch -- same text, same value, as before
  assert.deepEqual(store.durable[0].cells.E9, before[0].cells.E9);
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS', JSON.stringify(l3));
});

test('GH #332/#333: resize -- the dependency reader cannot place a reference to ANOTHER table; the incomplete read still refuses (the ignore predicate forgives only THIS table\'s own refs)', async () => {
  const other = book({ cells: { E9: cell(10, '=SUM(Other[Qty])') } });
  await NO_WRITE(createStore(other), resize(), {}, 'xlsx-structure-dependency-observation-incomplete');
});

test('GH #332/#333 (owner decision A): rename rewrites a structured reference to the NEW name, in the SAME callCommand, on the table\'s own sheet AND another sheet', async () => {
  const withRef = book({ cells: { E9: cell(10, '=SUM(Table1[Qty])') } });
  withRef[1].cells.B2 = cell(5, '=COUNT(Table1[Qty])');
  const store = createStore(withRef);
  const { result, m } = await run(store, task(rename()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1, 'one callCommand carries the rename AND the rewrite');
  assert.equal(tableOf(store).name, 'Sales');
  assert.equal(store.durable[0].cells.E9.f, '=SUM(Sales[Qty])');
  assert.equal(store.durable[1].cells.B2.f, '=COUNT(Sales[Qty])');
  assert.ok(result.execution.appliedOperations[0].evidence.evidence.formulasRewritten.some(a => a.includes('E9')));
  assert.ok(result.execution.appliedOperations[0].evidence.evidence.formulasRewritten.some(a => a.includes('B2')));
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS', JSON.stringify(l3));
});

test('GH #332/#333: a rename with an in-table [@Qty] relative reference needs no rewrite there -- nothing to change, not refused', async () => {
  const withRelative = book({ cells: { B2: cell(1, '=[@Qty]*2') } });
  const store = createStore(withRelative);
  const { result } = await run(store, task(rename()));
  assert.equal(result.ok, true, why(result));
  assert.equal(store.durable[0].cells.B2.f, '=[@Qty]*2', 'untouched: it never named the table');
});

test('GH #332/#333 MUTATION CONTROL: if the rewrite step is skipped, the OLD name stays in the formula and L3 catches it', async () => {
  const withRef = book({ cells: { E9: cell(10, '=SUM(Table1[Qty])') } });
  const store = createStore(withRef);
  const { result } = await run(store, task(rename()), { fault: 'tp-rename-rewrite-skipped' });
  assert.equal(result.ok, false, why(result));
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'FAIL');
  assert.ok(l3.evidence.evidence.cellsChangedOutsideTheRegion.some(x => x.includes('E9')));
});

test('REFUSALS: rename / resize while a structured reference the rewrite cannot safely reach (INDIRECT text) exists -- the measured #NAME?', async () => {
  const withIndirect = book({ cells: { E9: cell(10, '=INDIRECT("Table1[Qty]")') } });
  await NO_WRITE(createStore(withIndirect), rename(), {}, 'xlsx-table-structured-reference', e => assert.ok(e.formulas.some(x => x.includes('E9'))));
  // a structured reference whose table name is NOT any table the read found (here 'Other' names no real
  // table -- 'Other' is only a SHEET in this fixture) stays unresolvable, refused, named -- conservative,
  // never a guess about whose reference it is
  const other = book({ cells: { E9: cell(10, '=SUM(Other[Qty])') } });
  await NO_WRITE(createStore(other), rename(), {}, 'xlsx-structure-dependency-observation-incomplete');
});

// GH #332/#333 structured-ref-rewrite research probe (prod, eb2fca05 #ref): the dependency reader
// (dependencies.cjs) marks EVERY structured-reference formula in the whole workbook unobservable, not
// just the ones naming the table being renamed/resized -- so a DIFFERENT table's own dependent formula,
// unrelated to THIS rename/resize, was blocking it too. Forgiven only when the formula provably names a
// table the read actually found (not an external-workbook reference, and not an unresolvable name).
test('RENAME: a structured reference to a DIFFERENT, existing table no longer blocks -- the dependency is provably unrelated', async () => {
  const withOtherTableRef = book({ tables: [T1(), T1({ name: 'Other2', ref: 'F1:H5' })], cells: { E9: cell(10, '=SUM(Other2[Qty])') } });
  const store = createStore(withOtherTableRef);
  const { result, m } = await run(store, task(rename()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(store.durable[0].cells.E9.f, '=SUM(Other2[Qty])', 'untouched: it names a different table, not the one renamed');
});

test('RESIZE: a structured reference to a DIFFERENT, existing table no longer blocks', async () => {
  const withOtherTableRef = book({ tables: [T1(), T1({ name: 'Other2', ref: 'F1:H5' })], cells: { E9: cell(10, '=SUM(Other2[Qty])') } });
  const store = createStore(withOtherTableRef);
  const { result, m } = await run(store, task(resize()));
  assert.equal(result.ok, true, why(result));
  assert.equal(m.commands, 1);
  assert.equal(store.durable[0].cells.E9.f, '=SUM(Other2[Qty])');
});

// eb2fca05 v4 (LIVE review #ref): v1-v3's otherTableAxisIgnore only ever saw tableReads for the
// operation's OWN sheet (liveExtras asked tableSheets: [op.selector.sheet]) -- every prior "different
// table" test above puts that other table on the SAME sheet, which hid this. MEASURED live (prod, DS
// cell /26): the dependency reader marks formula:S!K2 / formula:S2!K2 / formula:S2!K3 all unobservable,
// and the forgiveness predicate never fires because `others` (from tableReads) is empty when the other
// table lives on a different sheet.
test('RENAME/RESIZE: a structured reference to a DIFFERENT table on ANOTHER SHEET no longer blocks -- MEASURED LIVE gap (eb2fca05 #ref)', async () => {
  const crossSheet = [
    { name: 'A', cells: { A1: cell('Item'), B1: cell('Qty'), C1: cell('Note'), A2: cell('a'), B2: cell(1), C2: cell('x'), A3: cell('b'), B3: cell(2), C3: cell('y'), A4: cell('c'), B4: cell(3), C4: cell('z'), A5: cell('d'), B5: cell(4), C5: cell('w'), E9: cell(10, '=SUM(Other2[Qty])') }, tables: [T1()], merges: [], objects: { charts: [], pivots: [], names: [] } },
    { name: 'S2', cells: { F1: cell('Item'), G1: cell('Qty'), H1: cell('Note'), F2: cell('x'), G2: cell(5), H2: cell('y') }, tables: [T1({ name: 'Other2', ref: 'F1:H2' })], objects: { charts: [], pivots: [], names: [] } },
  ];
  for (const step of [rename(), resize()]) {
    const store = createStore(crossSheet.map(s => JSON.parse(JSON.stringify(s))));
    const { result, m } = await run(store, task(step));
    assert.equal(result.ok, true, `${step.intent}: ${why(result)}`);
    assert.equal(m.commands, 1, step.intent);
    assert.equal(store.durable[0].cells.E9.f, '=SUM(Other2[Qty])', `${step.intent}: untouched (different table, different sheet)`);
  }
});

// eb2fca05 v3 (LIVE review #ref): v2's otherTableAxisIgnore forgave the whole axis the moment ANY known
// other table's name matched ANYWHERE in the raw formula text -- so a SECOND, unresolvable reference in
// the SAME formula (one `others.some()` never looked for) slipped through unnoticed, and a known name
// sitting INSIDE a string literal (never reachable by any rewrite) counted the same as a real reference.
test('RENAME/RESIZE: a known-other-table reference alongside an UNRESOLVABLE one in the SAME formula still blocks -- the unresolvable half is not waved through', async () => {
  const withMixed = book({ tables: [T1(), T1({ name: 'Other2', ref: 'F1:H5' })], cells: { E9: cell(10, '=SUM(Other2[Qty])+Ghost[Qty]') } });
  await NO_WRITE(createStore(withMixed), rename(), {}, 'xlsx-structure-dependency-observation-incomplete');
  await NO_WRITE(createStore(withMixed), resize(), {}, 'xlsx-structure-dependency-observation-incomplete');
});

test('RENAME/RESIZE: a known table name that occurs ONLY inside a string literal (INDIRECT text), alongside an unresolvable reference outside it, still blocks', async () => {
  const withLiteralOnly = book({ tables: [T1(), T1({ name: 'Other2', ref: 'F1:H5' })], cells: { E9: cell(10, '=Ghost[Qty]+INDIRECT("Other2[Qty]")') } });
  await NO_WRITE(createStore(withLiteralOnly), rename(), {}, 'xlsx-structure-dependency-observation-incomplete');
  await NO_WRITE(createStore(withLiteralOnly), resize(), {}, 'xlsx-structure-dependency-observation-incomplete');
});

test('MUTATION CONTROL: a formula naming a known other table ALSO sitting inside a literal must not be forgiven by scanning the whole formula text (unparsed) instead of its tokens', () => {
  // direct classify()-level check, same pattern as the existing different-table MUTATION CONTROL above:
  // 'Other2[' appears ONLY inside the string literal here -- a v2-shaped scan of the raw formula text
  // (not tokenize()'d) would find it and wrongly forgive; the fix must look outside literals only.
  const base = (deps, tables) => ({
    format: 'xlsx', complete: true, tier: 'full', runtime: { identity: RT24 }, sheets: [{ name: 'A' }, { name: 'Other' }],
    tableReads: { A: { tables } }, cells: { 'A\u0000E9': { f: '=Ghost[Qty]+INDIRECT("Other2[Qty]")' } }, formulas: {}, objects: { A: { pivots: [] }, Other: { pivots: [] } },
    dependencies: deps,
  });
  const op = fam.plan(rename(), 0).operation;
  const refused = fam.classify(base({ complete: false, unobservable: [{ axis: 'formula:A:E9', why: 'UNOBSERVABLE:external-or-structured-reference', host: 'A' }], refs: [] }, [T1(), T1({ name: 'Other2', ref: 'F1:H5' })]), op);
  assert.equal(refused.state, 'FAIL_CLOSED');
  assert.equal(refused.evidence.reason, 'xlsx-structure-dependency-observation-incomplete');
});

test('RENAME/RESIZE: an external-workbook reference anywhere still blocks (unresolvable, conservative)', async () => {
  const withExternal = book({ cells: { E9: cell(10, '=[1]Sheet1!A1') } });
  await NO_WRITE(createStore(withExternal), rename(), {}, 'xlsx-structure-dependency-observation-incomplete');
  await NO_WRITE(createStore(withExternal), resize(), {}, 'xlsx-structure-dependency-observation-incomplete');
});

test('MUTATION CONTROL: the different-table forgiveness must not widen to an unresolvable name -- flip the regex to always-true and the unresolvable-name refusal above would wrongly pass', () => {
  // direct classify()-level check (the obs() pattern this file already uses): an unobservable axis whose
  // formula names no table the read found stays un-ignored even under the widened predicate's OWN axis set
  const base = (deps, tables) => ({
    format: 'xlsx', complete: true, tier: 'full', runtime: { identity: RT24 }, sheets: [{ name: 'A' }, { name: 'Other' }],
    tableReads: { A: { tables } }, cells: { 'A\u0000E9': { f: '=SUM(Ghost[Qty])' } }, formulas: {}, objects: { A: { pivots: [] }, Other: { pivots: [] } },
    dependencies: deps,
  });
  const op = fam.plan(rename(), 0).operation;
  const refused = fam.classify(base({ complete: false, unobservable: [{ axis: 'formula:A:E9', why: 'UNOBSERVABLE:external-or-structured-reference', host: 'A' }], refs: [] }, [T1()]), op);
  assert.equal(refused.state, 'FAIL_CLOSED');
  assert.equal(refused.evidence.reason, 'xlsx-structure-dependency-observation-incomplete');
});

// GH #332/#333 (owner review #ref): `unobservable` is the FULL list the dependency reader produced --
// an axis appearing there is NOT proof it was the one still blocking (a forgiven axis stays in the list
// too). `blocking` narrows to the subset the ignore predicate did not forgive -- the actual cause.
test('dependencyConflict: `blocking` separates the un-forgiven axis from the full `unobservable` list', () => {
  const base = (deps, tables) => ({
    format: 'xlsx', complete: true, tier: 'full', runtime: { identity: RT24 }, sheets: [{ name: 'A' }, { name: 'Other' }],
    tableReads: { A: { tables } }, cells: { 'A\u0000E9': { f: '=SUM(Other2[Qty])' }, 'A\u0000E10': { f: '=SUM(Ghost[Qty])' } }, formulas: {}, objects: { A: { pivots: [] }, Other: { pivots: [] } },
    dependencies: deps,
  });
  const op = fam.plan(rename(), 0).operation;
  const unobservable = [
    { axis: 'formula:A:E9', why: 'UNOBSERVABLE:external-or-structured-reference', host: 'A' },
    { axis: 'formula:A:E10', why: 'UNOBSERVABLE:external-or-structured-reference', host: 'A' },
  ];
  const refused = fam.classify(base({ complete: false, unobservable, refs: [] }, [T1(), T1({ name: 'Other2', ref: 'F1:H5' })]), op);
  assert.equal(refused.state, 'FAIL_CLOSED');
  assert.equal(refused.evidence.reason, 'xlsx-structure-dependency-observation-incomplete');
  assert.deepEqual(refused.evidence.unobservable.map(u => u.axis), ['formula:A:E9', 'formula:A:E10'], 'the full list keeps the forgiven axis too');
  assert.deepEqual(refused.evidence.blocking.map(u => u.axis), ['formula:A:E10'], 'only the un-forgiven axis blocks');
});

test('REFUSALS: rename onto a name already taken on the sheet', async () => {
  await NO_WRITE(createStore(book({ tables: [T1(), T1({ name: 'Sales', ref: 'F1:H5' })] })), rename(), {}, 'xlsx-table-name-taken');
  await NO_WRITE(createStore(book({ tables: [T1(), T1({ name: 'SALES', ref: 'F1:H5' })] })), rename(), {}, 'xlsx-table-name-taken');
});

test('REFUSALS: resize -- another origin, a width change, a height under 2, an overlap, data / a merge / a pivot in the extension, a totals row', async () => {
  await NO_WRITE(createStore(book()), resize({ ref: 'B1:C7' }), {}, 'xlsx-table-resize-moves-origin');
  await NO_WRITE(createStore(book()), resize({ ref: 'A1:D7' }), {}, 'xlsx-table-resize-width-change-unmeasured');
  await NO_WRITE(createStore(book()), resize({ ref: 'A1:C1' }), {}, 'xlsx-table-resize-too-small');
  await NO_WRITE(createStore(book({ tables: [T1(), T1({ name: 'Two', ref: 'A7:C9' })] })), resize({ ref: 'A1:C8' }), {}, 'xlsx-table-resize-overlaps-table');
  await NO_WRITE(createStore(book({ cells: { A6: cell('x') } })), resize({ ref: 'A1:C7' }), {}, 'xlsx-table-resize-extension-has-data');
  await NO_WRITE(createStore(book({ merges: ['A6:B7'], cells: { A6: cell(''), A7: cell('') } })), resize({ ref: 'A1:C7' }), {}, 'xlsx-table-band-has-merge');
  await NO_WRITE(createStore(book({ tables: [T1({ totalsRow: true })] })), resize(), {}, 'xlsx-table-totals-row-unmeasured');
});

test('REFUSAL (synthetic observation -- the fake driver exposes pivot NAMES only): a pivot in the extension / the row under the table; a pivot read that cannot place it', () => {
  const obs = pivots => ({
    format: 'xlsx', complete: true, tier: 'full', runtime: { identity: RT19 }, sheets: [{ name: 'A' }],
    tableReads: { A: { tables: [T1()] } }, cells: { 'A\u0000A1': { v: 'Item', f: '', merge: 'A1' } }, formulas: {}, objects: { A: { pivots } },
    dependencies: { complete: true, refs: [], unobservable: [] },
  });
  const cls = (step, pivots) => fam.classify(obs(pivots), fam.plan(step, 0).operation);
  const hit = cls(resize({ ref: 'A1:C7' }), [{ name: 'P1', table: '$A$6:$C$9' }]);
  assert.equal(hit.state, 'FAIL_CLOSED');
  assert.equal(hit.evidence.reason, 'xlsx-table-band-has-pivot');
  const under = cls(totals(), [{ name: 'P1', table: '$A$6:$C$9' }]);
  assert.equal(under.evidence.reason, 'xlsx-table-band-has-pivot');
  assert.equal(cls(resize({ ref: 'A1:C7' }), [{ name: 'P1', table: '$A$20:$C$25' }]).state, 'APPLY', 'a pivot elsewhere does not block');
  assert.equal(cls(resize({ ref: 'A1:C7' }), [{ name: 'P1', table: 'garbage' }]).evidence.reason, 'xlsx-pivots-unobservable');
});

test('REFUSAL: an observation that cannot tell the tables is named, not guessed', async () => {
  await NO_WRITE(createStore(book()), style(), { tablesUnobservable: true }, 'xlsx-tables-unobservable');
});

// ---- the verification RED: a handler that answers ok while the state is wrong -----------------------------------

test('RED/GREEN: a Bridge write that reports ok but changed NOTHING fails the independent verification (all five intents)', async () => {
  for (const s of [style(), flag(), totals(), rename(), resize()]) {
    const store = createStore(book());
    const { result } = await run(store, task(s), { fault: 'tp-ok-but-nothing' });
    assert.equal(result.ok, false, `${s.intent}: a no-op that claims success must not pass`);
  }
});

test('RED/GREEN on /24 (GH #335 remainder): a totals-OFF Bridge write that reports ok but changed NOTHING fails the independent verification', async () => {
  const store = createStore(book({ tables: [T1_TOTALS_ON()] }));
  const { result } = await run(store, task(totals({ value: false })), { fault: 'tp-ok-but-nothing', runtime: RT24 });
  assert.equal(result.ok, false, 'a no-op that claims success must not pass');
});

test('the fake driver mirrors the real Bridge INDEPENDENTLY of classify(): a direct dispatch of the lowered command refuses on a runtime the capability has not been granted to, even though the operation passed plan() and is marked APPLY', async () => {
  const d = createFakeXlsxDriver(createStore(book({ tables: [T1_TOTALS_ON()] })), { runtime: RT19 });
  await d.open({ fileId: 'disposable' });
  const lowered = fam.lower(fam.plan(totals({ value: false }), 0).operation, { sheet: 'A', table: 'Table1', ref: 'A1:C6' });
  const r = await d.dispatch(null, lowered);
  assert.equal(r.ok, false);
  assert.equal(r.evidence.reason, 'BRIDGE:totals-off-not-measured');
  assert.equal(r.evidence.mutation, 'NONE');
});

test('RED: a write that also flips another field of the table / changes ANOTHER table / writes a cell outside is not accepted', async () => {
  for (const [fault, s, store] of [
    ['tp-collateral', style(), createStore(book())],
    ['tp-other-table', style(), createStore(book({ tables: [T1(), T1({ name: 'Two', ref: 'F1:H5' })] }))],
    ['tp-cell-side-effect', flag(), createStore(book())],
  ]) {
    const { result } = await run(store, task(s), { fault });
    assert.equal(result.ok, false, fault);
  }
});

test('RED: a rename that leaves a #NAME? error value in a formula cell is not accepted', async () => {
  const { result } = await run(createStore(book()), task(rename()), { fault: 'tp-rename-breaks-formula' });
  assert.equal(result.ok, false);
});

test('a handler refusal (the engine ignores the call, or the gate says CAPABILITY_MISSING) is a failed operation, never a success', async () => {
  const ignored = await run(createStore(book()), task(style()), { fault: 'ignore' });
  assert.equal(ignored.result.ok, false);
  const gated = await run(createStore(book()), task(style()), { tpWrite: 'gate-refused' });
  assert.equal(gated.result.ok, false);
});

// ---- the lowered command ------------------------------------------------------------------------------------

test('lower: ONE callCommand, the exact sheet check, the Bridge gate, ONE write with the property args; meta carries the capability marker', () => {
  const cases = [
    [fam.plan(style(), 0).operation, { property: 'style', name: 'TableStyleMedium9' }],
    [fam.plan(flag({ flag: 'columnStripes', value: true }), 0).operation, { property: 'flag', flag: 'columnStripes', value: true }],
    [fam.plan(totals(), 0).operation, { property: 'totalsRow', value: true }],
    [fam.plan(totals({ value: false }), 0).operation, { property: 'totalsRow', value: false }],
    [fam.plan(rename(), 0).operation, { property: 'name', name: 'Sales' }],
    [fam.plan(resize(), 0).operation, { property: 'ref', ref: 'A1:C7' }],
  ];
  for (const [op, args] of cases) {
    const c = fam.lower(op, { sheet: 'A', table: 'Table1', ref: 'A1:C5' });
    assert.equal(c.kind, 'callCommand');
    assert.equal(c.meta.family, 'table-props');
    assert.equal(c.meta.bridgeTableWrite, true);
    assert.match(c.body, /__eoTableGate\(\)/);
    assert.equal((c.body.match(/__eoTableWrite\(/g) || []).length, 1);
    assert.ok(c.body.includes(JSON.stringify({ sheet: 'A', table: 'Table1', ...args })), c.body);
    assert.doesNotThrow(() => new Function('Api', `return (async () => { const __eoTableGate = () => ({ ok: true }); const __eoTableWrite = () => ({ ok: true }); const __eoTableFail = w => w; ${c.body} })()`));
  }
});

test('sole step: a table property step cannot share a task with another step', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, { id: 'two', steps: [style(), flag()] });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
});

test('the driver second line: the SAME /19 + /21 list as the adapter; a table-props command is refused WITHOUT an editor elsewhere; an unmarked one is not', async () => {
  const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  assert.deepEqual([...fam.B25_BRIDGE_WRITE_RUNTIMES], [...drv.B25_BRIDGE_WRITE_RUNTIMES]);
  assert.deepEqual([...drv.B25_BRIDGE_WRITE_RUNTIMES], [RT19, '9.3.5|euro-office-bridge-v1/9.3.5.0/21', '9.3.5|euro-office-bridge-v1/9.3.5.0/22', '9.3.5|euro-office-bridge-v1/9.3.5.0/23', RT24, RT25, RT26, RT27, RT28, RT29]);
  const d = drv.createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const lowered = fam.lower(fam.plan(style(), 0).operation, { sheet: 'A', table: 'Table1', ref: 'A1:C5' });
  const r = await d.dispatch(null, lowered);
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b25-bridge-write-runtime-not-proven']);
  assert.equal(fam.lower(fam.plan(rename(), 0).operation, { sheet: 'A', table: 'Table1', ref: 'A1:C5' }).meta.bridgeTableWrite, true);
});

test('the in-editor prelude: gate, ONE write, fail -- parses, uses the /19-only runtime entries and the sheet.table.write/v1 id', () => {
  const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  const pre = drv.tableWriteBridgeWritePrelude();
  assert.ok(pre.includes('"sheet.table.write/v1"'));
  assert.ok(pre.includes('euro-office-bridge-v1/9.3.5.0/19') && !pre.includes('euro-office-bridge-v1/9.3.5.0/18'));
  assert.doesNotThrow(() => new Function('Api', `return (async () => { ${pre} return [typeof __eoTableGate, typeof __eoTableWrite, typeof __eoTableFail]; })()`));
});
