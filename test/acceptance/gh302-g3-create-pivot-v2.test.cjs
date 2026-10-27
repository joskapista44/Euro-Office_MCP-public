'use strict';

/**
 * GH #302 G3: create_pivot v2 -- n row / column / page fields, several data fields with functions, a row-axis layout, a
 * style name, onto an EMPTY pivotSheet (the G3 LIVE probe: a page area goes above `at` and a too-high `at` is shifted, so
 * the extent cannot be pre-computed). The 1 + 1 + Sum shape stays v1. Identity: the G2 axes, each data field's function
 * AND source field (GetSourceName), the row fields' layout form, the style; the source in session, the Bridge dataRef
 * after a reopen.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const pivots = require('../../src/adapters/xlsx/families/pivots.cjs');
const v2 = require('../../src/adapters/xlsx/families/pivots-v2.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const STEP = Object.freeze({ intent: 'create_pivot', name: 'Sales2', sheet: 'PS', range: 'A1:D7', rows: ['Region', 'Product'], columns: ['Quarter'], pages: [],
  data: [{ field: 'Amount', function: 'Sum' }, { field: 'Amount', function: 'Count' }], layout: 'Tabular', style: 'PivotStyleMedium9', pivotSheet: 'PV', at: 'A3' });
const stepOf = (extra = {}) => ({ ...STEP, ...extra });
const opOf = s => { const p = pivots.plan(s, 0); assert.equal(p.ok, true, p.detail); return p.operation; };

const HEADER = { 'PS\u0000A1': { v: 'Region', f: '', count: 0 }, 'PS\u0000B1': { v: 'Product', f: '', count: 0 }, 'PS\u0000C1': { v: 'Quarter', f: '', count: 0 }, 'PS\u0000D1': { v: 'Amount', f: '', count: 0 } };
const detail = (extra = {}) => ({ name: 'Sales2', source: 'PS!$A$1:$D$7', table: 'A3:H12', layout: { rows: ['Region', 'Product'], cols: ['Quarter'], data: ['Sum of Amount', 'Sum of Amount2'] },
  pages: [], values: { axis: 'columns', position: 2 }, dataFields: [{ name: 'Sum of Amount', function: 'Sum', field: 'Amount' }, { name: 'Sum of Amount2', function: 'Count', field: 'Amount' }],
  rowLayout: 'Tabular', style: 'PivotStyleMedium9', dataRef: 'PS!A1:D7', ...extra });
function obs({ pvUsed = null, pvPivots = [], other = { charts: [], names: [], pivots: [] }, pvCells = {} } = {}) {
  return { format: 'xlsx', complete: true, sheets: [{ name: 'PS', index: 0, used: 'A1:D7' }, { name: 'PV', index: 1, used: pvUsed }], cells: { ...HEADER, ...pvCells },
    objects: { PS: other, PV: { charts: [], names: [], pivots: pvPivots.length ? 'UNOBSERVABLE:pivot-layout-not-measured (pivot Sales2)' : [], ...(pvPivots.length ? { pivotDetails: pvPivots } : {}) } } };
}

test('G3 schema + plan: a v2 step is accepted; v1 (rows + columns + sum) still plans as v1; the mixes and unmeasured shapes are refused by name', () => {
  assert.equal(W.validateWriteRequest(req([stepOf()])).ok, true);
  const v1 = pivots.plan({ intent: 'create_pivot', name: 'P1', sheet: 'S', range: 'A1:C7', rows: 'reg', columns: 'prod', sum: 'amt', pivotSheet: 'H', at: 'H3' }, 0);
  assert.equal(v1.ok, true, v1.detail);
  assert.equal(v1.operation.v2, undefined);
  const bad = (extra, why) => { const p = pivots.plan(stepOf(extra), 0); assert.equal(p.ok, false, JSON.stringify(extra)); assert.match(p.detail, why); };
  bad({ sum: 'Amount' }, /does not support field\(s\): sum/);
  bad({ pivotSheet: 'PS' }, /another sheet than the source/);
  bad({ pages: ['Region'] }, /one of rows \/ columns \/ pages only/);
  bad({ data: [{ field: 'Amount', function: 'Median' }] }, /function one of/);
  bad({ data: [{ field: 'Amount', function: 'Sum' }, { field: 'Amount', function: 'Sum' }] }, /same function twice/);
  bad({ rows: [] }, /rows 1\.\.4/);
  bad({ layout: 'Grid' }, /layout must be one of/);
  bad({ style: 'MyStyle' }, /built-in pivot style/);
  bad({ data: [{ field: 'Amount', function: 'Sum', author: 'x' }] }, /data must be/);
  assert.equal(W.validateWriteRequest(req([stepOf({ data: [{ field: 'Amount', function: 'Median' }] })])).ok, false, 'the schema enum refuses it too');
});

test('G3 classify: an EMPTY pivotSheet + a known header -> APPLY; a non-empty pivotSheet / a field not in the header / an unread sheet -> named FAIL_CLOSED', () => {
  const op = opOf(stepOf());
  assert.equal(v2.classify(obs(), op).state, 'APPLY');
  const used = v2.classify(obs({ pvUsed: 'B2', pvCells: { 'PV\u0000B2': { v: 'x', f: '' } } }), op);
  assert.equal(used.state, 'FAIL_CLOSED');
  assert.equal(used.evidence.reason, 'xlsx-pivot-v2-destination-sheet-not-empty');
  assert.deepEqual(used.evidence.cells, ['B2']);
  // release LIVE run 1: a NEW empty sheet reports used 'A1' with a blank A1 -- that is empty
  assert.equal(v2.classify(obs({ pvUsed: 'A1', pvCells: { 'PV\u0000A1': { v: '', f: '' } } }), op).state, 'APPLY');
  const nofield = v2.classify(obs(), opOf(stepOf({ rows: ['Region', 'City'] })));
  assert.equal(nofield.evidence.reason, 'xlsx-pivot-field-not-in-header');
  const unread = obs(); unread.objects.PS = 'UNOBSERVABLE:threw';
  assert.equal(v2.classify(unread, op).evidence.reason, 'xlsx-pivot-observation-incomplete');
});

test('G3 classify: the same pivot already there -> SATISFIED (also after a reopen: the source by the Bridge dataRef); any differing part -> FAIL_CLOSED naming it', () => {
  const op = opOf(stepOf());
  assert.equal(v2.classify(obs({ pvUsed: 'A3:H12', pvPivots: [detail()] }), op).state, 'SATISFIED');
  assert.equal(v2.classify(obs({ pvUsed: 'A3:H12', pvPivots: [detail({ source: 'UNOBSERVABLE:pivot-layout-not-measured' })] }), op).state, 'SATISFIED', 'reopened: dataRef');
  for (const [extra, part] of [[{ style: 'PivotStyleLight16' }, 'style'], [{ rowLayout: 'Outline' }, 'layout'], [{ pages: ['Product'] }, 'pages'],
    [{ dataFields: [{ name: 'Sum of Amount', function: 'Sum', field: 'Amount' }] }, 'data'], [{ source: 'UNOBSERVABLE:x', dataRef: 'PS!A1:D6' }, 'source']]) {
    const c = v2.classify(obs({ pvUsed: 'A3:H12', pvPivots: [detail(extra)] }), op);
    assert.equal(c.state, 'FAIL_CLOSED', part);
    assert.equal(c.evidence.differs, part);
  }
});

test('G3 lower: one command -- empty-sheet re-check, insert, name, the field lists, each data field with its function, layout, style, refresh', () => {
  const body = v2.lower(opOf(stepOf())).body;
  assert.match(body, /destination-not-empty/);
  assert.match(body, /InsertPivotExistingWorksheet\(src\.GetRange\("A1:D7"\), host\.GetRange\("A3"\)\)/);
  assert.match(body, /p\.SetName\("Sales2"\)/);
  assert.match(body, /p\.AddFields\(\{"rows":\["Region","Product"\],"columns":\["Quarter"\]\}\)/);
  assert.match(body, /if \(d\.function !== 'Sum'\) f\.SetFunction\(d\.function\)/);
  assert.match(body, /p\.SetRowAxisLayout\("Tabular", false\)/);
  assert.match(body, /p\.SetStyleName\("PivotStyleMedium9"\)/);
  const bare = v2.lower(opOf(stepOf({ layout: undefined, style: undefined, columns: [] }))).body;
  assert.doesNotMatch(bare, /SetRowAxisLayout|SetStyleName|"columns"/);
});

test('G3 verify: L1 needs the whole identity; L3 keeps every other pivot', () => {
  const op = opOf(stepOf());
  assert.equal(v2.verifyPrimitive(op, null, { primary: obs({ pvUsed: 'A3:H12', pvPivots: [detail()] }) }).ok, true);
  assert.equal(v2.verifyPrimitive(op, null, { primary: obs({ pvUsed: 'A3:H12', pvPivots: [detail({ style: 'PivotStyleLight16' })] }) }).ok, false);
  const other = { charts: [], names: [], pivots: [{ name: 'Old', source: 'PS!$A$1:$D$7' }] };
  const kept = v2.verifyPreservation(op, { primary: obs({ other }) }, { primary: obs({ pvUsed: 'A3:H12', pvPivots: [detail()], other }) });
  assert.equal(kept.evidence.otherObjectsKept, true);
  const lost = v2.verifyPreservation(op, { primary: obs({ other }) }, { primary: obs({ pvUsed: 'A3:H12', pvPivots: [detail()] }) });
  assert.equal(lost.ok, false);
  // release LIVE run 2: the table's own cells APPEAR on the (empty before) pivotSheet -- not a foreign change; a source cell is
  const table = { 'PV\u0000A3': { v: 'Sum of Amount', f: '' }, 'PV\u0000C6': { v: '10', f: '' } };
  const grown = v2.verifyPreservation(op, { primary: obs() }, { primary: obs({ pvUsed: 'A3:H12', pvPivots: [detail()], pvCells: table }) });
  assert.equal(grown.ok, true, JSON.stringify(grown.evidence));
  const touched = obs({ pvUsed: 'A3:H12', pvPivots: [detail()], pvCells: table }); touched.cells = { ...touched.cells, 'PS\u0000D1': { v: 'Total', f: '', count: 0 } };
  assert.equal(v2.verifyPreservation(op, { primary: obs() }, { primary: touched }).ok, false, 'a changed source header is a foreign change');
});

test('G3 routing: the adapter plans a v2 step through pivots-v2 and reads every sheet\'s objects', () => {
  const p = planOperations({ id: 't', steps: [stepOf()] });
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operations[0].v2, true);
  assert.deepEqual(liveScope({ id: 't', steps: [stepOf()] }).objectSheets, ['*']);
});
