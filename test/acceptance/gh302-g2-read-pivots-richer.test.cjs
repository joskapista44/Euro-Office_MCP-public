'use strict';

/**
 * GH #302 G2: a richer read_pivots -- row / column / page field names, where the Values pseudo-field sits, every data
 * field with its function, the style name -- and a richer pivot no longer turns the whole sheet into a gap on the READ
 * surface (the write guards' pivots axis is unchanged). The driver's axes read runs against a fake Api shaped by the G2
 * LIVE probe (docs/evidence/gh302-g2-pivot-read-probe-20261026.json): the Values item's GetName throws, its GetIndex is -1.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { withProvenPivots } = require('../../src/adapters/xlsx/pivot-source.cjs');
const rt = require('../../src/server/read-tools.cjs');

const field = (name, index, orientation = 'Rows', position = 1) => ({ GetName: () => name, GetIndex: () => index, GetOrientation: () => orientation, GetPosition: () => position, GetLayoutForm: () => 'Tabular', GetLayoutCompactRow: () => false, GetLayoutSubtotals: () => true });
const VALUES = { GetName: () => { throw new Error("Cannot read properties of undefined (reading 'asc_getName')"); }, GetIndex: () => -1 };
const broken = { GetName: () => { throw new Error('boom'); }, GetIndex: () => 4 };
const dataField = (name, fn, field = 'Amount') => ({ GetName: () => name, GetFunction: () => fn, GetSourceName: () => field });
const range = a => ({ GetAddress: () => a, GetValue: () => [['']] });

function pivot({ rows = [field('Region', 1)], cols = [field('Quarter', 3), VALUES], pages = [field('Product', 2)], style = 'PivotStyleMedium9', pagesThrow = false, pivotFields = null } = {}) {
  return {
    GetName: () => 'PivotTable1', GetSource: () => { throw new Error("Cannot read properties of undefined (reading 'ws')"); }, GetTableRange1: () => range('A3:G8'), GetDataBodyRange: () => range('B6:G8'),
    GetRowFields: () => rows, GetColumnFields: () => cols, GetPageFields: () => { if (pagesThrow) throw new Error('no page fields'); return pages; },
    GetDataFields: () => [dataField('Sum of Amount', 'Sum'), dataField('Sum of Amount2', 'Count')], GetPivotFields: () => { if (pivotFields === 'throw') throw new Error('fields'); return pivotFields || [field('Region', 1, 'Rows'), field('Product', 2, 'Filters'), field('Quarter', 3, 'Columns')]; },
    GetStyleName: () => style,
  };
}
function run(p) {
  const ws = { GetName: () => 'PV', GetIndex: () => 0, GetUsedRange: () => null, GetAllCharts: () => [], GetAllPivotTables: () => [p], GetDefNames: () => [] };
  const Api = { GetSheets: () => [ws], GetSheet: () => ws, WorksheetFunction: { COUNT: () => 0 } };
  return new Function('Api', readQuery([], { objectSheets: ['PV'], usedRanges: false }))(Api).objects.PV.pivots[0];
}

test('G2 driver: rows / cols / pages by name, the Values pseudo-field as a POSITION (GetIndex -1), the style name', () => {
  const r = run(pivot());
  assert.deepEqual(r.axes, { rows: ['Region'], cols: ['Quarter'], pages: ['Product'], values: { axis: 'columns', position: 2 } });
  assert.equal(r.style, 'PivotStyleMedium9');
  assert.deepEqual(r.functions, [{ name: 'Sum of Amount', function: 'Sum', field: 'Amount' }, { name: 'Sum of Amount2', function: 'Count', field: 'Amount' }]);
});

test('G2 driver: Values on the row axis; a single data field -> no Values item -> values null', () => {
  assert.deepEqual(run(pivot({ rows: [field('Region', 1), VALUES], cols: [field('Quarter', 3)] })).axes.values, { axis: 'rows', position: 2 });
  assert.equal(run(pivot({ cols: [field('Quarter', 3)] })).axes.values, null);
});

test('G2 driver: a throwing item that is NOT the Values pseudo-field is its own named gap; a throwing list getter names the list', () => {
  const r = run(pivot({ cols: [field('Quarter', 3), broken], pagesThrow: true, pivotFields: 'throw' }));
  assert.deepEqual(r.axes.cols, ['Quarter', 'UNOBSERVABLE:field-name-unread']);
  assert.equal(r.axes.values, null);
  assert.match(r.axes.pages, /^UNOBSERVABLE:/, 'GetPageFields AND the orientation read both failed -> named');
});

test('G2 driver: GetPageFields throws on a pivot with NO page field (G3 probe) -> the pages are the Filters-oriented fields: none -> [], two -> by position', () => {
  const none = run(pivot({ pagesThrow: true, pivotFields: [field('Region', 1, 'Rows'), field('Product', 2, 'Hidden'), field('Quarter', 3, 'Columns')] }));
  assert.deepEqual(none.axes.pages, []);
  const two = run(pivot({ pagesThrow: true, pivotFields: [field('Region', 1, 'Rows'), field('Quarter', 3, 'Filters', 2), field('Product', 2, 'Filters', 1)] }));
  assert.deepEqual(two.axes.pages, ['Product', 'Quarter']);
});

const SHEETS = ['PS', 'PV'];
const rawRich = () => ({ name: 'PivotTable1', source: 'UNOBSERVABLE:THREW', table: 'A3:G8', layout: { rows: ['Region'], cols: 'UNOBSERVABLE:x', data: ['Sum of Amount', 'Sum of Amount2'] },
  functions: [{ name: 'Sum of Amount', function: 'Sum', field: 'Amount' }, { name: 'Sum of Amount2', function: 'Count', field: 'Amount' }], axes: { rows: ['Region'], cols: ['Quarter'], pages: ['Product'], values: { axis: 'columns', position: 2 } }, style: 'PivotStyleMedium9', proof: 'UNOBSERVABLE:pivot-fields-unread',
  fieldLayout: [{ name: 'Region', form: 'Tabular' }, { name: 'Product', form: 'Tabular' }, { name: 'Quarter', form: 'Tabular' }], dataRef: 'PS!A1:D7' });

test('G2 inventory: the pivots AXIS keeps naming the sheet (the write guards refuse as before); the READ rows list the pivot with its fields', () => {
  const out = withProvenPivots({ PV: { charts: [], names: [], pivots: [rawRich()] } }, 'any-runtime', SHEETS).PV;
  assert.equal(typeof out.pivots, 'string', 'the axis stays a gap -- write consumers unchanged');
  assert.deepEqual(out.pivotDetails, [{ name: 'PivotTable1', source: 'UNOBSERVABLE:pivot-runtime-not-verified', table: 'A3:G8',
    layout: { rows: ['Region'], cols: ['Quarter'], data: ['Sum of Amount', 'Sum of Amount2'] }, pages: ['Product'], values: { axis: 'columns', position: 2 },
    dataFields: [{ name: 'Sum of Amount', function: 'Sum', field: 'Amount' }, { name: 'Sum of Amount2', function: 'Count', field: 'Amount' }], rowLayout: 'Tabular', style: 'PivotStyleMedium9', dataRef: 'PS!A1:D7' }]);
});

test('G2 inventory: a raw pivot WITHOUT axes (an older read / the fake) builds no rows -- the existing shapes stay', () => {
  const { axes, style, ...old } = rawRich();
  assert.ok(axes && style);
  const out = withProvenPivots({ PV: { charts: [], names: [], pivots: [old] } }, 'any-runtime', SHEETS).PV;
  assert.equal(out.pivotDetails, undefined);
});

test('G2 projection: read_pivots lists the rows; a whole-sheet objects gap stays the sheet\'s reason', () => {
  const objects = withProvenPivots({ PV: { charts: [], names: [], pivots: [rawRich()] } }, 'any-runtime', SHEETS);
  const inv = { sheets: SHEETS.map((name, index) => ({ name, index })), objects: { ...objects, PS: 'UNOBSERVABLE:threw' } };
  const project = rt.FORMATS.xlsx.project;
  const r = project(inv, { operation: 'read_pivots', args: {} });
  assert.deepEqual(r.sheets.map(s => [s.sheet, Array.isArray(s.pivots) ? s.pivots.map(p => p.name) : null, s.reason || null]), [['PV', ['PivotTable1'], null], ['PS', null, 'UNOBSERVABLE:threw']]);
  assert.equal(r.sheets[0].pivots[0].style, 'PivotStyleMedium9');
  assert.equal(r.sheets[0].pivots[0].rowLayout, 'Tabular');
  assert.equal('dataRef' in r.sheets[0].pivots[0], false, 'the internal Bridge dataRef is not projected');
});
