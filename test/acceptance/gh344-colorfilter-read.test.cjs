'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { projectFilterRead } = require('../../src/adapters/xlsx/families/filter-read.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const rt = require('../../src/server/read-tools.cjs');

/** GH #344: read_filter resolves a COLOUR filter's colour through the Bridge /21 sheet.autoFilter.colorFilter.read/v1. */

const cell = { field: 2, on: true, operator: 'xlFilterCellColor', criteria1: null, criteria2: null };
const font = { field: 3, on: true, operator: 'xlFilterFontColor', criteria1: null, criteria2: null };
const value = { field: 1, on: true, operator: 'xlFilterValues', criteria1: ['a'], criteria2: null };
const inv = colours => ({ sheets: [{ name: 'S' }], filterReads: { S: { range: 'A1:C5', filters: [value, cell, font], ...(colours !== undefined ? { colours } : {}) } } });
const read = colours => projectFilterRead(inv(colours), { sheet: 'S' }).filters;
const COLS = { columns: [
  { field: 2, kind: 'cellColor', colour: { form: 'rgb', hex: 'FFC000' }, slot: 'fgColor' },
  { field: 3, kind: 'fontColor', colour: { form: 'theme', theme: 4, tint: -0.25 }, slot: 'fontColor' },
] };

test('the colours resolve in the capability\'s own forms; the value filter is untouched', () => {
  const [v, c, f] = read(COLS);
  assert.deepEqual(v, { field: 1, value: 'a' });
  assert.deepEqual(c.color, { form: 'rgb', hex: 'FFC000' });
  assert.equal(c.kind, 'cellColor');
  assert.deepEqual(f.color, { form: 'theme', theme: 4, tint: -0.25 });
  assert.equal(f.kind, 'fontColor');
  for (const x of [c, f]) {
    assert.equal(x.unsupported, true, 'a colour filter is still not a filter_range shape');
    // GH #344 WRITE shipped: the reason no longer claims there is no write path
    assert.match(x.reason, /^COLOUR_FILTER: written by filter_range \{ colour/);
    assert.equal('colorReason' in x, false);
  }
  const ix = read({ columns: [{ field: 2, kind: 'cellColor', colour: { form: 'indexed', indexed: 10 } }, COLS.columns[1]] })[1];
  assert.deepEqual(ix.color, { form: 'indexed', indexed: 10 });
});

test('NEG: every gap is named, the colour stays null, never guessed', () => {
  // a /19 runtime (no capability): the old COLOR_NOT_PUBLIC entry + the runtime's reason
  const g = read('UNOBSERVABLE:CAPABILITY_MISSING');
  for (const x of [g[1], g[2]]) { assert.equal(x.color, null); assert.match(x.reason, /^COLOR_NOT_PUBLIC/); assert.equal(x.colorReason, 'UNOBSERVABLE:CAPABILITY_MISSING'); }
  // not read at all (a filter_range L3 read): byte-identical to the pre-#344 entry
  const n = read(undefined);
  assert.equal(n[1].color, null); assert.equal('colorReason' in n[1], false);
  // a column the capability does not list
  const m = read({ columns: [COLS.columns[1]] });
  assert.equal(m[1].color, null); assert.equal(m[1].colorReason, 'UNOBSERVABLE:colour-filter-column-not-listed');
  assert.deepEqual(m[2].color, COLS.columns[1].colour);
  // the capability's kind disagrees with the public operator
  const k = read({ columns: [{ ...COLS.columns[0], kind: 'fontColor' }, COLS.columns[1]] });
  assert.equal(k[1].color, null); assert.equal(k[1].colorReason, 'UNOBSERVABLE:colour-filter-kind-mismatch:fontColor');
  // a dxf without a colour
  const z = read({ columns: [{ ...COLS.columns[0], colour: null }, COLS.columns[1]] });
  assert.equal(z[1].color, null); assert.match(z[1].reason, /^COLOR_NONE/);
});

test('readQuery: the colour read is generated only when asked, and runs only when a colour operator is present', () => {
  assert.doesNotMatch(readQuery([], { filterSheets: ['S'] }), /colorFilter/);
  const q = readQuery([], { filterSheets: ['S'], filterColor: true });
  assert.match(q, /sheet\.autoFilter\.colorFilter\.read\/v1/);
  assert.match(q, /if \(filters\.some\(f => f\.operator === 'xlFilterCellColor' \|\| f\.operator === 'xlFilterFontColor'\)\) return \{ range: addrOf\(afRange\), filters, colours: bridgeAxis/);
});

test('read_filter asks for the colours; filter_range\'s own L3 reads do not', () => {
  assert.deepEqual(rt.xlsxRequest('read_filter', { sheet: 'S' }), { windows: [], filterSheets: ['S'], filterColor: true });
  const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
  const scope = liveScope({ id: 't', steps: [{ intent: 'filter_range', sheet: 'S', range: 'A1:C5', field: 1, value: 'a' }] });
  assert.equal('filterColor' in scope, false);
});
