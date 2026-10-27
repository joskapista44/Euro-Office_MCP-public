'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery, inflateNestedTables } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const { FORMATS } = require('../../src/server/read-tools.cjs');

/**
 * read_tables returns the NESTED tables' cell content next to the top-level tables, bound to the holding cell as `cell.nestedTables`
 * (docs/GH-NESTED-TABLES-IN-READ-TABLES-DECISION.md). The REAL propertyQuery() string runs in a vm against a scripted Api; the product
 * projection runs on the normalized inventory. A cell without a nested table carries NO `nestedTables` key (existing outputs unchanged).
 */

// the PAGE-side result exactly as the callCommand callback would carry it (no inflation)
const runRaw = (doc, src = propertyQuery()) => {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${src} })()`, sandbox)));
};
// what readLive hands on: the page result after the driver's own inflation (richOutcome -> inflateNestedTables)
const run = (doc, src) => { const r = runRaw(doc, src); inflateNestedTables(r.document); return r; };
const paragraph = text => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetElementsCount: () => 0, GetElement: () => null, GetTextPr: () => null, GetRange: () => ({ GetTextPr: () => null }), GetStyle: () => null, GetParaPr: () => null });
const cellOf = (text, ...nested) => {
  const els = [paragraph(text), ...nested];
  return { GetContent: () => ({ GetElementsCount: () => els.length, GetElement: i => els[i], GetText: () => text }) };
};
const table = (tag, nestedAt = {}) => {
  const cells = [[0, 0], [0, 1], [1, 0], [1, 1]].map(([r, c]) => cellOf(`${tag} r${r}c${c}`, ...(nestedAt[`${r}${c}`] ? [].concat(nestedAt[`${r}${c}`]) : [])));
  return { GetClassType: () => 'table', GetRowsCount: () => 2, GetRow: r => ({ GetCellsCount: () => 2, GetCell: c => cells[r * 2 + c] }), GetTableTitle: () => null };
};
const brokenNested = () => {
  const cell = { GetContent: () => ({ GetElementsCount: () => 1, GetElement: () => paragraph('x'), GetText: () => { throw new Error('nested boom'); } }) };
  return { GetClassType: () => 'table', GetRowsCount: () => 1, GetRow: () => ({ GetCellsCount: () => 1, GetCell: () => cell }) };
};
const docOf = (body, all = body) => ({
  GetElementsCount: () => body.length, GetElement: i => body[i], GetAllTables: () => all, Search: () => [],
  GetAllBookmarksNames: () => [], GetBookmarkRange: () => null, GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
});
const texts = rows => rows.map(r => r.map(c => c.text));
const project = r => FORMATS.docx.project(normalizeInventory({ html: '<p>x</p>', paragraphs: r.paragraphs, document: r.document }), { operation: 'read_tables' });

test('vm: a nested table is returned on the HOLDING cell, rows / cells in order; the other cells and the top-level list are untouched', () => {
  const N = table('NIN');
  const T1 = table('T1', { '00': N }); const T2 = table('T2');
  const r = run(docOf([T1, T2], [T1, N, T2]));
  assert.equal(r.document.tables.length, 2);
  const holder = r.document.tables[0].rows[0][0];
  assert.equal(holder.nestedTables.length, 1);
  assert.deepEqual(texts(holder.nestedTables[0].rows), [['NIN r0c0', 'NIN r0c1'], ['NIN r1c0', 'NIN r1c1']]);
  assert.equal(r.document.tables[0].rows[0][1].nestedTables, undefined);
  assert.equal(r.document.tables[1].rows[1][1].nestedTables, undefined);
});

test('vm: NEG -- a document with no nested table has NO nestedTables key anywhere (bytes of the existing cell shape unchanged)', () => {
  const r = run(docOf([table('A'), table('B')]));
  assert.equal(JSON.stringify(r.document.tables).includes('nestedTables'), false);
});

test('vm: two nested tables in one cell are both returned, in document order; a nested table inside a nested table is returned (depth 2)', () => {
  const deep = table('DEEP');
  const N1 = table('N1', { '11': deep }); const N2 = table('N2');
  const T1 = table('T1', { '01': [N1, N2] });
  const r = run(docOf([T1], [T1, N1, deep, N2]));
  const holder = r.document.tables[0].rows[0][1];
  assert.deepEqual(holder.nestedTables.map(t => t.rows[0][0].text), ['N1 r0c0', 'N2 r0c0']);
  assert.equal(holder.nestedTables[0].rows[1][1].nestedTables[0].rows[0][0].text, 'DEEP r0c0');
});

test('vm: past the depth limit (3 nested levels) the holding cell names the limit instead of reading on', () => {
  const l4 = table('L4'); const l3 = table('L3', { '00': l4 }); const l2 = table('L2', { '00': l3 }); const l1 = table('L1', { '00': l2 }); const T = table('T', { '00': l1 });
  const r = run(docOf([T]));
  const d3 = r.document.tables[0].rows[0][0].nestedTables[0].rows[0][0].nestedTables[0].rows[0][0].nestedTables[0];
  assert.equal(d3.rows[0][0].text, 'L3 r0c0');
  assert.equal(d3.rows[0][0].nestedTables, 'UNOBSERVABLE:nested-depth-limit');
});

test('vm: a nested table whose read throws names itself on the HOLDING cell; the cell text, its siblings and the other tables stay readable', () => {
  const T1 = table('T1', { '00': brokenNested() }); const T2 = table('T2');
  const r = run(docOf([T1, T2]));
  assert.equal(r.error, undefined);
  assert.equal(r.document.tables.length, 2);
  assert.equal(r.document.tables[0].rows[0][0].text, 'T1 r0c0');
  assert.equal(r.document.tables[0].rows[0][0].nestedTables, 'UNOBSERVABLE:nested-table: nested boom');
  assert.equal(r.document.tables[0].rows[0][1].nestedTables, undefined);
});

test('product read_tables: nestedTables is projected (array or the named string), recursively; a cell without one has no key', () => {
  const N = table('NIN'); const T1 = table('T1', { '00': N, '11': brokenNested() });
  const out = project(run(docOf([T1], [T1, N])));
  const t = out.tables[0];
  assert.deepEqual(texts(t.rows[0][0].nestedTables[0].rows), [['NIN r0c0', 'NIN r0c1'], ['NIN r1c0', 'NIN r1c1']]);
  assert.equal(t.rows[1][1].nestedTables, 'UNOBSERVABLE:nested-table: nested boom');
  assert.equal('nestedTables' in t.rows[0][1], false);
});

test('vm: a nested table whose STRUCTURE cannot be walked fails the whole list by name (the GetAllTables index space of the later tables would be unknown) -- tables absent, tablesUnobservable set', () => {
  const unwalkable = { GetClassType: () => 'table', GetRowsCount: () => { throw new Error('structure boom'); } };
  const r = run(docOf([table('T1', { '00': unwalkable }), table('T2')]));
  assert.equal(r.document.tables, undefined);
  assert.equal(r.document.tablesUnobservable, 'UNOBSERVABLE:table-walk: element 0: structure boom');
});

test('vm: TRANSPORT -- on the page a top-level holding cell carries its nested tables as a JSON STRING (nestedTablesJson), never as nested objects (gnpq2 LIVE: the object form did not come back through the callCommand callback); the driver inflates it', () => {
  const N = table('NIN');
  const T1 = table('T1', { '00': N }); const T2 = table('T2');
  const raw = runRaw(docOf([T1, T2], [T1, N, T2]));
  const holder = raw.document.tables[0].rows[0][0];
  assert.equal(typeof holder.nestedTablesJson, 'string');
  assert.equal(holder.nestedTables, undefined, 'no nested object tree on the page-side result');
  const inflated = inflateNestedTables(JSON.parse(JSON.stringify(raw.document)));
  assert.deepEqual(texts(inflated.tables[0].rows[0][0].nestedTables[0].rows), [['NIN r0c0', 'NIN r0c1'], ['NIN r1c0', 'NIN r1c1']]);
  assert.equal(inflated.tables[0].rows[0][0].nestedTablesJson, undefined);
});

test('vm: inflateNestedTables NAMES an unparseable string instead of dropping it', () => {
  const doc = { tables: [{ rows: [[{ text: 'x', nestedTablesJson: '{broken' }]] }] };
  const out = inflateNestedTables(doc);
  assert.match(out.tables[0].rows[0][0].nestedTables, /^UNOBSERVABLE:nested-table-json: /);
});
