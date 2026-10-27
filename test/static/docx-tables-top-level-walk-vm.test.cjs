'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');

/**
 * The page-side tables list is the document's TOP-LEVEL tables (the HTML projection lists only those), found by the element walk
 * (gntv1 LIVE: the walk finds the top-level tables and marks the cell holding a nested one; GetAllTables() lists the nested table too, so the lengths differed
 * and every table family refused with table-read-shape-mismatch). The REAL propertyQuery() string runs in a vm against a scripted Api.
 * The bridge grid read keeps its GetAllTables() index space: a top-level table after a nested one is asked for under its GetAllTables() position.
 */

const run = (doc, src = propertyQuery()) => {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  const out = JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${src} })()`, sandbox)));
  out.__rdCalls = sandbox.__rdCalls || [];
  return out;
};
const spy = src => src.replace('const rd = (cap, args) =>', 'const rd = (cap, args) => { (globalThis.__rdCalls = globalThis.__rdCalls || []).push([cap.id, JSON.stringify(args)]); return null; }; const rdUnused = (cap, args) =>');

const paragraph = text => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetElementsCount: () => 0, GetElement: () => null, GetTextPr: () => null, GetRange: () => ({ GetTextPr: () => null }), GetStyle: () => null, GetParaPr: () => null });
const cellOf = (text, nested) => {
  const els = [paragraph(text), ...(nested ? [nested] : [])];
  return { GetContent: () => ({ GetElementsCount: () => els.length, GetElement: i => els[i], GetText: () => text }) };
};
const table = (first, nested) => {
  const cells = [cellOf(first, nested)];
  return { GetClassType: () => 'table', GetRowsCount: () => 1, GetRow: () => ({ GetCellsCount: () => cells.length, GetCell: c => cells[c] }), GetTableTitle: () => null };
};
const docOf = (body, all) => ({
  GetElementsCount: () => body.length, GetElement: i => body[i],
  GetAllTables: () => all, Search: () => [],
  GetAllBookmarksNames: () => [], GetBookmarkRange: () => null,
  GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
});
const texts = r => r.document.tables.map(t => t.rows[0][0].text);

test('vm: a nested table is NOT listed -- the list is the top-level tables in document order, the cell holding the nested one is still there', () => {
  const N = table('nested');
  const T1 = table('T1 holds nested', N); const T2 = table('T2'); const T3 = table('T3');
  const r = run(docOf([T1, paragraph('p'), T2, T3], [T1, N, T2, T3]));
  assert.equal(r.error, undefined);
  assert.deepEqual(texts(r), ['T1 holds nested', 'T2', 'T3']);
});

test('vm: NEG -- a document with no nested table is listed exactly as GetAllTables() lists it', () => {
  const T1 = table('A'); const T2 = table('B');
  const r = run(docOf([T1, paragraph('p'), T2], [T1, T2]));
  assert.deepEqual(texts(r), ['A', 'B']);
});

test('vm: the bridge grid read keeps the GetAllTables() index (T2 follows a nested table -> asked as table 2, not 1)', () => {
  const N = table('nested'); const T1 = table('T1', N); const T2 = table('T2');
  const r = run(docOf([T1, T2], [T1, N, T2]), spy(propertyQuery()));
  assert.equal(JSON.stringify(r.__rdCalls.filter(c => /effectiveGrid/.test(c[0])).map(c => JSON.parse(c[1]).table)), '[0,2]');
});

test('vm: a throwing element in the walk -> NAMED tablesUnobservable, tables absent (fail-closed), the rest of the query still read', () => {
  const T1 = table('A'); const T2 = table('B');
  const broken = { GetClassType: () => 'table', GetRowsCount: () => { throw new Error('walk boom'); } };
  const r = run(docOf([T1, broken, T2], [T1, T2]));
  assert.equal(r.error, undefined);
  assert.equal(r.document.tables, undefined);
  assert.equal(r.document.tablesUnobservable, 'UNOBSERVABLE:table-walk: element 1: walk boom');
});

test('inventory: tablesUnobservable passes through as the named reason; tables stays null (consumers fail closed as before)', () => {
  const inv = normalizeInventory({ html: '<p>x</p>', paragraphs: [], document: { tablesUnobservable: 'UNOBSERVABLE:table-walk: element 1: walk boom' } });
  assert.equal(inv.document.tables, null);
  assert.equal(inv.document.tablesUnobservable, 'UNOBSERVABLE:table-walk: element 1: walk boom');
  assert.equal(normalizeInventory({ html: '<p>x</p>', paragraphs: [], document: {} }).document.tablesUnobservable, undefined);
});
