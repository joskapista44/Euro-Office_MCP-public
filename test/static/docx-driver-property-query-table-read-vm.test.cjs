'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { PROPERTY_QUERY } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * W04-6 (a maintainer's LIVE catch, #ref): the offline acceptance suite (docx-w04-table-format.test.cjs
 * and every other table-family test) runs against test/fakes/fake-docx-driver.cjs, which NEVER
 * executes PROPERTY_QUERY -- it models the driver's OWN behaviour by hand. A bug INSIDE
 * PROPERTY_QUERY's own table-read branch (the real driver's in-document script) is therefore
 * invisible to 2041/2041 passing offline, and was only caught by a LIVE run: a TDZ ReferenceError
 * (BORDER_EDGES/BORDER_JSON_KEY/MARGIN_EDGES/tblFormatOf declared AFTER the `docInfo.tables =
 * (...).map(...)` call that invokes them synchronously) silently discarded `docInfo.tables`
 * ENTIRELY via the outer try/catch -- breaking every table family's read, not just
 * table-format's own fields.
 *
 * This test runs PROPERTY_QUERY itself, in a vm, against a minimal fake Api (one 1x1 table with
 * a ToJSON()) -- the SAME execution path the real driver's readLive() takes, never the fake
 * driver's hand-written model. It is RED on the bug (document.tables undefined) and GREEN on the
 * fix (document.tables present, including the table-format `format` field).
 */

function fakeApiWithOneTable(tblPrJson) {
  const para = { GetClassType: () => 'paragraph', GetElementsCount: () => 0, GetText: () => 'cell text', GetRange: () => ({ GetTextPr: () => null }), GetStyle: () => null, GetParaPr: () => null };
  const content = { GetElementsCount: () => 1, GetElement: () => para, GetText: () => 'cell text' };
  const table = {
    GetClassType: () => 'table',
    GetRowsCount: () => 1,
    GetRow: () => ({ GetCellsCount: () => 1, GetCell: () => ({ GetContent: () => content }) }),
    Table: {},
    ToJSON: () => JSON.stringify({ tblPr: tblPrJson }),
  };
  const doc = {
    GetElementsCount: () => 1, GetElement: () => table, // the body holds the table: the tables list is found by the element walk
    GetAllBookmarksNames: () => [], GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [table],
    GetDefaultTextPr: () => null,
    GetSections: () => [],
    GetAllStyles: () => [],
  };
  return { GetDocument: () => doc };
}

function runPropertyQuery(Api) {
  const script = new vm.Script(`(function(Api){${PROPERTY_QUERY}})`);
  return script.runInThisContext()(Api);
}

test('PROPERTY_QUERY, run in a vm (never the fake driver): document.tables is read, not silently discarded', () => {
  const result = runPropertyQuery(fakeApiWithOneTable({ tblBorders: {}, tblCellMar: {} }));
  assert.equal(result.error, undefined, JSON.stringify(result));
  assert.ok(Array.isArray(result.document && result.document.tables), `document.tables must be an array, got: ${JSON.stringify(result.document && result.document.tables)}`);
  assert.equal(result.document.tables.length, 1);
});

test("PROPERTY_QUERY's table-format read: a border edge set via ToJSON is carried through to document.tables[0].format.borders", () => {
  const tblPr = { tblBorders: { top: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 255, g: 0, b: 0 } } }, tblCellMar: {} };
  const result = runPropertyQuery(fakeApiWithOneTable(tblPr));
  assert.equal(result.error, undefined, JSON.stringify(result));
  const f = result.document.tables[0].format;
  assert.ok(f, 'format must be read, not absent');
  assert.deepEqual(f.borders.top, { value: 'single', sz: 4, space: 0, color: { auto: false, r: 255, g: 0, b: 0 } });
  assert.equal(f.borders.bottom, null, 'an edge with no entry in tblBorders is a successfully-read ABSENT value (null), not UNOBSERVABLE');
});

test("the start/end border key-mapping (MEASURED, #ref) is applied on read: ToJSON's 'start' key surfaces as the task-facing 'left' field", () => {
  const tblPr = { tblBorders: { start: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 0, g: 255, b: 0 } } }, tblCellMar: {} };
  const result = runPropertyQuery(fakeApiWithOneTable(tblPr));
  const f = result.document.tables[0].format;
  assert.deepEqual(f.borders.left, { value: 'single', sz: 4, space: 0, color: { auto: false, r: 0, g: 255, b: 0 } });
});

test('NEGATIVE CONTROL: a document with NO tables still reads cleanly (document.tables is an empty array, not absent)', () => {
  const doc = {
    GetElementsCount: () => 0, GetElement: () => null,
    GetAllBookmarksNames: () => [], GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetDefaultTextPr: () => null,
    GetSections: () => [],
    GetAllStyles: () => [],
  };
  const result = runPropertyQuery({ GetDocument: () => doc });
  assert.equal(result.error, undefined, JSON.stringify(result));
  assert.deepEqual(result.document.tables, []);
});
