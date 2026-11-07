'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery, richOutcome, inflateNestedTables, inflateTabsJson } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * GH #317 LIVE FAIL (a maintainer, docx-gh317-header-footer-tabs-product round 1): a paragraph's
 * format.tabs OBJECT ARRAY wipes the WHOLE surrounding callCommand result (header/footer's
 * sections came back null on the branch carrying this field, vs a readable result on main for
 * the SAME doc) -- the SAME transport-escape class gnpq2 already measured for a top-level
 * table's nested tables (nestedTablesJson). paragraphFormat() therefore hands tabs over as a
 * JSON string (format.tabsJson); inflateTabsJson (readLive -> richOutcome) turns it back into
 * format.tabs on the Node side.
 *
 * This runs the REAL page-side script (propertyQuery(), never the fake driver) in a vm against
 * a minimal fake Api, for BOTH consumers of the shared paragraphFormat(): a header part AND a
 * body-table cell -- a fix proven on one and not measured on the other would leave the second
 * silently unfixed.
 */

const QUERY = propertyQuery({ sectionVariantState: false });

function runRaw(doc) {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${QUERY} })()`, sandbox)));
}
// what readLive hands on: the page result after the driver's own inflation (richOutcome runs
// inflateNestedTables THEN inflateTabsJson, in that order -- the SAME order this test uses)
function run(doc) {
  const { document } = richOutcome({ ok: true, result: runRaw(doc) });
  return document;
}

function assertNoObjectValues(format) {
  for (const [k, v] of Object.entries(format)) assert.notEqual(typeof v, 'object', `format.${k} must be primitive/string on the page-side result, got ${typeof v}: ${JSON.stringify(v)}`);
}

function paragraphWithTabs(text, tabsRaw) {
  const paraPr = { GetJc: () => undefined, ToJSON: () => ({ tabs: tabsRaw }) };
  return {
    GetClassType: () => 'paragraph',
    GetElementsCount: () => 0,
    GetElement: () => null,
    GetText: () => text,
    GetRange: () => ({ GetTextPr: () => null }),
    GetStyle: () => null,
    GetParaPr: () => paraPr,
  };
}
const TABS_RAW = [{ pos: 5760, val: 'right' }]; // twips; the SAME shape ParaPr.ToJSON().tabs carries live
const TABS_OUT = [{ pos: 288, align: 'right' }]; // /20 twips -> points, the task-facing shape

function docWithHeaderParagraph(p) {
  const headerPart = { GetElementsCount: () => 1, GetElement: () => p };
  const sec = { GetHeader: v => (v === 'default' ? headerPart : null), GetFooter: () => null };
  return {
    GetElementsCount: () => 0, GetElement: () => null, GetAllTables: () => [], Search: () => [],
    GetAllBookmarksNames: () => [], GetBookmarkRange: () => null, GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
    GetSections: () => [sec], GetAllStyles: () => [], GetDefaultTextPr: () => null,
  };
}
function docWithTableCellParagraph(p) {
  const cell = { GetContent: () => ({ GetElementsCount: () => 1, GetElement: () => p, GetText: () => '' }) };
  const table = { GetClassType: () => 'table', GetRowsCount: () => 1, GetRow: () => ({ GetCellsCount: () => 1, GetCell: () => cell }), GetTableTitle: () => null };
  return {
    GetElementsCount: () => 1, GetElement: () => table, GetAllTables: () => [table], Search: () => [],
    GetAllBookmarksNames: () => [], GetBookmarkRange: () => null, GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null,
    GetSections: () => [], GetAllStyles: () => [], GetDefaultTextPr: () => null,
  };
}

test('TRANSPORT (header): the page-side result carries format.tabsJson (a string), never format.tabs (an object array) -- every format field stays a primitive/string', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithTabs('ZQW h', TABS_RAW)));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(typeof para.format.tabsJson, 'string');
  assert.equal('tabs' in para.format, false);
  assertNoObjectValues(para.format);
});

test('INFLATE (header): richOutcome turns format.tabsJson back into points-converted format.tabs; the JSON string key is gone afterwards', () => {
  const document = run(docWithHeaderParagraph(paragraphWithTabs('ZQW h', TABS_RAW)));
  const para = document.sections[0].header.default.paras[0];
  assert.deepEqual(para.format.tabs, TABS_OUT);
  assert.equal(para.format.tabsJson, undefined);
});

test('TRANSPORT + INFLATE (table cell): the SAME paragraphFormat() shared by table-cell reads -- tabsJson on the page, tabs after inflation, on both format and paraFormats', () => {
  const raw = runRaw(docWithTableCellParagraph(paragraphWithTabs('ZQW c', TABS_RAW)));
  const rawCell = raw.document.tables[0].rows[0][0];
  assert.equal(typeof rawCell.format.tabsJson, 'string');
  assert.equal('tabs' in rawCell.format, false);
  assertNoObjectValues(rawCell.format);
  const document = run(docWithTableCellParagraph(paragraphWithTabs('ZQW c', TABS_RAW)));
  const cell = document.tables[0].rows[0][0];
  assert.deepEqual(cell.format.tabs, TABS_OUT);
  assert.equal(cell.format.tabsJson, undefined);
  assert.deepEqual(cell.paraFormats[0].format.tabs, TABS_OUT);
  assert.equal(cell.paraFormats[0].format.tabsJson, undefined);
});

test('NEG: a paragraph with no tabs at all -- tabsJson is still a string ("[]"), never absent; inflates to an empty array, not UNOBSERVABLE', () => {
  const document = run(docWithHeaderParagraph(paragraphWithTabs('ZQW none', [])));
  assert.deepEqual(document.sections[0].header.default.paras[0].format.tabs, []);
});

test('inflateTabsJson NAMES an unparseable tabsJson string instead of dropping it (mirrors inflateNestedTables), on both a table cell and a header paragraph', () => {
  const doc1 = { tables: [{ rows: [[{ format: { tabsJson: '{broken' } }]] }] };
  inflateTabsJson(doc1);
  assert.match(doc1.tables[0].rows[0][0].format.tabs, /^UNOBSERVABLE:tabs-json: /);
  const doc2 = { sections: [{ header: { default: { paras: [{ format: { tabsJson: '{broken' } }] } }, footer: {} }] };
  inflateTabsJson(doc2);
  assert.match(doc2.sections[0].header.default.paras[0].format.tabs, /^UNOBSERVABLE:tabs-json: /);
});

test('inflateTabsJson reaches a NESTED table cell too, after inflateNestedTables has already turned its own JSON string back into objects (order-dependent -- richOutcome runs both, in this order)', () => {
  const nestedCellWithTabs = { text: 'nested', paragraphs: 1, simple: true, format: { tabsJson: JSON.stringify(TABS_OUT) }, paraFormats: [] };
  const holder = { text: 'holder', format: {}, paraFormats: [], nestedTablesJson: JSON.stringify([{ rows: [[nestedCellWithTabs]] }]) };
  const document = { tables: [{ rows: [[holder]] }] };
  inflateNestedTables(document);
  inflateTabsJson(document);
  const nested = document.tables[0].rows[0][0].nestedTables[0].rows[0][0];
  assert.deepEqual(nested.format.tabs, TABS_OUT);
  assert.equal(nested.format.tabsJson, undefined);
});
