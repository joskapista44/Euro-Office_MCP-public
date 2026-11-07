'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * GH #317 step 4 LIVE FAIL: the header/footer `runs` reader's first attempt shipped a per-run
 * `inline` TYPE ARRAY (the SAME walk the paragraph-level field already uses, applied to each run
 * too) -- one level DEEPER than the paragraph's own already-landed `inline` field. LIVE result: as
 * soon as ANY header/footer run carried that non-empty array, the callCommand transport dropped the
 * WHOLE surrounding result (`sections` came back unreadable) -- the SAME transport-escape class
 * already measured for nestedTablesJson and format.tabsJson (docx-tabs-json-transport-vm.test.cjs),
 * just one level deeper. The fix ships a PRIMITIVE boolean (`hasDrawing`) instead -- no Node-side
 * inflate needed, unlike tabsJson/nestedTablesJson.
 *
 * This runs the REAL page-side script (propertyQuery(), never the fake driver) in a vm against a
 * minimal fake Api, with a header paragraph whose FIRST run carries a drawing and SECOND run is
 * plain text -- the exact shape `position: 'start'` produces. If `hasDrawing` ever regresses back
 * into an array/object, this test catches it OFFLINE, before a LIVE round has to.
 */

const QUERY = propertyQuery({ sectionVariantState: false });

function runRaw(doc) {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${QUERY} })()`, sandbox)));
}

/** Every value on every run row must be a primitive (string/boolean/number), never an object or
 * array -- the EXACT class of value that wiped the whole callCommand result LIVE. */
function assertRunsPrimitivesOnly(runs) {
  for (const r of runs) {
    for (const [k, v] of Object.entries(r)) {
      assert.notEqual(typeof v, 'object', `run.${k} must be primitive on the page-side result, got ${typeof v}: ${JSON.stringify(v)}`);
    }
  }
}

const plainRun = text => ({
  GetClassType: () => 'run',
  GetText: () => text,
  GetRange: () => ({ GetTextPr: () => null }),
  ToJSON: () => ({ type: 'run', content: [] }),
});
const drawingRun = () => ({
  GetClassType: () => 'run',
  GetText: () => '',
  GetRange: () => ({ GetTextPr: () => null }),
  ToJSON: () => ({ type: 'run', content: [{ type: 'paraDrawing', drawingType: 'inline' }] }),
});

function paragraphWithRuns(text, runs) {
  return {
    GetClassType: () => 'paragraph',
    GetElementsCount: () => runs.length,
    GetElement: i => runs[i],
    GetText: () => text,
    GetRange: () => ({ GetTextPr: () => null }),
    GetStyle: () => null,
    GetParaPr: () => ({ GetJc: () => undefined, ToJSON: () => ({ tabs: [] }) }),
    // the paragraph-level `inline` field is a SEPARATE, pre-existing, already-landed mechanism --
    // kept plain here so this test stays scoped to the per-RUN field under test
    ToJSON: () => ({ type: 'paragraph', content: [] }),
  };
}

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

test('TRANSPORT: a header run carrying a drawing ships hasDrawing as a BOOLEAN -- every run row stays primitive-only, the transport-escape class this card already measured never triggers', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('ZQW317 start', [drawingRun(), plainRun('ZQW317 start')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.runs.length, 2);
  assertRunsPrimitivesOnly(para.runs);
  assert.equal(para.runs[0].hasDrawing, true);
  assert.equal(para.runs[1].hasDrawing, false);
  // the surrounding result is NOT dropped -- the exact symptom of the original bug
  assert.equal(para.text, 'ZQW317 start');
  assert.ok(Array.isArray(raw.document.sections), 'sections must survive the transport, not come back null/unreadable');
});

test('NEG: a header run with no drawing at all -- hasDrawing is false on every run, still primitive-only', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('ZQW317 plain', [plainRun('ZQW317 plain')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assertRunsPrimitivesOnly(para.runs);
  assert.equal(para.runs[0].hasDrawing, false);
});

test('the drawing run itself reads as empty text, never undefined/null -- the SAME distinction runHasDrawing relies on instead of text emptiness', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('ZQW317 start', [drawingRun(), plainRun('ZQW317 start')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.runs[0].text, '');
  assert.equal(typeof para.runs[0].hasDrawing, 'boolean');
});
