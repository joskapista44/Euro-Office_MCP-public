'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * GH #317 "real tab character" CANDIDATE, read side (HYPOTHESIS per the handoff, not yet LIVE-
 * confirmed): a maintainer's probe (internal tracker comment #ref) measured that after a header/footer
 * paragraph's text is written with a literal `\t` (lowering to a real `<w:tab/>` run-content
 * element), the paragraph's PER-RUN `GetText()` carries the tab but the paragraph's OWN `GetText()`
 * drops it. The fix (docx-driver.cjs `partOf`): when every child of the paragraph is a run and the
 * run read did not throw, build `paras[i].text` by JOINING the run texts instead of trusting the
 * paragraph object's own `GetText()`.
 *
 * This runs the REAL page-side script (propertyQuery(), never a fake driver) in a vm against a
 * minimal fake Api, mirroring docx-gh317-run-drawing-flag-transport-vm.test.cjs's own technique --
 * a fake paragraph whose OWN GetText() deliberately returns the tab-LESS text (reproducing the
 * measured discrepancy), while its run's GetText() carries the real tab. If the driver ever
 * regresses back to trusting the paragraph's own GetText(), this test fails RED.
 */

const QUERY = propertyQuery({ sectionVariantState: false });

function runRaw(doc) {
  const sandbox = { Api: { GetDocument: () => doc }, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${QUERY} })()`, sandbox)));
}

const plainRun = text => ({
  GetClassType: () => 'run',
  GetText: () => text,
  GetRange: () => ({ GetTextPr: () => null }),
  ToJSON: () => ({ type: 'run', content: [] }),
});

// A non-run paragraph child (e.g. a bookmark marker) -- used to prove the fallback path still
// trusts the paragraph's own GetText() when the paragraph is NOT run-only.
const nonRunChild = () => ({ GetClassType: () => 'bookmarkStart' });

function paragraphWithRuns(ownText, runs) {
  return {
    GetClassType: () => 'paragraph',
    GetElementsCount: () => runs.length,
    GetElement: i => runs[i],
    // the paragraph's OWN GetText() deliberately returns ownText -- which, in the run-only
    // (tab-dropping) case below, is a DIFFERENT, tab-less string than the join of its runs
    GetText: () => ownText,
    GetRange: () => ({ GetTextPr: () => null }),
    GetStyle: () => null,
    GetParaPr: () => ({ GetJc: () => undefined, ToJSON: () => ({ tabs: [] }) }),
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

test('GH317 tab-read: a run-only paragraph reports text by JOINING its run texts, not the paragraph\'s own lossy GetText()', () => {
  // the paragraph's OWN GetText() drops the tab (reproducing a maintainer's LIVE measurement) -- the run's
  // own GetText() carries it. The fixed driver must prefer the run join.
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('Sor1', [plainRun('\tSor1')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.text, '\tSor1', 'the paragraph text must carry the tab the run already carries');
  assert.equal(para.runs.length, 1);
  assert.equal(para.runs[0].text, '\tSor1');
});

test('GH317 tab-read: a two-run paragraph (tab run + trailing text run) joins in element order', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('Sor1', [plainRun('\t'), plainRun('Sor1')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.text, '\tSor1');
});

test('GH317 tab-read: a field-like empty-text run contributes nothing to the join (the pageNumber case stays unaffected)', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('Sor1', [plainRun('Sor1'), plainRun('')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.text, 'Sor1');
});

test('GH317 tab-read NEG: a paragraph with a non-run child falls back to the paragraph\'s own GetText() (the join is not trusted when not every child is a run)', () => {
  // paragraphWithRuns's GetElementsCount/GetElement are built from the `runs` array directly, so to
  // put a non-run child alongside a run, hand it the non-run fake as one of the "runs" -- the
  // driver's own per-child class check (not this test's naming) is what decides allRuns.
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('fallback-text', [nonRunChild(), plainRun('\tSor1')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.text, 'fallback-text', 'a non-run child must disable the join and keep the GetText() fallback');
});

test('GH317 tab-read NEG: a plain tab-less paragraph is unaffected (text equals the single run\'s text, as before)', () => {
  const raw = runRaw(docWithHeaderParagraph(paragraphWithRuns('plain', [plainRun('plain')])));
  const para = raw.document.sections[0].header.default.paras[0];
  assert.equal(para.text, 'plain');
});
