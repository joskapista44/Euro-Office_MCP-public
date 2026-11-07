'use strict';

/**
 * GH #310: a DOCX font size of 8 pt and below read back doubled (format_text fontSize 8 -> 16, the verify failed).
 * The editor's GetFontSize() returns HALF-POINTS; the driver halved it only when the value was above 16, so 8 pt (16
 * half-points) and every smaller size stayed doubled. This test runs the REAL PROPERTY_QUERY string (imported, not
 * retyped) via vm against a fake Api whose GetFontSize() answers in half-points, on both read paths: the
 * whole-paragraph format and the per-run format.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/** points -> the half-point value the editor's GetFontSize() returns */
const SIZES_PT = [12, 9, 8, 6, 4, 0.5];

function textPr(halfPoints) {
  return { GetBold: () => false, GetItalic: () => false, GetUnderline: () => false, GetStrikeout: () => false,
    GetFontFamily: () => 'Arial', GetFontSize: () => halfPoints, GetLanguage: () => 'hu-HU' };
}

function makeFakeApi() {
  const elements = SIZES_PT.map(pt => {
    const text = `SIZE_${pt}`;
    const pr = textPr(pt * 2);
    const run = { GetClassType: () => 'run', GetText: () => text, GetRange: () => ({ GetTextPr: () => pr }) };
    return { GetClassType: () => 'paragraph', GetText: () => text, GetElementsCount: () => 1, GetElement: () => run,
      GetNumbering: () => null, GetParaPr: () => null, GetStyle: () => null, Paragraph: null, __pr: pr };
  });
  const doc = {
    GetElementsCount: () => elements.length,
    GetElement: i => elements[i],
    Search: text => elements.filter(e => e.GetText() === text).map(e => ({ GetText: () => text, GetTextPr: () => e.__pr })),
    GetAllBookmarksNames: () => [],
    GetBookmarkRange: () => null,
    GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetSections: () => [],
    GetDefaultTextPr: () => null,
  };
  return { GetDocument: () => doc };
}

function runPropertyQuery(api, query = propertyQuery({})) {
  const sandbox = { Api: api, result: undefined };
  vm.createContext(sandbox);
  return JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${query} })()`, sandbox)));
}

/** The whole-paragraph size and the per-run sizes of every paragraph, in points. */
function sizesOf(r) {
  return SIZES_PT.map(pt => {
    const p = r.paragraphs.find(x => x.text === `SIZE_${pt}`);
    assert.ok(p, `paragraph SIZE_${pt} present`);
    return { pt, paragraph: p.format && p.format.fontSize, runs: (p.runs || []).map(x => x.format && x.format.fontSize) };
  });
}

test('GH #310: every size, 8 pt and below included, reads back in points on the paragraph and the run path', () => {
  const r = runPropertyQuery(makeFakeApi());
  assert.equal(r.error, undefined, `PROPERTY_QUERY threw: ${r.error}`);
  for (const s of sizesOf(r)) {
    assert.equal(s.paragraph, s.pt, `paragraph size ${s.pt} pt`);
    assert.deepEqual(s.runs, [s.pt], `run size ${s.pt} pt`);
  }
});

test('RED CONTROL: the old "> 16 ? / 2" rule, on the same query text, doubles 8 pt and below and only those', () => {
  const old = propertyQuery({}).split('v / 2;').join('v > 16 ? v / 2 : v;').split('n / 2;').join('n > 16 ? n / 2 : n;');
  assert.notEqual(old, propertyQuery({}), 'the control must actually rewrite the query');
  const wrong = sizesOf(runPropertyQuery(makeFakeApi(), old)).filter(s => s.paragraph !== s.pt).map(s => s.pt);
  assert.deepEqual(wrong, [8, 6, 4, 0.5]);
});
