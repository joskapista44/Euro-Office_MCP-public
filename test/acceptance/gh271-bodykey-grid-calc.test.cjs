'use strict';

/**
 * GH #271 (LIVE g271f + a no-write probe on the LARGE copy): two reads of the SAME, unwritten document differed in one
 * field only -- the calculated table grid (`gridCalc`, a bridge layout read) was null for 30/30 tables in one read and
 * set for 30/30 in the next. The body comparison of set_page_setup, edit_header/edit_footer, define_style/modify_style,
 * the comment and the image families included it, so all five LARGE rounds ended INTEGRAL_PRESERVATION_NOT_PROVEN
 * after a write that persisted. The body key compares the stored table content (rows, format), not the layout read.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const pageSetup = require('../../src/adapters/docx/families/page-setup.cjs');
const headerFooter = require('../../src/adapters/docx/families/header-footer.cjs');
const style = require('../../src/adapters/docx/families/style-definition.cjs');
const image = require('../../src/adapters/docx/families/image.cjs');

const cell = text => ({ text, paragraphs: [{ text }] });
const primary = (gridCalc, text = 'Qg cell') => ({
  blocks: [], images: [],
  document: { footnotes: [], endnotes: [], bookmarks: [], inMainBody: true, tables: [{ rows: [[cell(text), cell('Qg b')]], gridCalc, format: { width: null } }] },
});
const KEYS = { 'page-setup': pageSetup.bodyKey, 'header-footer': headerFooter.bodyKey, 'style-definition': style.bodyKey, image: image.bodyKey };

for (const [name, bodyKey] of Object.entries(KEYS)) {
  test(`${name}: an unread vs read calculated grid is not a body change`, () => {
    assert.equal(bodyKey(primary(null)), bodyKey(primary(Object.freeze([35.3, 35.3]))));
  });
  test(`RED CONTROL ${name}: a changed cell text is still a body change`, () => {
    assert.notEqual(bodyKey(primary(null)), bodyKey(primary(null, 'Qg CHANGED')));
  });
}
