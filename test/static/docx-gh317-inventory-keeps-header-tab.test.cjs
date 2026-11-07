'use strict';
// GH #317: a header/footer paragraph written with a real tab must read back WITH the tab through the
// inventory projection (the product's read_headers_footers and every header-footer terminal/guard
// compare against it). inventory.cjs used the body `norm`, which turns the tab into a space and trims
// it away -- a tabbed line never matched its own retry, and paragraph mode refused it
// (header-footer-paragraph-text-not-in-one-run, LIVE).
const test = require('node:test');
const assert = require('node:assert/strict');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');

const part = text => ({ paragraphs: 1, simple: true, paras: [{ text, format: {}, runs: [{ text }], inline: ['tab'] }], tables: [], withTables: false });

function headerText(text) {
  const inv = normalizeInventory({ html: '<p>x</p>', paragraphs: [], document: { sections: [{ header: { default: part(text) }, footer: {} }] } });
  return inv.document.sections[0].header.default.paras[0].text;
}

test('a leading and an inner tab survive the header/footer paragraph text projection', () => {
  assert.equal(headerText('\tSor1'), '\tSor1');
  assert.equal(headerText('Bal\tJobb'), 'Bal\tJobb');
});

test('other whitespace is still normalized exactly as before (no other header text changes)', () => {
  assert.equal(headerText('  Sor  1 \r\n'), 'Sor 1');
  assert.equal(headerText('Plain header'), 'Plain header');
});
