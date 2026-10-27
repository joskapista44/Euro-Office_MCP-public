'use strict';
// GH nested tables (gnl3a LIVE): parseBlocks matched a table lazily, so a NESTED table's own </table> ended the holding table -- its remaining rows
// then read as paragraphs OUTSIDE the tables, and replace_table_cell on the holding table failed its own L3 (text units outside the tables "changed").
const test = require('node:test');
const assert = require('node:assert/strict');
const { parseBlocks } = require('../../src/adapters/docx/inventory.cjs');

const nested = '<table><tr><td><p>NIN r0c0</p></td><td><p>NIN r0c1</p></td></tr></table>';
const html = `<p>before</p><table><tr><td><p>OUT r0c0</p>${nested}</td><td><p>OUT r0c1</p></td></tr><tr><td><p>OUT r1c0</p></td><td><p>OUT r1c1</p></td></tr></table><p>after</p>`;

test('a table holding a nested table is ONE block with its own rows; no paragraph of it leaks outside the tables', () => {
  const blocks = parseBlocks(html);
  assert.deepEqual(blocks.map(b => b.kind), ['paragraph', 'table', 'paragraph']);
  assert.deepEqual(blocks.map(b => b.kind === 'paragraph' ? b.text : null).filter(Boolean), ['before', 'after']);
  const t = blocks[1];
  assert.deepEqual(t.rows.map(r => r.length), [2, 2], 'only the holding table\'s OWN rows, not the nested one\'s');
  assert.equal(t.rows[1][1].text, 'OUT r1c1');
  assert.match(t.rows[0][0].text, /OUT r0c0/);
  assert.match(t.rows[0][0].text, /NIN r0c0/, 'the holding cell keeps the nested table\'s text, as the driver\'s cell text does');
});

test('NEG: a document without a nested table parses exactly as before', () => {
  const plain = '<p>a</p><table><tr><td><p>x</p></td><td><p>y</p></td></tr></table><p>b</p>';
  const blocks = parseBlocks(plain);
  assert.deepEqual(blocks.map(b => b.kind), ['paragraph', 'table', 'paragraph']);
  assert.deepEqual(blocks[1].rows, [[{ text: 'x', span: 1, rowspan: 1 }, { text: 'y', span: 1, rowspan: 1 }]]);
});

test('an unbalanced table falls back to the old match (no throw, no lost text)', () => {
  const blocks = parseBlocks('<table><tr><td><p>x</p></td></tr></table><table><tr><td>');
  assert.equal(blocks[0].kind, 'table');
});
