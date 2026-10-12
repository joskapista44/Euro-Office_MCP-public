'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { PROPERTY_QUERY } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const table = require('../../src/adapters/docx/families/table.cjs');

/**
 * W4c point 20-21 (column probe 1, 2026-09-24): an editor-made cell paragraph starts with an EMPTY
 * run in the default font, so the paragraph range's font family reads undefined; the family comes
 * from the NON-EMPTY runs, and runs that disagree are 'MIXED', never one invented family. The same
 * rule decides which family replace_table_cell (W4a) writes back on its one new run.
 */

const textOpts = { Numbering: false };
const run = (text, font) => ({ GetClassType: () => 'run', GetText: () => text, GetRange: () => ({ GetTextPr: () => ({ GetFontFamily: () => font }) }) });
const para = (runs, rangeFont) => ({ runs, GetClassType: () => 'paragraph', GetElementsCount: () => runs.length, GetElement: j => runs[j], GetRange: () => ({ GetTextPr: () => ({ GetFontFamily: () => rangeFont }) }) });

function extract(src, name) {
  const at = src.indexOf(`function ${name}(`);
  assert.ok(at >= 0, `${name} present`);
  let depth = 0;
  for (let i = src.indexOf('{', at); i < src.length; i += 1) {
    if (src[i] === '{') depth += 1;
    if (src[i] === '}') { depth -= 1; if (depth === 0) return src.slice(at, i + 1); }
  }
  throw new Error('unbalanced');
}
// eslint-disable-next-line no-new-func
const fontOf = new Function('textOpts', `${extract(PROPERTY_QUERY, 'cellFirstParagraphFont')}; return cellFirstParagraphFont;`)(textOpts);

const CASES = [
  ['a leading empty default-font run + Courier New', para([run('', 'Arial'), run('FC courier', 'Courier New')], undefined), 'Courier New'],
  ['Courier New + Arial runs', para([run('', 'Arial'), run('FM courier ', 'Courier New'), run('arial', 'Arial')], undefined), 'MIXED'],
  ['a plain cell (the default font on its run)', para([run('', 'Arial'), run('FP plain', 'Arial')], undefined), 'Arial'],
  ['no non-empty run -> the paragraph range', para([run('', 'Arial')], 'Arial'), 'Arial'],
  ['an unreadable family next to a readable one', para([run('a', undefined), run('b', 'Arial')], undefined), 'MIXED'],
];

test('the driver reads a cell\'s first-paragraph font family from its non-empty runs; disagreeing runs -> MIXED', () => {
  for (const [label, p, want] of CASES) assert.equal(fontOf(p), want, label);
});

test('NEGATIVE CONTROL: the paragraph-range read (the pre-W4c path) misses the uniform Courier New cell', () => {
  const p = CASES[0][1];
  const f = p.GetRange().GetTextPr().GetFontFamily();
  assert.notEqual(f, 'Courier New', 'the old path must be blind here, else this test proves nothing');
});

/** Runs replace_table_cell's command body against a one-cell document; returns the new run's calls. */
function replaceOnce(cellPara) {
  const calls = [];
  const content = { GetElementsCount: () => 1, GetElement: () => cellPara, GetText: () => cellPara.runs.map(r => r.GetText()).join('') + '\t', RemoveElement: () => {} };
  cellPara.RemoveAllElements = () => { calls.push('RemoveAllElements'); };
  cellPara.AddElement = () => { calls.push('AddElement'); };
  const cell = { GetContent: () => content };
  const tbl = { GetRowsCount: () => 1, GetRow: () => ({ GetCellsCount: () => 1, GetCell: () => cell }) };
  const Api = {
    GetDocument: () => ({ GetAllTables: () => [tbl] }),
    CreateRun: () => ({ AddText: t => calls.push(['AddText', t]), SetTextPr: () => calls.push('SetTextPr'), SetFontFamily: f => calls.push(['SetFontFamily', f]) }),
  };
  const text = cellPara.runs.map(r => r.GetText()).join('').replace(/\s+/g, ' ').trim();
  const body = table.commandBody({ family: 'table', action: 'cell', targetText: text, tableAnchor: null, text: 'NEW' });
  // eslint-disable-next-line no-new-func
  new Function('Api', body)(Api);
  return calls;
}

test('replace_table_cell writes the uniform family of the old non-empty runs back on its new run; MIXED -> no family is set', () => {
  const uniform = replaceOnce(para([run('', 'Arial'), run('FC courier', 'Courier New')], undefined));
  assert.ok(uniform.some(c => Array.isArray(c) && c[0] === 'AddText' && c[1] === 'NEW'), 'positive control: the body ran to the write');
  assert.deepEqual(uniform.filter(c => Array.isArray(c) && c[0] === 'SetFontFamily'), [['SetFontFamily', 'Courier New']]);
  const mixed = replaceOnce(para([run('', 'Arial'), run('FM courier ', 'Courier New'), run('arial', 'Arial')], undefined));
  assert.ok(mixed.includes('SetTextPr'));
  assert.equal(mixed.filter(c => Array.isArray(c) && c[0] === 'SetFontFamily').length, 0);
});
