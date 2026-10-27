'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { commandBody } = require('../../src/adapters/docx/families/style-definition.cjs');

/**
 * GH #292 S1 / S2: define_style / modify_style take the further measured fields (docs/evidence/gh292-style-definition-probe-20261021.json):
 * text { italic, underline, strike (true only), fontFamily }, paragraph { spacingBefore, indentRight, indentFirstLine (pt; negative = hanging),
 * keepNext, keepLines, pageBreakBefore (true only) }. Offline on the fake; the LIVE acceptance is gh292b-docx-style-fields-product.
 */
const FILE = '1300292';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const NAME = 'GH292 Fields';
const FIELDS = { text: { italic: true, underline: true, strike: true, fontFamily: 'Courier New' },
  paragraph: { spacingBefore: 12, indentRight: 36, indentFirstLine: -18, keepNext: true, keepLines: true, pageBreakBefore: true } };
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const req = steps => ({ fileId: FILE, task: { steps } });

function product(script = {}) {
  const store = createStore([{ text: 'Body', props: PLAIN }], { doc: { styles: {} } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readStyle = async name => { const r = await read.handler({ fileId: FILE, operation: 'read_styles', names: [name] }); return r.result.styles[0]; };
  return { store, tool, readStyle, writes: () => store.saves };
}

test('GH #292: define_style with every further field -> APPLIED; read_styles by name reads them back in points; a retry is SATISFIED with 0 writes; modify_style changes one of them', async () => {
  const p = product();
  const r = await p.tool.handler(req([{ intent: 'define_style', name: NAME, ...FIELDS }]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal(r.operations[0].status, 'APPLIED');
  const st = await p.readStyle(NAME);
  for (const [k, v] of Object.entries(FIELDS.text)) assert.equal(st.text[k], v, `text.${k}`);
  for (const [k, v] of Object.entries(FIELDS.paragraph)) assert.equal(st.paragraph[k], v, `paragraph.${k}`);
  const saves = p.writes();
  const again = await p.tool.handler(req([{ intent: 'define_style', name: NAME, ...FIELDS }]));
  assert.equal(again.ok, true); assert.equal(again.noOp, true); assert.equal(p.writes(), saves);
  const m = await p.tool.handler(req([{ intent: 'modify_style', name: NAME, paragraph: { indentFirstLine: 24 } }]));
  assert.equal(m.ok, true, JSON.stringify(m).slice(0, 400));
  assert.equal((await p.readStyle(NAME)).paragraph.indentFirstLine, 24);
});

test('GH #292: refused before dispatch -- a non-boolean flag (false is released, GH #292 S8), a blank font, an off-grid point value; still unknown: highlight / caps / lineSpacing / widowControl', () => {
  for (const s of [{ text: { italic: 'no' } }, { paragraph: { keepNext: 0 } }, { text: { fontFamily: '   ' } }, { paragraph: { spacingBefore: 12.1 } }, { paragraph: { indentFirstLine: -18.05 } },
    { text: { highlight: 'yellow' } }, { text: { caps: true } }, { paragraph: { lineSpacing: 1.5 } }, { paragraph: { widowControl: false } }]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'define_style', name: NAME, ...s }]));
    assert.equal(v.ok, false, JSON.stringify(s));
    assert.equal(v.dispatched, false, JSON.stringify(s));
  }
});

test('GH #292: a write that misses a requested field is never APPLIED (the editor ignores the command)', async () => {
  const p = product({ fault: 'ignore' });
  const r = await p.tool.handler(req([{ intent: 'define_style', name: NAME, ...FIELDS }]));
  assert.notEqual(r.ok, true);
  assert.notEqual((r.operations || [])[0] && r.operations[0].status, 'APPLIED');
});

test('GH #292: the REAL command body calls each measured setter with the measured arguments (twips, auto-spacing false, a negative first line)', () => {
  const calls = [];
  const rec = name => (...a) => calls.push([name, ...a]);
  const tp = Object.fromEntries(['SetItalic', 'SetUnderline', 'SetStrikeout', 'SetFontFamily'].map(n => [n, rec(n)]));
  const pp = Object.fromEntries(['SetSpacingBefore', 'SetIndRight', 'SetIndFirstLine', 'SetKeepNext', 'SetKeepLines', 'SetPageBreakBefore'].map(n => [n, rec(n)]));
  const style = { GetName: () => NAME, GetTextPr: () => tp, GetParaPr: () => pp };
  const doc = { GetStyle: () => ({ GetName: () => 'Normal' }), CreateStyle: () => style };
  const sandbox = { Api: { GetDocument: () => doc } };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'define', name: NAME, fields: FIELDS })} })()`, sandbox);
  assert.deepEqual(calls, [['SetItalic', true], ['SetUnderline', true], ['SetStrikeout', true], ['SetFontFamily', 'Courier New'],
    ['SetSpacingBefore', 240, false], ['SetIndRight', 720], ['SetIndFirstLine', -360], ['SetKeepNext', true], ['SetKeepLines', true], ['SetPageBreakBefore', true]]);
});
