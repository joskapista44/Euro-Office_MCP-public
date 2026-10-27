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
 * GH #292 S8: on a style, the flags take false and text.color takes 'auto' (docs/evidence/gh292s8-style-unset-probe-20261022.json:
 * SetItalic / SetUnderline / SetStrikeout / SetKeepNext / SetKeepLines / SetPageBreakBefore(false) persist as an EXPLICIT off -- saved
 * w:i w:val="0", w:u w:val="none", w:keepNext w:val="false", ToJSON false -- also after a true; SetColor(Api.AutoColor()) saves w:color
 * auto, ToJSON color.auto). NOT a removal: the style's TextPr / ParaPr has no Clear / Remove / Reset. Offline on the fake; the LIVE
 * acceptance is gh292e-docx-style-off-and-auto-product.
 */
const FILE = '1300295';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const NAME = 'GH292e Off';
const ON = { text: { italic: true, underline: true, strike: true, color: 'AA2222' }, paragraph: { keepNext: true, keepLines: true, pageBreakBefore: true } };
const OFF = { text: { italic: false, underline: false, strike: false, color: 'auto' }, paragraph: { keepNext: false, keepLines: false, pageBreakBefore: false } };
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const req = steps => ({ fileId: FILE, task: { steps } });

function product(script = {}) {
  const store = createStore([{ text: 'Body', props: PLAIN }], { doc: { styles: { Normal: { text: {}, paragraph: {} } } } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readStyle = async name => { const r = await read.handler({ fileId: FILE, operation: 'read_styles', names: [name] }); return r.result.styles[0]; };
  return { store, tool, readStyle, writes: () => store.saves };
}
const holds = (st, fields) => Object.entries(fields.text).every(([k, v]) => st.text[k] === v) && Object.entries(fields.paragraph).every(([k, v]) => st.paragraph[k] === v);

test('GH #292 S8: define_style with the flags false and color auto -> APPLIED, read back false / auto; a retry is SATISFIED with 0 writes', async () => {
  const p = product();
  const r = await p.tool.handler(req([{ intent: 'define_style', name: NAME, ...OFF }]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal(r.operations[0].status, 'APPLIED');
  assert.ok(holds(await p.readStyle(NAME), OFF));
  const saves = p.writes();
  const again = await p.tool.handler(req([{ intent: 'define_style', name: NAME, ...OFF }]));
  assert.equal(again.ok, true); assert.equal(again.noOp, true); assert.equal(p.writes(), saves);
});

test('GH #292 S8: modify_style turns every flag on and back off, and a colour back to auto', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([{ intent: 'define_style', name: NAME, ...ON }]))).ok, true);
  assert.ok(holds(await p.readStyle(NAME), ON));
  const m = await p.tool.handler(req([{ intent: 'modify_style', name: NAME, ...OFF }]));
  assert.equal(m.ok, true, JSON.stringify(m).slice(0, 400));
  assert.ok(holds(await p.readStyle(NAME), OFF));
});

test('GH #292 S8: still refused before dispatch -- a non-boolean flag, a colour that is neither hex nor exactly auto', () => {
  for (const s of [{ text: { italic: 'false' } }, { paragraph: { keepNext: null } }, { text: { color: 'Auto' } }, { text: { color: 'automatic' } }]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'define_style', name: NAME, ...s }]));
    assert.equal(v.ok, false, JSON.stringify(s));
    assert.equal(v.dispatched, false, JSON.stringify(s));
  }
});

test('GH #292 S8: the REAL command body passes false to each flag setter and calls SetColor(Api.AutoColor()) for auto', () => {
  const calls = [];
  const rec = name => (...a) => calls.push([name, ...a]);
  const tp = Object.fromEntries(['SetItalic', 'SetUnderline', 'SetStrikeout', 'SetColor'].map(n => [n, rec(n)]));
  const pp = Object.fromEntries(['SetKeepNext', 'SetKeepLines', 'SetPageBreakBefore'].map(n => [n, rec(n)]));
  const style = { GetName: () => NAME, GetTextPr: () => tp, GetParaPr: () => pp };
  const doc = { GetStyle: () => ({ GetName: () => 'Normal' }), CreateStyle: () => style };
  const sandbox = { Api: { GetDocument: () => doc, AutoColor: () => 'AUTO-COLOR', RGB: () => 'RGB' } };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'define', name: NAME, fields: OFF })} })()`, sandbox);
  assert.deepEqual(calls, [['SetColor', 'AUTO-COLOR'], ['SetItalic', false], ['SetUnderline', false], ['SetStrikeout', false],
    ['SetKeepNext', false], ['SetKeepLines', false], ['SetPageBreakBefore', false]]);
});
