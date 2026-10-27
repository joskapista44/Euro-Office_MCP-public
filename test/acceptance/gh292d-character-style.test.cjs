'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { commandBody, lower, plan } = require('../../src/adapters/docx/families/style-definition.cjs');

/**
 * GH #292 S5: define_style `type: 'character'` creates a character style (docs/evidence/gh292s5-character-style-probe-20261022.json:
 * CreateStyle(name, 'run') -> saved w:type="character", ToJSON styleType 'characterStyle'; every text setter persists; SetBasedOn takes
 * only a base of the same type; the ParaPr of a character style takes SetJc too -- stored, meaningless -- so the paragraph group is
 * refused on one). Offline on the fake; the LIVE acceptance is gh292d-docx-character-style-product.
 */
const FILE = '1300294';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const NAME = 'GH292d Emphasis';
const TEXT = { bold: true, italic: true, color: '2255AA', size: 12, fontFamily: 'Courier New' };
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const req = steps => ({ fileId: FILE, task: { steps } });
const cs = (extra = {}) => ({ intent: 'define_style', name: NAME, type: 'character', text: TEXT, ...extra });

function product(script = {}) {
  const store = createStore([{ text: 'Body', props: PLAIN }], { doc: { styles: { Normal: { text: {}, paragraph: {} }, 'Heading 1': { text: { bold: true }, paragraph: {} },
    'Default Paragraph Font': { type: 'character', text: {}, paragraph: {} } } } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readStyle = async name => { const r = await read.handler({ fileId: FILE, operation: 'read_styles', names: [name] }); return r.result.styles[0]; };
  return { store, tool, readStyle, writes: () => store.saves };
}

test('GH #292 S5: define_style type character -> APPLIED; read_styles says characterStyle with the text fields; a retry is SATISFIED with 0 writes', async () => {
  const p = product();
  const r = await p.tool.handler(req([cs()]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal(r.operations[0].status, 'APPLIED');
  const st = await p.readStyle(NAME);
  assert.equal(st.identity.styleType, 'characterStyle');
  for (const [k, v] of Object.entries(TEXT)) assert.equal(st.text[k], v, `text.${k}`);
  const saves = p.writes();
  const again = await p.tool.handler(req([cs()]));
  assert.equal(again.ok, true); assert.equal(again.noOp, true); assert.equal(p.writes(), saves);
});

test('GH #292 S5: a character style based on a character style (Default Paragraph Font, or one created before) is APPLIED; on a paragraph style it fails closed before any write', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([cs({ basedOn: 'Default Paragraph Font' })]))).ok, true);
  assert.equal((await p.readStyle(NAME)).identity.basedOnName, 'Default Paragraph Font');
  const on = await p.tool.handler(req([{ intent: 'define_style', name: 'GH292d On emphasis', type: 'character', basedOn: NAME }]));
  assert.equal(on.ok, true, JSON.stringify(on).slice(0, 400));
  const saves = p.writes();
  const bad = await p.tool.handler(req([{ intent: 'define_style', name: 'GH292d Wrong base', type: 'character', basedOn: 'Heading 1', text: { bold: true } }]));
  assert.notEqual(bad.ok, true);
  assert.match(JSON.stringify(bad), /define_style-base-style-type-mismatch:Heading 1:character/);
  const para = await p.tool.handler(req([{ intent: 'define_style', name: 'GH292d Para on char', basedOn: NAME, text: { bold: true } }]));
  assert.match(JSON.stringify(para), /define_style-base-style-type-mismatch/, 'a paragraph style on a character base, the other direction');
  assert.equal(p.writes(), saves);
});

test('GH #292 S5: a taken name of the other type is FAIL_CLOSED naming type; modify_style with a paragraph group on a character style fails closed; text on one is APPLIED', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([cs()]))).ok, true);
  const asPara = await p.tool.handler(req([{ intent: 'define_style', name: NAME, text: TEXT }]));
  assert.match(JSON.stringify(asPara), /define_style-name-already-exists-with-different-fields:GH292d Emphasis:type/);
  const m = await p.tool.handler(req([{ intent: 'modify_style', name: NAME, paragraph: { align: 'center' } }]));
  assert.match(JSON.stringify(m), /modify_style-paragraph-fields-on-a-character-style:GH292d Emphasis/);
  const mt = await p.tool.handler(req([{ intent: 'modify_style', name: NAME, text: { underline: true } }]));
  assert.equal(mt.ok, true, JSON.stringify(mt).slice(0, 400));
  assert.equal((await p.readStyle(NAME)).text.underline, true);
});

test('GH #292 S5: refused before dispatch -- a paragraph group on a character style, an unknown type, type on modify_style', () => {
  for (const s of [cs({ paragraph: { align: 'left' } }), cs({ type: 'table' }), cs({ type: 'run' }), { intent: 'modify_style', name: NAME, type: 'character', text: { bold: true } }]) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.ok, false, JSON.stringify(s));
    assert.equal(v.dispatched, false, JSON.stringify(s));
  }
  assert.match(plan(cs({ paragraph: { align: 'left' } }), 0).detail, /paragraph is not allowed on a character style/);
});

test('GH #292 S5: the REAL command body creates with CreateStyle(name, "run") and checks the base is a run style; an untyped meta stays a paragraph style', () => {
  const run = (meta, baseType) => {
    const calls = [];
    const style = { GetName: () => meta.name, GetTextPr: () => ({ SetBold: v => calls.push(['SetBold', v]) }), GetParaPr: () => ({}), SetBasedOn: b => calls.push(['SetBasedOn', b.GetName()]) };
    const base = { GetName: () => 'Base', GetType: () => baseType };
    const doc = { GetStyle: n => (n === 'Base' ? base : { GetName: () => 'Normal' }), CreateStyle: (n, t) => { calls.push(['CreateStyle', n, t]); return style; } };
    const sandbox = { Api: { GetDocument: () => doc } };
    vm.createContext(sandbox);
    vm.runInContext(`(function(){ ${commandBody(meta)} })()`, sandbox);
    return calls;
  };
  const op = plan(cs({ basedOn: 'Base', text: { bold: true } }), 0).operation;
  const meta = lower(op, { kind: 'style_definition', name: op.name, action: op.action, fields: op.fields }).meta;
  assert.deepEqual(run(meta, 'run'), [['CreateStyle', NAME, 'run'], ['SetBasedOn', 'Base'], ['SetBold', true]]);
  assert.deepEqual(run(meta, 'paragraph'), [], 'a paragraph base for a character style: nothing created');
  assert.deepEqual(run({ family: 'styleDefinition', action: 'styleDefinition', mode: 'define', name: 'X', fields: { text: { bold: true } } }, null), [['CreateStyle', 'X', 'paragraph'], ['SetBold', true]]);
});
