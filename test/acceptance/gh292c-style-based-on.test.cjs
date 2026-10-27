'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { commandBody, styleNamesOf } = require('../../src/adapters/docx/families/style-definition.cjs');
const { PROPERTY_QUERY, buildTargetedStyleQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * GH #292 S3: define_style takes `basedOn`, the NAME of an existing paragraph style (docs/evidence/gh292-style-definition-probe-20261021.json:
 * ApiStyle.SetBasedOn(doc.GetStyle(base)) persists -- saved w:basedOn, the reopened ToJSON basedOn = the base's styleId; CreateStyle's 3rd
 * argument stores nothing). The driver resolves that styleId back to a name (identity.basedOnName). Offline on the fake; the LIVE acceptance
 * is gh292c-docx-style-based-on-product.
 */
const FILE = '1300293';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const NAME = 'GH292c Based';
const BASE = 'Heading 1';
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const req = steps => ({ fileId: FILE, task: { steps } });
const ds = (extra = {}) => ({ intent: 'define_style', name: NAME, basedOn: BASE, text: { bold: true }, ...extra });

function product(script = {}) {
  const store = createStore([{ text: 'Body', props: PLAIN }], { doc: { styles: { Normal: { text: {}, paragraph: {} }, [BASE]: { text: { bold: true }, paragraph: {} } } } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readStyle = async name => { const r = await read.handler({ fileId: FILE, operation: 'read_styles', names: [name] }); return r.result.styles[0]; };
  return { store, tool, readStyle, writes: () => store.saves };
}

test('GH #292 S3: define_style basedOn -> APPLIED; read_styles names the base; a retry is SATISFIED with 0 writes; basedOn alone is a definition', async () => {
  const p = product();
  const r = await p.tool.handler(req([ds()]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal(r.operations[0].status, 'APPLIED');
  assert.equal((await p.readStyle(NAME)).identity.basedOnName, BASE);
  const saves = p.writes();
  const again = await p.tool.handler(req([ds()]));
  assert.equal(again.ok, true); assert.equal(again.noOp, true); assert.equal(p.writes(), saves);
  const only = await p.tool.handler(req([{ intent: 'define_style', name: 'GH292c Only base', basedOn: 'Normal' }]));
  assert.equal(only.ok, true, JSON.stringify(only).slice(0, 400));
  assert.equal((await p.readStyle('GH292c Only base')).identity.basedOnName, 'Normal');
});

test('GH #292 S3: a base that does not exist fails closed before any write; a taken name with another base is FAIL_CLOSED naming basedOn', async () => {
  const p = product();
  const r = await p.tool.handler(req([ds({ basedOn: 'GH292c No such style' })]));
  assert.notEqual(r.ok, true);
  assert.match(JSON.stringify(r), /define_style-base-style-does-not-exist:GH292c No such style/);
  assert.equal(p.writes(), 0);
  assert.equal((await p.readStyle(NAME)).gap, 'not-found');
  assert.equal((await p.tool.handler(req([ds()]))).ok, true);
  const other = await p.tool.handler(req([ds({ basedOn: 'Normal' })]));
  assert.notEqual(other.ok, true);
  assert.match(JSON.stringify(other), /define_style-name-already-exists-with-different-fields:GH292c Based:basedOn/);
});

test('GH #292 S3: a write that drops the base is never APPLIED (negative control)', async () => {
  const p = product({ fault: 'sd-basedon-dropped' });
  const r = await p.tool.handler(req([ds()]));
  assert.notEqual(r.ok, true);
  assert.notEqual((r.operations || [])[0] && r.operations[0].status, 'APPLIED');
});

test('GH #292 S3: refused before dispatch -- basedOn on modify_style, blank, naming the style itself, a non-string', () => {
  for (const s of [{ intent: 'modify_style', name: 'Normal', text: { bold: true }, basedOn: BASE }, ds({ basedOn: '  ' }), ds({ basedOn: NAME }), ds({ basedOn: 3 })]) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.ok, false, JSON.stringify(s));
    assert.equal(v.dispatched, false, JSON.stringify(s));
  }
});

test('GH #292 S3: the base is read by name too (styleNamesOf), so a session-created base is visible to the targeted read', () => {
  assert.deepEqual(styleNamesOf({ steps: [ds(), { intent: 'modify_style', name: 'X', text: { bold: true } }] }), [NAME, 'X', BASE]);
});

test('GH #292 S3: the REAL command body resolves the base BEFORE the create and calls SetBasedOn with it; a missing or non-paragraph base creates nothing', () => {
  const run = (baseType, baseName = BASE) => {
    const calls = [];
    const style = { GetName: () => NAME, GetTextPr: () => ({ SetBold: v => calls.push(['SetBold', v]) }), GetParaPr: () => ({}), SetBasedOn: b => calls.push(['SetBasedOn', b.GetName()]) };
    const base = { GetName: () => baseName, GetType: () => baseType };
    const doc = { GetStyle: n => (n === NAME ? { GetName: () => 'Normal' } : base), CreateStyle: (n, t) => { calls.push(['CreateStyle', n, t]); return style; } };
    const sandbox = { Api: { GetDocument: () => doc } };
    vm.createContext(sandbox);
    vm.runInContext(`(function(){ ${commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode: 'define', name: NAME, fields: { basedOn: BASE, text: { bold: true } } })} })()`, sandbox);
    return calls;
  };
  assert.deepEqual(run('paragraph'), [['CreateStyle', NAME, 'paragraph'], ['SetBasedOn', BASE], ['SetBold', true]]);
  assert.deepEqual(run('table'), [], 'a non-paragraph base: nothing created');
  assert.deepEqual(run('paragraph', 'Normal'), [], 'GetStyle fell back to another style (the Normal fallback): nothing created');
});

/** A fake editor for the REAL read queries: each style's ToJSON is a STRING with the measured shape (styleId + basedOn = a styleId). */
function readApi(list, enumerated) {
  const objs = Object.fromEntries(list.map(([name, styleId, basedOn]) => [name, { GetName: () => name, GetTextPr: () => ({}), GetParaPr: () => ({}), ToJSON: () => JSON.stringify({ name, styleId, basedOn, styleType: 'paragraph', rPr: {}, pPr: {} }) }]));
  const doc = { GetElementsCount: () => 0, GetElement: () => null, Search: () => [], GetAllBookmarksNames: () => [], GetBookmarkRange: () => null, GetSections: () => [], GetAllTables: () => [], GetAllImages: () => [],
    GetStyle: n => objs[n] || objs.Normal, GetAllStyles: () => enumerated.map(n => objs[n]) };
  return { Api: { GetDocument: () => doc } };
}
const runVm = (code, sandbox) => { vm.createContext(sandbox); return vm.runInContext(`(function(){ ${code} })()`, sandbox); };

test('GH #292 S3 (vm, the REAL queries): basedOn styleId -> the base NAME; a session-created base is found among the targeted names; an unknown id is named, never guessed', () => {
  const styles = [['Normal', '886', null], [BASE, '827', '886'], ['Fresh base', '1_3', null], [NAME, '1_4', '827'], ['On fresh', '1_5', '1_3'], ['Orphan', '1_6', '999']];
  const enumerated = ['Normal', BASE];
  const all = runVm(PROPERTY_QUERY, readApi(styles, enumerated));
  assert.equal(all.document.styles[BASE].identity.basedOnName, 'Normal');
  assert.equal(all.document.styles.Normal.identity.basedOnName, null);
  const t = runVm(buildTargetedStyleQuery([NAME, 'On fresh', 'Fresh base', 'Orphan']), readApi(styles, enumerated));
  assert.equal(t[NAME].identity.basedOnName, BASE, 'a built-in base, via the GetAllStyles fallback');
  assert.equal(t['On fresh'].identity.basedOnName, 'Fresh base', 'a session-created base, among the targeted names');
  assert.equal(t.Orphan.identity.basedOnName, 'UNOBSERVABLE:base-style-not-enumerated');
  assert.equal(t['Fresh base'].identity.basedOnName, null);
});
