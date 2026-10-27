'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { lower, plan } = require('../../src/adapters/docx/families/style-definition.cjs');

/**
 * GH #292 S4: modify_style `next` / `link` -- the word /4 Bridge WRITE docx.style.nextLink.write/v1 (owner GO 2026-10-22; measured:
 * docs/evidence/gh292s4-next-link-internal-probe-20261022.json, gh292s4b-next-link-internal-probe-run2-20261022.json). Offline on the
 * fake: the fake models the Bridge write on the accepted runtime (word /5 since the GH #292 S4 D-state, the handler byte-identical to /4);
 * on the retired /3 and /4 the write is refused before any dispatch.
 */
const FILE = '1300295';
const V5 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const req = steps => ({ fileId: FILE, task: { steps } });
const STYLES = { Normal: { text: {}, paragraph: {} }, 'Body Plain': { text: {}, paragraph: {} }, 'Heading 1': { text: { bold: true }, paragraph: {}, next: 'Normal', link: 'Heading 1 Char' },
  'Heading 1 Char': { type: 'character', text: { bold: true }, paragraph: {}, link: 'Heading 1' }, 'Body Char': { type: 'character', text: {}, paragraph: {} } };

function product(script = { runtime: V5 }) {
  const store = createStore([{ text: 'Body', props: PLAIN }], { doc: { styles: JSON.parse(JSON.stringify(STYLES)) } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, script) });
  const readStyle = async name => { const r = await read.handler({ fileId: FILE, operation: 'read_styles', names: [name] }); return r.result.styles[0]; };
  return { store, tool, readStyle, writes: () => store.saves };
}

test('GH #292 S4: next -> APPLIED on /5; read_styles names it; a retry is SATISFIED with 0 writes', async () => {
  const p = product();
  const r = await p.tool.handler(req([{ intent: 'modify_style', name: 'Body Plain', next: 'Heading 1' }]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal(r.operations[0].status, 'APPLIED');
  assert.equal((await p.readStyle('Body Plain')).identity.nextName, 'Heading 1');
  const saves = p.writes();
  const again = await p.tool.handler(req([{ intent: 'modify_style', name: 'Body Plain', next: 'Heading 1' }]));
  assert.equal(again.ok, true); assert.equal(again.noOp, true); assert.equal(p.writes(), saves);
});

test('GH #292 S4: link -> BOTH sides (L1 checks the pair, L3 lets only the target\'s own link change); the reverse direction is then SATISFIED', async () => {
  const p = product();
  const r = await p.tool.handler(req([{ intent: 'modify_style', name: 'Body Plain', link: 'Body Char' }]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  assert.equal((await p.readStyle('Body Plain')).identity.linkName, 'Body Char');
  assert.equal((await p.readStyle('Body Char')).identity.linkName, 'Body Plain');
  const back = await p.tool.handler(req([{ intent: 'modify_style', name: 'Body Char', link: 'Body Plain' }]));
  assert.equal(back.ok, true); assert.equal(back.noOp, true);
});

test('GH #292 S4: refused BEFORE any write, named: wrong types, a missing target, a style already linked elsewhere', async () => {
  const p = product();
  const saves = p.writes();
  for (const [step, re] of [
    [{ intent: 'modify_style', name: 'Body Char', next: 'Normal' }, /modify_style-next-type-mismatch:Body Char:Normal/],
    [{ intent: 'modify_style', name: 'Body Plain', next: 'Body Char' }, /modify_style-next-type-mismatch:Body Plain:Body Char/],
    [{ intent: 'modify_style', name: 'Body Plain', next: 'No Such Style' }, /modify_style-next-style-does-not-exist:No Such Style/],
    [{ intent: 'modify_style', name: 'Normal', link: 'Body Plain' }, /modify_style-link-type-mismatch:Normal:Body Plain/],
    [{ intent: 'modify_style', name: 'Body Plain', link: 'Heading 1 Char' }, /modify_style-link-already-linked-elsewhere:Body Plain:Heading 1 Char/],
  ]) {
    const r = await p.tool.handler(req([step]));
    assert.notEqual(r.ok, true, JSON.stringify(step));
    assert.match(JSON.stringify(r), re);
  }
  assert.equal(p.writes(), saves, 'no write');
});

test('GH #292 S4 D-state: on the retired word /3 (no next / link WRITE) and /4 the step is refused, named, before any dispatch', async () => {
  for (const old of ['3', '4']) {
    const p = product({ runtime: `9.3.5|euro-office-bridge-word-v1/9.3.5.0/${old}` });
    const r = await p.tool.handler(req([{ intent: 'modify_style', name: 'Body Plain', next: 'Heading 1' }]));
    assert.notEqual(r.ok, true, old);
    assert.match(JSON.stringify(r), /docx-style-next-link-bridge-write-runtime-not-proven/, old);
    assert.equal((await p.readStyle('Body Plain')).identity.nextName, null, old);
  }
});

test('GH #292 S4: the verification catches a setter that does not take, and a link written on one side only (L1 never APPLIED)', async () => {
  for (const fault of ['sd-nextlink-dropped', 'sd-link-one-side']) {
    const p = product({ runtime: V5, fault });
    const step = fault === 'sd-nextlink-dropped' ? { intent: 'modify_style', name: 'Body Plain', next: 'Heading 1' } : { intent: 'modify_style', name: 'Body Plain', link: 'Body Char' };
    const r = await p.tool.handler(req([step]));
    assert.notEqual(r.ok, true, fault);
    assert.notEqual(r.operations && r.operations[0] && r.operations[0].status, 'APPLIED', fault);
  }
});

test('GH #292 S4 schema / plan: define_style next|link, a blank name, a self-link and an unknown field are refused before any read', async () => {
  const p = product();
  for (const step of [{ intent: 'define_style', name: 'X1', next: 'Normal' }, { intent: 'modify_style', name: 'Body Plain', next: ' ' },
    { intent: 'modify_style', name: 'Body Plain', link: 'Body Plain' }, { intent: 'modify_style', name: 'Body Plain', nextStyle: 'Normal' }]) {
    const r = await p.tool.handler(req([step]));
    assert.notEqual(r.ok, true, JSON.stringify(step));
  }
  assert.equal(plan({ intent: 'define_style', name: 'X1', next: 'Normal' }, 0).ok, false);
  assert.equal(plan({ intent: 'modify_style', name: 'Body Plain', next: 'Normal' }, 0).ok, true);
});

test('GH #292 S4 command body (vm): the Bridge gate runs BEFORE any public setter -- a gate refusal writes nothing; a write receipt failure is returned, named', () => {
  const op = plan({ intent: 'modify_style', name: 'Body Plain', next: 'Heading 1', text: { bold: true } }, 0).operation;
  const cmd = lower(op, { kind: 'style_definition', name: 'Body Plain', action: 'modify', fields: op.fields, styleNextLinkBridge: true });
  assert.equal(cmd.meta.bridgeStyleNextLink, true);
  const run = (gate, write) => {
    const setBold = []; const style = { GetName: () => 'Body Plain', GetTextPr: () => ({ SetBold: v => setBold.push(v) }), GetParaPr: () => null };
    const ctx = { Api: { GetDocument: () => ({ GetStyle: () => style }) }, __eoS4Gate: gate, __eoS4Write: write };
    const r = vm.runInNewContext(`(function(){ ${cmd.body} })()`, ctx);
    return { r: r === undefined ? undefined : JSON.parse(JSON.stringify(r)), setBold };
  };
  const refused = run(() => ({ ok: false, reason: 'PATCH_VERSION_MISMATCH' }), () => { throw new Error('must not write'); });
  assert.deepEqual(refused.r, { ok: false, stage: 'bridge-write-gate', reason: 'PATCH_VERSION_MISMATCH', mutation: 'NONE' });
  assert.deepEqual(refused.setBold, [], 'no public setter ran');
  const failed = run(() => ({ ok: true, bridge: {} }), () => ({ ok: false, reason: 'BRIDGE:write-not-observed', uncertain: true, mutation: 'UNKNOWN' }));
  assert.deepEqual(failed.r, { stage: 'bridge-write', field: 'next', ok: false, reason: 'BRIDGE:write-not-observed', uncertain: true, mutation: 'UNKNOWN' });
  const ok = run(() => ({ ok: true, bridge: {} }), (b, args) => (args.field === 'next' && args.target === 'Heading 1' && args.style === 'Body Plain' ? { ok: true } : { ok: false }));
  assert.deepEqual(ok.r, { ok: true }); assert.deepEqual(ok.setBold, [true]);
});
