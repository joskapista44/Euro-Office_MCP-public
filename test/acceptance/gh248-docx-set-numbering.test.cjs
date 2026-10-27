'use strict';

const test = require('../helpers/trim-skip.cjs').testFor('set_numbering'); // release 3 trim: skipped while set_numbering is not in the allowlist
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
// GH #293 L7: the set_numbering / set_list_type co-member guards read the word /5 Bridge's whole affected set -- the released paths run on /5
const L7_V5 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5';
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

/**
 * GH #248 (DOCX P1 WRITE; the release itself is an OWNER DECISION, alongside #228..#266): set_numbering served by docx_execute_task -- code / start / suffix on the numbering a list item ALREADY sits on.
 * Design and the measured reason for every rule: docs/GH248-DOCX-SET-NUMBERING-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on
 * test/fakes/fake-docx-driver.cjs (per numbering and level format slots, the live model). The LIVE acceptance is test/live/gh248-docx-set-numbering-product.live.cjs (not run by the author).
 *
 * The rule the card asks for: a definition other items sit on is NEVER changed silently -- here it fails closed BEFORE dispatch, with a named reason. READ<->WRITE symmetry: read_numbering returns the
 * written format (numFmt / lvlText / lvlJc / suff / start) per item, and the numId that shows who shares a definition.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const SOLO_ITEM = 'GH248 solo item';
const SHARED_A = 'GH248 shared item A';
const SHARED_B = 'GH248 shared item B';
const CHILD = 'GH248 nested child';
const PLAIN_PARA = 'GH248 plain paragraph';
const item = (text, numId, level = 0) => ({ text, props: PLAIN, list: { type: 'numbered', level, numId } });
const doc = () => [item(SOLO_ITEM, 'n1'), item(SHARED_A, 'n2'), item(SHARED_B, 'n2'), item(CHILD, 'n2', 1), { text: PLAIN_PARA, props: PLAIN }];
const sn = (fields, text = SOLO_ITEM, extra = {}) => ({ intent: 'set_numbering', selector: { text }, ...fields, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product(formats = {}, script = {}) {
  const store = createStore(doc(), { doc: { numberingFormats: formats } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, runtime: L7_V5, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const numberingOf = async text => { const r = await read.handler({ fileId: FILE, operation: 'read_numbering' }); assert.equal(r.ok, true, JSON.stringify(r)); const it = r.result.items.find(x => x.text === text); assert.ok(it, text); return it.numbering; };
  return { store, tool, numberingOf, writes: () => store.saves };
}
const snapshot = s => JSON.stringify([s.durable.map(p => [p.text, p.list || null]), s.durableDoc.numberingFormats]);

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: set_numbering is in both allowlists and the published schema; each code, start and suffix alone or together, with an optional step id, is accepted', () => {
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('set_numbering'));
  assert.ok(T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes('set_numbering'));
  assert.ok(D.inputSchema().properties.task.properties.steps.items.oneOf.some(s => s.properties.intent.const === 'set_numbering'));
  for (const code of ['1.', '1)', 'a.', 'I.', 'bullet', 'none']) assert.equal(verdict([sn({ code })]).ok, true, code);
  for (const fields of [{ start: 1 }, { start: 9999 }, { suffix: 'tab' }, { suffix: 'space' }, { suffix: 'none' }, { code: 'a.', start: 3, suffix: 'space' }]) assert.equal(verdict([sn(fields)]).ok, true, JSON.stringify(fields));
  assert.equal(verdict([sn({ code: '1.' }, SOLO_ITEM, { id: 'n1' })]).ok, true);
});

test('SCHEMA: custom / symbol (not released), a bad code or suffix, an out-of-range start, no field, a blank text, a ref selector and a caller-chosen sharedDefinition are refused pre-dispatch', () => {
  const refused = [
    [sn({ custom: { format: 'decimal', text: '%1.', align: 'left' } }), 'UNKNOWN_FIELD'], [sn({ code: 'bullet', symbol: '*' }), 'UNKNOWN_FIELD'],
    [sn({ code: '2.' }), 'REFUSED_SHAPE'], [sn({ code: 'A.' }), 'REFUSED_SHAPE'], [sn({ suffix: 'nothing' }), 'REFUSED_SHAPE'],
    [sn({ start: 0 }), 'REFUSED_SHAPE'], [sn({ start: -3 }), 'REFUSED_SHAPE'], [sn({ start: 1.5 }), 'REFUSED_SHAPE'], [sn({ start: 10000 }), 'REFUSED_SHAPE'], [sn({ start: '3' }), 'REFUSED_SHAPE'],
    [sn({}), 'REFUSED_SHAPE'], [sn({ code: '1.' }, '   '), 'REFUSED_SHAPE'],
    [{ intent: 'set_numbering', selector: { ref: 's0' }, code: '1.' }, 'UNKNOWN_FIELD'], [{ intent: 'set_numbering', code: '1.' }, 'REFUSED_SHAPE'],
    [sn({ code: '1.' }, SOLO_ITEM, { sharedDefinition: 'propagate' }), 'UNKNOWN_FIELD'], [sn({ code: '1.' }, SOLO_ITEM, { level: 1 }), 'UNKNOWN_FIELD'],
  ];
  for (const [step, code] of refused) {
    const v = verdict([step]);
    assert.equal(v.code, code, `${JSON.stringify(step)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: set_numbering is SOLO -- next to any other step (or a second set_numbering) the task is refused pre-dispatch', () => {
  assert.equal(verdict([sn({ code: '1.' }), sn({ start: 2 }, SHARED_A)]).ok, false);
  const v = verdict([sn({ code: '1.' }), { intent: 'rewrite_paragraph', selector: { text: PLAIN_PARA }, text: 'changed' }]);
  assert.equal(v.ok, false);
  assert.equal(v.dispatched, false);
});

// ---- WRITE + READ-BACK on a definition the item has to itself ---------------------------------------------------------------

test('WRITE + READ-BACK: code / start / suffix on an item that is alone on its numbering apply once; read_numbering returns the written format; the retry writes nothing', async () => {
  const p = product();
  const r = await p.tool.handler(req([sn({ code: '1)', start: 3, suffix: 'space' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const n = await p.numberingOf(SOLO_ITEM);
  assert.equal(n.format.numFmt, 'decimal');
  assert.equal(n.format.lvlText, '%1)');
  assert.equal(n.format.lvlJc, 'left');
  assert.equal(n.format.start, 3);
  assert.equal(n.format.suff, 'space');
  assert.equal(p.store.durableDoc.numberingFormats.n2, undefined, 'a different numbering is untouched');
  const writes = p.writes();
  const again = await p.tool.handler(req([sn({ code: '1)', start: 3, suffix: 'space' })]));
  assert.equal(again.ok, true, JSON.stringify(again));
  assert.equal(p.writes(), writes);
});

test('WRITE: every one of the six codes lands its (numFmt, lvlText) on a lone item', async () => {
  const want = { '1.': ['decimal', '%1.'], '1)': ['decimal', '%1)'], 'a.': ['lowerLetter', '%1.'], 'I.': ['upperRoman', '%1.'], bullet: ['bullet', '-'], none: ['none', ''] };
  for (const [code, [numFmt, lvlText]] of Object.entries(want)) {
    const p = product();
    assert.equal((await p.tool.handler(req([sn({ code })]))).ok, true, code);
    const f = (await p.numberingOf(SOLO_ITEM)).format;
    assert.deepEqual([f.numFmt, f.lvlText], [numFmt, lvlText], code);
  }
});

// ---- the card's rule: a definition other items sit on is never changed silently -------------------------------------------

test('SHARED: an item whose numbering and level another item sits on -> fail-closed BEFORE dispatch with the named reason, the numbering / level / co-member count named, the document and every format unchanged, 0 writes', async () => {
  const p = product({ n2: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' }] });
  const before = snapshot(p.store);
  for (const [fields, text] of [[{ code: 'a.' }, SHARED_A], [{ start: 5 }, SHARED_B], [{ suffix: 'space' }, SHARED_A]]) {
    const r = await p.tool.handler(req([sn(fields, text)]));
    assert.equal(r.ok, false, JSON.stringify(fields));
    assert.match(JSON.stringify(r), /set-numbering-shared-definition:n2:level0:1/, 'the reason names the numbering, the level and the number of other items');
    assert.equal(snapshot(p.store), before);
    assert.equal(p.writes(), 0);
  }
});

test('SHARED: only items at the SAME level count -- a nested child alone at its level of a shared numbering is written, and the parents\' level keeps its format', async () => {
  const p = product({ n2: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' }] });
  const r = await p.tool.handler(req([sn({ code: 'a.' }, CHILD)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(p.store.durableDoc.numberingFormats.n2[1].numFmt, 'lowerLetter');
  assert.equal(p.store.durableDoc.numberingFormats.n2[0].numFmt, 'decimal', 'the parents\' level is unchanged');
});

test('FAULT: a write at one level that also rewrote ANOTHER level of the same numbering is never APPLIED (the other-levels check of L3)', async () => {
  const p = product({ n2: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' }] }, { fault: 'numbering-touch-other-level' });
  const r = await p.tool.handler(req([sn({ code: 'a.' }, CHILD)]));
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r), /INTEGRAL_PRESERVATION_NOT_PROVEN/);
});

test('SHARED: when the shared definition ALREADY holds every requested value the step is SATISFIED with no write (nothing to change, nothing to refuse)', async () => {
  const p = product({ n2: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left', suff: 'tab' }] });
  const r = await p.tool.handler(req([sn({ code: '1.', suffix: 'tab' }, SHARED_A)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(p.writes(), 0);
});

test('FAIL-CLOSED: a plain (non-list) paragraph and an absent text are refused with a named reason, nothing written', async () => {
  for (const [text, re] of [[PLAIN_PARA, /not-listed/], ['GH248 no such text', /not.found/i]]) {
    const p = product();
    const before = snapshot(p.store);
    const r = await p.tool.handler(req([sn({ code: '1.' }, text)]));
    assert.equal(r.ok, false);
    assert.match(JSON.stringify(r), re);
    assert.equal(snapshot(p.store), before);
    assert.equal(p.writes(), 0);
  }
});

// ---- the adapter default is unchanged -------------------------------------------------------------------------------------------------

test('ADAPTER: the default policy still propagates to the co-members (W-07), only an explicit sharedDefinition: "refuse" fails closed; an unknown policy never plans', () => {
  assert.equal(planOperations({ steps: [{ intent: 'set_numbering', selector: { text: SHARED_A }, code: '1.' }] }).operations[0].sharedDefinition, undefined);
  assert.equal(planOperations({ steps: [{ intent: 'set_numbering', selector: { text: SHARED_A }, code: '1.', sharedDefinition: 'refuse' }] }).operations[0].sharedDefinition, 'refuse');
  assert.equal(planOperations({ steps: [{ intent: 'set_numbering', selector: { text: SHARED_A }, code: '1.', sharedDefinition: 'maybe' }] }).ok, false);
});
