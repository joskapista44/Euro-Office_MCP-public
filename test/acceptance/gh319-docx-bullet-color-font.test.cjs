'use strict';

const test = require('../helpers/trim-skip.cjs').testFor('set_numbering'); // release 3 trim: same gate gh248's own file uses
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

/**
 * GH #319 (LIVE-measured, gh319-bulletcolorfont-probe-20261028): the bullet GLYPH's own colour
 * and typeface, additive to GH #248's released set_numbering (code/start/suffix). The write path
 * (through `level.GetTextPr().SetColor`/`SetFontFamily`, no direct setter on the level) is
 * mirrored in test/fakes/fake-docx-driver.cjs's `applySetNumbering` -- the SAME shared-slot shape
 * `format`/`start`/`suffix` already use. Offline only, same house style as gh248's own file; the
 * LIVE acceptance is test/live/docx-gh319-bullet-color-font-probe.live.cjs (not run by the author).
 *
 * The card's own rule: `bulletColor`/`bulletFont` are CODE:'bullet'-ONLY, in the SAME step --
 * the glyph's colour/typeface is a property of the template being WRITTEN here, never of an
 * existing numbering whose code the caller leaves untouched.
 */

const FILE = '1303627';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const SOLO_ITEM = 'GH319 solo bullet item';
const OTHER_CODE_ITEM = 'GH319 decimal item';
const PLAIN_PARA = 'GH319 plain paragraph';
const item = (text, numId, type = 'bullet', level = 0) => ({ text, props: PLAIN, list: { type, level, numId } });
const doc = () => [item(SOLO_ITEM, 'n1'), item(OTHER_CODE_ITEM, 'n3', 'numbered'), { text: PLAIN_PARA, props: PLAIN }];
const sn = (fields, text = SOLO_ITEM, extra = {}) => ({ intent: 'set_numbering', selector: { text }, ...fields, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product(formats = {}, script = {}) {
  const store = createStore(doc(), { doc: { numberingFormats: formats } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const numberingOf = async text => { const r = await read.handler({ fileId: FILE, operation: 'read_numbering' }); assert.equal(r.ok, true, JSON.stringify(r)); const it = r.result.items.find(x => x.text === text); assert.ok(it, text); return it.numbering; };
  return { store, tool, numberingOf, writes: () => store.saves };
}

// ---- the step schema (pre-dispatch) ---------------------------------------------------------------------------------------------

test('SCHEMA: bulletColor / bulletFont, alone or together, WITH code: "bullet" in the same step, are accepted', () => {
  assert.ok(D.inputSchema().properties.task.properties.steps.items.oneOf.some(s => s.properties.intent.const === 'set_numbering' && s.properties.bulletColor && s.properties.bulletFont));
  for (const fields of [
    { code: 'bullet', bulletColor: '#0000FF' },
    { code: 'bullet', bulletFont: 'Wingdings' },
    { code: 'bullet', bulletColor: '#ff00aa', bulletFont: 'Symbol' },
    { code: 'bullet', bulletColor: '#0000FF', start: 1, suffix: 'tab' },
  ]) assert.equal(verdict([sn(fields)]).ok, true, JSON.stringify(fields));
});

test('SCHEMA: bulletColor / bulletFont WITHOUT code: "bullet" in the same step -- refused pre-dispatch, named, nothing written', () => {
  const refused = [
    sn({ bulletColor: '#0000FF' }),             // no code at all
    sn({ code: '1.', bulletColor: '#0000FF' }), // a DIFFERENT code
    sn({ code: 'none', bulletFont: 'Symbol' }), // a DIFFERENT code, the other field
  ];
  for (const step of refused) {
    const v = verdict([step]);
    assert.equal(v.ok, false, JSON.stringify(step));
    assert.equal(v.dispatched, false);
  }
  // bulletFont ALONE (no code/start/suffix at all) is caught by the schema-level "at least one of
  // code, start, suffix" constraint BEFORE the request ever reaches the adapter's plan() -- same
  // refusal outcome, a more generic message; the adapter's OWN, field-specific message is reached
  // only once a code IS present (the cases below, where that earlier gate does not fire).
  const v2 = verdict([sn({ bulletFont: 'Wingdings' })]);
  assert.equal(v2.ok, false);
  const v3 = verdict([sn({ code: 'a.', bulletColor: '#0000FF' })]);
  assert.match(JSON.stringify(v3), /bulletColor is only valid with code: 'bullet'/);
  const v4 = verdict([sn({ code: 'none', bulletFont: 'Symbol' })]);
  assert.match(JSON.stringify(v4), /bulletFont is only valid with code: 'bullet'/);
});

test('SCHEMA: a malformed bulletColor (not #RRGGBB) or a blank bulletFont is REFUSED_SHAPE, named', () => {
  for (const bad of ['0000FF', '#00F', '#GGGGGG', 'blue', '#0000FF00']) {
    const v = verdict([sn({ code: 'bullet', bulletColor: bad })]);
    assert.equal(v.ok, false, bad);
  }
  const v = verdict([sn({ code: 'bullet', bulletFont: '   ' })]);
  assert.equal(v.ok, false);
  assert.match(JSON.stringify(v), /bulletFont must be a non-empty string/);
});

// ---- WRITE + READ-BACK -----------------------------------------------------------------------------------------------------------

test('WRITE + READ-BACK: code: "bullet" + bulletColor + bulletFont on a solo item lands both; read_numbering returns them; the retry writes nothing', async () => {
  const p = product();
  const r = await p.tool.handler(req([sn({ code: 'bullet', bulletColor: '#0000FF', bulletFont: 'Wingdings' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const n = await p.numberingOf(SOLO_ITEM);
  assert.equal(n.format.numFmt, 'bullet');
  assert.equal(n.format.color, '#0000FF');
  assert.equal(n.format.font, 'Wingdings');
  const writes = p.writes();
  const again = await p.tool.handler(req([sn({ code: 'bullet', bulletColor: '#0000FF', bulletFont: 'Wingdings' })]));
  assert.equal(again.ok, true, JSON.stringify(again));
  assert.equal(p.writes(), writes, 'SATISFIED: already holding the requested value writes nothing');
});

test('WRITE: bulletColor alone (no bulletFont) lands only the colour; the glyph font key stays absent (UNOBSERVABLE), never guessed', async () => {
  const p = product();
  const r = await p.tool.handler(req([sn({ code: 'bullet', bulletColor: '#FF00AA' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const n = await p.numberingOf(SOLO_ITEM);
  assert.equal(n.format.color, '#FF00AA');
  assert.equal(n.format.font, 'UNOBSERVABLE:font-not-set');
});

test('WRITE: bulletColor is case-insensitive on compare (the fake stores verbatim, L1 uppercases before matching)', async () => {
  const p = product();
  const r = await p.tool.handler(req([sn({ code: 'bullet', bulletColor: '#aa00ff' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const n = await p.numberingOf(SOLO_ITEM);
  assert.equal(n.format.color, '#AA00FF', 'stored uppercased, same normalizeWanted-style convention as formatting.cjs');
});

test('FAULT: the glyph colour write silently did not land (L1 catches it, never reports ok on an unobserved value)', async () => {
  const p = product({}, { fault: 'numbering-bulletcolor-not-written' });
  const r = await p.tool.handler(req([sn({ code: 'bullet', bulletColor: '#0000FF' })]));
  assert.equal(r.ok, false, JSON.stringify(r));
});

test('FAULT: the glyph font write silently did not land (L1 catches it)', async () => {
  const p = product({}, { fault: 'numbering-bulletfont-not-written' });
  const r = await p.tool.handler(req([sn({ code: 'bullet', bulletFont: 'Wingdings' })]));
  assert.equal(r.ok, false, JSON.stringify(r));
});

test('FAIL-CLOSED: a plain (non-list) paragraph with code: "bullet" + bulletColor is refused with a named reason, nothing written', async () => {
  const p = product();
  const r = await p.tool.handler(req([sn({ code: 'bullet', bulletColor: '#0000FF' }, PLAIN_PARA)]));
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r), /not-listed/);
  assert.equal(p.writes(), 0);
});

// ---- the adapter plan() itself (pure, no driver) ---------------------------------------------------------------------------------

test('ADAPTER: plan() carries bulletColor (uppercased) and bulletFont (verbatim) into operation.fields, only alongside code: "bullet"', () => {
  const r = planOperations({ steps: [{ intent: 'set_numbering', selector: { text: SOLO_ITEM }, code: 'bullet', bulletColor: '#0000ff', bulletFont: 'Wingdings' }] });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations[0].fields.bulletColor, '#0000FF');
  assert.deepEqual(r.operations[0].fields.bulletFont, 'Wingdings');
  const bad = planOperations({ steps: [{ intent: 'set_numbering', selector: { text: SOLO_ITEM }, code: '1.', bulletColor: '#0000FF' }] });
  assert.equal(bad.ok, false);
});
