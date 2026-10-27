'use strict';

const test = require('../helpers/trim-skip.cjs').testFor('define_style'); // release 3 trim: skipped while define_style is not in the allowlist
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #264 (DOCX P1 WRITE; the release itself is an OWNER DECISION, alongside #228..#257b): define_style served by docx_execute_task -- the adapter's own intent
 * (families/style-definition.cjs; the GitHub issue calls it define_paragraph_style) with EXACTLY the six fields it measured end to end: text { bold, size, color } and paragraph
 * { align, indentLeft, spacingAfter }. Design and the measured reason for every rule: docs/GH264-DOCX-DEFINE-STYLE-DECISION.md. Offline only: the product handler runs the REAL
 * docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs, which stores the style in the editor's RAW units (half-points / twips, as the real file does).
 * The LIVE acceptance is test/live/gh264-docx-define-style-product.live.cjs (not run by the author).
 *
 * READ-BACK: GH #224b gave read_styles a `names` argument (the name-targeted doc.GetStyle read), so a defined custom style IS readable through the product -- see the read-back test below (on the
 * fake, which models the GetAllStyles blindness: the style is NOT in the plain enumeration and IS returned by name). The write is also verified by the adapter's own name-targeted read (L1).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const OUTSIDE = 'GH264 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
const NAME = 'GH264 Report';
const WANTED = { text: { bold: true, size: 16, color: '2255AA' }, paragraph: { align: 'center', indentLeft: 18, spacingAfter: 12 } }; // the values the adapter path measured LIVE
const RAW = { text: { bold: true, size: 32, color: '2255AA' }, paragraph: { align: 'center', indentLeft: 360, spacingAfter: 240 } }; // half-points / twips, as stored
const ds = (name, fields = WANTED, extra = {}) => ({ intent: 'define_style', name, ...fields, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product(styles = {}) {
  const store = createStore(doc(), { doc: { styles } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readStyles = async names => { const r = await read.handler({ fileId: FILE, operation: 'read_styles', ...(names ? { names } : {}) }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.styles; };
  return { store, tool, readStyles, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: define_style is in both allowlists and the published schema; the measured values, each group alone, a name only with one group, an optional step id and every align are accepted', () => {
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('define_style'));
  assert.ok(T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes('define_style'));
  assert.ok(D.inputSchema().properties.task.properties.steps.items.oneOf.some(s => s.properties.intent.const === 'define_style'));
  for (const fields of [WANTED, { text: WANTED.text }, { paragraph: WANTED.paragraph }, { text: { color: 'ff0000' } }, { text: { bold: false } }, { text: { size: 0.5 } }, { text: { size: 10.5 } },
    { paragraph: { indentLeft: 0, spacingAfter: 0 } }, { paragraph: { indentLeft: 0.25, spacingAfter: 18.75 } }, ...['left', 'right', 'center', 'justify'].map(align => ({ paragraph: { align } }))]) {
    const r = verdict([ds(NAME, fields)]);
    assert.equal(r.ok, true, `${JSON.stringify(fields)} -> ${r.code} ${r.reason}`);
  }
  assert.equal(verdict([{ ...ds(NAME), id: 'd1' }]).ok, true);
});

test('SCHEMA: the name -- missing, empty, blank, over 128 characters, a control character, a non-string -- is refused pre-dispatch', () => {
  for (const name of [undefined, '', '   ', 'x'.repeat(129), 'bad\nname', 'bad\tname', 'bad\u0000name', 123, null, { x: 1 }]) {
    const step = { intent: 'define_style', ...(name === undefined ? {} : { name }), ...WANTED };
    const v = verdict([step]);
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(name)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
  assert.equal(verdict([ds('x'.repeat(128))]).ok, true, 'exactly 128');
});

test('SCHEMA: only the measured fields -- highlight / caps / spaceBefore (the format_text name) / lineSpacing / widowControl / tabs / borders / a style-level theme colour are UNKNOWN_FIELD, a malformed basedOn REFUSED_SHAPE (GH #292 released italic / underline / strike / fontFamily / spacingBefore / indentRight / indentFirstLine / keep*); an empty group or no group is REFUSED_SHAPE', () => {
  for (const [group, key] of [['text', 'highlight'], ['text', 'language'], ['text', 'vertAlign'], ['text', 'caps'],
    ['paragraph', 'spaceBefore'], ['paragraph', 'lineSpacing'], ['paragraph', 'shading'], ['paragraph', 'tabs'], ['paragraph', 'borderTop'], ['paragraph', 'widowControl']]) {
    const v = verdict([ds(NAME, { [group]: { bold: true, ...(group === 'paragraph' ? { align: 'left' } : {}), [key]: group === 'text' ? 'x' : 1 } })]);
    assert.equal(v.code, 'UNKNOWN_FIELD', `${group}.${key} -> ${v.code}`);
    assert.equal(v.dispatched, false);
  }
  for (const extra of [{ basedOn: null }, { basedOn: 7 }, { basedOn: '' }, { basedOn: ' ' }, { basedOn: NAME }, { type: 'table' }, { type: 'character' } /* WANTED has a paragraph group: not on a character style (GH #292 S5) */]) { // GH #292 S3: basedOn is a released field -- a wrong value is a shape refusal
    const v = verdict([ds(NAME, WANTED, extra)]);
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
  for (const extra of [{ styleId: '5_1' }, { link: 'x' }]) {
    const v = verdict([ds(NAME, WANTED, extra)]);
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
  }
  for (const fields of [{}, { text: {} }, { paragraph: {} }, { text: {}, paragraph: {} }]) {
    const v = verdict([ds(NAME, fields)]);
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(fields)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: value shapes -- color must be a BARE 6-hex or exactly "auto" (GH #292 S8), size on the 0.5-point grid within 0.5..1638, indentLeft / spacingAfter on the 0.25-point grid within 0..1000, align one of four, bold a boolean', () => {
  const cases = [
    [{ text: { color: '#2255AA' } }], [{ text: { color: 'AUTO' } }], [{ text: { color: '2255A' } }], [{ text: { color: 'GGGGGG' } }], [{ text: { color: 0x2255aa } }],
    [{ text: { size: 0 } }], [{ text: { size: -1 } }], [{ text: { size: 1639 } }], [{ text: { size: '16' } }], [{ text: { size: 10.3 } }], [{ text: { size: 0.25 } }], [{ text: { size: 16.1 } }],
    [{ text: { bold: 'true' } }], [{ text: { bold: 1 } }],
    [{ paragraph: { align: 'both' } }], [{ paragraph: { align: 'Center' } }], [{ paragraph: { align: '' } }],
    [{ paragraph: { indentLeft: -1 } }], [{ paragraph: { indentLeft: 1001 } }], [{ paragraph: { indentLeft: 18.05 } }], [{ paragraph: { spacingAfter: 0.1 } }], [{ paragraph: { spacingAfter: -0.25 } }], [{ paragraph: { spacingAfter: '12' } }],
  ];
  for (const [fields] of cases.map(c => [c[0]])) {
    const v = verdict([ds(NAME, fields)]);
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(fields)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SOLO: define_style must be the only step of its task (the proof is the final style list and body; a name defined twice or beside another step is refused)', () => {
  const others = [{ intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'X' }, ds('GH264 Other'), ds(NAME)];
  for (const other of others) {
    for (const steps of [[ds(NAME), other], [other, ds(NAME)]]) {
      const v = verdict(steps);
      assert.equal(v.code, 'REFUSED_SHAPE', `${other.intent}: ${v.code} ${v.reason}`);
      assert.match(v.reason, /must be the only step/);
      assert.equal(v.dispatched, false);
    }
  }
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('WRITE: the style is created with every requested field in the editor\'s raw units (16pt -> 32 half-points, 18pt / 12pt -> 360 / 240 twips); L1 / L3 / whole PASS; every other style and the body are untouched', async () => {
  const p = product({ Other: { text: { bold: true }, paragraph: {} } });
  const before = JSON.stringify([p.store.durable.map(x => x.text), p.store.durableDoc.styles.Other]);
  const r = await p.tool.handler(req([ds(NAME)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.operations[0].l1, 'PASS');
  assert.equal(r.operations[0].l3, 'PASS');
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes() >= 1, true);
  assert.deepEqual(p.store.durableDoc.styles[NAME], RAW);
  assert.equal(JSON.stringify([p.store.durable.map(x => x.text), p.store.durableDoc.styles.Other]), before, 'the body and the other style are untouched');
});

test('READ-BACK (GH #224b): the defined custom style is NOT in the plain read_styles enumeration (GetAllStyles blindness, modelled by the fake) and IS returned by NAME with the six fields in points -- the product read-back the release was held for', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ds(NAME)]))).ok, true);
  assert.equal((await p.readStyles()).some(s => s.name === NAME), false, 'the enumeration is blind to a CreateStyle-d style');
  const byName = await p.readStyles([NAME]);
  assert.equal(byName.length, 1);
  assert.deepEqual(byName[0], { name: NAME, identity: { name: NAME, basedOn: null, next: null, link: null, styleType: 'paragraphStyle', basedOnName: null, nextName: null, linkName: null }, text: { bold: true, size: 16, color: '2255AA', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' }, paragraph: { align: 'center', indentLeft: 18, spacingAfter: 12, spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' } });
  assert.deepEqual(await p.readStyles(['GH264 Missing']), [{ name: 'GH264 Missing', gap: 'not-found' }]);
});

test('WRITE: a text-only and a paragraph-only definition leave the other group empty (inherited), not defaulted', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ds('GH264 TextOnly', { text: { color: 'FF0000' } })]))).ok, true);
  assert.deepEqual(p.store.durableDoc.styles['GH264 TextOnly'].text, { color: 'FF0000' });
  assert.equal(Object.keys(p.store.durableDoc.styles['GH264 TextOnly'].paragraph || {}).length, 0);
  assert.equal((await p.tool.handler(req([ds('GH264 ParaOnly', { paragraph: { align: 'justify' } })]))).ok, true);
  assert.deepEqual(p.store.durableDoc.styles['GH264 ParaOnly'].paragraph, { align: 'justify' });
});

test('RETRY IS IDEMPOTENT (GH #264): the SAME task again is SATISFIED / noOp with ZERO new writes -- the name exists AND every requested field reads back as requested, so it is our own retry, not a stranger\'s name', async () => {
  const p = product();
  const task = req([ds(NAME)]);
  assert.equal((await p.tool.handler(task)).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves, 'no write on the retry');
  assert.deepEqual(p.store.durableDoc.styles[NAME], RAW, 'the style is untouched');
});

test('A TAKEN NAME with DIFFERENT fields fails closed with a NAMED reason listing the differing fields, never writes and never overwrites; an identical pre-existing style is SATISFIED', async () => {
  const existing = { text: { bold: true, size: 32, color: '2255AA' }, paragraph: { align: 'center', indentLeft: 360, spacingAfter: 240 } };
  const p = product({ [NAME]: existing });
  const same = await p.tool.handler(req([ds(NAME)]));
  assert.equal(same.ok, true, JSON.stringify(same));
  assert.deepEqual(same.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), 0);
  const other = await p.tool.handler(req([ds(NAME, { text: { size: 15, color: 'FF0000' }, paragraph: { align: 'left' } })]));
  assert.equal(other.ok, false);
  assert.equal(other.code, 'OPERATION_FAILED');
  assert.match(other.reason, /define_style-name-already-exists-with-different-fields:GH264 Report:text\.size,text\.color,paragraph\.align/);
  assert.deepEqual(other.operations.map(o => o.status), ['FAILED']);
  assert.equal(p.writes(), 0);
  assert.deepEqual(p.store.durableDoc.styles[NAME], existing, 'the existing definition is untouched');
});

test('PRE-DISPATCH refusals never open a session: basedOn, an unmeasured (highlight) field, a blank name and an off-grid size refuse before any driver opens', async () => {
  for (const step of [ds(NAME, WANTED, { basedOn: NAME }), ds(NAME, { text: { highlight: 'yellow' } }), ds('   '), ds(NAME, { text: { size: 10.3 } })]) {
    const p = product();
    const r = await p.tool.handler(req([step]));
    assert.ok(['UNKNOWN_FIELD', 'REFUSED_SHAPE'].includes(r.code), `${JSON.stringify(step)} -> ${r.code}`);
    assert.equal(r.dispatched, false);
    assert.equal(p.opens(), 0);
  }
});
