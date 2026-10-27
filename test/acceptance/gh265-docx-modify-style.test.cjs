'use strict';

const test = require('../helpers/trim-skip.cjs').testFor('modify_style'); // release 3 trim: skipped while modify_style is not in the allowlist
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #265 (DOCX P1 WRITE; the release itself is an OWNER DECISION, alongside #228..#264): modify_style served by docx_execute_task -- the SAME step shape and the SAME six fields as define_style
 * (GH #264), on an EXISTING style. Design and the measured reason for every rule: docs/GH265-DOCX-MODIFY-STYLE-DECISION.md. Offline only: the product handler runs the REAL docx adapter /
 * executor / lifecycle on test/fakes/fake-docx-driver.cjs, which stores the style in the editor's RAW units (half-points / twips). The LIVE acceptance is
 * test/live/gh265-docx-modify-style-product.live.cjs (not run by the author).
 *
 * The product read_styles check below runs on the FAKE, which lists every style; on the real DS 9.3.5 read_styles lists only the built-in styles (GetAllStyles is blind to CreateStyle,
 * GH #224), so the read-back proof exists for a modified BUILT-IN style (Heading 1, measured) and not for a custom one. The adapter's L3 proves every OTHER style and the body unchanged; it does
 * NOT compare the modified style's own unnamed fields -- the tests below pin that they survive on the fake, the LIVE harness measures it on the real file.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const OUTSIDE = 'GH265 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
const NAME = 'GH265 Report';
const OTHER = 'GH265 Other';
const RAW = () => ({ text: { bold: true, size: 32, color: '2255AA' }, paragraph: { align: 'center', indentLeft: 360, spacingAfter: 240 } }); // 16pt / 18pt / 12pt in half-points / twips
const OTHER_RAW = () => ({ text: { bold: false, size: 20 }, paragraph: { indentLeft: 40 } });
const ms = (name, fields, extra = {}) => ({ intent: 'modify_style', name, ...fields, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product(styles = { [NAME]: RAW(), [OTHER]: OTHER_RAW() }) {
  const store = createStore(doc(), { doc: { styles } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readStyles = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_styles' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.styles; };
  return { store, tool, readStyles, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.styles]);

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: modify_style is in both allowlists and the published schema; the same six fields as define_style are accepted, each group alone, with an optional step id', () => {
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('modify_style'));
  assert.ok(T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes('modify_style'));
  assert.ok(D.inputSchema().properties.task.properties.steps.items.oneOf.some(s => s.properties.intent.const === 'modify_style'));
  for (const fields of [{ text: { bold: true, size: 20, color: 'FF0000' }, paragraph: { align: 'right', indentLeft: 0, spacingAfter: 6.25 } }, { text: { size: 26 } }, { text: { bold: false } }, { paragraph: { align: 'justify' } }, { paragraph: { spacingAfter: 0 } }]) {
    const r = verdict([ms(NAME, fields)]);
    assert.equal(r.ok, true, `${JSON.stringify(fields)} -> ${r.code} ${r.reason}`);
  }
  assert.equal(verdict([{ ...ms(NAME, { text: { size: 26 } }), id: 'm1' }]).ok, true);
});

test('SCHEMA: the same refusals as define_style -- a blank / over-long / control-character name, basedOn, an unmeasured field, an empty group or none, a "#" colour, an off-grid size are refused pre-dispatch', () => {
  const refused = [
    [{ name: '   ', text: { size: 20 } }, 'REFUSED_SHAPE'], [{ name: 'x'.repeat(129), text: { size: 20 } }, 'REFUSED_SHAPE'], [{ name: 'bad\nname', text: { size: 20 } }, 'REFUSED_SHAPE'], [{ text: { size: 20 } }, 'REFUSED_SHAPE'],
    [{ name: NAME, basedOn: 'Normal', text: { size: 20 } }, 'UNKNOWN_FIELD'], [{ name: NAME, text: { highlight: 'yellow' } }, 'UNKNOWN_FIELD'], [{ name: NAME, paragraph: { spaceBefore: 6 } }, 'UNKNOWN_FIELD'],
    [{ name: NAME }, 'REFUSED_SHAPE'], [{ name: NAME, text: {} }, 'REFUSED_SHAPE'], [{ name: NAME, paragraph: {} }, 'REFUSED_SHAPE'],
    [{ name: NAME, text: { color: '#FF0000' } }, 'REFUSED_SHAPE'], [{ name: NAME, text: { color: 'auto ' } }, 'REFUSED_SHAPE'], [{ name: NAME, text: { size: 10.3 } }, 'REFUSED_SHAPE'], [{ name: NAME, paragraph: { indentLeft: 18.05 } }, 'REFUSED_SHAPE'],
  ];
  for (const [step, code] of refused) {
    const v = verdict([{ intent: 'modify_style', ...step }]);
    assert.equal(v.code, code, `${JSON.stringify(step)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SOLO: modify_style must be the only step of its task (beside another modify_style, a define_style or any other intent it is refused)', () => {
  const others = [{ intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'X' }, ms(OTHER, { text: { size: 20 } }), { intent: 'define_style', name: 'GH265 New', text: { size: 20 } }];
  for (const other of others) {
    for (const steps of [[ms(NAME, { text: { size: 26 } }), other], [other, ms(NAME, { text: { size: 26 } })]]) {
      const v = verdict(steps);
      assert.equal(v.code, 'REFUSED_SHAPE', `${other.intent}: ${v.code} ${v.reason}`);
      assert.match(v.reason, /must be the only step/);
      assert.equal(v.dispatched, false);
    }
  }
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('WRITE: ONLY the named field changes (size 16pt -> 20pt = 40 half-points); the style\'s own UNNAMED fields (bold, colour, the whole paragraph group), every other style and the body are untouched; L1 / L3 / whole PASS', async () => {
  const p = product();
  const beforeOther = JSON.stringify(p.store.durableDoc.styles[OTHER]);
  const beforeText = JSON.stringify(p.store.durable.map(x => x.text));
  const r = await p.tool.handler(req([ms(NAME, { text: { size: 20 } })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.operations[0].l1, 'PASS');
  assert.equal(r.operations[0].l3, 'PASS');
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes() >= 1, true);
  assert.deepEqual(p.store.durableDoc.styles[NAME], { ...RAW(), text: { bold: true, size: 40, color: '2255AA' } });
  assert.equal(JSON.stringify(p.store.durableDoc.styles[OTHER]), beforeOther, 'the other style is untouched');
  assert.equal(JSON.stringify(p.store.durable.map(x => x.text)), beforeText, 'the body text is untouched');
});

test('WRITE: every field at once, an explicit bold:false, and a paragraph-only change each land exactly in raw units and leave the rest', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ms(NAME, { text: { bold: false, size: 10.5, color: 'FF0000' }, paragraph: { align: 'right', indentLeft: 0.25, spacingAfter: 0 } })]))).ok, true);
  assert.deepEqual(p.store.durableDoc.styles[NAME], { text: { bold: false, size: 21, color: 'FF0000' }, paragraph: { align: 'right', indentLeft: 5, spacingAfter: 0 } });
  const q = product();
  assert.equal((await q.tool.handler(req([ms(NAME, { paragraph: { spacingAfter: 6 } })]))).ok, true);
  assert.deepEqual(q.store.durableDoc.styles[NAME], { ...RAW(), paragraph: { align: 'center', indentLeft: 360, spacingAfter: 120 } });
});

test('RETRY IS IDEMPOTENT (measured MODIFY_RETRY, unlike define_style): the SAME task again is SATISFIED / noOp with ZERO new writes', async () => {
  const p = product();
  const task = req([ms(NAME, { text: { size: 20, color: 'FF0000' } })]);
  assert.equal((await p.tool.handler(task)).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves, 'no write on the retry');
});

test('A MISSING NAME fails CLOSED (modify_style-name-does-not-exist) with no write and NO style created -- modify never turns into define', async () => {
  const p = product();
  const before = snapshot(p.store);
  const r = await p.tool.handler(req([ms('GH265 Ghost', { text: { size: 20 } })]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'OPERATION_FAILED');
  assert.match(r.reason, /modify_style-name-does-not-exist/);
  assert.deepEqual(r.operations.map(o => o.status), ['FAILED']);
  assert.equal(p.writes(), 0);
  assert.equal(snapshot(p.store), before, 'nothing was created or changed');
});

test('READ SYMMETRY on the fake (which lists every style; the real DS 9.3.5 lists only built-ins): read_styles shows the modified size in POINTS and the untouched colour / paragraph group; the other style is unchanged', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ms(NAME, { text: { size: 20 } })]))).ok, true);
  const styles = await p.readStyles();
  const mine = styles.find(s => s.name === NAME);
  assert.ok(mine, 'the modified style is listed on the fake');
  assert.equal(mine.text.size, 20);
  assert.equal(mine.text.color, '2255AA');
  assert.equal(mine.paragraph.align, 'center');
  const other = styles.find(s => s.name === OTHER);
  assert.equal(other.text.size, 10, 'the other style: 20 half-points = 10pt, untouched');
});

test('PRE-DISPATCH refusals never open a session: basedOn, an unmeasured (highlight) field, a blank name and an off-grid indent refuse before any driver opens', async () => {
  for (const step of [ms(NAME, { text: { size: 20 } }, { basedOn: 'Normal' }), ms(NAME, { text: { highlight: 'yellow' } }), ms('   ', { text: { size: 20 } }), ms(NAME, { paragraph: { indentLeft: 18.05 } })]) {
    const p = product();
    const r = await p.tool.handler(req([step]));
    assert.ok(['UNKNOWN_FIELD', 'REFUSED_SHAPE'].includes(r.code), `${JSON.stringify(step)} -> ${r.code}`);
    assert.equal(r.dispatched, false);
    assert.equal(p.opens(), 0);
  }
});
