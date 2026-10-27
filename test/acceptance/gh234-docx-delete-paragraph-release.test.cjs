'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #234 (DOCX P1 WRITE): delete_paragraph served by docx_execute_task (the B commit). Design and the measured reason for every schema
 * decision: docs/GH234-PARAGRAPH-DELETE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor /
 * lifecycle on test/fakes/fake-docx-driver.cjs; the LIVE acceptance is test/live/docx-delete-paragraph-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const PLAIN = { ...RICH, bold: false, italic: false };
const ALPHA = 'GH234 ALPHA';
const BETA = 'GH234 BETA';
const GAMMA = 'GH234 GAMMA';
const DELTA = 'GH234 DELTA';
const TWIN = 'GH234 TWIN';
const doc = () => [
  { text: 'Heading', level: 1, props: PLAIN },
  { text: ALPHA, props: RICH },
  { text: BETA, props: PLAIN },
  { text: GAMMA, props: RICH },
  { text: DELTA, props: PLAIN },
  { text: TWIN, props: PLAIN },
  { text: 'Closing paragraph.', props: PLAIN },
  { text: TWIN, props: PLAIN },
];
const del = (text, before, after, extra = {}) => ({ intent: 'delete_paragraph', selector: { text, before, after }, ...extra });
const BETA_DEL = () => del(BETA, ALPHA, GAMMA);
const req = steps => ({ fileId: FILE, task: { steps } });
const texts = store => store.durable.map(x => x.text);

function product(script = {}, paragraphs = doc()) {
  const store = createStore(paragraphs);
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  return { store, tool, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------

test('SCHEMA: the exact-text target with BOTH neighbours is accepted, with the optional kind and precondition qualifiers', () => {
  for (const sel of [{ text: BETA, before: ALPHA, after: GAMMA }, { text: BETA, before: ALPHA, after: GAMMA, kind: 'paragraph' },
    { text: BETA, before: ALPHA, after: GAMMA, precondition: { text: BETA } }, { text: BETA, before: 'x'.repeat(80), after: 'y' }]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'delete_paragraph', selector: sel, id: 'd1' }]));
    assert.equal(v.ok, true, `${JSON.stringify(sel)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: every selector key outside {text, before, after, kind, precondition} is UNKNOWN_FIELD, 0 dispatch; a delete step has no `text` of its own', () => {
  const extras = { heading: 'Heading', level: 1, fingerprint: 'abc', headingPath: [{ level: 1, text: 'Heading' }], contains: true,
    sectionHeading: 'Heading', paragraphOrdinal: 1, ref: 's0', index: 1, position: 1, offset: 1, paragraphIndex: 1, bookmark: 'b', tableCell: { row: 1 } };
  for (const [k, val] of Object.entries(extras)) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'delete_paragraph', selector: { text: BETA, before: ALPHA, after: GAMMA, [k]: val } }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', `${k}: ${v.code}`);
    assert.equal(v.dispatched, false, k);
  }
  for (const [k, val] of Object.entries({ text: 'x', dependsOn: ['s0'], expectedPreservation: { bold: true }, format: { bold: true }, destinationSelector: { text: BETA } })) {
    const v = D.validateDocxWriteRequest(req([{ ...BETA_DEL(), [k]: val }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', `step ${k}: ${v.code}`);
  }
  for (const sel of [{ heading: 'Heading', level: 1 }, { fingerprint: 'abc' }]) assert.equal(D.validateDocxWriteRequest(req([{ intent: 'delete_paragraph', selector: sel }])).code, 'UNKNOWN_FIELD');
});

test('SCHEMA: text, before and after are ALL required; empty (= document edge), over-80 and non-string neighbours are REFUSED_SHAPE', () => {
  const bad = [
    { before: ALPHA, after: GAMMA }, { text: BETA, after: GAMMA }, { text: BETA, before: ALPHA }, { text: BETA },
    { text: BETA, before: '', after: GAMMA }, { text: BETA, before: ALPHA, after: '' }, { text: BETA, before: '', after: '' },
    { text: BETA, before: 'x'.repeat(81), after: GAMMA }, { text: BETA, before: ALPHA, after: 'y'.repeat(81) },
    { text: BETA, before: 5, after: GAMMA }, { text: BETA, before: ALPHA, after: GAMMA, kind: 'table' }, { text: '', before: ALPHA, after: GAMMA },
  ];
  for (const sel of bad) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'delete_paragraph', selector: sel }]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(sel)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
  assert.equal(D.validateDocxWriteRequest(req([{ intent: 'delete_paragraph' }])).code, 'REFUSED_SHAPE');
});

test('SCHEMA: a delete task is that ONE step -- next to any other step (or a second delete) it is REFUSED_SHAPE', () => {
  const ins = { intent: 'insert_paragraph_after', selector: { text: ALPHA }, text: 'X' };
  for (const steps of [[BETA_DEL(), ins], [ins, BETA_DEL()], [BETA_DEL(), del(DELTA, GAMMA, TWIN)], [BETA_DEL(), { intent: 'rewrite_paragraph', selector: { text: ALPHA }, text: 'y' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
  assert.equal(D.validateDocxWriteRequest(req([ins, { ...ins, intent: 'insert_paragraph_before' }])).ok, true, 'the rule is about delete only');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid delete', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([BETA_DEL(), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the delete step with text / before / after required', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  const d = items.oneOf.find(s => s.properties.intent.const === 'delete_paragraph');
  assert.deepEqual(d.required, ['intent', 'selector']);
  assert.deepEqual(d.properties.selector.required, ['text', 'before', 'after']);
  assert.equal(d.properties.selector.properties.before.maxLength, 80);
  assert.match(D.createDocxWriteTool({}).description, /delete_paragraph/);
});

// ---- the product handler, end to end on the fake transport ---------------------------------------------------------------

test('SURVIVING ANCHOR: the target is gone and its two neighbours now stand next to each other; the receipt is APPLIED with L1 / whole-task PASS and NOTHING else changed', async () => {
  const p = product();
  const before = p.store.durable.map(x => ({ text: x.text, props: JSON.stringify(x.props) }));
  const r = await p.tool.handler(req([BETA_DEL()]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.tool, 'docx_execute_task');
  assert.deepEqual(r.operations.map(o => [o.status, o.l1]), [['APPLIED', 'PASS']]);
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(r.noOp, false);
  assert.equal(p.writes() >= 1, true);
  const after = p.store.durable.map(x => ({ text: x.text, props: JSON.stringify(x.props) }));
  assert.deepEqual(after, before.filter(x => x.text !== BETA), 'the sequence is the old one minus the target; every other paragraph keeps its text AND run properties');
  const t = texts(p.store);
  assert.equal(t.indexOf(GAMMA), t.indexOf(ALPHA) + 1, 'the former neighbours are adjacent');
});

test('NEIGHBOUR TEXT is a PREFIX of at most 80 characters: a long neighbour works with its first characters, and its full text finds nothing (0 writes)', async () => {
  const LONG = `${'L'.repeat(60)} middle ${'M'.repeat(60)}`;
  const mk = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: `${LONG} one`, props: PLAIN }, { text: BETA, props: PLAIN }, { text: `${LONG} two`, props: PLAIN }, { text: DELTA, props: PLAIN }];
  const p = product({}, mk());
  const r = await p.tool.handler(req([del(BETA, `${LONG} one`.slice(0, 80), `${LONG} two`.slice(0, 80))]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(texts(p.store).includes(BETA), false);
  const q = product({}, mk());
  const full = await q.tool.handler(req([del(BETA, `${LONG} one`, `${LONG} two`)]));
  assert.equal(full.dispatched, false, 'over 80 characters is refused pre-dispatch');
  assert.equal(q.writes(), 0);
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes (the surviving anchor, not the absence alone, is the proof)', async () => {
  const p = product();
  const task = req([BETA_DEL()]);
  assert.equal((await p.tool.handler(task)).ok, true);
  const saves = p.writes(); const seq = texts(p.store);
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
  assert.deepEqual(texts(p.store), seq);
});

test('ABSENCE ALONE IS NOT ENOUGH: a target that is not there, with neighbours that are NOT adjacent, fails closed (the typo case with a wrong pair)', async () => {
  const p = product();
  const seq = texts(p.store);
  const r = await p.tool.handler(req([del('GH234 TYPO', ALPHA, GAMMA)]));   // ALPHA and GAMMA are not adjacent (BETA is between)
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(r.noOp, false);
  assert.deepEqual(texts(p.store), seq);
  assert.equal(p.writes(), 0);
  const r2 = await p.tool.handler(req([del('GH234 TYPO', 'GH234 NOWHERE', GAMMA)]));
  assert.equal(r2.ok, false);
  assert.equal(p.writes(), 0);
});

test('KNOWN LIMIT (named in the decision doc): a target that was never there, with neighbours that ARE adjacent, reads SATISFIED / noOp / 0 writes -- after a deletion the text itself leaves no evidence, only the adjacency does', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([BETA_DEL()]))).ok, true);
  const saves = p.writes();
  const r = await p.tool.handler(req([del('GH234 NEVER EXISTED', ALPHA, GAMMA)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves, 'nothing is written -- the claim is only "this pair is adjacent and no such paragraph exists"');
});

test('WRONG CONTEXT never deletes and never reads as done: the target exists but the named neighbours do not match -> ok:false, ZERO writes, the document unchanged', async () => {
  for (const [b, a] of [['GH234 WRONG', GAMMA], [ALPHA, 'GH234 WRONG'], [GAMMA, DELTA]]) {   // the last pair IS adjacent elsewhere: the step may read SATISFIED, the task must still fail
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([del(BETA, b, a)]));
    assert.equal(r.ok, false, `${b} | ${a}: ${JSON.stringify(r)}`);
    assert.equal(r.noOp, false);
    assert.deepEqual(texts(p.store), seq);
    assert.equal(p.writes(), 0);
  }
});

test('AMBIGUOUS TARGET fails closed: two paragraphs with the target text -> ZERO writes, the document unchanged -- the neighbours do NOT disambiguate a duplicate (the editor command finds its target by exact text)', async () => {
  for (const sel of [del(TWIN, DELTA, 'Closing'), del(TWIN, 'Closing', 'GH234')]) {
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([sel]));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
    assert.deepEqual(texts(p.store), seq);
    assert.equal(p.writes(), 0);
  }
});

test('MISSING TARGET / kind mismatch / PRECONDITION mismatch fail closed: refused, ZERO writes', async () => {
  for (const sel of [{ text: 'GH234 NOT THERE', before: ALPHA, after: GAMMA }, { text: BETA, before: ALPHA, after: GAMMA, kind: 'heading' }, { text: BETA, before: ALPHA, after: GAMMA, precondition: { text: 'something else' } }]) {
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([{ intent: 'delete_paragraph', selector: sel }]));
    assert.equal(r.ok, false, JSON.stringify(sel) + JSON.stringify(r));
    assert.deepEqual(texts(p.store), seq);
    assert.equal(p.writes(), 0);
  }
});

test('EDITOR FAULT never reads as APPLIED: the command silently does nothing -> not ok, no APPLIED step, nothing saved', async () => {
  const p = product({ fault: 'ignore' });
  const seq = texts(p.store);
  const r = await p.tool.handler(req([BETA_DEL()]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
  assert.deepEqual(texts(p.store), seq);
  assert.equal(p.writes(), 0);
});

test('PRE-DISPATCH refusals never open a session: a refused delete task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[del(BETA, '', GAMMA)], [{ intent: 'delete_paragraph', selector: { text: BETA } }], [BETA_DEL(), BETA_DEL()], [{ intent: 'delete_paragraph', selector: { heading: 'Heading', level: 1 } }]]) {
    const r = await p.tool.handler(req(steps));
    assert.equal(r.dispatched, false);
  }
  assert.equal(p.opens(), 0);
});

// ---- the release gate ------------------------------------------------------------------------------------------------------

test('GATE: the DOCX allowlist is exactly the accepted set (both halves agree); a one-sided change withholds only docx_execute_task', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']);
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  const pre = ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'define_style'];
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#234 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  assert.ok(T.capabilityInventory().some(r => r.TOOL === 'docx:delete_paragraph' && r.EXPOSE_IN_PRODUCT_MCP === 'YES'));
});

test('GH #293 N5 finding: a delete of a LIST ITEM (not a released target kind) is never SATISFIED -- the surviving-context proof matched the whole list block by containment and reported a deletion that did not happen', async () => {
  const L = { type: 'numbered', level: 0, numId: 'n1' };
  const paragraphs = [{ text: 'GH293 intro', props: PLAIN }, { text: 'GH293 A1', props: PLAIN, list: L }, { text: 'GH293 A2', props: PLAIN, list: L }, { text: 'GH293 between', props: PLAIN }];
  const p = product({}, paragraphs);
  const before = JSON.stringify(texts(p.store));
  const r = await p.tool.handler(req([del('GH293 A2', 'GH293 A1', 'GH293 between')]));
  assert.notEqual(r.ok, true, JSON.stringify(r).slice(0, 400));
  assert.notEqual((r.operations || [])[0] && (r.operations || [])[0].status, 'SATISFIED');
  assert.match(JSON.stringify(r), /delete-target-present-as-an-unsupported-unit:list_item/);
  assert.equal(JSON.stringify(texts(p.store)), before, 'nothing deleted');
  assert.equal(p.writes(), 0);
});
