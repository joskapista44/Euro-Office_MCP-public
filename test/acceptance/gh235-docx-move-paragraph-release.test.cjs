'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #235 (DOCX P1 WRITE): move_paragraph_before / move_paragraph_after served by docx_execute_task (the B commit). Design and the measured
 * reason for every schema decision: docs/GH235-PARAGRAPH-MOVE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx
 * adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs; the LIVE acceptance is test/live/docx-move-paragraph-product.live.cjs
 * (not run by the author).
 */

const FILE = '1303626';
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Arial', fontSize: 26, language: 'hu-HU' };
const PLAIN = { ...RICH, bold: false, italic: false, fontSize: 22 };
const A = 'GH235 A';
const B = 'GH235 B';
const C = 'GH235 C';
const DD = 'GH235 D';
const E = 'GH235 E';
const doc = () => [
  { text: 'Heading', level: 1, props: PLAIN },
  { text: A, props: PLAIN },
  { text: B, props: RICH },
  { text: C, props: PLAIN },
  { text: DD, props: PLAIN },
  { text: E, props: PLAIN },
];
const mv = (side, text, dest, extra = {}) => ({ intent: `move_paragraph_${side}`, selector: { text }, destinationSelector: { text: dest }, ...extra });
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
/** `seq` without `text`, then `text` directly `side` of `dest` (the one expected result of a move). */
const expected = (seq, side, text, dest) => { const rest = seq.filter(t => t !== text); const i = rest.indexOf(dest); return [...rest.slice(0, side === 'before' ? i : i + 1), text, ...rest.slice(side === 'before' ? i : i + 1)]; };

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------

test('SCHEMA: both intents accept the exact-text source and destination, with the optional kind / precondition (source) and kind (destination)', () => {
  for (const side of ['before', 'after']) {
    for (const [sel, dst] of [[{ text: B }, { text: DD }], [{ text: B, kind: 'paragraph' }, { text: DD, kind: 'paragraph' }], [{ text: B, precondition: { text: B } }, { text: 'Heading', kind: 'heading' }]]) {
      const v = D.validateDocxWriteRequest(req([{ intent: `move_paragraph_${side}`, selector: sel, destinationSelector: dst, id: 'm1' }]));
      assert.equal(v.ok, true, `${side} ${JSON.stringify([sel, dst])} -> ${v.code} ${v.reason}`);
    }
  }
});

test('SCHEMA: every key outside the two small selectors is UNKNOWN_FIELD, 0 dispatch (a heading / fingerprint selector APPLIES then reports ok:false: measured on both selectors)', () => {
  const extras = { heading: 'Heading', level: 1, fingerprint: 'abc', headingPath: [{ level: 1, text: 'Heading' }], contains: true, before: 'x', after: 'y',
    sectionHeading: 'Heading', paragraphOrdinal: 1, ref: 's0', index: 1, position: 1, offset: 1, paragraphIndex: 1, bookmark: 'b', tableCell: { row: 1 } };
  for (const [k, val] of Object.entries(extras)) {
    for (const which of ['selector', 'destinationSelector']) {
      const step = mv('after', B, DD);
      step[which] = { ...step[which], [k]: val };
      const v = D.validateDocxWriteRequest(req([step]));
      assert.equal(v.code, 'UNKNOWN_FIELD', `${which}.${k}: ${v.code}`);
      assert.equal(v.dispatched, false);
    }
  }
  assert.equal(D.validateDocxWriteRequest(req([{ ...mv('after', B, DD), destinationSelector: { text: DD, precondition: { text: DD } } }])).code, 'UNKNOWN_FIELD', 'the contract has no precondition on a destination');
  for (const [k, val] of Object.entries({ text: 'x', dependsOn: ['s0'], expectedPreservation: { bold: true }, format: { bold: true } })) {
    assert.equal(D.validateDocxWriteRequest(req([{ ...mv('after', B, DD), [k]: val }])).code, 'UNKNOWN_FIELD', `step ${k}`);
  }
  for (const [sel, dst] of [[{ heading: 'Heading', level: 1 }, { text: DD }], [{ text: B }, { heading: 'Heading', level: 1 }]]) {
    assert.equal(D.validateDocxWriteRequest(req([{ intent: 'move_paragraph_after', selector: sel, destinationSelector: dst }])).code, 'UNKNOWN_FIELD');
  }
});

test('SCHEMA: missing / empty / non-string text, a missing destination, a bad kind, and a move onto itself are REFUSED_SHAPE', () => {
  const cases = [
    { intent: 'move_paragraph_after', selector: { text: B } },
    { intent: 'move_paragraph_after', destinationSelector: { text: DD } },
    { intent: 'move_paragraph_after', selector: {}, destinationSelector: { text: DD } },
    { intent: 'move_paragraph_after', selector: { text: B }, destinationSelector: {} },
    { intent: 'move_paragraph_after', selector: { text: '' }, destinationSelector: { text: DD } },
    { intent: 'move_paragraph_after', selector: { text: B }, destinationSelector: { text: 5 } },
    { intent: 'move_paragraph_after', selector: { text: B, kind: 'table' }, destinationSelector: { text: DD } },
    { intent: 'move_paragraph_before', selector: { text: B }, destinationSelector: { text: B } },
  ];
  for (const s of cases) { const v = D.validateDocxWriteRequest(req([s])); assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
});

test('SCHEMA: a move task is that ONE step -- next to any other step (or a second move) it is REFUSED_SHAPE; a delete task still is too', () => {
  const ins = { intent: 'insert_paragraph_after', selector: { text: A }, text: 'X' };
  for (const steps of [[mv('after', B, DD), ins], [ins, mv('after', B, DD)], [mv('after', B, DD), mv('after', A, E)], [mv('after', B, DD), mv('after', A, DD)]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
  assert.equal(D.validateDocxWriteRequest(req([mv('after', B, DD), { intent: 'delete_paragraph', selector: { text: C, before: A, after: DD } }])).code, 'REFUSED_SHAPE');
  assert.equal(D.validateDocxWriteRequest(req([ins, { ...ins, intent: 'insert_paragraph_before' }])).ok, true, 'the rule is about delete / move only');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid move', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([mv('after', B, DD), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries both move steps with selector and destinationSelector required', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  for (const intent of ['move_paragraph_before', 'move_paragraph_after']) {
    const s = items.oneOf.find(x => x.properties.intent.const === intent);
    assert.deepEqual(s.required, ['intent', 'selector', 'destinationSelector']);
    assert.deepEqual(s.properties.selector.required, ['text']);
    assert.deepEqual(s.properties.destinationSelector.required, ['text']);
  }
  assert.match(D.createDocxWriteTool({}).description, /move_paragraph_before, move_paragraph_after/);
});

// ---- the product handler, end to end on the fake transport ---------------------------------------------------------------

test('DIRECTIONAL ADJACENCY + OLD POSITION CLOSED: AFTER puts the paragraph directly after the destination, BEFORE directly before; its old neighbours close up; ONLY the order changes', async () => {
  for (const [side, text, dest] of [['after', B, DD], ['before', DD, B], ['after', A, E], ['before', E, 'Heading']]) {
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([mv(side, text, dest)]));
    assert.equal(r.ok, true, `${side} ${text}: ${JSON.stringify(r)}`);
    assert.equal(r.tool, 'docx_execute_task');
    assert.deepEqual(r.operations.map(o => [o.status, o.l1]), [['APPLIED', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(r.noOp, false);
    assert.equal(p.writes() >= 1, true);
    const after = texts(p.store);
    assert.deepEqual(after, expected(seq, side, text, dest), `${side} ${text} -> ${dest}: the whole sequence`);
    assert.equal(after.length, seq.length, 'a move never adds or drops a paragraph');
    assert.equal(new Set(after).size, after.length, 'no duplicate');
  }
});

test('PRESERVATION (offline model): the moved paragraph keeps its text AND run properties, and every other paragraph is untouched', async () => {
  const p = product();
  const byText = () => Object.fromEntries(p.store.durable.map(x => [x.text, JSON.stringify(x.props)]));
  const before = byText();
  const r = await p.tool.handler(req([mv('after', B, DD)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(byText(), before, 'text -> run properties map is identical, bold + italic source included');
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes and NO duplicate (both directions)', async () => {
  for (const side of ['after', 'before']) {
    const p = product();
    const task = req([mv(side, B, DD)]);
    assert.equal((await p.tool.handler(task)).ok, true);
    const saves = p.writes(); const seq = texts(p.store);
    const retry = await p.tool.handler(task);
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
    assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
    assert.equal(p.writes(), saves, `${side}: no write on the retry`);
    assert.deepEqual(texts(p.store), seq);
  }
});

test('ALREADY THERE: moving a paragraph to the side of the destination it already stands on is SATISFIED / noOp / 0 writes', async () => {
  for (const [side, text, dest] of [['after', B, A], ['before', B, C]]) {
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([mv(side, text, dest)]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.noOp, true);
    assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
    assert.equal(p.writes(), 0);
    assert.deepEqual(texts(p.store), seq);
  }
});

test('RETRY with the OPPOSITE side applies the move back (it is a new, different request), never a duplicate', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([mv('after', B, DD)]))).ok, true);     // A C D B E
  const back = await p.tool.handler(req([mv('before', B, DD)]));                // B directly before D -> A C B D E
  assert.equal(back.ok, true, JSON.stringify(back));
  assert.deepEqual(texts(p.store), ['Heading', A, C, B, DD, E]);
});

test('FAIL-CLOSED: missing source / missing destination / duplicated source / duplicated destination / kind mismatch / source precondition mismatch -> ZERO writes, the document unchanged', async () => {
  const dupB = () => [...doc(), { text: B, props: RICH }];
  const dupD = () => [...doc(), { text: DD, props: PLAIN }];
  const cases = [
    [doc(), mv('after', 'GH235 NOWHERE', DD)], [doc(), mv('after', B, 'GH235 NOWHERE')], [dupB(), mv('after', B, DD)], [dupD(), mv('after', B, DD)],
    [doc(), { intent: 'move_paragraph_after', selector: { text: B, kind: 'heading' }, destinationSelector: { text: DD } }],
    [doc(), { intent: 'move_paragraph_after', selector: { text: B, precondition: { text: 'something else' } }, destinationSelector: { text: DD } }],
  ];
  for (const [paras, step] of cases) {
    const p = product({}, paras);
    const seq = texts(p.store);
    const r = await p.tool.handler(req([step]));
    assert.equal(r.ok, false, JSON.stringify(step) + JSON.stringify(r));
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
    assert.deepEqual(texts(p.store), seq);
    assert.equal(p.writes(), 0);
  }
});

test('EDITOR FAULT never reads as APPLIED: the command silently does nothing -> not ok, no APPLIED step, nothing saved', async () => {
  const p = product({ fault: 'ignore' });
  const seq = texts(p.store);
  const r = await p.tool.handler(req([mv('after', B, DD)]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
  assert.deepEqual(texts(p.store), seq);
  assert.equal(p.writes(), 0);
});

test('PRE-DISPATCH refusals never open a session: a refused move task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[{ intent: 'move_paragraph_after', selector: { heading: 'Heading', level: 1 }, destinationSelector: { text: DD } }], [mv('after', B, B)], [mv('after', B, DD), mv('after', A, E)]]) {
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
  const pre = ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'define_style'];
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#235 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  for (const i of ['move_paragraph_before', 'move_paragraph_after']) assert.ok(T.capabilityInventory().some(r => r.TOOL === `docx:${i}` && r.EXPOSE_IN_PRODUCT_MCP === 'YES'), i);
});
