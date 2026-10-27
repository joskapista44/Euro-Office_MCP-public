'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #233 (DOCX P1 WRITE): insert_paragraph_before / insert_paragraph_after served by docx_execute_task (the B commit). Design and the
 * measured reasons for every schema decision: docs/GH233-PARAGRAPH-INSERT-PRODUCT-DECISION.md. Offline only: the product handler runs the
 * REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs; the LIVE acceptance is
 * test/live/docx-insert-paragraph-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const PLAIN = { ...RICH, bold: false, italic: false };
const ALPHA = 'GH233 ALPHA';
const BETA = 'GH233 BETA';
const GAMMA = 'GH233 GAMMA';
const TWIN = 'GH233 TWIN';
const NEW = 'GH233 NEW PARAGRAPH';
const doc = () => [
  { text: 'Heading', level: 1, props: PLAIN },
  { text: ALPHA, props: RICH },
  { text: BETA, props: PLAIN },
  { text: GAMMA, props: PLAIN },
  { text: TWIN, props: PLAIN },
  { text: TWIN, props: PLAIN },
  { text: 'Closing paragraph.', props: PLAIN },
];
const ins = (side, anchor, text = NEW, extra = {}) => ({ intent: `insert_paragraph_${side}`, selector: { text: anchor }, text, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const texts = store => store.durable.map(x => x.text);

function product(script = {}) {
  const store = createStore(doc());
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  return { store, tool, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
/** `seq` with `text` placed directly `side` of `anchor` (the one expected index). */
const expected = (seq, side, anchor, text) => { const i = seq.indexOf(anchor); return [...seq.slice(0, side === 'before' ? i : i + 1), text, ...seq.slice(side === 'before' ? i : i + 1)]; };

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------

test('SCHEMA: both intents accept the exact-text anchor, with the optional kind and precondition qualifiers', () => {
  for (const side of ['before', 'after']) {
    for (const sel of [{ text: ALPHA }, { text: ALPHA, kind: 'paragraph' }, { text: 'Heading', kind: 'heading' }, { text: ALPHA, precondition: { text: ALPHA } }, { text: ALPHA, kind: 'paragraph', precondition: { text: ALPHA } }]) {
      const v = D.validateDocxWriteRequest(req([{ intent: `insert_paragraph_${side}`, selector: sel, text: NEW, id: 'i1' }]));
      assert.equal(v.ok, true, `${side} ${JSON.stringify(sel)} -> ${v.code} ${v.reason}`);
    }
  }
});

test('SCHEMA: every selector key outside {text, kind, precondition} is UNKNOWN_FIELD, 0 dispatch (each omission is a measured defect, see the decision doc section 2)', () => {
  const extras = { heading: 'Heading', level: 1, fingerprint: 'abc', headingPath: [{ level: 1, text: 'Heading' }], contains: true, before: 'x', after: 'y',
    sectionHeading: 'Heading', paragraphOrdinal: 1, ref: 's0', index: 1, position: 1, offset: 1, paragraphIndex: 1, bookmark: 'b', tableCell: { row: 1 } };
  for (const [k, val] of Object.entries(extras)) {
    for (const side of ['before', 'after']) {
      const v = D.validateDocxWriteRequest(req([{ intent: `insert_paragraph_${side}`, selector: { text: ALPHA, [k]: val }, text: NEW }]));
      assert.equal(v.code, 'UNKNOWN_FIELD', `${side} ${k}: ${v.code}`);
      assert.equal(v.dispatched, false, k);
    }
  }
});

test('SCHEMA: a selector without `text` is REFUSED_SHAPE (the whole-task check reads the anchor by its text); a heading / fingerprint anchor is UNKNOWN_FIELD', () => {
  for (const sel of [{}, { kind: 'heading' }, { precondition: { text: ALPHA } }]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'insert_paragraph_after', selector: sel, text: NEW }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(sel));
    assert.equal(v.dispatched, false);
  }
  for (const sel of [{ heading: 'Heading', level: 1 }, { fingerprint: 'abc' }]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'insert_paragraph_after', selector: sel, text: NEW }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(sel));
  }
});

test('SCHEMA: malformed values are refused pre-dispatch (empty / blank / multi-line / non-string / missing text, bad kind, step keys outside the schema)', () => {
  const cases = [
    [ins('after', ALPHA, ''), 'REFUSED_SHAPE'], [ins('after', ALPHA, '   '), 'REFUSED_SHAPE'], [ins('after', ALPHA, '\t'), 'REFUSED_SHAPE'],
    [ins('after', ALPHA, 'line one\nline two'), 'REFUSED_SHAPE'], [ins('before', ALPHA, 'a\rb'), 'REFUSED_SHAPE'],
    [ins('after', ALPHA, 5), 'REFUSED_SHAPE'], [{ intent: 'insert_paragraph_after', selector: { text: ALPHA } }, 'REFUSED_SHAPE'],
    [{ intent: 'insert_paragraph_after', selector: { text: ALPHA, kind: 'table' }, text: NEW }, 'REFUSED_SHAPE'],
    [ins('after', ALPHA, NEW, { dependsOn: ['s0'] }), 'UNKNOWN_FIELD'], [ins('after', ALPHA, NEW, { expectedPreservation: { bold: true } }), 'UNKNOWN_FIELD'],
    [ins('after', ALPHA, NEW, { format: { bold: true } }), 'UNKNOWN_FIELD'], [ins('after', ALPHA, NEW, { destinationSelector: { text: BETA } }), 'UNKNOWN_FIELD'],
  ];
  for (const [s, code] of cases) { const v = D.validateDocxWriteRequest(req([s])); assert.equal(v.code, code, `${JSON.stringify(s)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
});

test('SCHEMA: two inserts on the SAME anchor and the SAME side in one task are REFUSED_SHAPE; other sides / anchors / chained anchors are fine', () => {
  assert.equal(D.validateDocxWriteRequest(req([ins('after', ALPHA, 'A1'), ins('after', ALPHA, 'A2')])).code, 'REFUSED_SHAPE');
  assert.equal(D.validateDocxWriteRequest(req([ins('before', ALPHA, 'A1'), ins('before', ALPHA, 'A2')])).code, 'REFUSED_SHAPE');
  for (const steps of [[ins('after', ALPHA, 'A1'), ins('before', ALPHA, 'A0')], [ins('after', ALPHA, 'A1'), ins('after', BETA, 'B1')], [ins('after', ALPHA, 'A1'), ins('after', 'A1', 'A2')]]) {
    assert.equal(D.validateDocxWriteRequest(req(steps)).ok, true, JSON.stringify(steps.map(s => [s.intent, s.selector.text])));
  }
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid insert (whole task refused, step index named)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([ins('after', ALPHA), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema offers rewrite_paragraph, the two insert steps and (GH #234) delete_paragraph', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  assert.deepEqual(items.oneOf.map(s => s.properties.intent.const), ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']);
  for (const s of items.oneOf.filter(x => /^insert_paragraph_/.test(x.properties.intent.const))) assert.deepEqual(s.required, ['intent', 'selector', 'text']);
  assert.deepEqual(items.oneOf[1].properties.selector.required, ['text']);
  assert.match(D.createDocxWriteTool({}).description, /insert_paragraph_before, insert_paragraph_after/);
});

// ---- the product handler, end to end on the fake transport ---------------------------------------------------------------

test('DIRECTIONAL READBACK: AFTER puts the new paragraph directly after its anchor, BEFORE directly before; the receipt is APPLIED with L1 / whole-task PASS and ONLY that paragraph is new', async () => {
  for (const side of ['after', 'before']) {
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([ins(side, BETA)]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.tool, 'docx_execute_task');
    assert.deepEqual(r.operations.map(o => [o.status, o.l1]), [['APPLIED', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(r.noOp, false);
    assert.equal(p.writes() >= 1, true);
    assert.deepEqual(texts(p.store), expected(seq, side, BETA, NEW), `${side}: the whole sequence is the old one plus the one new paragraph at the one index`);
  }
});

test('PRESERVATION: the anchor and every other paragraph keep their text AND run properties (placement NONE: nothing else is checked by the adapter, this test is the neighbour proof offline)', async () => {
  const p = product();
  const before = p.store.durable.map(x => ({ text: x.text, props: JSON.stringify(x.props) }));
  const r = await p.tool.handler(req([ins('after', ALPHA)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = p.store.durable.map(x => ({ text: x.text, props: JSON.stringify(x.props) }));
  assert.deepEqual(after.filter(x => x.text !== NEW), before, 'everything that was there is untouched, bold + italic anchor included');
  assert.equal(after.length, before.length + 1);
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes and NO duplicate paragraph (both directions)', async () => {
  for (const side of ['after', 'before']) {
    const p = product();
    const task = req([ins(side, BETA)]);
    assert.equal((await p.tool.handler(task)).ok, true);
    const savesAfterFirst = p.writes();
    const seq = texts(p.store);
    const retry = await p.tool.handler(task);
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
    assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
    assert.equal(p.writes(), savesAfterFirst, `${side}: no write on the retry`);
    assert.deepEqual(texts(p.store), seq);
    assert.equal(texts(p.store).filter(t => t === NEW).length, 1, 'never a duplicate');
  }
});

test('RETRY with the OPPOSITE side, or a text that already exists elsewhere, fails closed: no second copy is ever written', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ins('after', BETA)]))).ok, true);
  const seq = texts(p.store); const saves = p.writes();
  for (const task of [req([ins('before', BETA)]), req([ins('after', ALPHA, NEW)]), req([ins('after', ALPHA, GAMMA)]), req([ins('after', BETA, BETA)])]) {
    const r = await p.tool.handler(task);
    assert.equal(r.ok, false, JSON.stringify(task));
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
  }
  assert.deepEqual(texts(p.store), seq);
  assert.equal(p.writes(), saves);
});

test('MULTI-STEP: two inserts on two anchors (and a chained anchor) apply in one task; the retry is SATISFIED for every step', async () => {
  const p = product();
  const task = req([ins('after', ALPHA, 'A1'), ins('before', GAMMA, 'G0'), ins('after', 'A1', 'A2')]);
  const r = await p.tool.handler(task);
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED', 'APPLIED', 'APPLIED']);
  assert.deepEqual(texts(p.store).slice(0, 7), ['Heading', ALPHA, 'A1', 'A2', BETA, 'G0', GAMMA]);
  const saves = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED', 'SATISFIED', 'SATISFIED']);
  assert.equal(p.writes(), saves);
});

test('AMBIGUOUS ANCHOR fails closed: two paragraphs with the anchor text -> refused, ZERO writes, the document unchanged -- also when before/after were meant to disambiguate (they are not in the schema)', async () => {
  const p = product();
  const seq = texts(p.store);
  const r = await p.tool.handler(req([ins('after', TWIN)]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.deepEqual(texts(p.store), seq);
  assert.equal(p.writes(), 0);
  assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
  const withContext = await p.tool.handler(req([{ intent: 'insert_paragraph_after', selector: { text: TWIN, before: GAMMA }, text: NEW }]));
  assert.equal(withContext.code, 'UNKNOWN_FIELD');
  assert.equal(p.writes(), 0);
});

test('MISSING ANCHOR / kind mismatch / PRECONDITION mismatch fail closed: refused, ZERO writes', async () => {
  for (const sel of [{ text: 'GH233 NOT IN THE DOCUMENT' }, { text: ALPHA, kind: 'heading' }, { text: ALPHA, precondition: { text: 'something else' } }]) {
    const p = product();
    const seq = texts(p.store);
    const r = await p.tool.handler(req([{ intent: 'insert_paragraph_after', selector: sel, text: NEW }]));
    assert.equal(r.ok, false, JSON.stringify(sel) + JSON.stringify(r));
    assert.deepEqual(texts(p.store), seq);
    assert.equal(p.writes(), 0);
  }
});

test('A heading is a valid anchor: insert BEFORE the first heading puts the paragraph at the top of the document', async () => {
  const p = product();
  const seq = texts(p.store);
  const r = await p.tool.handler(req([{ intent: 'insert_paragraph_before', selector: { text: 'Heading', kind: 'heading' }, text: NEW }]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(texts(p.store), [NEW, ...seq]);
});

test('EDITOR FAULTS never read as APPLIED: the command silently does nothing, or inserts the paragraph twice -> not ok, no APPLIED step, nothing saved', async () => {
  for (const fault of ['ignore', 'insert-twice']) {
    const p = product({ fault });
    const seq = texts(p.store);
    const r = await p.tool.handler(req([ins('after', BETA)]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.deepEqual(texts(p.store), seq, fault);
    assert.equal(p.writes(), 0, fault);
  }
});

test('PRE-DISPATCH refusals never open a session: a refused insert task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[{ intent: 'insert_paragraph_after', selector: { heading: 'Heading', level: 1 }, text: NEW }], [ins('after', ALPHA, 'A\nB')], [ins('after', ALPHA, 'x'), ins('after', ALPHA, 'y')]]) {
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
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: ['rewrite_paragraph', 'define_style'], format: 'docx' }).ok, false, 'the pre-#233 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: ['rewrite_paragraph', 'define_style'], format: 'xlsx' }).ok, true);
  const rows = T.capabilityInventory().filter(r => r.FORMAT === 'DOCX' && r.EXPOSE_IN_PRODUCT_MCP === 'YES').map(r => r.TOOL).sort();
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(rows, ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
});
