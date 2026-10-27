'use strict';

const test = require('node:test');
const testSetListType = require('../helpers/trim-skip.cjs').testFor('set_list_type'); // release 3 trim: skipped while set_list_type is withheld
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
// GH #293 L7: the set_numbering / set_list_type co-member guards read the word /5 Bridge's whole affected set -- the released paths run on /5
const L7_V5 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5';
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #236 (DOCX P1 WRITE): apply_list / set_list_level / set_list_type served by docx_execute_task (the B commit). Design and the measured
 * reason for every schema decision: docs/GH236-LIST-APPLY-TYPE-LEVEL-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx
 * adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs (whose numberings model the real "a level's type is shared by every item on
 * it" semantics); the READ side is the product docx_read read_numbering on the SAME store. The LIVE acceptance is
 * test/live/docx-list-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const RICH = { bold: true, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const PLAIN = { ...RICH, bold: false };
const P1 = 'GH236 PLAIN ONE';
const P2 = 'GH236 PLAIN TWO';
const L1 = 'GH236 LIST ONE';       // bullet@0 on n1, shares n1@0 with L2
const L2 = 'GH236 LIST TWO';       // bullet@0 on n1
const L3 = 'GH236 LIST THREE';     // bullet@1 on n1 -- the ONLY item on n1@1
const doc = () => [
  { text: 'Heading', level: 1, props: PLAIN },
  { text: P1, props: RICH },
  { text: P2, props: PLAIN },
  { text: L1, props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } },
  { text: L2, props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } },
  { text: L3, props: PLAIN, list: { type: 'bullet', level: 1, numId: 'n1' } },
  { text: 'Closing paragraph.', props: PLAIN },
];
const al = (text, type, level) => ({ intent: 'apply_list', selector: { text }, type, ...(level === undefined ? {} : { level }) });
const sl = (text, level) => ({ intent: 'set_list_level', selector: { text }, level });
const st = (text, type) => ({ intent: 'set_list_type', selector: { text }, type });
const req = steps => ({ fileId: FILE, task: { steps } });

function product(script = {}, paragraphs = doc()) {
  const store = createStore(paragraphs);
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, runtime: L7_V5, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const numbering = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_numbering' }); assert.equal(r.ok, true, JSON.stringify(r)); return Object.fromEntries(r.result.items.map(i => [i.text, i.numbering])); };
  return { store, tool, numbering, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
const listOf = store => Object.fromEntries(store.durable.map(x => [x.text, x.list ? { ...x.list } : null]));

// ---- the step schemas (pre-dispatch) -------------------------------------------------------------------------------

test('SCHEMA: the three intents accept exactly their measured shapes (selector { text }, type bullet|numbered, level 0-8, apply_list level optional)', () => {
  const ok = [al(P1, 'bullet'), al(P1, 'numbered', 0), al(P1, 'bullet', 8), sl(L1, 0), sl(L1, 8), { ...al(P1, 'bullet'), id: 'a1' }];
  for (const s of ok) { const v = D.validateDocxWriteRequest(req([s])); assert.equal(v.ok, true, `${JSON.stringify(s)} -> ${v.code} ${v.reason}`); }
});

test('SCHEMA: every selector key beyond { text } is UNKNOWN_FIELD, 0 dispatch (the family\'s plan accepts a selector of exactly { text } or { ref })', () => {
  const extras = { kind: 'paragraph', before: 'x', after: 'y', contains: true, heading: 'Heading', fingerprint: 'abc', headingPath: [{ level: 1, text: 'Heading' }],
    sectionHeading: 'Heading', paragraphOrdinal: 1, precondition: { text: P1 }, ref: 's0', index: 1, position: 1, offset: 1, paragraphIndex: 1, bookmark: 'b', tableCell: { row: 1 } };
  for (const intent of [al(P1, 'bullet'), sl(L1, 1)]) {
    for (const [k, val] of Object.entries(extras)) {
      const v = D.validateDocxWriteRequest(req([{ ...intent, selector: { text: intent.selector.text, [k]: val } }]));
      assert.equal(v.code, 'UNKNOWN_FIELD', `${intent.intent} ${k}: ${v.code}`);
      assert.equal(v.dispatched, false);
    }
  }
});

test('SCHEMA: step keys outside each intent\'s own shape are UNKNOWN_FIELD (type on set_list_level, level on set_list_type, text / dependsOn / format anywhere)', () => {
  const cases = [{ ...sl(L1, 1), type: 'bullet' }, { ...al(P1, 'bullet'), text: 'x' }, { ...al(P1, 'bullet'), dependsOn: ['s0'] },
    { ...sl(L1, 1), format: { bold: true } }];
  for (const s of cases) { const v = D.validateDocxWriteRequest(req([s])); assert.equal(v.code, 'UNKNOWN_FIELD', `${JSON.stringify(s)} -> ${v.code}`); assert.equal(v.dispatched, false); }
});

test('SCHEMA: missing / invalid type or level, blank text and a missing selector are REFUSED_SHAPE, 0 dispatch', () => {
  const bad = [
    { intent: 'apply_list', selector: { text: P1 } }, { intent: 'apply_list', selector: { text: P1 }, type: 'dotted' }, { intent: 'apply_list', selector: { text: P1 }, type: 'bullet', level: 9 },
    { intent: 'apply_list', selector: { text: P1 }, type: 'bullet', level: -1 }, { intent: 'apply_list', selector: { text: P1 }, type: 'bullet', level: 1.5 },
    { intent: 'apply_list', selector: { text: P1 }, type: 'bullet', level: '1' },
    { intent: 'set_list_level', selector: { text: L1 } }, { intent: 'set_list_level', selector: { text: L1 }, level: 9 },
    al('   ', 'bullet'), al('', 'bullet'), { intent: 'apply_list', type: 'bullet' }, { intent: 'set_list_level', selector: {}, level: 1 },
  ];
  for (const s of bad) { const v = D.validateDocxWriteRequest(req([s])); assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
});

test('SCHEMA: a list task is that ONE step -- any second step (another list step, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const ins = { intent: 'insert_paragraph_after', selector: { text: P1 }, text: 'X' };
  for (const steps of [[al(P1, 'bullet'), al(P2, 'bullet')], [al(P1, 'bullet'), sl(P1, 2)], [ins, al(P1, 'bullet')], [sl(L1, 1), ins],
    [al(P1, 'bullet'), { intent: 'rewrite_paragraph', selector: { text: P2 }, text: 'y' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

testSetListType('WHY ONE STEP (the measured hazard, run BELOW the product validation): a verified apply_list next to a set_list_type the adapter cannot prove -> the task is not ok, yet the SAVE of the verified step carries the unproven shared-level change into the document', async () => {
  const store = createStore(doc());
  const exec = D.createDocxWriteExecution({ configSource: { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'K' } }; } }, gate: createSessionGate(),
    runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
  const r = await exec({ fileId: FILE, task: { steps: [al(P1, 'bullet'), st(L1, 'numbered')] } });
  assert.equal(r.ok, false);
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED', 'DISPATCHED_OUTCOME_UNKNOWN']);
  assert.equal(store.saves, 1, 'the verified step triggered a save');
  const list = listOf(store);
  assert.equal(list[L1].type, 'numbered', 'the UNPROVEN step landed in the saved document');
  assert.equal(list[L2].type, 'numbered', 'and it re-templated a level another item sits on');
  assert.equal(D.validateDocxWriteRequest(req([al(P1, 'bullet'), st(L1, 'numbered')])).code, 'REFUSED_SHAPE', 'the product refuses that task (GH #236b: set_list_type is released, but SOLO -- and the shared level is stopped before dispatch)');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid list step', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.equal(others.has('set_list_type'), false);   // (indent / outdent / remove_list joined with GH #237, set_list_type with GH #236b)
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([al(P1, 'bullet'), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the three list steps with their required fields', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  const by = i => items.oneOf.find(x => x.properties.intent.const === i);
  assert.deepEqual(by('apply_list').required, ['intent', 'selector', 'type']);
  assert.deepEqual(by('set_list_level').required, ['intent', 'selector', 'level']);
  if (by('set_list_type')) assert.deepEqual(by('set_list_type').required, ['intent', 'selector', 'type'], 'set_list_type is released with GH #236b');
  else assert.equal(by('set_list_type'), undefined, 'set_list_type is withheld by the release 3 trim');
  assert.deepEqual(by('apply_list').properties.selector.required, ['text']);
  assert.match(D.createDocxWriteTool({}).description, /apply_list, set_list_level/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_numbering on the SAME store -----------

test('READ SYMMETRY apply_list: the item becomes bullet / numbered at the asked level, read_numbering says so, a NEW numbering identity; no other item changed; the receipt carries L1 + L3 PASS', async () => {
  for (const [type, level] of [['bullet', undefined], ['numbered', 2]]) {
    const p = product();
    const before = await p.numbering();
    const r = await p.tool.handler(req([al(P1, type, level)]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    const after = await p.numbering();
    const n = after[P1];
    assert.deepEqual([n.present, n.type, n.level], [true, type, level === undefined ? 0 : level]);
    assert.ok(n.numId && !Object.values(before).some(b => b && b.numId === n.numId), 'a NEW numbering identity');
    for (const t of Object.keys(before)) if (t !== P1) assert.deepEqual(after[t], before[t], `${t} untouched`);
  }
});

test('READ SYMMETRY set_list_level: only that item\'s level moves (same numbering identity); the other items of the numbering keep type and level', async () => {
  const p = product();
  const before = await p.numbering();
  const r = await p.tool.handler(req([sl(L3, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const after = await p.numbering();
  assert.equal(after[L3].level, 2);
  assert.equal(after[L3].numId, before[L3].numId, 'same numbering');
  for (const t of Object.keys(before)) if (t !== L3) assert.deepEqual(after[t], before[t], `${t} untouched`);
  const shared = await product().tool.handler(req([sl(L1, 3)]));   // an item that SHARES its numbering level with another item moves alone
  assert.equal(shared.ok, true, JSON.stringify(shared));
});

testSetListType('RELEASED set_list_type (GH #236b): an item alone on its numbering level is re-templated (read_numbering says so); an item whose level another item sits on fails closed BEFORE dispatch with a named reason -- zero writes, nothing changed', async () => {
  const lone = product();
  const before = await lone.numbering();
  const r = await lone.tool.handler(req([st(L3, 'numbered')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await lone.numbering();
  assert.equal(after[L3].type, 'numbered');
  assert.equal(after[L3].numId, before[L3].numId, 'same numbering identity');
  const own = n => (n ? [n.present, n.type, n.level, n.numId] : n);   // the numbering's 9-slot levelFormats array of an item on the SAME numbering at another level moves with the re-templated level; its own type / level / numId do not
  for (const t of Object.keys(before)) if (t !== L3) assert.deepEqual(own(after[t]), own(before[t]), `${t} untouched`);
  for (const target of [L1, L2]) {
    const p = product();
    const beforeList = listOf(p.store);
    const refused = await p.tool.handler(req([st(target, 'numbered')]));
    assert.equal(refused.ok, false, JSON.stringify(refused));
    assert.match(JSON.stringify(refused), /list-shared-definition:[^:"]+:level\d+:1/, 'the numbering, the level and the number of other items are named');
    assert.equal(p.writes(), 0);
    assert.deepEqual(listOf(p.store), beforeList);
  }
  const caller = D.validateDocxWriteRequest(req([{ ...st(L3, 'numbered'), sharedDefinition: 'propagate' }]));
  assert.equal(caller.code, 'UNKNOWN_FIELD', 'the caller cannot choose the policy');
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes, for every intent', async () => {
  for (const step of [al(P1, 'numbered', 1), sl(L3, 2)]) {
    const p = product();
    const task = req([step]);
    assert.equal((await p.tool.handler(task)).ok, true);
    const saves = p.writes(); const list = listOf(p.store);
    const retry = await p.tool.handler(task);
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
    assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
    assert.equal(p.writes(), saves, `${step.intent}: no write on the retry`);
    assert.deepEqual(listOf(p.store), list);
  }
});

test('FAIL-CLOSED preconditions: set_list_level / set_list_type on a paragraph that is NOT a list item, a missing or duplicated target -> refused, ZERO writes, nothing changed', async () => {
  const dup = () => [...doc(), { text: P1, props: PLAIN }];
  const cases = [[doc(), sl(P1, 1)], [doc(), st(P1, 'numbered')], [doc(), al('GH236 NOT THERE', 'bullet')], [dup(), al(P1, 'bullet')]];
  for (const [paras, step] of cases) {
    const p = product({}, paras);
    const list = listOf(p.store);
    const r = await p.tool.handler(req([step]));
    assert.equal(r.ok, false, JSON.stringify(step) + JSON.stringify(r));
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
    assert.deepEqual(listOf(p.store), list);
    assert.equal(p.writes(), 0);
  }
});

test('EDITOR FAULTS never read as APPLIED: a silent no-op, a wrong level, a write that also moves another item, and a collateral re-template of another level -> not ok, nothing saved', async () => {
  const cases = [['ignore', al(P1, 'bullet')], ['list-wrong-level', al(P1, 'bullet')], ['list-touch-other', al(P1, 'bullet')], ['list-retemplate-other-level', sl(L3, 2)]];
  for (const [fault, step] of cases) {
    const p = product({ fault });
    const list = listOf(p.store);
    const r = await p.tool.handler(req([step]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.deepEqual(listOf(p.store), list, fault);
    assert.equal(p.writes(), 0, fault);
  }
});

test('KNOWN LIMIT (named in the decision doc): a heading is a valid target for apply_list -- the family accepts it; the schema cannot tell a heading from a paragraph by text', async () => {
  const p = product();
  const r = await p.tool.handler(req([al('Heading', 'bullet')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(listOf(p.store).Heading.type, 'bullet');
});

test('PRE-DISPATCH refusals never open a session: a refused list task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[{ intent: 'apply_list', selector: { text: P1, kind: 'paragraph' }, type: 'bullet' }], [al(P1, 'dotted')], [al(P1, 'bullet'), sl(P1, 2)], [sl(L1, 9)]]) {
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
  const pre = ['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'define_style'];
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#236 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  for (const i of ['apply_list', 'set_list_level']) assert.ok(T.capabilityInventory().some(r => r.TOOL === `docx:${i}` && r.EXPOSE_IN_PRODUCT_MCP === 'YES'), i);
  assert.equal(T.capabilityInventory().some(r => r.TOOL === 'docx:set_list_type' && r.EXPOSE_IN_PRODUCT_MCP === 'YES'), W.DOCX_WRITE_ALLOWLIST.includes('set_list_type'), 'set_list_type is released with GH #236b');
});
