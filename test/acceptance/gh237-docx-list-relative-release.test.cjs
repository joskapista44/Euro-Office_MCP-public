'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #237 (DOCX P1 WRITE): indent_list_item / outdent_list_item / remove_list served by docx_execute_task (the B commit). Design and the measured
 * reason for every rule: docs/GH237-LIST-INDENT-OUTDENT-REMOVE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter /
 * executor / lifecycle on test/fakes/fake-docx-driver.cjs (its numberings model "a level's type is shared by every item on it", its runtime gate
 * is the REAL driver's internalGateOf); the READ side is the product docx_read read_numbering on the SAME store. The LIVE acceptance is
 * test/live/docx-list-relative-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const P1 = 'GH237 PLAIN ONE';
const P2 = 'GH237 PLAIN TWO';
const L1 = 'GH237 LIST ONE';       // bullet@0 on n1, shares n1@0 with L2
const L2 = 'GH237 LIST TWO';       // bullet@0 on n1
const L3 = 'GH237 LIST THREE';     // bullet@1 on n1
const doc = () => [
  { text: 'Heading', level: 1, props: PLAIN },
  { text: P1, props: PLAIN },
  { text: P2, props: PLAIN },
  { text: L1, props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } },
  { text: L2, props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } },
  { text: L3, props: PLAIN, list: { type: 'bullet', level: 1, numId: 'n1' } },
  { text: 'Closing paragraph.', props: PLAIN },
];
const al = (text, type, level) => ({ intent: 'apply_list', selector: { text }, type, ...(level === undefined ? {} : { level }) });
const sl = (text, level) => ({ intent: 'set_list_level', selector: { text }, level });
const ind = text => ({ intent: 'indent_list_item', selector: { text } });
const out = text => ({ intent: 'outdent_list_item', selector: { text } });
const rm = text => ({ intent: 'remove_list', selector: { text } });
const req = steps => ({ fileId: FILE, task: { steps } });

function product(script = {}, paragraphs = doc(), storeOpts = {}) {
  const store = createStore(paragraphs, storeOpts);
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const numbering = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_numbering' }); assert.equal(r.ok, true, JSON.stringify(r)); return Object.fromEntries(r.result.items.map(i => [i.text, i.numbering])); };
  return { store, tool, numbering, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
const listOf = store => Object.fromEntries(store.durable.map(x => [x.text, x.list ? { ...x.list } : null]));

// ---- the step schemas and the pair rule (pre-dispatch) ----------------------------------------------------------------

test('SCHEMA: remove_list { text } alone, and the one released relative shape [apply_list | set_list_level, indent | outdent] on the SAME text, are accepted', () => {
  const ok = [[rm(L1)], [al(P1, 'bullet'), ind(P1)], [al(P1, 'numbered', 3), out(P1)], [sl(L3, 2), ind(L3)], [sl(L3, 0), out(L3)], [al(P1, 'bullet', 8), out(P1)], [{ ...rm(L1), id: 'r1' }]];
  for (const steps of ok) { const v = D.validateDocxWriteRequest(req(steps)); assert.equal(v.ok, true, `${JSON.stringify(steps.map(s => s.intent))} -> ${v.code} ${v.reason}`); }
});

test('SCHEMA: every key beyond { text } on the selector, and every step key beyond intent / id / selector, is UNKNOWN_FIELD (0 dispatch)', () => {
  const extras = { kind: 'paragraph', before: 'x', contains: true, heading: 'Heading', fingerprint: 'abc', precondition: { text: L1 }, ref: 's0', index: 1 };
  for (const mk of [ind, out, rm]) {
    for (const [k, val] of Object.entries(extras)) {
      const steps = mk === rm ? [{ ...rm(L1), selector: { text: L1, [k]: val } }] : [al(P1, 'bullet'), { ...mk(P1), selector: { text: P1, [k]: val } }];
      const v = D.validateDocxWriteRequest(req(steps));
      assert.equal(v.code, 'UNKNOWN_FIELD', `${mk(L1).intent} ${k}: ${v.code}`);
      assert.equal(v.dispatched, false);
    }
    for (const [k, val] of Object.entries({ type: 'bullet', level: 1, text: 'x', dependsOn: ['s0'] })) {
      const steps = mk === rm ? [{ ...rm(L1), [k]: val }] : [al(P1, 'bullet'), { ...mk(P1), [k]: val }];
      assert.equal(D.validateDocxWriteRequest(req(steps)).code, 'UNKNOWN_FIELD', `${mk(L1).intent} step ${k}`);
    }
  }
});

test('SCHEMA: blank text and a missing selector are REFUSED_SHAPE', () => {
  for (const s of [rm('   '), rm(''), { intent: 'remove_list' }, { intent: 'remove_list', selector: {} }]) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(s));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a RELATIVE step is refused unless its task is exactly [absolute anchor, relative] on the same text (a standalone one has no provable terminal level)', () => {
  const bad = [
    [ind(L3)], [out(L3)],                                          // standalone
    [al(P1, 'bullet'), ind(P2)], [sl(L3, 2), out(L1)],             // another text
    [al(P1, 'bullet'), ind(P1), ind(P1)], [al(P1, 'bullet'), ind(P1), out(P1)],  // more than one relative step
    [ind(P1), al(P1, 'bullet')],                                   // relative first
    [rm(L3), ind(L3)],                                             // not an absolute anchor (a relative step with no absolute anchor)
    [al(P1, 'bullet'), sl(P1, 1), ind(P1)],                        // anchor + anchor + relative
    [al(P1, 'bullet', 8), ind(P1)], [sl(L3, 8), ind(L3)],          // indent beyond level 8
    [ind(L3), { intent: 'insert_paragraph_after', selector: { text: P1 }, text: 'X' }],
  ];
  for (const steps of bad) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: remove_list is that ONE step; the #236 list rule is unchanged (an anchor next to anything but its relative step is still refused)', () => {
  for (const steps of [[rm(L1), rm(L2)], [rm(L1), al(P1, 'bullet')], [al(P1, 'bullet'), rm(P1)], [al(P1, 'bullet'), al(P2, 'bullet')]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
  }
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid step (set_numbering joined with GH #248)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.ok(others.size >= 0);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([rm(L1), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the three steps; the relative ones and remove_list take only intent / id / selector', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  for (const intent of ['indent_list_item', 'outdent_list_item', 'remove_list']) {
    const s = items.oneOf.find(x => x.properties.intent.const === intent);
    assert.deepEqual(s.required, ['intent', 'selector']);
    assert.deepEqual(Object.keys(s.properties).sort(), ['id', 'intent', 'selector']);
    assert.deepEqual(s.properties.selector.required, ['text']);
  }
  assert.match(D.createDocxWriteTool({}).description, /indent_list_item, outdent_list_item, remove_list/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_numbering on the SAME store -----------------

test('remove_list READ SYMMETRY: the item leaves its list (read_numbering present:false, no numId), the OTHER item of the same numbering keeps type / level / numId; L1 + L3 PASS', async () => {
  const p = product();
  const before = await p.numbering();
  const r = await p.tool.handler(req([rm(L1)]));      // L1 shares n1@0 with L2
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  assert.equal(r.wholeTask, 'PASS');
  const after = await p.numbering();
  assert.deepEqual(after[L1], { present: false });
  for (const t of Object.keys(before)) if (t !== L1) assert.deepEqual(after[t], before[t], `${t} untouched`);
  assert.equal(listOf(p.store)[L1], null);
});

test('remove_list RETRY: the SAME task again is SATISFIED / noOp / 0 writes; on a paragraph that is not a list it is SATISFIED from the start', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([rm(L1)]))).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(req([rm(L1)]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
  const plain = product();
  const r = await plain.tool.handler(req([rm(P1)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.equal(plain.writes(), 0);
});

test('remove_list BRIDGE GATE, first line (classification): a runtime the Bridge removal is not proven on -> FAILED docx-numbering-remove-bridge-write-runtime-not-proven, CERTAIN, ZERO writes, nothing changed, NO fallback', async () => {
  for (const script of [{ gateOverride: { numberingRemoveBridge: false } }, { runtime: '9.3.4.60|other' }]) {
    const p = product(script);
    const list = listOf(p.store);
    const r = await p.tool.handler(req([rm(L1)]));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.equal(r.code, 'OPERATION_FAILED');
    assert.match(r.reason, /docx-numbering-remove-bridge-write-runtime-not-proven/);
    assert.equal(r.mutationCertainty, 'CERTAIN');
    assert.equal(r.operations[0].failure.stage, 'CLASSIFICATION');
    assert.equal(p.writes(), 0);
    assert.deepEqual(listOf(p.store), list, 'the document is untouched -- there is no direct-internal path to fall back to');
  }
});

test('remove_list BRIDGE GATE, second line (dispatch): classification open but the runtime at dispatch is not proven -> refused before the write, UNCERTAIN, ZERO writes', async () => {
  const p = product({ runtime: '9.3.4.60|other', gateOverride: { numberingRemoveBridge: true } });
  const list = listOf(p.store);
  const r = await p.tool.handler(req([rm(L1)]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.match(r.reason, /DISPATCH_REFUSED_NOT_PROVEN_PRE_DISPATCH/);
  assert.equal(r.mutationCertainty, 'UNCERTAIN');
  assert.equal(p.writes(), 0);
  assert.deepEqual(listOf(p.store), list);
});

test('READ SYMMETRY indent / outdent: the level moves by exactly one in read_numbering, the numbering identity stays, every other item is untouched; retry is SATISFIED for both steps / 0 writes', async () => {
  for (const [steps, level] of [[[al(P1, 'bullet'), ind(P1)], 1], [[al(P1, 'numbered', 3), out(P1)], 2], [[sl(L3, 2), ind(L3)], 3], [[sl(L3, 3), out(L3)], 2]]) {
    const p = product();
    const before = await p.numbering();
    const target = steps[0].selector.text;
    const r = await p.tool.handler(req(steps));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS'], ['APPLIED', 'PASS', 'PASS']]);
    const after = await p.numbering();
    assert.equal(after[target].level, level, steps.map(s => s.intent).join('+'));
    if (steps[0].intent === 'set_list_level') assert.equal(after[target].numId, before[target].numId, 'a level move keeps the numbering');
    for (const t of Object.keys(before)) if (t !== target) assert.deepEqual(after[t], before[t], `${t} untouched`);
    const saves = p.writes();
    const retry = await p.tool.handler(req(steps));
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
    assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED', 'SATISFIED']);
    assert.equal(p.writes(), saves);
  }
});

test('outdent from level 0 is the Bridge removal: [set_list_level 0, outdent] takes the item out of its list (read_numbering present:false); the retry is SATISFIED', async () => {
  const p = product();
  const before = await p.numbering();
  const r = await p.tool.handler(req([sl(L1, 0), out(L1)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.numbering();
  assert.deepEqual(after[L1], { present: false });
  for (const t of Object.keys(before)) if (t !== L1) assert.deepEqual(after[t], before[t], `${t} untouched`);
  const retry = await p.tool.handler(req([sl(L1, 0), out(L1)]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
});

test('DESTINATION-LEVEL TYPE CHECK survives the product path: an indent INTO a level whose format differs gives the item that level\'s type (read live, not guessed)', async () => {
  const paras = [{ text: 'Heading', level: 1, props: PLAIN },
    { text: 'N0', props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } }, { text: 'N1', props: PLAIN, list: { type: 'bullet', level: 1, numId: 'n1' } }, { text: 'Tail', props: PLAIN }];
  const numberings = { n1: ['bullet', 'bullet', 'numbered', 'bullet', 'bullet', 'bullet', 'bullet', 'bullet', 'bullet'] };
  const p = product({}, paras, { doc: { numberings } });
  const r = await p.tool.handler(req([sl('N1', 1), ind('N1')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const n = (await p.numbering()).N1;
  assert.deepEqual([n.type, n.level], ['numbered', 2], 'the type is the destination level\'s');
  assert.equal(n.levelFormats[2], 13);
});

test('DESTINATION-LEVEL GUARDS fail closed BEFORE any dispatch through the product: the per-level format read unreadable, disagreeing with the item, or an unknown destination format -> ZERO writes, nothing changed', async () => {
  const cases = [[{ levelFormatsUnreadable: true }, /eo-internal-list-level-format-unavailable/], [{ levelFormatsDisagree: true }, /eo-internal-list-level-format-disagrees/], [{ levelFormatUnknownAt: 2 }, /list-destination-level-type-unknown/]];
  for (const [script, why] of cases) {
    const p = product(script);
    const list = listOf(p.store);
    const r = await p.tool.handler(req([sl(L3, 1), ind(L3)]));    // L3 is on level 1: the destination is level 2
    assert.equal(r.ok, false, `${JSON.stringify(script)}: ${JSON.stringify(r)}`);
    assert.match(r.reason, why);
    assert.equal(r.operations[r.operations.length - 1].status === 'APPLIED', false, 'the relative step never ran');
    assert.deepEqual(listOf(p.store), list, 'the list state is unchanged');
    // a destination-format guard that only the SECOND step needs lets the anchor (here a set_list_level onto the level the item is already on) run first:
    // that is a verified no-op write, visible in the receipt as APPLIED next to the FAILED relative step
    assert.equal(p.writes(), r.operations[0].status === 'APPLIED' ? 1 : 0);
  }
});

test('WHY A PAIR (run BELOW the product validation): a standalone relative step is FAILED list-relative-level-unprovable, and an anchor on ANOTHER text leaves the anchor saved next to a failed step', async () => {
  const store = createStore(doc());
  const exec = D.createDocxWriteExecution({ configSource: { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'K' } }; } }, gate: createSessionGate(),
    runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
  const alone = await exec({ fileId: FILE, task: { steps: [ind(L3)] } });
  assert.equal(alone.ok, false);
  assert.match(alone.reason, /list-relative-level-unprovable/);
  assert.equal(store.saves, 0);
  const other = await exec({ fileId: FILE, task: { steps: [al(P1, 'bullet'), ind(P2)] } });
  assert.equal(other.ok, false);
  assert.deepEqual(other.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(store.saves, 1, 'the verified anchor was saved next to the failed relative step');
  assert.equal(D.validateDocxWriteRequest(req([al(P1, 'bullet'), ind(P2)])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('FAIL-CLOSED: a missing or duplicated remove_list target, and editor faults (a silent no-op; the write also moves another item), never read as APPLIED and save nothing', async () => {
  const dup = () => [...doc(), { text: L1, props: PLAIN }];
  for (const [paras, script, steps] of [[doc(), {}, [rm('GH237 NOT THERE')]], [dup(), {}, [rm(L1)]], [doc(), { fault: 'ignore' }, [rm(L1)]], [doc(), { fault: 'list-touch-other' }, [rm(L1)]],
    [doc(), { fault: 'list-wrong-level' }, [sl(L3, 2), ind(L3)]]]) {
    const p = product(script, paras);
    const list = listOf(p.store);
    const r = await p.tool.handler(req(steps));
    assert.equal(r.ok, false, `${JSON.stringify(script)} ${steps.map(s => s.intent)}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false);
    assert.deepEqual(listOf(p.store), list);
    assert.equal(p.writes(), 0);
  }
});

test('PRE-DISPATCH refusals never open a session: a refused list-relative task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[ind(L3)], [al(P1, 'bullet'), ind(P2)], [rm(L1), rm(L2)], [{ ...rm(L1), selector: { text: L1, kind: 'paragraph' } }], [sl(L3, 8), ind(L3)]]) {
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
  const pre = [...W.DOCX_WRITE_ALLOWLIST].slice(0, 9);
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#237 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  for (const i of ['indent_list_item', 'outdent_list_item', 'remove_list']) assert.ok(T.capabilityInventory().some(r => r.TOOL === `docx:${i}` && r.EXPOSE_IN_PRODUCT_MCP === 'YES'), i);
});
