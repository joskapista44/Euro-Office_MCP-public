'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #254 (DOCX P1 WRITE): merge_table_cells served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH254-TABLE-MERGE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs (its MergeCells model follows
 * the merge/split probe 1); the READ side is the product docx_read read_tables on the SAME store. The LIVE acceptance is test/live/docx-table-merge-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH254 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
// T0 3 x 4 (the merge target); T1 shares 'GH254 r1c1' with T0; T2 has a colspan owner (row 1, cols 1-2); T3 has a vertical merge over rows 1-2 (column 0)
const grid = (tag, rows, cols) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(`${tag} r${r}c${c}`))) });
const tables = () => [
  grid('GH254', 3, 4),
  { rows: [[cell('GH254 x0'), cell('GH254 x1')], [cell('GH254 y0'), cell('GH254 r1c1')]] },
  { rows: [[cell('GH254 F0'), cell('GH254 G0'), cell('GH254 H0'), cell('GH254 I0')], [cell('GH254 F1'), { ...cell('GH254 M1a', 'GH254 M1b'), span: 2 }, cell('GH254 I1')], [cell('GH254 F2'), cell('GH254 G2'), cell('GH254 H2'), cell('GH254 I2')]] },
  { rows: [[cell('GH254 V0'), cell('GH254 W0'), cell('GH254 X0')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH254 VTop', props: { ...PLAIN } }] }, cell('GH254 W1'), cell('GH254 X1')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH254 W2'), cell('GH254 X2')]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const mg = (start, end, anchor, merged, rows, cols, extra = {}) => ({ intent: 'merge_table_cells', selector: tc(start, anchor), endSelector: { tableCell: { text: end } }, terminalMergedText: merged, terminalMergedSpan: { rows, cols }, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

function product(script = {}, tabs = tables()) {
  const store = createStore(doc(), { doc: { tables: tabs } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  const readText = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_text' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.blocks.filter(b => b.kind !== 'table').map(b => b.text); };
  return { store, tool, readTables, readText, writes: () => store.saves };
}
const durable = store => JSON.stringify(store.durableDoc.tables);
const texts = rows => rows.map(r => r.map(c => c.text));

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: selector { tableCell: { text, tableAnchor } } (the anchor REQUIRED), endSelector { tableCell: { text } }, terminalMergedText and terminalMergedSpan are accepted, with an optional step id', () => {
  for (const steps of [[mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)], [{ ...mg('GH254 r1c0', 'GH254 r2c0', 'GH254 r0c0', 'GH254 r1c0 GH254 r2c0', 2, 1), id: 'm1' }],
    [mg('GH254 r0c1', 'GH254 r1c2', 'GH254 r2c0', 'GH254 r0c1 GH254 r0c2 GH254 r1c1 GH254 r1c2', 2, 2)]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: a missing anchor / endSelector / terminalMergedText / terminalMergedSpan, blank texts, an invalid span, equal corners, an anchor equal to a corner, a merged text equal to a corner, a one-cell span are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2);
  const bad = [
    mg('GH254 r1c1', 'GH254 r1c2', undefined, 'GH254 r1c1 GH254 r1c2', 1, 2),
    { ...ok, endSelector: undefined }, { ...ok, terminalMergedText: undefined }, { ...ok, terminalMergedSpan: undefined }, { ...ok, selector: {} },
    { ...ok, endSelector: {} }, { ...ok, endSelector: { tableCell: {} } }, { ...ok, terminalMergedSpan: {} }, { ...ok, terminalMergedSpan: { rows: 1 } }, { ...ok, terminalMergedSpan: { cols: 2 } },
    { ...ok, terminalMergedSpan: { rows: 0, cols: 2 } }, { ...ok, terminalMergedSpan: { rows: 1, cols: -1 } }, { ...ok, terminalMergedSpan: { rows: 1.5, cols: 2 } }, { ...ok, terminalMergedSpan: { rows: '1', cols: 2 } }, { ...ok, terminalMergedSpan: { rows: 1, cols: 1001 } },
    { ...ok, terminalMergedSpan: { rows: 1, cols: 1 } },
    mg('   ', 'GH254 r1c2', 'GH254 r0c0', 'x y', 1, 2), mg('GH254 r1c1', '', 'GH254 r0c0', 'x y', 1, 2), mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', '   ', 1, 2), mg('GH254 r1c1', 'GH254 r1c2', '   ', 'x y', 1, 2),
    mg('GH254 r1c1', 'GH254 r1c1', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c1', 1, 2),
    mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r1c1', 'GH254 r1c1 GH254 r1c2', 1, 2), mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r1c2', 'GH254 r1c1 GH254 r1c2', 1, 2), mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r1c1 GH254 r1c2', 'GH254 r1c1 GH254 r1c2', 1, 2),
    mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1', 1, 2),
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: every other selector form / alias / extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } form, an endSelector anchor, terminalTableAnchor, kind, position, splitRows, an extra span key, dependsOn', () => {
  const base = mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2);
  for (const extra of [{ selector: { text: 'GH254 r1c1' } }, { selector: { ...base.selector, kind: 'paragraph' } }, { selector: { tableCell: { text: 'GH254 r1c1', tableAnchor: 'GH254 r0c0', column: 1 } } },
    { endSelector: { tableCell: { text: 'GH254 r1c2', tableAnchor: 'GH254 r0c0' } } }, { endSelector: { text: 'GH254 r1c2' } }, { endSelector: { tableCell: { text: 'GH254 r1c2' }, kind: 'x' } },
    { terminalTableAnchor: 'GH254 r0c0' }, { terminalMergedSpan: { rows: 1, cols: 2, area: 2 } }, { position: 'last' }, { text: 'x' }, { splitRows: 2 }, { dependsOn: ['s0'] }, { terminalColumnCount: 3 }]) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a merge task is that ONE step -- a second step (another merge, a column step, a rewrite, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const one = mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2);
  const two = mg('GH254 r1c0', 'GH254 r2c0', 'GH254 r0c0', 'GH254 r1c0 GH254 r2c0', 2, 1);
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  const col = { intent: 'delete_table_column', selector: tc('GH254 r1c3', 'GH254 r0c0'), terminalColumnCount: 3 };
  for (const steps of [[one, two], [one, col], [rw, one], [one, rw]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): a merge whose successor fails leaves the FIRST merge saved; two merges whose regions touch fail the second AFTER the save', async () => {
  const run = async steps => {
    const store = createStore(doc(), { doc: { tables: tables() } });
    const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
    return { r: await exec({ fileId: FILE, task: { steps } }), store };
  };
  const failSecond = await run([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2), mg('GH254 Not There', 'GH254 r2c3', 'GH254 r0c0', 'GH254 Not There GH254 r2c3', 1, 2)]);
  assert.equal(failSecond.r.ok, false);
  assert.deepEqual(failSecond.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(failSecond.store.saves, 1, 'the first merge was saved next to the failed second step');
  assert.equal(failSecond.store.durableDoc.tables[0].rows[1].length, 3);
  assert.equal(D.validateDocxWriteRequest(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2), mg('GH254 r2c0', 'GH254 r2c1', 'GH254 r0c0', 'GH254 r2c0 GH254 r2c1', 1, 2)])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid merge (split #255, formatting #256-#258)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the merge step: selector.tableCell with text AND tableAnchor required, endSelector.tableCell text only, terminalMergedText, terminalMergedSpan { rows, cols } (integers 1..1000), nothing else', () => {
  const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === 'merge_table_cells');
  assert.deepEqual(s.required, ['intent', 'selector', 'endSelector', 'terminalMergedText', 'terminalMergedSpan']);
  assert.deepEqual(Object.keys(s.properties).sort(), ['endSelector', 'id', 'intent', 'selector', 'terminalMergedSpan', 'terminalMergedText']);
  assert.deepEqual(s.properties.selector.properties.tableCell.required, ['text', 'tableAnchor']);
  assert.deepEqual(Object.keys(s.properties.endSelector.properties.tableCell.properties), ['text']);
  assert.deepEqual(s.properties.endSelector.properties.tableCell.required, ['text']);
  assert.deepEqual(s.properties.terminalMergedSpan.required, ['rows', 'cols']);
  assert.deepEqual(s.properties.terminalMergedSpan.properties.rows, { type: 'integer', minimum: 1, maximum: 1000 });
  assert.equal(s.additionalProperties, false);
  assert.match(D.createDocxWriteTool({}).description, /merge_table_cells/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

test('READ SYMMETRY (horizontal 1x2, row 1 of T0): the row reads one cell fewer with the joined text, every other row / table / the text outside unchanged; L1 + L3 PASS', async () => {
  const p = product();
  const b = await p.readTables();
  const outside = await p.readText();
  const r = await p.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes(), 1);
  const a = await p.readTables();
  assert.deepEqual(texts(a[0].rows)[1], ['GH254 r1c0', 'GH254 r1c1 GH254 r1c2', 'GH254 r1c3']);
  assert.deepEqual(a[0].rows[0], b[0].rows[0]); assert.deepEqual(a[0].rows[2], b[0].rows[2]);
  assert.deepEqual(a.slice(1), b.slice(1), 'the other tables are untouched');
  assert.deepEqual(await p.readText(), outside);
});

test('READ SYMMETRY (vertical 2x1, rows 1-2 of column 0): the owner holds both texts, the covered cell reads empty, rows keep their cell count; every other cell unchanged', async () => {
  const p = product();
  const b = await p.readTables();
  const r = await p.tool.handler(req([mg('GH254 r1c0', 'GH254 r2c0', 'GH254 r0c0', 'GH254 r1c0 GH254 r2c0', 2, 1)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const a = await p.readTables();
  assert.deepEqual(a[0].rows.map(row => row.length), b[0].rows.map(row => row.length));
  assert.equal(a[0].rows[1][0].text, 'GH254 r1c0 GH254 r2c0');
  assert.equal(a[0].rows[2][0].text, '');
  assert.deepEqual(texts(a[0].rows)[0], texts(b[0].rows)[0]);
  assert.deepEqual(texts(a[0].rows)[1].slice(1), texts(b[0].rows)[1].slice(1));
  assert.deepEqual(texts(a[0].rows)[2].slice(1), texts(b[0].rows)[2].slice(1));
  assert.deepEqual(a.slice(1), b.slice(1));
});

test('READ SYMMETRY (2x2 region): the owner holds the four texts in ROW-MAJOR order; the other cells of the table unchanged', async () => {
  const p = product();
  const b = await p.readTables();
  const merged = 'GH254 r0c1 GH254 r0c2 GH254 r1c1 GH254 r1c2';
  const r = await p.tool.handler(req([mg('GH254 r0c1', 'GH254 r1c2', 'GH254 r2c0', merged, 2, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const a = await p.readTables();
  assert.equal(a[0].rows[0].filter(c => c.text === merged).length, 1);
  assert.deepEqual(texts(a[0].rows)[2], texts(b[0].rows)[2], 'the row below the region is untouched');
  assert.deepEqual(a.slice(1), b.slice(1));
});

test('CORNER ORDER: the END corner may come first (the region is the bounding rectangle of the two corners), the merged text stays row-major', async () => {
  const p = product();
  const r = await p.tool.handler(req([mg('GH254 r1c2', 'GH254 r1c1', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(texts((await p.readTables())[0].rows)[1][1], 'GH254 r1c1 GH254 r1c2');
});

test('RETRY: the same task again is SATISFIED / noOp / 0 writes (the corners are gone, the merged text is there once as an owner of the declared span)', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]))).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
});

test('NAMED LIMIT (measured, the #251 / #253 survivor question): corners that NEVER existed (typos) read SATISFIED when the terminal text is an owner of the declared span in the anchor table -- the proof is the terminal identity, not the corners\' past', async () => {
  const done = product();
  assert.equal((await done.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]))).ok, true);
  const snap = durable(done.store); const saves = done.writes();
  const typo = await done.tool.handler(req([mg('GH254 r1c1 typo', 'GH254 r1c2 typo', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]));
  assert.equal(typo.ok, true, JSON.stringify(typo));
  assert.equal(typo.noOp, true);
  assert.deepEqual(typo.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(done.writes(), saves); assert.equal(durable(done.store), snap);
  const wrongSpan = await done.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 3)]));
  assert.equal(wrongSpan.ok, false, 'the declared span must match the owner');
  assert.match(wrongSpan.reason, /merge-terminal-unprovable/);
});

test('TABLE IDENTITY: the anchor decides the table; a corner text two tables hold merges in the ANCHORED table only', async () => {
  const p = product();
  const before = await p.readTables();
  const r = await p.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.deepEqual(after[1], before[1], 'the twin in the other table is untouched');
  assert.equal(after[1].rows[1][1].text, 'GH254 r1c1');
});

test('GH #291 T-07: a region that wholly contains a vertical merge -> APPLIED: one owner of the region span, the paragraphs in row-major owner order', async () => {
  const p = product();
  const r = await p.tool.handler(req([mg('GH254 V0', 'GH254 W2', 'GH254 X0', 'GH254 V0 GH254 W0 GH254 VTop GH254 W1 GH254 W2', 3, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const vt = p.store.durableDoc.tables.find(t => t.rows.flat().some(c => c.paras.some(x => x.text === 'GH254 VTop'))).rows;
  assert.deepEqual([vt[0][0].span, vt[0][0].rowspan, vt[0][0].paras.map(x => x.text).join(' ')], [2, 3, 'GH254 V0 GH254 W0 GH254 VTop GH254 W1 GH254 W2']);
});

test('REGION GATE (fail closed, 0 writes, the document unchanged): partial cover of a merged cell, a wrong terminal text / span, a missing corner, an anchor inside the region, a missing / ambiguous anchor', async () => {
  const dupAnchor = [{ rows: [[cell('GH254 Anchor'), cell('GH254 X'), cell('GH254 Y')], [cell('GH254 Z'), cell('GH254 Q'), cell('GH254 R')]] }, { rows: [[cell('GH254 Anchor'), cell('GH254 K')]] }];
  for (const [tabs, steps, why] of [
    [tables(), [mg('GH254 F1', 'GH254 G2', 'GH254 F0', 'GH254 F1 GH254 M1a GH254 M1b GH254 F2 GH254 G2', 2, 2)], /merge-region-partially-covers-a-merged-cell/],
    // GH #291 T-07: a region wholly containing a vertical merge is no longer refused (positive test below)
    [tables(), [mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 wrong text', 1, 2)], /merge-terminal-contract-mismatch/],
    [tables(), [mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 3)], /merge-terminal-contract-mismatch/],
    [tables(), [mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 2, 1)], /merge-terminal-contract-mismatch/],
    [tables(), [mg('GH254 Not There', 'GH254 r1c2', 'GH254 r0c0', 'GH254 Not There GH254 r1c2', 1, 2)], /merge-terminal-unprovable|merge-corner-not-unique/],
    [tables(), [mg('GH254 r0c1', 'GH254 r1c2', 'GH254 r0c2', 'GH254 r0c1 GH254 r0c2 GH254 r1c1 GH254 r1c2', 2, 2)], /table-anchor-in-merge-region/],
    [tables(), [mg('GH254 r1c1', 'GH254 r1c2', 'GH254 No Such', 'GH254 r1c1 GH254 r1c2', 1, 2)], /table-anchor-not-found/],
    [dupAnchor, [mg('GH254 Q', 'GH254 R', 'GH254 Anchor', 'GH254 Q GH254 R', 1, 2)], /table-anchor-ambiguous/],
  ]) {
    const p = product({}, tabs);
    const snap = durable(p.store);
    const r = await p.tool.handler(req(steps));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.match(r.reason, why, r.reason);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
});

test('GUARD BOUNDARY: a vertical merge OUTSIDE the region (another column / row) is no obstacle; a colspan owner entirely OUTSIDE the region is no obstacle', async () => {
  const v = product();
  const rv = await v.tool.handler(req([mg('GH254 W0', 'GH254 X0', 'GH254 V0', 'GH254 W0 GH254 X0', 1, 2)]));
  assert.equal(rv.ok, true, JSON.stringify(rv));
  const h = product();
  const rh = await h.tool.handler(req([mg('GH254 F2', 'GH254 G2', 'GH254 F0', 'GH254 F2 GH254 G2', 1, 2)]));
  assert.equal(rh.ok, true, JSON.stringify(rh));
});

test('FAIL-CLOSED editor faults: a silent no-op, lost content, joined paragraphs, a re-formatted paragraph, a wrong span, a cell / table touched -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'merge-lose-content', 'merge-join-paragraphs', 'merge-reformat', 'merge-wrong-span', 'merge-touch-cell', 'merge-touch-other-table']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([mg('GH254 r1c1', 'GH254 r1c2', 'GH254 r0c0', 'GH254 r1c1 GH254 r1c2', 1, 2)]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
});
