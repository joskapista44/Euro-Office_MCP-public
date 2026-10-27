'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const MS = require('../../src/adapters/docx/families/table-merge-split.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #255 (DOCX P1 WRITE): split_table_cell served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH255-TABLE-SPLIT-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs (its Split model follows
 * the W05-1a/1b measurements); the READ side is the product docx_read read_tables on the SAME store. The LIVE acceptance is test/live/docx-table-split-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH255 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
const grid = (tag, rows, cols) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(`${tag} r${r}c${c}`))) });
// T0 3 x 3 plain; T1 shares 'GH255 r1c1' with T0; T2 has a merged owner (row 1, cols 0-1, TWO paragraphs); T3 has a vertical merge (rows 1-2, column 0)
const tables = () => [
  grid('GH255', 3, 3),
  { rows: [[cell('GH255 x0'), cell('GH255 x1')], [cell('GH255 y0'), cell('GH255 r1c1')]] },
  { rows: [[cell('GH255 F0'), cell('GH255 G0'), cell('GH255 H0')], [{ ...cell('GH255 M1a', 'GH255 M1b'), span: 2 }, cell('GH255 H1')], [cell('GH255 F2'), cell('GH255 G2'), cell('GH255 H2')]] },
  { rows: [[cell('GH255 V0'), cell('GH255 W0')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH255 VTop', props: { ...PLAIN } }] }, cell('GH255 W1')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH255 W2')]] },
];
const MERGED = 'GH255 M1a GH255 M1b';
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const sp = (text, anchor, splitRows, splitCols, extra = {}) => ({ intent: 'split_table_cell', selector: tc(text, anchor), splitRows, splitCols, ...extra });
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
const nonEmpty = rows => rows.flat().map(c => c.text).filter(t => t !== '');
const count = rows => rows.reduce((n, r) => n + r.length, 0);

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: { tableCell: { text, tableAnchor } } (the anchor REQUIRED) plus splitRows 1..2 and splitCols 1..3 is accepted, with an optional step id', () => {
  for (const steps of [[sp('GH255 r1c1', 'GH255 r0c0', 1, 2)], [sp('GH255 r1c1', 'GH255 r0c0', 2, 1)], [sp('GH255 r0c0', 'GH255 r2c2', 2, 2)], [{ ...sp('GH255 r1c0', 'GH255 r0c2', 1, 3), id: 's1' }], [sp(MERGED, 'GH255 F0', 1, 2)]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: a missing anchor / splitRows / splitCols, blank texts, a value outside the MEASURED sets, a 1x1 split, an anchor equal to the target are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = sp('GH255 r1c1', 'GH255 r0c0', 1, 2);
  const bad = [
    sp('GH255 r1c1', undefined, 1, 2), { ...ok, splitRows: undefined }, { ...ok, splitCols: undefined }, { ...ok, selector: {} }, { ...ok, selector: { tableCell: { tableAnchor: 'GH255 r0c0' } } },
    { ...ok, splitRows: 0 }, { ...ok, splitRows: 5 }, { ...ok, splitRows: 20 }, { ...ok, splitRows: 1.5 }, { ...ok, splitRows: '1' }, { ...ok, splitRows: -1 },
    { ...ok, splitCols: 0 }, { ...ok, splitCols: 7 }, { ...ok, splitCols: 20 }, { ...ok, splitCols: 2.5 }, { ...ok, splitCols: '2' },
    sp('GH255 r1c1', 'GH255 r0c0', 1, 1),
    sp('   ', 'GH255 r0c0', 1, 2), sp('', 'GH255 r0c0', 1, 2), sp('GH255 r1c1', '   ', 1, 2), sp('GH255 r1c1', 'GH255 r1c1', 1, 2),
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: every other selector form / alias / extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } form, endSelector, terminalTableAnchor, the merge fields, kind, position, dependsOn', () => {
  const base = sp('GH255 r1c1', 'GH255 r0c0', 1, 2);
  for (const extra of [{ selector: { text: 'GH255 r1c1' } }, { selector: { ...base.selector, kind: 'paragraph' } }, { selector: { tableCell: { text: 'GH255 r1c1', tableAnchor: 'GH255 r0c0', column: 1 } } },
    { endSelector: { tableCell: { text: 'GH255 r1c2' } } }, { terminalTableAnchor: 'GH255 r0c0' }, { terminalMergedText: 'x' }, { terminalMergedSpan: { rows: 1, cols: 2 } }, { position: 'last' }, { text: 'x' }, { dependsOn: ['s0'] }, { terminalColumnCount: 3 }]) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a split task is that ONE step -- a second step (another split, a merge, a rewrite, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const one = sp('GH255 r1c1', 'GH255 r0c0', 1, 2);
  const two = sp('GH255 r2c0', 'GH255 r0c0', 2, 1);
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  const merge = { intent: 'merge_table_cells', selector: tc('GH255 r1c0', 'GH255 r0c0'), endSelector: { tableCell: { text: 'GH255 r2c0' } }, terminalMergedText: 'GH255 r1c0 GH255 r2c0', terminalMergedSpan: { rows: 2, cols: 1 } };
  for (const steps of [[one, two], [one, merge], [merge, one], [rw, one], [one, rw]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): a split whose successor fails leaves the FIRST split saved', async () => {
  const store = createStore(doc(), { doc: { tables: tables() } });
  const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
  const r = await exec({ fileId: FILE, task: { steps: [sp('GH255 r1c1', 'GH255 r0c0', 1, 2), sp('GH255 Not There', 'GH255 r0c0', 1, 2)] } });
  assert.equal(r.ok, false);
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(store.saves, 1, 'the first split was saved next to the failed second step');
  assert.equal(store.durableDoc.tables[0].rows[1].length, 4);
  assert.equal(D.validateDocxWriteRequest(req([sp('GH255 r1c1', 'GH255 r0c0', 1, 2), sp('GH255 r2c0', 'GH255 r0c0', 2, 1)])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid split (formatting #256-#258)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([sp('GH255 r1c1', 'GH255 r0c0', 1, 2), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the split step: selector.tableCell with text AND tableAnchor required, splitRows / splitCols required integers whose MAXIMA are the adapter\'s measured sets, nothing else', () => {
  const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === 'split_table_cell');
  assert.deepEqual(s.required, ['intent', 'selector', 'splitRows', 'splitCols']);
  assert.deepEqual(Object.keys(s.properties).sort(), ['id', 'intent', 'selector', 'splitCols', 'splitRows']);
  assert.deepEqual(s.properties.selector.properties.tableCell.required, ['text', 'tableAnchor']);
  assert.deepEqual(s.properties.splitRows, { type: 'integer', minimum: 1, maximum: Math.max(...MS.SPLIT_ROWS_MEASURED) }, 'the schema maximum IS the adapter\'s measured set');
  assert.deepEqual(s.properties.splitCols, { type: 'integer', minimum: 1, maximum: Math.max(...MS.SPLIT_COLS_MEASURED) });
  assert.deepEqual([...MS.SPLIT_ROWS_MEASURED], [1, 2, 3, 4]); assert.deepEqual([...MS.SPLIT_COLS_MEASURED], [1, 2, 3, 4, 5, 6]); // GH #291 S2 / S2x
  assert.equal(s.additionalProperties, false);
  assert.match(D.createDocxWriteTool({}).description, /split_table_cell/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

for (const [name, target, anchor, rows, cols] of [['NEW 1x2', 'GH255 r1c1', 'GH255 r0c0', 1, 2], ['NEW 2x1', 'GH255 r1c1', 'GH255 r0c0', 2, 1], ['NEW 2x2', 'GH255 r0c0', 'GH255 r2c2', 2, 2], ['NEW 1x3', 'GH255 r1c0', 'GH255 r0c2', 1, 3]]) {
  test(`READ SYMMETRY (${name} on a plain cell): the non-empty cells keep their TEXT and READING ORDER, the target is there once, the table GREW with empty cells, every other table and the text outside unchanged; L1 + L3 PASS`, async () => {
    const p = product();
    const b = await p.readTables();
    const outside = await p.readText();
    const r = await p.tool.handler(req([sp(target, anchor, rows, cols)]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(p.writes(), 1);
    const a = await p.readTables();
    assert.deepEqual(nonEmpty(a[0].rows), nonEmpty(b[0].rows), 'content and order unchanged');
    assert.equal(a[0].rows.flat().filter(c => c.text === target).length, 1);
    assert.ok(count(a[0].rows) > count(b[0].rows), 'the table grew');
    assert.ok(a[0].rows.flat().filter(c => c.text === '').length >= rows * cols - 1, 'the new cells are empty');
    assert.deepEqual(a.slice(1), b.slice(1), 'the other tables are untouched');
    assert.deepEqual(await p.readText(), outside);
  });
}

test('RESTORE 1x2 of a horizontally merged owner (TWO paragraphs): the row regains its cell, ALL the content stays in the LEFT cell, the stored owner has no gridSpan; the other rows / tables unchanged', async () => {
  const p = product();
  const b = await p.readTables();
  const r = await p.tool.handler(req([sp(MERGED, 'GH255 F0', 1, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const a = await p.readTables();
  assert.equal(a[2].rows[1].length, b[2].rows[1].length + 1);
  assert.equal(a[2].rows[1][0].text, MERGED);
  assert.equal(a[2].rows[1][1].text, '');
  assert.equal(a[2].rows[1][2].text, 'GH255 H1');
  assert.equal(a[2].rows[1][0].paraFormats.length, 2, 'both paragraphs stay in the left cell');
  assert.deepEqual(a[2].rows[0], b[2].rows[0]); assert.deepEqual(a[2].rows[2], b[2].rows[2]);
  assert.deepEqual([a[0], a[1], a[3]], [b[0], b[1], b[3]]);
  assert.equal(p.store.durableDoc.tables[2].rows[1].every(c => (c.span || 1) === 1), true, 'no cell of the row keeps a gridSpan');
});

test('RESTORE 2x1 of a vertical merge: the stored owner loses its rowspan, the covered cell is a plain cell again; the rows keep their cell count', async () => {
  const p = product();
  const b = await p.readTables();
  const r = await p.tool.handler(req([sp('GH255 VTop', 'GH255 V0', 2, 1)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const a = await p.readTables();
  assert.deepEqual(a[3].rows.map(row => row.length), b[3].rows.map(row => row.length));
  assert.deepEqual(a[3].rows.map(row => row[0].text), ['GH255 V0', 'GH255 VTop', '']);
  const d = p.store.durableDoc.tables[3].rows;
  assert.equal(d.flat().some(c => c.vContinue || (c.rowspan || 1) > 1), false, 'no vertical merge is left in the saved document');
  assert.deepEqual([a[0], a[1], a[2]], [b[0], b[1], b[2]]);
});

test('RESTORE: the requested shape must be the owner\'s span -- 2x1 / 1x3 / 2x2 on a 1x2 owner, 1x2 on a 2x1 owner are FAILED (split-shape-is-not-the-cell-span), 0 writes, the document unchanged', async () => {
  for (const [text, anchor, rows, cols] of [[MERGED, 'GH255 F0', 2, 1], [MERGED, 'GH255 F0', 1, 3], [MERGED, 'GH255 F0', 2, 2], ['GH255 VTop', 'GH255 V0', 1, 2], ['GH255 VTop', 'GH255 V0', 2, 2]]) {
    const p = product();
    const snap = durable(p.store);
    const r = await p.tool.handler(req([sp(text, anchor, rows, cols)]));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.match(r.reason, /split-shape-is-not-the-cell-span/);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
});

test('RETRY: the same task again is SATISFIED / noOp / 0 writes (the target is 1x1 at the top-left of an empty block) -- after a NEW split and after a RESTORE', async () => {
  for (const [step] of [[sp('GH255 r1c1', 'GH255 r0c0', 1, 2)], [sp(MERGED, 'GH255 F0', 1, 2)]]) {
    const p = product();
    assert.equal((await p.tool.handler(req([step]))).ok, true);
    const saves = p.writes(); const snap = durable(p.store);
    const retry = await p.tool.handler(req([step]));
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
    assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
    assert.equal(p.writes(), saves); assert.equal(durable(p.store), snap);
  }
});

test('NAMED LIMIT (measured, the #251 / #253 / #254 survivor question): a plain cell that ALREADY has empty 1x1 cells to its right reads SATISFIED for a NEW 1x2 split -- the proof is the closed state, not that WE created the empty cell', async () => {
  const tabs = tables(); tabs[0].rows.forEach(row => row.splice(2, 0, cell('')));          // a column of empty cells: r1c1 is followed by an empty cell already (an uneven row would be uneven-logical-width)
  const p = product({}, tabs);
  const snap = durable(p.store);
  const r = await p.tool.handler(req([sp('GH255 r1c1', 'GH255 r0c0', 1, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), 0); assert.equal(durable(p.store), snap);
});

test('TABLE IDENTITY: the anchor decides the table; a target text two tables hold is split in the ANCHORED table only; a missing / ambiguous anchor and a missing target are FAILED with 0 writes', async () => {
  const p = product();
  const before = await p.readTables();
  const r = await p.tool.handler(req([sp('GH255 r1c1', 'GH255 r0c0', 1, 2)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.deepEqual(after[1], before[1], 'the twin in the other table is untouched');
  const dupAnchor = [{ rows: [[cell('GH255 Anchor'), cell('GH255 X')], [cell('GH255 Y'), cell('GH255 Z')]] }, { rows: [[cell('GH255 Anchor'), cell('GH255 K')]] }];
  for (const [tabs, steps, why] of [
    [tables(), [sp('GH255 r1c1', 'GH255 No Such', 1, 2)], /table-anchor-not-found/], [dupAnchor, [sp('GH255 Z', 'GH255 Anchor', 1, 2)], /table-anchor-ambiguous/],
    [tables(), [sp('GH255 Not There', 'GH255 r0c0', 1, 2)], /split-target-not-found/],
  ]) {
    const q = product({}, tabs);
    const snap = durable(q.store);
    const x = await q.tool.handler(req(steps));
    assert.equal(x.ok, false, JSON.stringify(x));
    assert.match(x.reason, why, x.reason);
    assert.equal(q.writes(), 0);
    assert.equal(durable(q.store), snap);
  }
});

test('FAIL-CLOSED editor faults (a multi-paragraph merged owner): a silent no-op, joined paragraphs, a re-formatted paragraph, content in the wrong cell, a wrong grid -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'split-join-paragraphs', 'split-reformat', 'split-content-wrong-cell', 'split-wrong-grid']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([sp(MERGED, 'GH255 F0', 1, 2)]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
});
