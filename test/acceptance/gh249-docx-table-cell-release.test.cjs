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
 * GH #249 (DOCX P1 WRITE): replace_table_cell served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH249-TABLE-CELL-REPLACE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on
 * test/fakes/fake-docx-driver.cjs (its table model follows the Wave 4a / W06-3 LIVE probes: the driver's public table read, a merged cell colspan,
 * a vertical merge that the HTML projection renders as rowspan + an omitted continuation cell); the READ side is the product docx_read read_tables
 * on the SAME store. The LIVE acceptance is test/live/docx-table-cell-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Courier New', fontSize: 28, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const richCell = text => ({ span: 1, paras: [{ text, props: { ...RICH }, alignment: '2', style: 'Table Text' }] });
const OUTSIDE = 'GH249 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }, { text: 'GH249 Twin', props: PLAIN }];
// T0 PLAIN (3 x 2, one merged two-paragraph cell), T1 TWINS (shares the text 'GH249 Twin' with T0), T2 VMERGE (a vertical merge), T3 NESTED (a cell that is not plain paragraphs of runs)
const tables = () => [
  { rows: [[cell('GH249 Name'), cell('GH249 Qty')], [richCell('GH249 Rich'), cell('GH249 Twin')], [{ ...cell('GH249 M1', 'GH249 M2'), span: 2 }]] },
  { rows: [[cell('GH249 Item'), cell('GH249 Price')], [cell('GH249 Pear'), cell('GH249 Twin')]] },
  { rows: [[cell('GH249 VA'), cell('GH249 VB')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH249 VTop', props: { ...PLAIN } }] }, cell('GH249 V5')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH249 V7')]] },
  { rows: [[cell('GH249 NA'), { ...cell('GH249 Nested'), nonPlain: true }]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const rc = (text, newText, tableAnchor) => ({ intent: 'replace_table_cell', selector: tc(text, tableAnchor), text: newText });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

function product(script = {}, tabs = tables(), paragraphs = doc()) {
  const store = createStore(paragraphs, { doc: { tables: tabs } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  const readText = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_text' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.blocks.filter(b => b.kind !== 'table').map(b => b.text); };
  return { store, tool, readTables, readText, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
/** every cell of every table, with its place: { 'ti:ri:ci': cell } */
const cellsOf = tabs => Object.fromEntries(tabs.flatMap((t, ti) => t.rows.flatMap((row, ri) => row.map((c, ci) => [`${ti}:${ri}:${ci}`, c]))));
const shapeOf = tabs => JSON.stringify(tabs.map(t => t.rows.map(r => r.length)));
const durable = store => JSON.stringify(store.durableDoc.tables);

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: the one canonical selector { tableCell: { text [, tableAnchor] } } is accepted, with an optional step id', () => {
  for (const steps of [[rc('GH249 Rich', 'X')], [rc('GH249 Rich', 'X', 'GH249 Name')], [{ ...rc('GH249 Rich', 'X'), id: 'c1' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: every other selector form and every extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } and combined forms, kind / precondition / heading / fingerprint / ref, physical row / column hints', () => {
  const sel = { tableCell: { text: 'GH249 Rich' } };
  const bad = [
    { selector: { text: 'GH249 Rich' } }, { selector: { ...sel, text: 'GH249 Rich' } }, { selector: { ...sel, kind: 'paragraph' } },
    { selector: { ...sel, precondition: { text: 'GH249 Rich' } } }, { selector: { ...sel, heading: 'Heading' } }, { selector: { ...sel, fingerprint: 'abc' } },
    { selector: { ...sel, ref: 's0' } }, { selector: { ...sel, before: 'x' } }, { selector: { ...sel, after: 'x' } },
    { selector: { tableCell: { text: 'GH249 Rich', row: 1 } } }, { selector: { tableCell: { text: 'GH249 Rich', column: 0 } } },
    { selector: { tableCell: { text: 'GH249 Rich', index: 0 } } }, { selector: { tableCell: { text: 'GH249 Rich', contains: true } } },
    { selector: { tableCell: { text: 'GH249 Rich', tableIndex: 0 } } },
    { selector: sel, before: 'x' }, { selector: sel, dependsOn: ['s0'] }, { selector: sel, expectedPreservation: { bold: true } },
  ];
  for (const extra of bad) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'replace_table_cell', text: 'X', ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a missing selector / tableCell / text, and blank, tab, line-break or unchanged texts, are REFUSED_SHAPE (0 dispatch); the anchor must be another cell that survives the replacement', () => {
  const bad = [
    { intent: 'replace_table_cell', text: 'X' }, { intent: 'replace_table_cell', selector: {}, text: 'X' }, { intent: 'replace_table_cell', selector: { tableCell: { tableAnchor: 'GH249 Name' } }, text: 'X' },
    { intent: 'replace_table_cell', selector: tc('GH249 Rich') },
    rc('   ', 'X'), rc('', 'X'), rc('GH249 Rich', '   '), rc('GH249 Rich', ''), rc('GH249 Rich', 'a\tb'), rc('GH249 Rich', 'a\nb'), rc('GH249 Rich', 'a\r\nb'),
    rc('GH249 Rich', 'GH249 Rich'), rc('GH249 Rich', '  GH249   Rich '),                   // the same after whitespace normalization
    rc('GH249 Rich', 'X', 'GH249 Rich'), rc('GH249 Rich', 'X', 'X'), rc('GH249 Rich', 'X', '   '),   // anchor == old text / new text / blank
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a replace_table_cell task is that ONE step -- any second step (another cell, the same cell again, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  for (const steps of [[rc('GH249 Rich', 'X'), rc('GH249 Pear', 'Y')], [rc('GH249 Rich', 'X'), rc('X', 'Z')], [rc('GH249 Rich', 'X'), rc('GH249 Rich', 'Z')], [rc('GH249 Rich', 'X'), rw], [rw, rc('GH249 Rich', 'X')]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): the same cell twice saves the first step and then FAILS the second; a chain on the same cell is APPLIED twice and the whole-task check fails AFTER the save', async () => {
  const run = async steps => {
    const store = createStore(doc(), { doc: { tables: tables() } });
    const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
    return { r: await exec({ fileId: FILE, task: { steps } }), store };
  };
  const twice = await run([rc('GH249 Pear', 'GH249 Plum'), rc('GH249 Pear', 'GH249 Fig')]);
  assert.equal(twice.r.ok, false);
  assert.deepEqual(twice.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(twice.store.saves, 1, 'the first step was saved next to the failed second one');
  assert.equal(twice.store.durableDoc.tables[1].rows[1][0].paras[0].text, 'GH249 Plum');
  const chain = await run([rc('GH249 Pear', 'GH249 Plum'), rc('GH249 Plum', 'GH249 Fig')]);
  assert.equal(chain.r.ok, false);
  assert.deepEqual(chain.r.operations.map(o => o.status), ['APPLIED', 'APPLIED']);
  assert.equal(chain.r.wholeTask, 'FAIL');
  assert.equal(chain.store.saves >= 1, true, 'saved, yet the receipt is ok:false');
  assert.equal(D.validateDocxWriteRequest(req([rc('GH249 Pear', 'GH249 Plum'), rc('GH249 Plum', 'GH249 Fig')])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid table-cell step (the row / column / merge families are #250-#255; format_table joined with GH #256)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  for (const intent of ['insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table']) {
    assert.equal(others.has(intent), false, `${intent} is released by the release-3 table chain`);
  }
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([rc('GH249 Rich', 'X'), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the table-cell step: selector.tableCell required, text required inside it, tableAnchor optional, nothing else', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  const s = items.oneOf.find(x => x.properties.intent.const === 'replace_table_cell');
  assert.deepEqual(s.required, ['intent', 'selector', 'text']);
  assert.deepEqual(Object.keys(s.properties).sort(), ['id', 'intent', 'selector', 'text']);
  assert.deepEqual(s.properties.selector.required, ['tableCell']);
  assert.deepEqual(Object.keys(s.properties.selector.properties), ['tableCell']);
  assert.deepEqual(s.properties.selector.properties.tableCell.required, ['text']);
  assert.deepEqual(Object.keys(s.properties.selector.properties.tableCell.properties).sort(), ['tableAnchor', 'text']);
  assert.equal(s.properties.selector.additionalProperties, false);
  assert.equal(s.properties.selector.properties.tableCell.additionalProperties, false);
  assert.match(D.createDocxWriteTool({}).description, /replace_table_cell/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

test('READ SYMMETRY: the named cell reads the new text, EVERY other cell (text / simple / format / paraFormats) and every table shape read unchanged; text outside the tables is unchanged; L1 + L3 PASS; retry SATISFIED / 0 writes', async () => {
  const p = product();
  const before = await p.readTables();
  const outside = await p.readText();
  const r = await p.tool.handler(req([rc('GH249 Pear', 'GH249 Plum')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes(), 1);
  const after = await p.readTables();
  assert.equal(shapeOf(after), shapeOf(before));
  const b = cellsOf(before); const a = cellsOf(after);
  assert.equal(a['1:1:0'].text, 'GH249 Plum');
  for (const k of Object.keys(b)) if (k !== '1:1:0') assert.deepEqual(a[k], b[k], `${k} untouched`);
  assert.deepEqual(a['1:1:0'].format, b['1:1:0'].format, 'the target keeps its format');
  assert.deepEqual(await p.readText(), outside);
  const retry = await p.tool.handler(req([rc('GH249 Pear', 'GH249 Plum')]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), 1, 'the retry saved nothing');
});

test('FORMAT: a formatted cell keeps bold / italic / font / size / style / alignment of its first paragraph (read_tables format + the durable model)', async () => {
  const p = product();
  const before = cellsOf(await p.readTables())['0:1:0'];
  assert.equal(before.format.bold, true);
  const r = await p.tool.handler(req([rc('GH249 Rich', 'GH249 Richer')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = cellsOf(await p.readTables())['0:1:0'];
  assert.equal(after.text, 'GH249 Richer');
  assert.deepEqual(after.format, before.format);
  const d = p.store.durableDoc.tables[0].rows[1][0].paras;
  assert.equal(d.length, 1);
  assert.deepEqual([d[0].props.bold, d[0].props.italic, d[0].props.fontFamily, d[0].props.fontSize, d[0].style, d[0].alignment], [true, true, 'Courier New', 28, 'Table Text', '2']);
});

test('MERGED CELL (horizontal): the cell is named by its paragraphs joined with a space; it becomes ONE paragraph and keeps its colspan; the table shape and every other cell are unchanged', async () => {
  const p = product();
  const before = await p.readTables();
  assert.equal(cellsOf(before)['0:2:0'].text, 'GH249 M1 GH249 M2');
  const r = await p.tool.handler(req([rc('GH249 M1 GH249 M2', 'GH249 Merged')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.equal(shapeOf(after), shapeOf(before));
  const a = cellsOf(after); const b = cellsOf(before);
  assert.equal(a['0:2:0'].text, 'GH249 Merged');
  assert.equal(a['0:2:0'].paraFormats.length, 1, 'a multi-paragraph cell collapses to ONE paragraph (named, measured LIVE in Wave 4a)');
  for (const k of Object.keys(b)) if (k !== '0:2:0') assert.deepEqual(a[k], b[k], `${k} untouched`);
  const d = p.store.durableDoc.tables[0].rows[2][0];
  assert.equal(d.span, 2, 'colspan kept');
  assert.equal(d.paras.length, 1);
});

test('TABLE IDENTITY: a cell text that two tables hold is AMBIGUOUS without an anchor (FAILED, 0 writes, nothing changed); with tableAnchor it applies in the named table ONLY; the other twin stays', async () => {
  const noAnchor = product();
  const snap = durable(noAnchor.store);
  const r0 = await noAnchor.tool.handler(req([rc('GH249 Twin', 'GH249 Twin Two')]));
  assert.equal(r0.ok, false, JSON.stringify(r0));
  assert.match(r0.reason, /table-cell-ambiguous/);
  assert.equal(noAnchor.writes(), 0);
  assert.equal(durable(noAnchor.store), snap);
  for (const [anchor, ti, ri, ci] of [['GH249 Item', 1, 1, 1], ['GH249 Qty', 0, 1, 1]]) {
    const p = product();
    const r = await p.tool.handler(req([rc('GH249 Twin', 'GH249 Twin Two', anchor)]));
    assert.equal(r.ok, true, `${anchor}: ${JSON.stringify(r)}`);
    const a = cellsOf(await p.readTables());
    assert.equal(a[`${ti}:${ri}:${ci}`].text, 'GH249 Twin Two');
    const other = ti === 1 ? '0:1:1' : '1:1:1';
    assert.equal(a[other].text, 'GH249 Twin', 'the twin in the other table is untouched');
    const retry = await p.tool.handler(req([rc('GH249 Twin', 'GH249 Twin Two', anchor)]));
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
  }
});

test('TABLE IDENTITY fail-closed: an anchor no table has, or that several tables have, FAILED with 0 writes; a text that is in no cell (a paragraph outside the tables) is not a cell', async () => {
  const dupAnchor = [{ rows: [[cell('GH249 Anchor'), cell('GH249 A1')]] }, { rows: [[cell('GH249 Anchor'), cell('GH249 B1')]] }];
  for (const [tabs, steps, why] of [
    [tables(), [rc('GH249 Pear', 'X', 'GH249 No Such Cell')], /table-anchor-not-found/],
    [dupAnchor, [rc('GH249 A1', 'X', 'GH249 Anchor')], /table-anchor-ambiguous/],
    [tables(), [rc('GH249 Not There', 'X')], /table-cell-terminal-unprovable/],
    [tables(), [rc(OUTSIDE, 'X')], /table-cell-terminal-unprovable/],
  ]) {
    const p = product({}, tabs);
    const snap = durable(p.store);
    const r = await p.tool.handler(req(steps));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.match(r.reason, why);
    assert.equal(r.operations[0].status, 'FAILED');
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
});

test('SHAPE fail-closed: a cell that is not plain paragraphs of runs is FAILED before any write; GH #291: a vertical merge in the TARGET table is no longer refused (owner and other cells apply); a nested cell in ANOTHER table is no obstacle', async () => {
  for (const text of ['GH249 V5', 'GH249 VTop']) {
    const p = product();
    const r = await p.tool.handler(req([rc(text, 'GH249 X')]));
    assert.equal(r.ok, true, `${text}: ${JSON.stringify(r)}`);
    assert.equal(r.operations[0].status, 'APPLIED');
  }
  for (const [text, why] of [['GH249 Nested', /table-cell-not-plain-paragraphs-of-runs/]]) {
    const p = product();
    const snap = durable(p.store);
    const r = await p.tool.handler(req([rc(text, 'X')]));
    assert.equal(r.ok, false, `${text}: ${JSON.stringify(r)}`);
    assert.match(r.reason, why);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
  const p = product();
  const r = await p.tool.handler(req([rc('GH249 Pear', 'GH249 Plum')]));     // tables 2 and 3 carry a vertical merge / a nested cell
  assert.equal(r.ok, true, JSON.stringify(r));
});

test('RETRY: SATISFIED (noOp, 0 writes) once the old text is gone and the new text is in a cell; old AND new present is FAILED (conflict), not a guess', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([rc('GH249 Pear', 'GH249 Plum')]))).ok, true);
  const conflict = product();
  const snap = durable(conflict.store);
  const r = await conflict.tool.handler(req([rc('GH249 Pear', 'GH249 Name')]));           // 'GH249 Name' is already a cell (another table, no anchor)
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.match(r.reason, /table-cell-conflicting-old-and-new-state/);
  assert.equal(conflict.writes(), 0);
  assert.equal(durable(conflict.store), snap);
  const scoped = product();
  const okScoped = await scoped.tool.handler(req([rc('GH249 Pear', 'GH249 Name', 'GH249 Item')]));   // with the anchor the other table's 'GH249 Name' is out of scope
  assert.equal(okScoped.ok, true, JSON.stringify(okScoped));
});

test('NAMED LIMIT (measured): without an anchor the retry check counts across ALL tables, so a new text that another table already holds reads SATISFIED when the old text is absent -- the post-state holds, the receipt says so, nothing is written', async () => {
  const p = product();
  const snap = durable(p.store);
  const r = await p.tool.handler(req([rc('GH249 Never Existed', 'GH249 Name')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), 0);
  assert.equal(durable(p.store), snap);
});

test('FAIL-CLOSED editor faults: a silent no-op, the wrong cell written, the append instead of the replacement, an extra paragraph, another cell touched, another cell\'s format changed, dropped text properties, a changed style, an unmerge, a kept second paragraph, an added row, text changed outside the table -- never APPLIED, nothing saved', async () => {
  // each fault runs on the cell where it BITES: dropping the text properties on the formatted cell, changing the style on a cell that is not already 'Table Text'
  for (const [fault, target] of [['ignore', 'GH249 Rich'], ['cell-wrong-cell', 'GH249 Rich'], ['cell-append', 'GH249 Rich'], ['cell-extra-empty-paragraph', 'GH249 Rich'], ['cell-touch-other', 'GH249 Rich'],
    ['cell-touch-other-format', 'GH249 Rich'], ['cell-drop-props', 'GH249 Rich'], ['cell-drop-style', 'GH249 Pear'], ['cell-add-row', 'GH249 Rich'], ['cell-change-outside', 'GH249 Rich']]) {
    const p = product({ fault });
    const r = await p.tool.handler(req([rc(target, `${target} replaced`)]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
  const m = product({ fault: 'cell-unmerge' });
  const r = await m.tool.handler(req([rc('GH249 M1 GH249 M2', 'GH249 Merged')]));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.equal(m.writes(), 0);
  const k = product({ fault: 'cell-keep-rest' });
  const r2 = await k.tool.handler(req([rc('GH249 M1 GH249 M2', 'GH249 Merged')]));
  assert.equal(r2.ok, false, JSON.stringify(r2));
  assert.equal(k.writes(), 0);
});

test('PRE-DISPATCH refusals never open a session: a refused table-cell task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[{ intent: 'replace_table_cell', selector: { text: 'GH249 Rich' }, text: 'X' }], [rc('GH249 Rich', 'a\tb')], [rc('GH249 Rich', 'GH249 Rich')], [rc('GH249 Rich', 'X'), rc('GH249 Pear', 'Y')], [rc('GH249 Rich', 'X', 'GH249 Rich')]]) {
    const r = await p.tool.handler(req(steps));
    assert.equal(r.dispatched, false);
  }
  assert.equal(p.opens(), 0);
});

// ---- the release gate ------------------------------------------------------------------------------------------------------

test('GATE: the DOCX allowlist is exactly the accepted set (both halves agree); a one-sided change withholds only docx_execute_task', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  assert.equal(W.DOCX_WRITE_ALLOWLIST.includes('replace_table_cell'), true);
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  const pre = [...W.DOCX_WRITE_ALLOWLIST].filter(i => i !== 'replace_table_cell');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#249 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  assert.ok(T.capabilityInventory().some(r => r.TOOL === 'docx:replace_table_cell' && r.EXPOSE_IN_PRODUCT_MCP === 'YES'));
});
