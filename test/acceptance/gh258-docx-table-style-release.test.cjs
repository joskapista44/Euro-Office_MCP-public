'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #258 (DOCX P1 WRITE): format_table with the `style` field ONLY, served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH258-TABLE-STYLE-BY-NAME-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs (its table-style model
 * follows W04-4: the name resolves against the document's table-style list); the READ side is the product docx_read read_tables (format.style = { id, name }) on the SAME store. The LIVE acceptance is
 * test/live/docx-table-style-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH258 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
const grid = (tag, rows, cols) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(`${tag} r${r}c${c}`))) });
// T0 3 x 3; T1 shares 'GH258 r1c1' with T0; T2 is a third plain table
const tables = () => [grid('GH258', 3, 3), { rows: [[cell('GH258 x0'), cell('GH258 x1')], [cell('GH258 y0'), cell('GH258 r1c1')]] }, grid('GH258C', 2, 2)];
const STYLES = [{ name: 'Table Grid', styleId: 701 }, { name: 'Light List', styleId: 702 }];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const ft = (text, tableAnchor, style, extra = {}) => ({ intent: 'format_table', table: tc(text, tableAnchor), style, ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

function product(script = {}, tabs = tables(), styles = STYLES) {
  const store = createStore(doc(), { doc: { tables: tabs, tableStyles: styles } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true, ...script });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  const readText = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_text' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.blocks.filter(b => b.kind !== 'table').map(b => b.text); };
  return { store, tool, readTables, readText, writes: () => store.saves };
}
const durable = store => JSON.stringify(store.durableDoc.tables);
const styleOf = t => t.format.style;

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: table { tableCell: { text [, tableAnchor] } } plus a style NAME is accepted, with an optional step id', () => {
  for (const steps of [[ft('GH258 r0c0', undefined, 'Table Grid')], [ft('GH258 r1c1', 'GH258 r0c0', 'Light List')], [{ ...ft('GH258C r0c0', undefined, 'Table Grid'), id: 'f1' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: a missing table / style, blank texts, an empty style, a non-string style are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = ft('GH258 r0c0', undefined, 'Table Grid');
  const bad = [{ ...ok, table: undefined }, { ...ok, style: undefined }, { ...ok, style: '' }, { ...ok, style: '   ' }, { ...ok, style: 7 }, { ...ok, style: null }, { ...ok, table: {} }, { ...ok, table: { tableCell: {} } },
    ft('   ', undefined, 'Table Grid'), ft('', undefined, 'Table Grid'), ft('GH258 r0c0', '   ', 'Table Grid'), { intent: 'format_table', style: 'Table Grid' }, { intent: 'format_table', table: tc('GH258 r0c0') }];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: every other field / form is UNKNOWN_FIELD (0 dispatch) -- the bare { text } table form, selector, position, dependsOn, an extra tableCell key (borders / shading / cellMargins / indent / width joined this schema in the release-3 merge)', () => {
  const base = ft('GH258 r0c0', undefined, 'Table Grid');
  const border = { style: 'single', size: 4, color: '#FF0000' };
  for (const extra of [{ table: { text: 'GH258 r0c0' } },
    { table: { tableCell: { text: 'GH258 r0c0', column: 1 } } }, { selector: tc('GH258 r0c0') }, { position: 'last' }, { dependsOn: ['s0'] }, { text: 'x' }]) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
  for (const axes of [{ shading: '#FF0000' }, { borders: { top: border } }, { cellMargins: { top: 1, right: 1, bottom: 1, left: 1 } }, { indent: 100 }, { width: 9000 }]) {
    assert.equal(D.validateDocxWriteRequest(req([{ ...base, ...axes }])).ok, true, `${JSON.stringify(axes)} together with a style is ONE step (the merged union)`);
    assert.equal(D.validateDocxWriteRequest(req([{ intent: 'format_table', table: tc('GH258 r0c0'), ...axes }])).ok, true, `${JSON.stringify(axes)} alone (style-less) is the #256/#257 form`);
  }
});

test('SCHEMA: a format_table task is that ONE step -- a second step (another table, a merge, a rewrite, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const one = ft('GH258 r0c0', undefined, 'Table Grid');
  const two = ft('GH258C r0c0', undefined, 'Light List');
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  const split = { intent: 'split_table_cell', selector: tc('GH258 r1c0', 'GH258 r0c0'), splitRows: 1, splitCols: 2 };
  for (const steps of [[one, two], [one, split], [split, one], [rw, one], [one, rw]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid format_table (the other table fields and the styles are not released here)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([ft('GH258 r0c0', undefined, 'Table Grid'), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the format_table step: table.tableCell (text required, tableAnchor optional) and style OPTIONAL (the union with the #256/#257 axes after the release-3 merge)', () => {
  const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === 'format_table');
  assert.deepEqual(s.required, ['intent', 'table']);
  assert.deepEqual(Object.keys(s.properties).sort(), ['borders', 'cellMargins', 'id', 'indent', 'intent', 'shading', 'style', 'table', 'width']);
  assert.deepEqual(s.properties.table.required, ['tableCell']);
  assert.deepEqual(s.properties.table.properties.tableCell.required, ['text']);
  assert.deepEqual(Object.keys(s.properties.table.properties.tableCell.properties).sort(), ['tableAnchor', 'text']);
  assert.equal(s.properties.style.type, 'string');
  assert.equal(s.additionalProperties, false);
  assert.match(D.createDocxWriteTool({}).description, /format_table/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

test('READ SYMMETRY (style by name): the table reads { id, name } of the style, every cell text, every other table (style AND content) and the text outside unchanged; L1 + L3 PASS; one save', async () => {
  const p = product();
  const b = await p.readTables();
  const outside = await p.readText();
  const r = await p.tool.handler(req([ft('GH258 r0c0', undefined, 'Table Grid')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes(), 1);
  const a = await p.readTables();
  assert.deepEqual(styleOf(a[0]), { id: '701', name: 'Table Grid' });
  assert.notEqual(styleOf(b[0]).name, 'Table Grid', 'the round proves something: the table was not styled before');
  assert.deepEqual(a[0].rows, b[0].rows, 'every cell of the table unchanged');
  assert.deepEqual({ ...a[0].format, style: null }, { ...b[0].format, style: null }, 'every OTHER field of the table format unchanged');
  assert.deepEqual(a.slice(1), b.slice(1), 'the other tables are untouched');
  assert.deepEqual(await p.readText(), outside);
});

test('RESTYLE: a table that already has one style takes another by name (Table Grid -> Light List); the first style is replaced, not stacked', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH258 r0c0', undefined, 'Table Grid')]))).ok, true);
  const r = await p.tool.handler(req([ft('GH258 r0c0', undefined, 'Light List')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.deepEqual(styleOf((await p.readTables())[0]), { id: '702', name: 'Light List' });
  assert.equal(p.writes(), 2);
});

test('RETRY: the same task again is SATISFIED / noOp / 0 writes (the table already carries the style)', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH258 r0c0', undefined, 'Table Grid')]))).ok, true);
  const saves = p.writes(); const snap = durable(p.store);
  const retry = await p.tool.handler(req([ft('GH258 r0c0', undefined, 'Table Grid')]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves); assert.equal(durable(p.store), snap);
});

test('NAMED LIMIT (the #251-#255 survivor question): a table that ALREADY carries the style reads SATISFIED -- the proof is the closed state, not that WE set it', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH258C r0c0', undefined, 'Light List')]))).ok, true);
  const saves = p.writes();
  const again = await p.tool.handler(req([ft('GH258C r1c1', undefined, 'Light List')]));       // ANOTHER cell of the same table names it
  assert.equal(again.noOp, true);
  assert.deepEqual(again.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
});

test('STYLE RESOLUTION (fail closed, 0 writes, the document unchanged): an unknown name, the right name in the WRONG CASE (the match is exact), two styles with the same name, an unreadable style list', async () => {
  for (const [script, styles, name, why] of [
    [{}, STYLES, 'No Such Style', /table-style-not-found/], [{}, STYLES, 'table grid', /table-style-not-found/], [{}, STYLES, 'Table Grid ', /table-style-not-found|table-style-ambiguous/],
    [{}, [...STYLES, { name: 'Light List', styleId: 703 }], 'Light List', /table-style-ambiguous/], [{ tableStylesUnreadable: true }, STYLES, 'Table Grid', /table-style-list-unobservable|unobservable/],
  ]) {
    const p = product(script, tables(), styles);
    const snap = durable(p.store);
    const r = await p.tool.handler(req([ft('GH258 r0c0', undefined, name)]));
    assert.equal(r.ok, false, `${name}: ${JSON.stringify(r)}`);
    assert.match(r.reason, why, r.reason);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
});

test('TABLE IDENTITY: the anchor decides the table; a cell text two tables hold styles the ANCHORED table only; a missing / ambiguous cell is FAILED with 0 writes', async () => {
  const p = product();
  const before = await p.readTables();
  const r = await p.tool.handler(req([ft('GH258 r1c1', 'GH258 r0c0', 'Table Grid')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.equal(styleOf(after[0]).name, 'Table Grid');
  assert.deepEqual(after[1], before[1], 'the twin table is untouched');
  const bare = product();
  const amb = await bare.tool.handler(req([ft('GH258 r1c1', undefined, 'Table Grid')]));
  assert.equal(amb.ok, false, 'the bare twin text names two tables');
  assert.match(amb.reason, /table-anchor-ambiguous/);
  assert.equal(bare.writes(), 0);
  const miss = product();
  const m = await miss.tool.handler(req([ft('GH258 Not There', undefined, 'Table Grid')]));
  assert.equal(m.ok, false);
  assert.match(m.reason, /table-anchor-not-found/);
  assert.equal(miss.writes(), 0);
});

test('vertical merge (#270 G4/G8, g270v3 G9; GH #291): a merge-free target next to a merged neighbour applies; the merged table itself is now styled too, the merge kept', async () => {
  const withMerge = () => { const tabs = tables(); tabs.push({ rows: [[{ span: 1, rowspan: 2, paras: [{ text: 'GH258 VTop', props: { ...PLAIN } }] }, cell('GH258 W0')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH258 W1')]] }); return tabs; };
  const ok = product({}, withMerge());
  const r = await ok.tool.handler(req([ft('GH258 r0c0', undefined, 'Table Grid')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.operations[0].status, 'APPLIED');
  const p = product({}, withMerge());
  const m = await p.tool.handler(req([ft('GH258 VTop', undefined, 'Table Grid')]));
  assert.equal(m.ok, true, JSON.stringify(m));
  assert.equal(m.operations[0].status, 'APPLIED');
  const merged = p.store.durableDoc.tables[p.store.durableDoc.tables.length - 1];
  assert.deepEqual([merged.rows[0][0].rowspan, merged.rows[1][0].vContinue], [2, true]);
});

test('FAIL-CLOSED editor faults: a silent no-op, ANOTHER field of the table changed, ANOTHER table changed, the body text changed -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'tf-other-field-changed', 'tf-other-table', 'tf-body-change']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([ft('GH258 r0c0', undefined, 'Table Grid')]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
});
