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
 * GH #252 (DOCX P1 WRITE): insert_table_column_before / insert_table_column_after served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH252-TABLE-COLUMN-INSERT-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs
 * (its column model follows the Wave 4c column probe 1); the READ side is the product docx_read read_tables on the SAME store. The LIVE acceptance is
 * test/live/docx-table-column-insert-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH252 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
// T0: 3 x 3; T1 shares 'GH252 B1' with T0 (2 x 2); T2 has a horizontal merge in its last row (cols 1-2); T3 has a vertical merge
const tables = () => [
  { rows: [[cell('GH252 A0'), cell('GH252 B0'), cell('GH252 C0')], [cell('GH252 A1'), cell('GH252 B1'), cell('GH252 C1')], [cell('GH252 A2'), cell('GH252 B2'), cell('GH252 C2')]] },
  { rows: [[cell('GH252 D0'), cell('GH252 E0')], [cell('GH252 D1'), cell('GH252 B1')]] },
  { rows: [[cell('GH252 F0'), cell('GH252 G0'), cell('GH252 H0')], [cell('GH252 F1'), { ...cell('GH252 M1'), span: 2 }]] },
  { rows: [[cell('GH252 V0'), cell('GH252 W0')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH252 VTop', props: { ...PLAIN } }] }, cell('GH252 W1')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH252 W2')]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const ins = (intent, text, n, tableAnchor, extra = {}) => ({ intent, selector: tc(text, tableAnchor), terminalColumnCount: n, ...extra });
const before = (text, n, a, e) => ins('insert_table_column_before', text, n, a, e);
const after = (text, n, a, e) => ins('insert_table_column_after', text, n, a, e);
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

function product(script = {}, tabs = tables()) {
  const store = createStore(doc(), { doc: { tables: tabs } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  const readText = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_text' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.blocks.filter(b => b.kind !== 'table').map(b => b.text); };
  return { store, tool, readTables, readText, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
const durable = store => JSON.stringify(store.durableDoc.tables);

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: { tableCell: { text [, tableAnchor] } } plus terminalColumnCount is accepted for both intents, with an optional step id', () => {
  for (const steps of [[before('GH252 B0', 4)], [after('GH252 B0', 4, 'GH252 A2')], [{ ...after('GH252 C0', 4), id: 'c1' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: a missing / invalid terminalColumnCount, a blank text, an anchor equal to the target cell are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = after('GH252 B0', 4);
  const bad = [
    { ...ok, terminalColumnCount: undefined }, { ...ok, terminalColumnCount: 1 }, { ...ok, terminalColumnCount: 0 }, { ...ok, terminalColumnCount: -1 }, { ...ok, terminalColumnCount: 1001 }, { ...ok, terminalColumnCount: 2.5 }, { ...ok, terminalColumnCount: '4' },
    { intent: 'insert_table_column_after', terminalColumnCount: 4 }, { ...ok, selector: {} }, { ...ok, selector: { tableCell: { tableAnchor: 'GH252 A0' } } },
    after('   ', 4), after('', 4), after('GH252 B0', 4, 'GH252 B0'), after('GH252 B0', 4, '   '),
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: every other selector form / alias / extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } form, terminalTableAnchor, kind, position, a text, row / column hints', () => {
  const base = after('GH252 B0', 4);
  for (const extra of [{ selector: { text: 'GH252 B0' } }, { selector: { ...base.selector, kind: 'paragraph' } }, { selector: { tableCell: { text: 'GH252 B0', tableAnchor: 'GH252 A0', column: 1 } } },
    { terminalTableAnchor: 'GH252 A0' }, { position: 'last' }, { text: 'x' }, { dependsOn: ['s0'] }, { before: 'x' }, { terminalRowCount: 3 }]) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a column insert task is that ONE step -- a second step (another column, another table, a row step, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  const row = { intent: 'insert_table_row_after', selector: tc('GH252 B0'), terminalRowCount: 4 };
  for (const steps of [[after('GH252 B0', 4), before('GH252 B0', 5)], [after('GH252 B0', 4), after('GH252 D0', 3)], [after('GH252 B0', 4), row], [rw, after('GH252 B0', 4)]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): an insert whose successor fails leaves the FIRST column inserted and saved', async () => {
  const store = createStore(doc(), { doc: { tables: tables() } });
  const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
  const r = await exec({ fileId: FILE, task: { steps: [after('GH252 B0', 4, 'GH252 A2'), after('GH252 Not There', 5, 'GH252 A2')] } });
  assert.equal(r.ok, false);
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(store.saves, 1, 'the first insert was saved next to the failed second step');
  assert.equal(store.durableDoc.tables[0].rows[0].length, 4);
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid column insert (merge / split #254-#255)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([after('GH252 B0', 4), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries both column steps: selector.tableCell (text required), terminalColumnCount (integer 2..1000) required, nothing else', () => {
  for (const intent of ['insert_table_column_before', 'insert_table_column_after']) {
    const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === intent);
    assert.deepEqual(s.required, ['intent', 'selector', 'terminalColumnCount']);
    assert.deepEqual(Object.keys(s.properties).sort(), ['id', 'intent', 'selector', 'terminalColumnCount']);
    assert.deepEqual(s.properties.terminalColumnCount, { type: 'integer', minimum: 2, maximum: 1000 });
    assert.deepEqual(s.properties.selector.required, ['tableCell']);
    assert.deepEqual(s.properties.selector.properties.tableCell.required, ['text']);
    assert.equal(s.additionalProperties, false);
  }
  assert.match(D.createDocxWriteTool({}).description, /insert_table_column_before/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

const texts = rows => rows.map(r => r.map(c => c.text));
for (const [name, mk, target, col] of [['before the first column', before, 'GH252 A1', 0], ['after the first column', after, 'GH252 A1', 1], ['before the middle column', before, 'GH252 B1', 1], ['after the last column', after, 'GH252 C1', 3]]) {
  test(`READ SYMMETRY (${name}): width 3 -> 4, one EMPTY cell at the boundary in EVERY row, every other cell in order, every other table and the text outside read unchanged; L1 + L3 PASS`, async () => {
    const p = product();
    const b = await p.readTables();
    const outside = await p.readText();
    const r = await p.tool.handler(req([mk(target, 4, 'GH252 A2' === target ? 'GH252 A0' : 'GH252 A2')]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(p.writes(), 1);
    const a = await p.readTables();
    const want = texts(b[0].rows).map(row => { const x = row.slice(); x.splice(col, 0, ''); return x; });
    assert.deepEqual(texts(a[0].rows), want, 'one empty cell at the declared boundary in every row, the rest in order');
    assert.deepEqual(a.slice(1), b.slice(1), 'the other tables are untouched');
    assert.deepEqual(await p.readText(), outside);
  });
}

test('RETRY: the same task again is SATISFIED / noOp / 0 writes (the declared side holds an empty column and the width is terminal)', async () => {
  for (const mk of [before, after]) {
    const p = product();
    assert.equal((await p.tool.handler(req([mk('GH252 B1', 4, 'GH252 A2')]))).ok, true);
    const saves = p.writes();
    const retry = await p.tool.handler(req([mk('GH252 B1', 4, 'GH252 A2')]));
    assert.equal(retry.ok, true, JSON.stringify(retry));
    assert.equal(retry.noOp, true);
    assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
    assert.equal(p.writes(), saves);
  }
});

test('NAMED LIMIT (measured): the SATISFIED proof is "the declared-side column is EMPTY and the width is terminal" -- NOT provenance; an empty column that was already there reads SATISFIED, 0 writes', async () => {
  const tabs = tables(); for (const row of tabs[0].rows) row.splice(2, 0, cell(''));   // an empty column already sits after B (width 4)
  const p = product({}, tabs);
  const snap = durable(p.store);
  const r = await p.tool.handler(req([after('GH252 B1', 4, 'GH252 A2')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), 0);
  assert.equal(durable(p.store), snap);
});

test('TABLE IDENTITY: the optional anchor decides the table; a target text two tables hold is ambiguous without it and applies in the ANCHORED table only; a missing target / anchor, a wrong count are FAILED with 0 writes', async () => {
  const amb = product();
  const snap = durable(amb.store);
  const a = await amb.tool.handler(req([after('GH252 B1', 4)]));
  assert.equal(a.ok, false, JSON.stringify(a));
  assert.match(a.reason, /table-column-target-ambiguous/);
  assert.equal(amb.writes(), 0);
  assert.equal(durable(amb.store), snap);
  const p = product();
  const b = await p.readTables();
  const r = await p.tool.handler(req([after('GH252 B1', 4, 'GH252 A2')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const af = await p.readTables();
  assert.equal(af[0].rows[0].length, 4);
  assert.deepEqual(af[1], b[1], 'the twin in the other table is untouched');
  for (const [steps, why] of [
    [[after('GH252 Not There', 4)], /table-column-target-not-found/],
    [[after('GH252 B1', 4, 'GH252 No Such')], /table-anchor-not-found/],
    [[after('GH252 B0', 3, 'GH252 A2')], /table-column-terminal-unprovable/], [[after('GH252 B0', 5, 'GH252 A2')], /table-column-terminal-unprovable/],
  ]) {
    const q = product();
    const s = durable(q.store);
    const x = await q.tool.handler(req(steps));
    assert.equal(x.ok, false, JSON.stringify(x));
    assert.match(x.reason, why);
    assert.equal(x.operations[0].status, 'FAILED');
    assert.equal(q.writes(), 0);
    assert.equal(durable(q.store), s);
  }
});

test('COLSPAN (fail-closed boundary): an insert boundary a colspan straddles is FAILED before any write; a boundary at the edge of the span applies', async () => {
  for (const [steps, why] of [
    [[after('GH252 G0', 4, 'GH252 F1')], /colspan-straddles-insert-boundary/],
    [[before('GH252 H0', 4, 'GH252 F1')], /colspan-straddles-insert-boundary/],
  ]) {
    const p = product();
    const snap = durable(p.store);
    const r = await p.tool.handler(req(steps));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.match(r.reason, why);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
});

test('COLSPAN TARGET (GH #291 H3, LIVE gh291-h-colspan-probe-20261022.json): an insert after / before a target that spans columns is APPLIED at the span\'s own edge in every row, the span kept', async () => {
  for (const [step, row0, row1] of [
    [after('GH252 M1', 4), ['GH252 F0', 'GH252 G0', 'GH252 H0', ''], ['GH252 F1', 'GH252 M1', '']],
    [before('GH252 M1', 4), ['GH252 F0', '', 'GH252 G0', 'GH252 H0'], ['GH252 F1', '', 'GH252 M1']],
  ]) {
    const p = product();
    const r = await p.tool.handler(req([step]));
    assert.equal(r.ok, true, JSON.stringify(r));
    const t = p.store.durableDoc.tables.find(x => x.rows.flat().some(c => c.paras[0].text === 'GH252 M1')).rows;
    assert.deepEqual(t.map(row => row.map(c => c.paras.map(x => x.text).join(' '))), [row0, row1]);
    assert.equal(t[1].find(c => c.paras[0].text === 'GH252 M1').span, 2);
  }
});

test('VERTICAL MERGE (GH #291 T-04): an insert in a table that holds a vertical merge is APPLIED, the merge kept on its own column; a vertical merge in ANOTHER table is no obstacle', async () => {
  const p = product();
  const r = await p.tool.handler(req([after('GH252 V0', 3)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const vt = p.store.durableDoc.tables.find(t => t.rows.flat().some(c => c.paras[0].text === 'GH252 VTop')).rows;
  assert.deepEqual([vt[1][0].rowspan, vt[2][0].vContinue, vt.flat().filter(c => c.vContinue).length, vt[0].length], [2, true, 1, 3]);
  assert.equal((await product().tool.handler(req([after('GH252 A1', 4, 'GH252 A2')]))).ok, true);
});

test('FAIL-CLOSED editor faults: a silent no-op, a stale column index, the wrong side, two columns, a wide / non-empty new cell, a reordered row, another cell / table / outside text touched -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'col-stale-index', 'col-wrong-side', 'col-twice', 'col-new-span2', 'col-new-not-empty', 'col-reorder', 'col-touch-cell', 'col-touch-other-table', 'col-change-outside']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([after('GH252 B1', 4, 'GH252 A2')]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
});

test('PRE-DISPATCH refusals never open a session: a refused column-insert task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[after('GH252 B0', 1)], [{ intent: 'insert_table_column_after', selector: { text: 'GH252 B0' }, terminalColumnCount: 4 }], [after('GH252 B0', 4, 'GH252 B0')], [after('GH252 B0', 4), before('GH252 B0', 5)]]) {
    const r = await p.tool.handler(req(steps));
    assert.equal(r.dispatched, false);
  }
  assert.equal(p.opens(), 0);
});

// ---- the release gate ------------------------------------------------------------------------------------------------------

test('GATE: the DOCX allowlist is exactly the accepted set (both halves agree); a one-sided change withholds only docx_execute_task', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('insert_table_column_before') && W.DOCX_WRITE_ALLOWLIST.includes('insert_table_column_after'));
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  const pre = [...W.DOCX_WRITE_ALLOWLIST].filter(i => !i.startsWith('insert_table_column_'));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#252 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  for (const i of ['insert_table_column_before', 'insert_table_column_after']) assert.ok(T.capabilityInventory().some(r => r.TOOL === `docx:${i}` && r.EXPOSE_IN_PRODUCT_MCP === 'YES'), i);
});
