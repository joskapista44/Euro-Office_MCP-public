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
 * GH #253 (DOCX P1 WRITE): delete_table_column served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH253-TABLE-COLUMN-DELETE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs
 * (its column model follows the Wave 4c column probe 1 and the W-06/2 LIVE acceptance); the READ side is the product docx_read read_tables on the SAME store. The LIVE acceptance is
 * test/live/docx-table-column-delete-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH253 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
// T0 3 x 3; T1 shares 'GH253 B1' with T0; T2 has a colspan row (cols 1-2); T3 has a vertical merge over rows 1-2 (column 0)
const tables = () => [
  { rows: [[cell('GH253 A0'), cell('GH253 B0'), cell('GH253 C0')], [cell('GH253 A1'), cell('GH253 B1'), cell('GH253 C1')], [cell('GH253 A2'), cell('GH253 B2'), cell('GH253 C2')]] },
  { rows: [[cell('GH253 D0'), cell('GH253 E0')], [cell('GH253 D1'), cell('GH253 B1')]] },
  { rows: [[cell('GH253 F0'), cell('GH253 G0'), cell('GH253 H0')], [cell('GH253 F1'), { ...cell('GH253 M1'), span: 2 }]] },
  { rows: [[cell('GH253 V0'), cell('GH253 W0')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH253 VTop', props: { ...PLAIN } }] }, cell('GH253 W1')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH253 W2')]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const del = (text, n, tableAnchor, extra = {}) => ({ intent: 'delete_table_column', selector: tc(text, tableAnchor), terminalColumnCount: n, ...extra });
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

test('SCHEMA: { tableCell: { text, tableAnchor } } (the surviving anchor is REQUIRED) plus terminalColumnCount is accepted, with an optional step id', () => {
  for (const steps of [[del('GH253 B1', 2, 'GH253 A0')], [{ ...del('GH253 C0', 2, 'GH253 A2'), id: 'd1' }], [del('GH253 A1', 1, 'GH253 B0')]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: a missing anchor, a missing / invalid terminalColumnCount, a blank text, an anchor equal to the target cell are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = del('GH253 B1', 2, 'GH253 A0');
  const bad = [
    del('GH253 B1', 2), { ...ok, terminalColumnCount: undefined }, { ...ok, terminalColumnCount: 0 }, { ...ok, terminalColumnCount: -1 }, { ...ok, terminalColumnCount: 1001 }, { ...ok, terminalColumnCount: 2.5 }, { ...ok, terminalColumnCount: '2' },
    { intent: 'delete_table_column', terminalColumnCount: 2 }, { ...ok, selector: {} }, { ...ok, selector: { tableCell: { tableAnchor: 'GH253 A0' } } },
    del('   ', 2, 'GH253 A0'), del('', 2, 'GH253 A0'), del('GH253 B1', 2, 'GH253 B1'), del('GH253 B1', 2, '   '),
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: every other selector form / alias / extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } form, terminalTableAnchor, kind, position, a text, row / column hints, terminalRowCount', () => {
  const base = del('GH253 B1', 2, 'GH253 A0');
  for (const extra of [{ selector: { text: 'GH253 B1' } }, { selector: { ...base.selector, kind: 'paragraph' } }, { selector: { tableCell: { text: 'GH253 B1', tableAnchor: 'GH253 A0', column: 1 } } },
    { terminalTableAnchor: 'GH253 A0' }, { position: 'last' }, { text: 'x' }, { dependsOn: ['s0'] }, { before: 'x' }, { terminalRowCount: 3 }]) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a delete task is that ONE step -- a second step (another column, another table, an insert, a row step, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  const ins = { intent: 'insert_table_column_after', selector: tc('GH253 A0'), terminalColumnCount: 3 };
  const row = { intent: 'delete_table_row', selector: tc('GH253 A1', 'GH253 A0'), terminalRowCount: 2 };
  for (const steps of [[del('GH253 C0', 2, 'GH253 A0'), del('GH253 B0', 1, 'GH253 A0')], [del('GH253 C0', 2, 'GH253 A0'), del('GH253 E0', 1, 'GH253 D0')], [del('GH253 C0', 2, 'GH253 A0'), ins], [del('GH253 C0', 2, 'GH253 A0'), row], [rw, del('GH253 C0', 2, 'GH253 A0')]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): a delete whose successor fails leaves the FIRST column gone and saved; a later step that deletes the column holding the earlier step\'s anchor fails with table-anchor-not-found AFTER the save', async () => {
  const run = async steps => {
    const store = createStore(doc(), { doc: { tables: tables() } });
    const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
    return { r: await exec({ fileId: FILE, task: { steps } }), store };
  };
  const failSecond = await run([del('GH253 C0', 2, 'GH253 A0'), del('GH253 Not There', 1, 'GH253 A0')]);
  assert.equal(failSecond.r.ok, false);
  assert.deepEqual(failSecond.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(failSecond.store.saves, 1, 'the first delete was saved next to the failed second step');
  assert.equal(failSecond.store.durableDoc.tables[0].rows[0].length, 2);
  const anchorGone = await run([del('GH253 C0', 2, 'GH253 A0'), del('GH253 B0', 1, 'GH253 A0')]);
  assert.deepEqual(anchorGone.r.operations.map(o => o.status), ['APPLIED', 'APPLIED'], 'both applied on this shape: the SECOND step uses the surviving A0');
  const gone = await run([del('GH253 A0', 2, 'GH253 C0'), del('GH253 B0', 1, 'GH253 A0')]);
  assert.equal(gone.r.ok, false, JSON.stringify(gone.r));
  assert.deepEqual(gone.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.match(gone.r.reason, /table-anchor-not-found/);
  assert.equal(D.validateDocxWriteRequest(req([del('GH253 A0', 2, 'GH253 C0'), del('GH253 B0', 1, 'GH253 A0')])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid column delete (merge / split #254-#255)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([del('GH253 C0', 2, 'GH253 A0'), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the delete step: selector.tableCell with text AND tableAnchor required, terminalColumnCount (integer 1..1000) required, nothing else', () => {
  const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === 'delete_table_column');
  assert.deepEqual(s.required, ['intent', 'selector', 'terminalColumnCount']);
  assert.deepEqual(Object.keys(s.properties).sort(), ['id', 'intent', 'selector', 'terminalColumnCount']);
  assert.deepEqual(s.properties.terminalColumnCount, { type: 'integer', minimum: 1, maximum: 1000 });
  assert.deepEqual(s.properties.selector.required, ['tableCell']);
  assert.deepEqual(s.properties.selector.properties.tableCell.required, ['text', 'tableAnchor']);
  assert.equal(s.additionalProperties, false);
  assert.match(D.createDocxWriteTool({}).description, /delete_table_column/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

const texts = rows => rows.map(r => r.map(c => c.text));
for (const [name, target, anchor, col] of [['first column', 'GH253 A1', 'GH253 B0', 0], ['middle column', 'GH253 B1', 'GH253 A0', 1], ['last column', 'GH253 C1', 'GH253 A0', 2]]) {
  test(`READ SYMMETRY (delete the ${name}): width 3 -> 2, the target column is gone from EVERY row, every other cell in order, every other table and the text outside the tables read unchanged; L1 + L3 PASS`, async () => {
    const p = product();
    const b = await p.readTables();
    const outside = await p.readText();
    const r = await p.tool.handler(req([del(target, 2, anchor)]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(p.writes(), 1);
    const a = await p.readTables();
    assert.deepEqual(texts(a[0].rows), texts(b[0].rows).map(row => row.filter((_, k) => k !== col)), 'every surviving cell, in order, unchanged');
    assert.deepEqual(a[0].rows, b[0].rows.map(row => row.filter((_, k) => k !== col)), 'format / paraFormats of the survivors unchanged too');
    assert.deepEqual(a.slice(1), b.slice(1), 'the other tables are untouched');
    assert.deepEqual(await p.readText(), outside);
  });
}

test('RETRY: the same task again is SATISFIED / noOp / 0 writes (the target is absent from the anchor table and the width is terminal)', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([del('GH253 B1', 2, 'GH253 A0')]))).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(req([del('GH253 B1', 2, 'GH253 A0')]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
});

test('NAMED LIMIT (measured, the #251 survivor question): a column that was NEVER there reads SATISFIED when the anchor table exists and holds the terminal width -- the missing column alone is not evidence of OUR deletion; the proof is anchor + absence + width', async () => {
  const tabs = tables(); tabs[0].rows.forEach(row => row.splice(1, 1));        // GH253 B* is already gone: 2 columns
  const typo = product({}, tabs);
  const snap = durable(typo.store);
  const r = await typo.tool.handler(req([del('GH253 B1 (typo)', 2, 'GH253 A0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(typo.writes(), 0);
  assert.equal(durable(typo.store), snap);
  const start = product({}, tabs);                                              // the same typo with the START width is not SATISFIED, nor applied
  const s = await start.tool.handler(req([del('GH253 B1 (typo)', 1, 'GH253 A0')]));
  assert.equal(s.ok, false, JSON.stringify(s));
  assert.match(s.reason, /table-column-terminal-unprovable/);
  assert.equal(start.writes(), 0);
});

test('TABLE IDENTITY: the anchor decides the table; a target text two tables hold is deleted in the ANCHORED table only; an anchor in the target column, a missing / ambiguous anchor, a missing target, a wrong width are FAILED with 0 writes', async () => {
  const p = product();
  const before = await p.readTables();
  const r = await p.tool.handler(req([del('GH253 B1', 2, 'GH253 A0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.equal(after[0].rows[0].length, 2);
  assert.deepEqual(after[1], before[1], 'the twin in the other table is untouched');
  const dupAnchor = [{ rows: [[cell('GH253 Anchor'), cell('GH253 X')], [cell('GH253 Y'), cell('GH253 Z')]] }, { rows: [[cell('GH253 Anchor'), cell('GH253 Q')]] }];
  for (const [tabs, steps, why] of [
    [tables(), [del('GH253 B1', 2, 'GH253 B2')], /table-anchor-in-target-column/],
    [tables(), [del('GH253 B1', 2, 'GH253 No Such')], /table-anchor-not-found/], [dupAnchor, [del('GH253 Z', 1, 'GH253 Anchor')], /table-anchor-ambiguous/],
    [tables(), [del('GH253 Not There', 2, 'GH253 A0')], /table-column-terminal-unprovable/],
    [tables(), [del('GH253 B1', 3, 'GH253 A0')], /delete-target-present-but-column-count-not-terminal-plus-one/], [tables(), [del('GH253 B1', 4, 'GH253 A0')], /delete-target-present-but-column-count-not-terminal-plus-one/],
    [tables(), [del('GH253 B1', 1, 'GH253 A0')], /delete-target-present-but-column-count-not-terminal-plus-one/],
  ]) {
    const q = product({}, tabs);
    const snap = durable(q.store);
    const x = await q.tool.handler(req(steps));
    assert.equal(x.ok, false, JSON.stringify(x));
    assert.match(x.reason, why);
    assert.equal(x.operations[0].status, 'FAILED');
    assert.equal(q.writes(), 0);
    assert.equal(durable(q.store), snap);
  }
});

test('COLSPAN (destructive boundary): a target that spans several columns, or a column a colspan COVERS, is FAILED before any write (the merged content would be lost); a column outside every span, in the same table, applies', async () => {
  for (const [steps, why] of [
    [[del('GH253 M1', 2, 'GH253 F0')], /target-spans-columns/],
    [[del('GH253 G0', 2, 'GH253 F1')], /colspan-covers-deleted-column/],
    [[del('GH253 H0', 2, 'GH253 F1')], /colspan-covers-deleted-column/],
  ]) {
    const p = product();
    const snap = durable(p.store);
    const r = await p.tool.handler(req(steps));
    assert.equal(r.ok, false, JSON.stringify(r));
    assert.match(r.reason, why);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
  const edge = product();
  const e = await edge.tool.handler(req([del('GH253 F0', 2, 'GH253 H0')]));       // column 0 sits left of the span (cols 1-2)
  assert.equal(e.ok, true, JSON.stringify(e));
  assert.equal((await edge.readTables())[2].rows[0].length, 2);
});

test('VERTICAL MERGE (W-06/2 LIVE-measured: the delete is supported): a column outside the merge, the merge\'s OWN column and a neighbour all apply; the merge column\'s content is lost with the column (that is the delete); a vertical merge in ANOTHER table is no obstacle', async () => {
  for (const [target, anchor, width] of [['GH253 W0', 'GH253 V0', 1], ['GH253 VTop', 'GH253 W0', 1], ['GH253 W1', 'GH253 VTop', 1]]) {
    const p = product();
    const r = await p.tool.handler(req([del(target, width, anchor)]));
    assert.equal(r.ok, true, `${target}: ${JSON.stringify(r)}`);
    assert.equal((await p.readTables())[3].rows[0].length, 1, target);
  }
  const own = product();
  await own.tool.handler(req([del('GH253 VTop', 1, 'GH253 W0')]));
  assert.deepEqual(texts((await own.readTables())[3].rows), [['GH253 W0'], ['GH253 W1'], ['GH253 W2']], 'the merge column (VTop and its continuation) is gone, the other column intact');
  assert.equal((await product().tool.handler(req([del('GH253 B1', 2, 'GH253 A0')]))).ok, true);
});

test('LAST COLUMNS: a delete that leaves ONE column applies (terminalColumnCount 1); the only column of a table cannot be deleted (its anchor would sit in the deleted column)', async () => {
  const two = [{ rows: [[cell('GH253 X0'), cell('GH253 Y0')], [cell('GH253 P0'), cell('GH253 Q0')]] }];
  const p = product({}, two);
  const r = await p.tool.handler(req([del('GH253 X0', 1, 'GH253 Y0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal((await p.readTables())[0].rows[0].length, 1);
  const solo = product({}, [{ rows: [[cell('GH253 S0')], [cell('GH253 S1')]] }]);
  const x = await solo.tool.handler(req([del('GH253 S0', 1, 'GH253 S1')]));
  assert.equal(x.ok, false, JSON.stringify(x));
  assert.match(x.reason, /table-anchor-in-target-column/);
  assert.equal(solo.writes(), 0);
});

test('FAIL-CLOSED editor faults: a silent no-op, a stale column index, ANOTHER column deleted, a reordered row, a cell / table / outside text touched -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'col-stale-index', 'col-delete-other', 'col-touch-cell', 'col-touch-other-table', 'col-change-outside']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([del('GH253 B1', 2, 'GH253 A0')]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
  const wide = [{ rows: [[cell('GH253 A0'), cell('GH253 B0'), cell('GH253 C0'), cell('GH253 D0')], [cell('GH253 A1'), cell('GH253 B1'), cell('GH253 C1'), cell('GH253 D1')]] }];
  const reorder = product({ fault: 'col-reorder' }, wide);
  assert.equal((await reorder.tool.handler(req([del('GH253 B1', 3, 'GH253 A0')]))).ok, false);
  assert.equal(reorder.writes(), 0);
  const left = [{ rows: [[cell('GH253 A0'), cell('GH253 B0'), cell('GH253 C0')], [{ ...cell('GH253 P1'), span: 2 }, cell('GH253 Q1')]] }];   // a colspan LEFT of the target: logical != physical
  const phys = product({ fault: 'col-physical-index' }, left);
  assert.equal((await phys.tool.handler(req([del('GH253 C0', 2, 'GH253 A0')]))).ok, false, 'the physical-index model mis-deletes the colspan row');
  assert.equal(phys.writes(), 0);
  assert.equal((await product({}, left).tool.handler(req([del('GH253 C0', 2, 'GH253 A0')]))).ok, true, 'the same delete without the fault applies (logical column identity)');
});

test('PRE-DISPATCH refusals never open a session: a refused column-delete task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[del('GH253 B1', 2)], [{ intent: 'delete_table_column', selector: { text: 'GH253 B1' }, terminalColumnCount: 2 }], [del('GH253 B1', 0, 'GH253 A0')],
    [del('GH253 B1', 2, 'GH253 B1')], [del('GH253 C0', 2, 'GH253 A0'), del('GH253 B0', 1, 'GH253 A0')]]) {
    const r = await p.tool.handler(req(steps));
    assert.equal(r.dispatched, false);
  }
  assert.equal(p.opens(), 0);
});

// ---- the release gate ------------------------------------------------------------------------------------------------------

test('GATE: the DOCX allowlist is exactly the accepted set (both halves agree); a one-sided change withholds only docx_execute_task', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('delete_table_column'));
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  const pre = [...W.DOCX_WRITE_ALLOWLIST].filter(i => i !== 'delete_table_column');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#253 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  assert.ok(T.capabilityInventory().some(r => r.TOOL === 'docx:delete_table_column' && r.EXPOSE_IN_PRODUCT_MCP === 'YES'));
});
