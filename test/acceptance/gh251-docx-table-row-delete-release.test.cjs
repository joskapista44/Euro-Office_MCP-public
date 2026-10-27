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
 * GH #251 (DOCX P1 WRITE): delete_table_row served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH251-TABLE-ROW-DELETE-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs
 * (its row model follows the Wave 4b / M5 / W06-3 LIVE probes); the READ side is the product docx_read read_tables on the SAME store. The LIVE acceptance is
 * test/live/docx-table-row-delete-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH251 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
// T0 (4 rows: A0, A1, a horizontally merged two-paragraph row, A3), T1 shares 'GH251 B1' with T0, T2 has a vertical merge over rows 1-2
const tables = () => [
  { rows: [[cell('GH251 A0'), cell('GH251 B0')], [cell('GH251 A1'), cell('GH251 B1')], [{ ...cell('GH251 M1', 'GH251 M2'), span: 2 }], [cell('GH251 A3'), cell('GH251 B3')]] },
  { rows: [[cell('GH251 C0'), cell('GH251 D0')], [cell('GH251 C1'), cell('GH251 B1')]] },
  { rows: [[cell('GH251 V0'), cell('GH251 W0')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH251 VTop', props: { ...PLAIN } }] }, cell('GH251 W1')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH251 W2')], [cell('GH251 V3'), cell('GH251 W3')]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const del = (text, n, tableAnchor, extra = {}) => ({ intent: 'delete_table_row', selector: tc(text, tableAnchor), terminalRowCount: n, ...extra });
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

test('SCHEMA: { tableCell: { text, tableAnchor } } (the surviving anchor is REQUIRED) plus terminalRowCount is accepted, with an optional step id', () => {
  for (const steps of [[del('GH251 A1', 3, 'GH251 A0')], [{ ...del('GH251 A3', 3, 'GH251 B0'), id: 'd1' }], [del('GH251 A1', 1, 'GH251 A0')]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: a missing anchor, a missing / invalid terminalRowCount, a blank text, an anchor equal to the target cell are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = del('GH251 A1', 3, 'GH251 A0');
  const bad = [
    del('GH251 A1', 3), { ...ok, terminalRowCount: undefined }, { ...ok, terminalRowCount: 0 }, { ...ok, terminalRowCount: -1 }, { ...ok, terminalRowCount: 1001 }, { ...ok, terminalRowCount: 2.5 }, { ...ok, terminalRowCount: '3' },
    { intent: 'delete_table_row', terminalRowCount: 3 }, { ...ok, selector: {} }, { ...ok, selector: { tableCell: { tableAnchor: 'GH251 A0' } } },
    del('   ', 3, 'GH251 A0'), del('', 3, 'GH251 A0'), del('GH251 A1', 3, 'GH251 A1'), del('GH251 A1', 3, '   '),
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: every other selector form / alias / extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } form, terminalTableAnchor, kind, position, a text, row / column hints', () => {
  const base = del('GH251 A1', 3, 'GH251 A0');
  for (const extra of [{ selector: { text: 'GH251 A1' } }, { selector: { ...base.selector, kind: 'paragraph' } }, { selector: { tableCell: { text: 'GH251 A1', tableAnchor: 'GH251 A0', row: 1 } } },
    { terminalTableAnchor: 'GH251 A0' }, { position: 'last' }, { text: 'x' }, { dependsOn: ['s0'] }, { before: 'x' }]) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a delete task is that ONE step -- a second step (another row, another table, an insert, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  const ins = { intent: 'insert_table_row_after', selector: tc('GH251 B0'), terminalRowCount: 4 };
  for (const steps of [[del('GH251 A0', 3, 'GH251 B3'), del('GH251 A1', 2, 'GH251 B3')], [del('GH251 A0', 3, 'GH251 B3'), del('GH251 C0', 1, 'GH251 D1')], [del('GH251 A0', 3, 'GH251 B3'), ins], [rw, del('GH251 A0', 3, 'GH251 B3')]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): a delete whose successor fails leaves the FIRST row gone and saved; a later step that deletes the row holding the earlier step\'s anchor fails with table-anchor-not-found AFTER the save', async () => {
  const run = async steps => {
    const store = createStore(doc(), { doc: { tables: tables() } });
    const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
    return { r: await exec({ fileId: FILE, task: { steps } }), store };
  };
  const failSecond = await run([del('GH251 A0', 3, 'GH251 B3'), del('GH251 Not There', 2, 'GH251 B3')]);
  assert.equal(failSecond.r.ok, false);
  assert.deepEqual(failSecond.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(failSecond.store.saves, 1, 'the first delete was saved next to the failed second step');
  assert.equal(failSecond.store.durableDoc.tables[0].rows.length, 3);
  const anchorGone = await run([del('GH251 A0', 3, 'GH251 B3'), del('GH251 A3', 2, 'GH251 B0')]);
  assert.deepEqual(anchorGone.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.match(anchorGone.r.reason, /table-anchor-not-found/);
  assert.equal(D.validateDocxWriteRequest(req([del('GH251 A0', 3, 'GH251 B3'), del('GH251 A3', 2, 'GH251 B0')])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid row delete (merge / split #254-#255)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([del('GH251 A0', 3, 'GH251 B3'), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries the delete step: selector.tableCell with text AND tableAnchor required, terminalRowCount (integer 1..1000) required, nothing else', () => {
  const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === 'delete_table_row');
  assert.deepEqual(s.required, ['intent', 'selector', 'terminalRowCount']);
  assert.deepEqual(Object.keys(s.properties).sort(), ['id', 'intent', 'selector', 'terminalRowCount']);
  assert.deepEqual(s.properties.terminalRowCount, { type: 'integer', minimum: 1, maximum: 1000 });
  assert.deepEqual(s.properties.selector.required, ['tableCell']);
  assert.deepEqual(s.properties.selector.properties.tableCell.required, ['text', 'tableAnchor']);
  assert.equal(s.additionalProperties, false);
  assert.match(D.createDocxWriteTool({}).description, /delete_table_row/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

for (const [target, anchor, at] of [['GH251 A0', 'GH251 B3', 0], ['GH251 A1', 'GH251 A0', 1], ['GH251 M1 GH251 M2', 'GH251 A0', 2], ['GH251 A3', 'GH251 A0', 3]]) {
  test(`READ SYMMETRY (delete ${target}): rows 4 -> 3, the target row is gone, EVERY other row (text / simple / format / paraFormats) in order, every other table and the text outside the tables read unchanged; L1 + L3 PASS`, async () => {
    const p = product();
    const before = await p.readTables();
    const outside = await p.readText();
    const r = await p.tool.handler(req([del(target, 3, anchor)]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(p.writes(), 1);
    const after = await p.readTables();
    assert.deepEqual(after[0].rows, before[0].rows.filter((_, i) => i !== at), 'every surviving row, in order, unchanged');
    assert.deepEqual(after.slice(1), before.slice(1), 'the other tables are untouched');
    assert.deepEqual(await p.readText(), outside);
  });
}

test('RETRY: the same task again is SATISFIED / noOp / 0 writes (the target is absent from the anchor table and the count is terminal)', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([del('GH251 A1', 3, 'GH251 A0')]))).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(req([del('GH251 A1', 3, 'GH251 A0')]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
});

test('NAMED LIMIT (measured, the #234 survivor question): a row that was NEVER there reads SATISFIED when the anchor table exists and holds the terminal count -- the missing row alone is not evidence of OUR deletion; the proof is anchor + absence + count', async () => {
  const tabs = tables(); tabs[0].rows.splice(1, 1);                 // GH251 A1 is already gone: 3 rows
  const typo = product({}, tabs);
  const snap = durable(typo.store);
  const r = await typo.tool.handler(req([del('GH251 A1 (typo)', 3, 'GH251 A0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.noOp, true);
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(typo.writes(), 0);
  assert.equal(durable(typo.store), snap);
  const start = product({}, tabs);                                  // the same typo with the START count is not SATISFIED, nor applied
  const s = await start.tool.handler(req([del('GH251 A1 (typo)', 2, 'GH251 A0')]));
  assert.equal(s.ok, false, JSON.stringify(s));
  assert.match(s.reason, /table-row-terminal-unprovable/);
  assert.equal(start.writes(), 0);
});

test('TABLE IDENTITY: the anchor decides the table; a target text two tables hold is deleted in the ANCHORED table only; an anchor in the target row, a missing / ambiguous anchor, a missing target are FAILED with 0 writes', async () => {
  const p = product();
  const before = await p.readTables();
  const r = await p.tool.handler(req([del('GH251 B1', 3, 'GH251 A0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.equal(after[0].rows.length, 3);
  assert.deepEqual(after[1], before[1], 'the twin in the other table is untouched');
  const dupAnchor = [{ rows: [[cell('GH251 Anchor'), cell('GH251 X')], [cell('GH251 Y'), cell('GH251 Z')]] }, { rows: [[cell('GH251 Anchor'), cell('GH251 Q')]] }];
  for (const [tabs, steps, why] of [
    [tables(), [del('GH251 A3', 3, 'GH251 B3')], /table-anchor-in-target-row/],
    [tables(), [del('GH251 A1', 3, 'GH251 No Such')], /table-anchor-not-found/], [dupAnchor, [del('GH251 Y', 1, 'GH251 Anchor')], /table-anchor-ambiguous/],
    [tables(), [del('GH251 Not There', 3, 'GH251 A0')], /table-row-terminal-unprovable/],
    [tables(), [del('GH251 A1', 4, 'GH251 A0')], /delete-target-present-but-row-count-not-terminal-plus-one/], [tables(), [del('GH251 A1', 5, 'GH251 A0')], /delete-target-present-but-row-count-not-terminal-plus-one/],
    [tables(), [del('GH251 A1', 2, 'GH251 A0')], /delete-target-present-but-row-count-not-terminal-plus-one/],
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

test('VERTICAL MERGE (destructive boundary): a delete of the merge OWNER row is FAILED before any write; GH #291 T-06: its CONTINUATION row applies (the merge ends, the owner content kept); a delete outside the span (above / below), in the same table, applies', async () => {
  const cont = product();
  const rc = await cont.tool.handler(req([del('GH251 W2', 3, 'GH251 W0')]));
  assert.equal(rc.ok, true, JSON.stringify(rc));
  const vt = cont.store.durableDoc.tables[2].rows;
  assert.deepEqual([vt.length, vt[1][0].paras[0].text, vt[1][0].rowspan, vt.flat().some(c => c.vContinue)], [3, 'GH251 VTop', undefined, false]);
  for (const target of ['GH251 VTop', 'GH251 W1']) {
    const p = product();
    const snap = durable(p.store);
    const r = await p.tool.handler(req([del(target, 3, 'GH251 W0')]));
    assert.equal(r.ok, false, `${target}: ${JSON.stringify(r)}`);
    assert.match(r.reason, /row-delete-inside-a-vertical-merge/);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
  for (const [target, anchor] of [['GH251 V0', 'GH251 W3'], ['GH251 V3', 'GH251 W0']]) {
    const p = product();
    const r = await p.tool.handler(req([del(target, 3, anchor)]));
    assert.equal(r.ok, true, `${target}: ${JSON.stringify(r)}`);
    assert.equal((await p.readTables())[2].rows.length, 3);
  }
  const other = product();     // a vertical merge in ANOTHER table is no obstacle
  assert.equal((await other.tool.handler(req([del('GH251 A1', 3, 'GH251 A0')]))).ok, true);
});

test('LAST ROWS: a delete that leaves ONE row applies (terminalRowCount 1); the only row of a table cannot be deleted (its anchor would sit in the deleted row)', async () => {
  const one = [{ rows: [[cell('GH251 ONLY0'), cell('GH251 ONLY1')], [cell('GH251 X0'), cell('GH251 X1')]] }];
  const p = product({}, one);
  const r = await p.tool.handler(req([del('GH251 ONLY0', 1, 'GH251 X0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal((await p.readTables())[0].rows.length, 1);
  const solo = product({}, [{ rows: [[cell('GH251 S0'), cell('GH251 S1')]] }]);
  const x = await solo.tool.handler(req([del('GH251 S0', 1, 'GH251 S1')]));
  assert.equal(x.ok, false, JSON.stringify(x));
  assert.match(x.reason, /table-anchor-in-target-row/);
  assert.equal(solo.writes(), 0);
});

test('FAIL-CLOSED editor faults: a silent no-op, a stale row index, ANOTHER row deleted, a reordered table, another cell / table / outside text touched -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'row-stale-index', 'row-delete-other', 'row-reorder', 'row-touch-cell', 'row-touch-other-table', 'row-change-outside']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([del('GH251 A1', 3, 'GH251 A0')]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
});

test('PRE-DISPATCH refusals never open a session: a refused row-delete task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[del('GH251 A1', 3)], [{ intent: 'delete_table_row', selector: { text: 'GH251 A1' }, terminalRowCount: 3 }], [del('GH251 A1', 0, 'GH251 A0')],
    [del('GH251 A1', 3, 'GH251 A1')], [del('GH251 A0', 3, 'GH251 B3'), del('GH251 A1', 2, 'GH251 B3')]]) {
    const r = await p.tool.handler(req(steps));
    assert.equal(r.dispatched, false);
  }
  assert.equal(p.opens(), 0);
});

// ---- the release gate ------------------------------------------------------------------------------------------------------

test('GATE: the DOCX allowlist is exactly the accepted set (both halves agree); a one-sided change withholds only docx_execute_task', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('delete_table_row'));
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  const pre = [...W.DOCX_WRITE_ALLOWLIST].filter(i => i !== 'delete_table_row');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#251 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  assert.ok(T.capabilityInventory().some(r => r.TOOL === 'docx:delete_table_row' && r.EXPOSE_IN_PRODUCT_MCP === 'YES'));
});
