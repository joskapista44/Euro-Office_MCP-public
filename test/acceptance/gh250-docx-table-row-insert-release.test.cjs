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
 * GH #250 (DOCX P1 WRITE): insert_table_row_before / insert_table_row_after served by docx_execute_task (the B commit). Design and the measured reason for every rule:
 * docs/GH250-TABLE-ROW-INSERT-PRODUCT-DECISION.md. Offline only: the product handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs
 * (its row model follows the Wave 4b / M5 LIVE probes); the READ side is the product docx_read read_tables on the SAME store. The LIVE acceptance is
 * test/live/docx-table-row-insert-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Courier New', fontSize: 28, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const richCell = text => ({ span: 1, paras: [{ text, props: { ...RICH }, alignment: '2', style: 'Table Text' }] });
const OUTSIDE = 'GH250 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
// T0 (4 rows: a formatted row, a plain row, a horizontally merged two-paragraph row, a last row), T1 TWINS (shares 'GH250 B1' with T0), T2 VMERGE (rows 1-2 of column 0), T3 holds a 'GH250 Shared'
const tables = () => [
  { rows: [[richCell('GH250 A0'), cell('GH250 B0')], [cell('GH250 A1'), cell('GH250 B1')], [{ ...cell('GH250 M1', 'GH250 M2'), span: 2 }], [cell('GH250 A3'), cell('GH250 B3')]] },
  { rows: [[cell('GH250 C0'), cell('GH250 D0')], [cell('GH250 C1'), cell('GH250 B1')]] },
  { rows: [[cell('GH250 V0'), cell('GH250 W0')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH250 VTop', props: { ...PLAIN } }] }, cell('GH250 W1')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH250 W2')], [cell('GH250 V3'), cell('GH250 W3')]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const row = (side, text, n, tableAnchor, extra = {}) => ({ intent: `insert_table_row_${side}`, selector: tc(text, tableAnchor), terminalRowCount: n, ...extra });
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

test('SCHEMA: the #249 selector { tableCell: { text [, tableAnchor] } } plus the required terminalRowCount is accepted, before and after, with an optional step id', () => {
  for (const steps of [[row('before', 'GH250 A1', 5)], [row('after', 'GH250 A1', 5, 'GH250 A0')], [{ ...row('after', 'GH250 A3', 5, 'GH250 A0'), id: 'r1' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.ok, true, `${JSON.stringify(steps)} -> ${v.code} ${v.reason}`);
  }
});

test('SCHEMA: every other selector form / alias / extra key is UNKNOWN_FIELD (0 dispatch) -- the bare { text } form, terminalTableAnchor, kind, position, a text, row / column hints', () => {
  const base = row('after', 'GH250 A1', 5);
  const bad = [
    { selector: { text: 'GH250 A1' } }, { selector: { tableCell: { text: 'GH250 A1' }, text: 'GH250 A1' } }, { selector: { ...base.selector, kind: 'paragraph' } },
    { selector: { tableCell: { text: 'GH250 A1', row: 1 } } }, { selector: { tableCell: { text: 'GH250 A1', index: 0 } } }, { selector: { tableCell: { text: 'GH250 A1', contains: true } } },
    { terminalTableAnchor: 'GH250 A0' }, { position: 'last' }, { text: 'x' }, { count: 2 }, { dependsOn: ['s0'] }, { before: 'x' },
  ];
  for (const extra of bad) {
    const v = D.validateDocxWriteRequest(req([{ ...base, ...extra }]));
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(extra));
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a missing or invalid terminalRowCount, a missing selector / tableCell / text, and an anchor equal to the target cell are REFUSED_SHAPE (0 dispatch)', () => {
  const ok = row('after', 'GH250 A1', 5);
  const bad = [
    { intent: 'insert_table_row_after', selector: tc('GH250 A1') }, { ...ok, terminalRowCount: 1 }, { ...ok, terminalRowCount: 0 }, { ...ok, terminalRowCount: -3 }, { ...ok, terminalRowCount: 1001 },
    { ...ok, terminalRowCount: 5.5 }, { ...ok, terminalRowCount: '5' }, { ...ok, terminalRowCount: null },
    { intent: 'insert_table_row_after', terminalRowCount: 5 }, { ...ok, selector: {} }, { ...ok, selector: { tableCell: { tableAnchor: 'GH250 A0' } } },
    row('after', '   ', 5), row('after', '', 5), row('after', 'GH250 A1', 5, 'GH250 A1'), row('after', 'GH250 A1', 5, '   '),
  ];
  for (const s of bad) {
    const v = D.validateDocxWriteRequest(req([s]));
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(s)} -> ${v.code}: ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a row-insert task is that ONE step -- a second step (the same table, another table, any other intent) is REFUSED_SHAPE, 0 dispatch', () => {
  const rw = { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'y' };
  for (const steps of [[row('after', 'GH250 A0', 5, 'GH250 B0'), row('after', 'GH250 A3', 6, 'GH250 B0')], [row('after', 'GH250 A0', 5, 'GH250 B0'), row('before', 'GH250 C1', 3, 'GH250 C0')],
    [row('after', 'GH250 A0', 5), rw], [rw, row('after', 'GH250 A0', 5)], [row('after', 'GH250 A0', 5, 'GH250 B0'), { intent: 'replace_table_cell', selector: tc('GH250 B3'), text: 'x' }]]) {
    const v = D.validateDocxWriteRequest(req(steps));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(steps.map(s => s.intent)));
    assert.equal(v.dispatched, false);
  }
});

test('WHY ONE STEP (the measured hazard, run BELOW the product validation): a row insert followed by a step that fails leaves the FIRST row saved next to a failed task; anchor-less row steps cannot be grouped at all', async () => {
  const run = async steps => {
    const store = createStore(doc(), { doc: { tables: tables() } });
    const exec = D.createDocxWriteExecution({ configSource: cfg, gate: createSessionGate(), runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) }, createDriver: () => createFakeDocxDriver(store, { saveObserver: true }) });
    return { r: await exec({ fileId: FILE, task: { steps } }), store };
  };
  const partial = await run([row('after', 'GH250 A0', 5, 'GH250 B0'), row('after', 'GH250 Not There', 6, 'GH250 B0')]);
  assert.equal(partial.r.ok, false);
  assert.deepEqual(partial.r.operations.map(o => o.status), ['APPLIED', 'FAILED']);
  assert.equal(partial.store.saves, 1, 'the first row was saved next to the failed second step');
  assert.equal(partial.store.durableDoc.tables[0].rows.length, 5);
  const anchorless = await run([row('after', 'GH250 A0', 5), row('after', 'GH250 A3', 6)]);
  assert.equal(anchorless.r.ok, false);
  assert.match(anchorless.r.reason, /row-operations-without-table-anchor-cannot-be-grouped/);
  assert.equal(anchorless.store.saves, 0);
  assert.equal(D.validateDocxWriteRequest(req([row('after', 'GH250 A0', 5, 'GH250 B0'), row('after', 'GH250 Not There', 6, 'GH250 B0')])).code, 'REFUSED_SHAPE', 'the product schema refuses that task shape');
});

test('ALLOWLIST: every OTHER routed docx intent is still REFUSED_NOT_ALLOWLISTED pre-dispatch, even next to a valid row insert (merge / split #254-#255; delete_table_row joined with #251)', () => {
  const others = D.docxRoutedIntents();
  for (const i of W.DOCX_WRITE_ALLOWLIST) others.delete(i);
  assert.equal(others.size, 0); // release 3 trim + GH #266/#269e/#268d: every previously-withheld intent (define_style, modify_style, set_numbering, set_list_type, add_footnote, add_endnote, add_link, add_bookmark, resolve_comment, reply_comment, add_comment_on_match, set_image_size, set_image_wrap) is now released -- the array is current, nothing withheld
  assert.deepEqual([...others].sort(), []);
  for (const intent of others) {
    const v = D.validateDocxWriteRequest(req([row('after', 'GH250 A0', 5), { intent }]));
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', `${intent}: ${v.code}`);
    assert.equal(v.stepIndex, 1);
    assert.equal(v.wholeTask, true);
  }
});

test('SCHEMA: the published inputSchema carries both row steps: selector.tableCell and terminalRowCount (integer 2..1000) required, optional extendMerge (only true, GH #291 T-05), nothing else', () => {
  const items = D.inputSchema().properties.task.properties.steps.items;
  for (const intent of ['insert_table_row_before', 'insert_table_row_after']) {
    const s = items.oneOf.find(x => x.properties.intent.const === intent);
    assert.deepEqual(s.required, ['intent', 'selector', 'terminalRowCount']);
    assert.deepEqual(Object.keys(s.properties).sort(), ['extendMerge', 'id', 'intent', 'selector', 'terminalRowCount']); // GH #291 T-05: the explicit merge-extending mode
    assert.deepEqual(s.properties.extendMerge, { const: true });
    assert.deepEqual(s.properties.terminalRowCount, { type: 'integer', minimum: 2, maximum: 1000 });
    assert.deepEqual(s.properties.selector.required, ['tableCell']);
    assert.equal(s.additionalProperties, false);
  }
  assert.match(D.createDocxWriteTool({}).description, /insert_table_row_before, insert_table_row_after/);
});

// ---- the product handler, end to end on the fake transport; the READ side is read_tables on the SAME store --------------------------

const POSITIONS = [['before', 'GH250 A0', 0], ['after', 'GH250 A0', 1], ['before', 'GH250 A1', 1], ['after', 'GH250 A1', 2], ['before', 'GH250 A3', 3], ['after', 'GH250 A3', 4]];
for (const [side, target, at] of POSITIONS) {
  test(`READ SYMMETRY (${side} ${target}): rows 4 -> 5, the new row sits at index ${at}, EMPTY, with the target row's cell count; every OTHER row (text / simple / format / paraFormats), every other table and the text outside the tables read unchanged; L1 + L3 PASS`, async () => {
    const p = product();
    const before = await p.readTables();
    const outside = await p.readText();
    const r = await p.tool.handler(req([row(side, target, 5, 'GH250 B0')]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    assert.equal(r.wholeTask, 'PASS');
    assert.equal(p.writes(), 1);
    const after = await p.readTables();
    assert.equal(after[0].rows.length, 5);
    const targetRow = before[0].rows.findIndex(rw => rw.some(c => c.text === target));
    assert.deepEqual(after[0].rows[at].map(c => c.text), before[0].rows[targetRow].map(() => ''), 'the new row is empty and has the target row\'s cell count');
    const kept = after[0].rows.filter((_, i) => i !== at);
    assert.deepEqual(kept, before[0].rows, 'every original row, in order, unchanged');
    assert.deepEqual(after.slice(1), before.slice(1), 'the other tables are untouched');
    assert.deepEqual(await p.readText(), outside);
  });
}

test('RETRY: the same task again is SATISFIED / noOp / 0 writes; a PRE-EXISTING empty row on the declared side with the claimed count reads SATISFIED too (named limit: the proof is count + emptiness, not provenance)', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([row('after', 'GH250 A1', 5, 'GH250 B0')]))).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(req([row('after', 'GH250 A1', 5, 'GH250 B0')]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves);
  const tabs = tables(); tabs[0].rows.splice(2, 0, [cell(''), cell('')]);       // 5 rows: an empty row already sits after GH250 A1
  const pre = product({}, tabs);
  const same = await pre.tool.handler(req([row('after', 'GH250 A1', 5, 'GH250 B0')]));
  assert.equal(same.ok, true, JSON.stringify(same));
  assert.equal(same.noOp, true);
  assert.equal(pre.writes(), 0, 'no write: the claimed terminal state already holds');
  const more = await pre.tool.handler(req([row('after', 'GH250 A1', 6, 'GH250 B0')]));    // the count says one more is wanted -> a second empty row is inserted
  assert.equal(more.ok, true, JSON.stringify(more));
  assert.equal((await pre.readTables())[0].rows.length, 6);
});

test('MERGED ROW (horizontal): a row insert beside a colspan row names the row by its merged cell\'s joined text; the new row takes the target row\'s CELL COUNT and is empty; the merged row itself is unchanged', async () => {
  const p = product();
  const before = await p.readTables();
  const r = await p.tool.handler(req([row('after', 'GH250 M1 GH250 M2', 5, 'GH250 B0')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  const after = await p.readTables();
  assert.equal(after[0].rows[3].length, before[0].rows[2].length);
  assert.deepEqual(after[0].rows[3].map(c => c.text), ['']);
  assert.deepEqual(after[0].rows[2], before[0].rows[2]);
  assert.equal(p.store.durableDoc.tables[0].rows[2][0].span, 2, 'the merged row keeps its colspan');
});

test('TABLE IDENTITY: a cell text two tables hold is AMBIGUOUS without an anchor (FAILED, 0 writes); with tableAnchor the row is inserted in the named table ONLY', async () => {
  const noAnchor = product();
  const snap = durable(noAnchor.store);
  const r0 = await noAnchor.tool.handler(req([row('after', 'GH250 B1', 5)]));
  assert.equal(r0.ok, false, JSON.stringify(r0));
  assert.match(r0.reason, /table-row-target-ambiguous/);
  assert.equal(noAnchor.writes(), 0);
  assert.equal(durable(noAnchor.store), snap);
  for (const [anchor, ti, n] of [['GH250 A0', 0, 5], ['GH250 C0', 1, 3]]) {
    const p = product();
    const before = await p.readTables();
    const r = await p.tool.handler(req([row('after', 'GH250 B1', n, anchor)]));
    assert.equal(r.ok, true, `${anchor}: ${JSON.stringify(r)}`);
    const after = await p.readTables();
    assert.equal(after[ti].rows.length, n);
    assert.deepEqual(after[1 - ti], before[1 - ti], 'the other table is untouched');
  }
});

test('FAIL-CLOSED, 0 writes, nothing changed: a wrong terminalRowCount (+2, unchanged, fewer), a cell that is not there, an anchor no table / several tables have', async () => {
  const dupAnchor = [{ rows: [[cell('GH250 Anchor'), cell('GH250 X')], [cell('GH250 Y'), cell('GH250 Z')]] }, { rows: [[cell('GH250 Anchor'), cell('GH250 Q')]] }];
  for (const [tabs, steps, why] of [
    [tables(), [row('after', 'GH250 A1', 6, 'GH250 B0')], /table-row-terminal-unprovable/], [tables(), [row('after', 'GH250 A1', 4, 'GH250 B0')], /table-row-terminal-unprovable/],
    [tables(), [row('after', 'GH250 A1', 3, 'GH250 B0')], /table-row-terminal-unprovable/],
    [tables(), [row('after', 'GH250 Not There', 5, 'GH250 B0')], /table-row-target-not-found/], [tables(), [row('after', 'GH250 A1', 5, 'GH250 No Such')], /table-anchor-not-found/],
    [dupAnchor, [row('after', 'GH250 Y', 3, 'GH250 Anchor')], /table-anchor-ambiguous/],
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

test('VERTICAL MERGE (rowspan safety boundary): an insert that would CUT a vertical merge, or whose target row sits inside one, is FAILED before any write; an insert on a boundary that cuts none (above / below the span) applies', async () => {
  const A = 'GH250 W0';
  for (const [side, target, why] of [['after', 'GH250 W1', /row-insert-cuts-a-vertical-merge/], ['before', 'GH250 W2', /row-insert-cuts-a-vertical-merge/], ['after', 'GH250 VTop', /row-insert-cuts-a-vertical-merge/]]) {
    const p = product();
    const snap = durable(p.store);
    const r = await p.tool.handler(req([row(side, target, 5, A)]));
    assert.equal(r.ok, false, `${side} ${target}: ${JSON.stringify(r)}`);
    assert.match(r.reason, why);
    assert.equal(p.writes(), 0);
    assert.equal(durable(p.store), snap);
  }
  for (const [side, target] of [['before', 'GH250 V0'], ['after', 'GH250 V0'], ['before', 'GH250 V3'], ['after', 'GH250 V3']]) {
    const p = product();
    const r = await p.tool.handler(req([row(side, target, 5, A)]));
    assert.equal(r.ok, true, `${side} ${target}: ${JSON.stringify(r)}`);
    assert.equal((await p.readTables())[2].rows.length, 5);
  }
  const other = product();     // a vertical merge in ANOTHER table is no obstacle
  assert.equal((await other.tool.handler(req([row('after', 'GH250 A1', 5, 'GH250 B0')]))).ok, true);
});

test('FAIL-CLOSED editor faults: a silent no-op, a stale row index, the wrong side, a wrong-shaped / non-empty / doubled new row, a reordered table, another cell / table / outside text touched -- never APPLIED, nothing saved', async () => {
  for (const fault of ['ignore', 'row-stale-index', 'row-wrong-side', 'row-new-shape', 'row-new-not-empty', 'row-twice', 'row-reorder', 'row-touch-cell', 'row-touch-other-table', 'row-change-outside']) {
    const p = product({ fault });
    const r = await p.tool.handler(req([row('after', 'GH250 A1', 5, 'GH250 B0')]));
    assert.equal(r.ok, false, `${fault}: ${JSON.stringify(r)}`);
    assert.equal((r.operations || []).some(o => o.status === 'APPLIED'), false, fault);
    assert.equal(p.writes(), 0, fault);
  }
});

test('PRE-DISPATCH refusals never open a session: a refused row-insert task opens zero drivers', async () => {
  const p = product();
  for (const steps of [[{ intent: 'insert_table_row_after', selector: { text: 'GH250 A1' }, terminalRowCount: 5 }], [row('after', 'GH250 A1', 1)], [{ intent: 'insert_table_row_after', selector: tc('GH250 A1') }],
    [row('after', 'GH250 A1', 5, 'GH250 A1')], [row('after', 'GH250 A0', 5, 'GH250 B0'), row('after', 'GH250 A3', 6, 'GH250 B0')]]) {
    const r = await p.tool.handler(req(steps));
    assert.equal(r.dispatched, false);
  }
  assert.equal(p.opens(), 0);
});

// ---- the release gate ------------------------------------------------------------------------------------------------------

test('GATE: the DOCX allowlist is exactly the accepted set (both halves agree); a one-sided change withholds only docx_execute_task', () => {
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...T.ACCEPTED_DOCX_WRITE_ALLOWLIST]);
  for (const i of ['insert_table_row_before', 'insert_table_row_after']) assert.ok(W.DOCX_WRITE_ALLOWLIST.includes(i), i);
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const xlsx = [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST];
  const pre = [...W.DOCX_WRITE_ALLOWLIST].filter(i => !/^insert_table_row_/.test(i));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'docx' }).ok, false, 'the pre-#250 list no longer matches');
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: xlsx, docxAllowlist: pre, format: 'xlsx' }).ok, true);
  for (const i of ['insert_table_row_before', 'insert_table_row_after']) assert.ok(T.capabilityInventory().some(r => r.TOOL === `docx:${i}` && r.EXPOSE_IN_PRODUCT_MCP === 'YES'), i);
});
