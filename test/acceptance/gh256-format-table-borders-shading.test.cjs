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
 * GH #256 (DOCX P1 WRITE, table-level borders and shading; the release itself is an OWNER DECISION, alongside #228..#232): format_table served by
 * docx_execute_task for the TABLE-LEVEL borders and shading axes ONLY. Design and the measured reason for every rule:
 * docs/GH256-FORMAT-TABLE-BORDERS-SHADING-DECISION.md. The table is named like replace_table_cell's cell (table.cjs cellSelector, #249): the one canonical
 * { tableCell: { text [, tableAnchor] } } under the adapter's own step key `table`. Offline only: the product handler runs the REAL docx adapter /
 * executor / lifecycle on test/fakes/fake-docx-driver.cjs (applyTableFormat models the measured write persistence and the READ shape); the READ side is
 * the product docx_read read_tables (GH #222) on the SAME store. The LIVE acceptance is test/live/gh256-format-table-borders-shading-product.live.cjs
 * (not run by the author). cellMargins / indent / width (#257) and style (#258, merged in release 3) joined this schema later.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH256 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
// T0 TARGET (shares 'GH256 Twin' with T1), T1 TWINS, T2 VMERGE (a vertical merge: tableRead refuses the whole document's table identity)
const tables = (withVmerge = false) => [
  { rows: [[cell('GH256 Name'), cell('GH256 Qty')], [cell('GH256 Alpha'), cell('GH256 Twin')]] },
  { rows: [[cell('GH256 Item'), cell('GH256 Price')], [cell('GH256 Pear'), cell('GH256 Twin')]] },
  ...(withVmerge ? [{ rows: [[cell('GH256 VA'), cell('GH256 VB')], [{ span: 1, rowspan: 2, paras: [{ text: 'GH256 VTop', props: { ...PLAIN } }] }, cell('GH256 V5')], [{ span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }, cell('GH256 V7')]] }] : []),
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const SINGLE = (size, color) => ({ style: 'single', size, color });
const ft = (text, fields, tableAnchor) => ({ intent: 'format_table', table: tc(text, tableAnchor), ...fields });
const WANTED = { borders: { top: SINGLE(4, '#FF0000'), left: SINGLE(8, '#00FF00'), insideH: SINGLE(2, '#0000FF'), bottom: { style: 'none' } }, shading: '#330066' };
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

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
const hex = c => ({ r: parseInt(c.slice(1, 3), 16), g: parseInt(c.slice(3, 5), 16), b: parseInt(c.slice(5, 7), 16) });
/** what read_tables returns for a written single border: value / sz (the SAME eighths) / space / colour as { auto, r, g, b } */
const readBorder = (size, color) => ({ value: 'single', sz: size, space: 0, color: { auto: false, ...hex(color) } });
const NONE_READ = { value: 'none', sz: 0, space: 0, color: { auto: true, r: 0, g: 0, b: 0 } };
const UNSET_BORDERS = { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
const UNSET_REST = { cellMargins: { top: null, right: null, bottom: null, left: null }, indent: null, width: null, style: { id: null, name: null } };

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: format_table is in the allowlist and the published schema; borders (any of the six edges, single or none), shading, both, an optional tableAnchor and step id are accepted', () => {
  assert.ok(W.DOCX_WRITE_ALLOWLIST.includes('format_table'));
  assert.ok(T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes('format_table'));
  assert.ok(D.inputSchema().properties.task.properties.steps.items.oneOf.some(s => s.properties.intent.const === 'format_table'));
  for (const steps of [
    [ft('GH256 Name', WANTED)], [ft('GH256 Name', { shading: '#330066' })], [ft('GH256 Name', { borders: { top: SINGLE(4, '#FF0000') } })],
    [ft('GH256 Name', { borders: { top: SINGLE(1, '#ff00aa'), bottom: SINGLE(96, '#000000'), left: { style: 'none' }, right: SINGLE(4, '#FFFFFF'), insideH: SINGLE(4, '#123456'), insideV: { style: 'none' } } })],
    [ft('GH256 Twin', { shading: '#112233' }, 'GH256 Name')], [{ ...ft('GH256 Name', WANTED), id: 't1' }],
  ]) {
    const r = verdict(steps);
    assert.equal(r.ok, true, `${JSON.stringify(steps)} -> ${r.code} ${r.reason}`);
  }
});

test('SCHEMA: the table must be named by the ONE canonical { tableCell: { text [, tableAnchor] } } -- a bare { text }, an extra key, a physical index or a missing table are refused pre-dispatch', () => {
  const base = { shading: '#330066' };
  for (const table of [undefined, { text: 'GH256 Name' }, { tableCell: { text: 'GH256 Name' }, text: 'GH256 Name' }, { tableCell: { text: 'GH256 Name', row: 1 } }, { tableCell: { text: 'GH256 Name', tableIndex: 0 } }, { tableCell: {} }, { tableCell: { text: '' } }, 'GH256 Name', null]) {
    const v = verdict([{ intent: 'format_table', ...(table === undefined ? {} : { table }), ...base }]);
    assert.ok(['REFUSED_SHAPE', 'UNKNOWN_FIELD'].includes(v.code), `${JSON.stringify(table)} -> ${v.code}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: border edges -- a non-single/none style, a non-integer or out-of-range size, a bare / named colour, a none edge that also sets size or color, an unknown edge are refused pre-dispatch', () => {
  const cases = [
    [{ top: { style: 'double', size: 4, color: '#FF0000' } }, 'REFUSED_SHAPE'], [{ top: { style: 'dashed', size: 4, color: '#FF0000' } }, 'REFUSED_SHAPE'],
    [{ top: SINGLE(0, '#FF0000') }, 'REFUSED_SHAPE'], [{ top: SINGLE(97, '#FF0000') }, 'REFUSED_SHAPE'], [{ top: SINGLE(0.5, '#FF0000') }, 'REFUSED_SHAPE'], [{ top: SINGLE(-4, '#FF0000') }, 'REFUSED_SHAPE'], [{ top: SINGLE('4', '#FF0000') }, 'REFUSED_SHAPE'],
    [{ top: SINGLE(4, 'FF0000') }, 'REFUSED_SHAPE'], [{ top: SINGLE(4, 'red') }, 'REFUSED_SHAPE'], [{ top: SINGLE(4, '#FF00') }, 'REFUSED_SHAPE'],
    [{ top: { style: 'single', size: 4 } }, 'REFUSED_SHAPE'], [{ top: { style: 'single', color: '#FF0000' } }, 'REFUSED_SHAPE'], [{ top: { size: 4, color: '#FF0000' } }, 'REFUSED_SHAPE'],
    [{ top: { style: 'none', size: 4 } }, 'REFUSED_SHAPE'], [{ top: { style: 'none', color: '#FF0000' } }, 'REFUSED_SHAPE'], [{ top: { style: 'single', size: 4, color: '#FF0000', space: 0 } }, 'REFUSED_SHAPE'],
    [{ middle: SINGLE(4, '#FF0000') }, 'UNKNOWN_FIELD'], [{ start: SINGLE(4, '#FF0000') }, 'UNKNOWN_FIELD'], [{ top: null }, 'REFUSED_SHAPE'], [{ top: 'single' }, 'REFUSED_SHAPE'],
  ];
  for (const [borders, code] of cases) { const v = verdict([ft('GH256 Name', { borders })]); assert.equal(v.code, code, `${JSON.stringify(borders)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
});

test('SCHEMA: an empty borders object, a step naming no axis, and a malformed shading are REFUSED_SHAPE pre-dispatch; `style` joined with the release-3 schema merge (GH #258): a table style NAME together with borders / shading is one step, `styleId` / `tableStyle` stay UNKNOWN_FIELD', () => {
  for (const fields of [{ borders: {} }, {}, { shading: 'FF0000' }, { shading: 'red' }, { shading: '#FF00' }, { shading: 330066 }, { shading: '' }]) {
    const v = verdict([ft('GH256 Name', fields)]);
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(fields)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
  for (const fields of [{ styleId: 701 }, { tableStyle: 'Table Grid' }]) {
    const v = verdict([ft('GH256 Name', { shading: '#330066', ...fields })]);
    assert.equal(v.code, 'UNKNOWN_FIELD', JSON.stringify(fields));
    assert.equal(v.dispatched, false);
  }
});

test('SOLO: format_table must be the only step of its task (the proof is a property of the FINAL table; any other step can change the cell or the table it names)', () => {
  const others = [{ intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'X' }, { intent: 'replace_table_cell', selector: tc('GH256 Name'), text: 'N' }, ft('GH256 Item', { shading: '#112233' })];
  for (const other of others) {
    for (const steps of [[ft('GH256 Name', { shading: '#330066' }), other], [other, ft('GH256 Name', { shading: '#330066' })]]) {
      const v = verdict(steps);
      assert.equal(v.code, 'REFUSED_SHAPE', `${other.intent}: ${v.code} ${v.reason}`);
      assert.match(v.reason, /must be the only step/);
      assert.equal(v.dispatched, false);
    }
  }
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('WRITE + READ SYMMETRY: borders and shading land on the named table; read_tables returns the SAME size in eighths (as `sz`), the style as `value`, the colour as { auto, r, g, b }, shading as { val:clear, fill }; every other axis, table, cell and paragraph is untouched', async () => {
  const p = product();
  const beforeText = await p.readText();
  const r = await p.tool.handler(req([ft('GH256 Name', WANTED)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.operations[0].l1, 'PASS');
  assert.equal(r.operations[0].l3, 'PASS');
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes() >= 1, true);
  const t = await p.readTables();
  assert.deepEqual(t[0].format.borders, { top: readBorder(4, '#FF0000'), bottom: NONE_READ, left: readBorder(8, '#00FF00'), right: null, insideH: readBorder(2, '#0000FF'), insideV: null });
  assert.deepEqual(t[0].format.shading, { val: 'clear', fill: hex('#330066'), color: { auto: false, ...hex('#330066') } });
  for (const [k, v] of Object.entries(UNSET_REST)) assert.deepEqual(t[0].format[k], v, `${k} (#257 / not requested) is untouched`);
  assert.deepEqual(t[1].format.borders, UNSET_BORDERS, 'the twin table is untouched');
  assert.equal(t[1].format.shading, null);
  assert.deepEqual((await p.readText()), beforeText, 'the paragraphs outside the tables are untouched');
  assert.deepEqual(t[0].rows.map(row => row.map(c => c.text)), [['GH256 Name', 'GH256 Qty'], ['GH256 Alpha', 'GH256 Twin']], 'the cells of the formatted table are untouched');
});

test('INDEPENDENT AXES: a second shading-only request leaves the borders of the first request exactly as they were; a borders-only request leaves the shading', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH256 Name', { borders: { top: SINGLE(4, '#FF0000') } })]))).ok, true);
  const second = await p.tool.handler(req([ft('GH256 Name', { shading: '#112233' })]));
  assert.equal(second.ok, true, JSON.stringify(second));
  let t = await p.readTables();
  assert.deepEqual(t[0].format.borders.top, readBorder(4, '#FF0000'), 'the first request\'s border survives a shading-only request');
  assert.deepEqual(t[0].format.shading.fill, hex('#112233'));
  const third = await p.tool.handler(req([ft('GH256 Name', { borders: { left: SINGLE(8, '#00FF00') } })]));
  assert.equal(third.ok, true, JSON.stringify(third));
  t = await p.readTables();
  assert.deepEqual(t[0].format.borders.left, readBorder(8, '#00FF00'));
  assert.deepEqual(t[0].format.borders.top, readBorder(4, '#FF0000'), 'a second edge does not overwrite the first');
  assert.deepEqual(t[0].format.shading.fill, hex('#112233'), 'a borders-only request leaves the shading');
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes (idempotent, restart-safe)', async () => {
  const p = product();
  const task = req([ft('GH256 Name', WANTED)]);
  assert.equal((await p.tool.handler(task)).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves, 'no write on the retry');
});

test('TABLE IDENTITY (the #249 pattern): a cell text shared by two tables is ambiguous and fails CLOSED with no write; tableAnchor (another cell of the intended table) selects exactly one; an unknown cell text fails closed', async () => {
  const p = product();
  const amb = await p.tool.handler(req([ft('GH256 Twin', { shading: '#112233' })]));
  assert.equal(amb.ok, false);
  assert.equal(amb.operations[0].status, 'FAILED');
  assert.match(amb.reason, /table-anchor-ambiguous/);
  assert.equal(p.writes(), 0);
  const miss = await p.tool.handler(req([ft('GH256 nowhere', { shading: '#112233' })]));
  assert.equal(miss.ok, false);
  assert.match(miss.reason, /table-anchor-not-found/);
  assert.equal(p.writes(), 0);
  const ok = await p.tool.handler(req([ft('GH256 Twin', { shading: '#112233' }, 'GH256 Item')]));
  assert.equal(ok.ok, true, JSON.stringify(ok));
  const t = await p.readTables();
  assert.deepEqual(t[1].format.shading.fill, hex('#112233'), 'the anchored table (the one holding GH256 Item) got it');
  assert.equal(t[0].format.shading, null, 'the other table holding the same cell text is untouched');
});

test('PER-TABLE (#270) and GH #291: a merge-free TARGET table is formatted next to a vertically merged neighbour; a vertically merged TARGET table is now formatted too, the merge kept; a write that breaks the merge is never APPLIED', async () => {
  const ok = product({}, tables(true));
  const applied = await ok.tool.handler(req([ft('GH256 Name', { shading: '#112233' })]));
  assert.equal(applied.ok, true, JSON.stringify(applied));
  assert.equal(applied.operations[0].status, 'APPLIED');
  const p = product({}, tables(true));
  const before = JSON.stringify(p.store.durableDoc.tables.map(t => t.rows.map(r => r.map(c => [c.rowspan || 1, !!c.vContinue]))));
  const r = await p.tool.handler(req([ft('GH256 VA', { shading: '#112233' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.operations[0].status, 'APPLIED');
  assert.equal(JSON.stringify(p.store.durableDoc.tables.map(t => t.rows.map(r0 => r0.map(c => [c.rowspan || 1, !!c.vContinue])))), before, 'every vertical merge kept');
  const broken = product({ fault: 'tf-break-vmerge' }, tables(true));
  const rb = await broken.tool.handler(req([ft('GH256 VA', { shading: '#112233' })]));
  assert.notEqual(rb.ok, true, JSON.stringify(rb));
  assert.notEqual(rb.operations[0].status, 'APPLIED', JSON.stringify(rb.operations[0]));
});

test('PRE-DISPATCH refusals never open a session: a bare { text } table, a double border style and a step with no axis refuse before any driver opens', async () => {
  for (const step of [{ intent: 'format_table', table: { text: 'GH256 Name' }, shading: '#330066' }, ft('GH256 Name', { borders: { top: { style: 'double', size: 4, color: '#FF0000' } } }), ft('GH256 Name', {})]) {
    const p = product();
    const r = await p.tool.handler(req([step]));
    assert.ok(['REFUSED_SHAPE', 'UNKNOWN_FIELD'].includes(r.code), `${JSON.stringify(step)} -> ${r.code}`);
    assert.equal(r.dispatched, false);
    assert.equal(p.opens(), 0);
  }
});
