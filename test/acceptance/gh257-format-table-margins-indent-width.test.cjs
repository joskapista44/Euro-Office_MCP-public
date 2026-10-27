'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #257 (DOCX P1 WRITE, table-level cellMargins / indent / width; the release itself is an OWNER DECISION, alongside #228..#256): the format_table step
 * (GH #256) gains the three TWIPS axes. Design and the measured reason for every rule: docs/GH257-FORMAT-TABLE-MARGINS-INDENT-WIDTH-DECISION.md. The unit is
 * TWIPS on write AND on read (read_tables returns { type:'dxa', w } for the margins and the width, a bare number for the indent). Offline only: the product
 * handler runs the REAL docx adapter / executor / lifecycle on test/fakes/fake-docx-driver.cjs (applyTableFormat models the measured write persistence and the
 * READ shape); the READ side is the product docx_read read_tables (GH #222) on the SAME store. The LIVE acceptance is
 * test/live/gh257-format-table-margins-indent-width-product.live.cjs (not run by the author). `style` (GH #258) joined this schema in the release-3 merge.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const OUTSIDE = 'GH257 outside paragraph';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: OUTSIDE, props: PLAIN }];
const tables = () => [
  { rows: [[cell('GH257 Name'), cell('GH257 Qty')], [cell('GH257 Alpha'), cell('GH257 Twin')]] },
  { rows: [[cell('GH257 Item'), cell('GH257 Price')], [cell('GH257 Pear'), cell('GH257 Twin')]] },
];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const ft = (text, fields, tableAnchor) => ({ intent: 'format_table', table: tc(text, tableAnchor), ...fields });
// the values docs/evidence/docx-w04-table-format-live-acceptance-rebased-PASS-20261012.json measured on the adapter path
const MARGINS = { top: 50, right: 60, bottom: 70, left: 80 };
const WANTED = { cellMargins: MARGINS, indent: 100, width: 2000 };
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));
const dxa = w => ({ type: 'dxa', w });
const readMargins = m => ({ top: dxa(m.top), right: dxa(m.right), bottom: dxa(m.bottom), left: dxa(m.left) });

function product(tabs = tables()) {
  const store = createStore(doc(), { doc: { tables: tabs } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  return { store, tool, readTables, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: cellMargins (all four edges), indent and width are accepted, each alone and together, next to borders / shading and with a tableAnchor', () => {
  for (const fields of [WANTED, { cellMargins: MARGINS }, { indent: 100 }, { width: 2000 }, { indent: 1 }, { indent: 0 }, { cellMargins: { top: 1, right: 1, bottom: 1, left: 1 } }, { cellMargins: { top: 0, right: 0, bottom: 0, left: 0 } },
    { ...WANTED, shading: '#330066', borders: { top: { style: 'single', size: 4, color: '#FF0000' } } }, { cellMargins: { top: 31680, right: 31680, bottom: 31680, left: 31680 }, indent: 31680, width: 31680 }]) {
    const r = verdict([ft('GH257 Name', fields)]);
    assert.equal(r.ok, true, `${JSON.stringify(fields)} -> ${r.code} ${r.reason}`);
  }
  assert.equal(verdict([ft('GH257 Twin', WANTED, 'GH257 Item')]).ok, true);
});

test('SCHEMA: cellMargins needs ALL FOUR edges as non-negative integers (0 is accepted since GH #257b) -- a missing / extra / negative / fractional / string / out-of-range edge is refused pre-dispatch', () => {
  const cases = [
    [{ top: 1, right: 1, bottom: 1 }, 'REFUSED_SHAPE'], [{ top: 1, right: 1, bottom: 1, left: 1, start: 1 }, 'UNKNOWN_FIELD'], [{}, 'REFUSED_SHAPE'],
    [{ top: -1, right: 1, bottom: 1, left: 1 }, 'REFUSED_SHAPE'], [{ top: 1.5, right: 1, bottom: 1, left: 1 }, 'REFUSED_SHAPE'], [{ top: '1', right: 1, bottom: 1, left: 1 }, 'REFUSED_SHAPE'],
    [{ top: 31681, right: 1, bottom: 1, left: 1 }, 'REFUSED_SHAPE'], [{ top: null, right: 1, bottom: 1, left: 1 }, 'REFUSED_SHAPE'], [null, 'REFUSED_SHAPE'], [50, 'REFUSED_SHAPE'],
  ];
  for (const [cellMargins, code] of cases) { const v = verdict([ft('GH257 Name', { cellMargins })]); assert.equal(v.code, code, `${JSON.stringify(cellMargins)} -> ${v.code} ${v.reason}`); assert.equal(v.dispatched, false); }
});

test('SCHEMA: indent is a non-negative and width a POSITIVE integer (twips; a zero indent is accepted since GH #257b, a zero width is not) -- a negative / zero width / fractional / string / unit-carrying / out-of-range value is refused pre-dispatch', () => {
  for (const [fields, code] of [
    [{ indent: -1 }, 'REFUSED_SHAPE'], [{ indent: 1.5 }, 'REFUSED_SHAPE'], [{ indent: '100' }, 'REFUSED_SHAPE'], [{ indent: 31681 }, 'REFUSED_SHAPE'], [{ indent: { w: 100 } }, 'REFUSED_SHAPE'], [{ indent: null }, 'REFUSED_SHAPE'],
    [{ width: 0 }, 'REFUSED_SHAPE'], [{ width: -5 }, 'REFUSED_SHAPE'], [{ width: 2000.5 }, 'REFUSED_SHAPE'], [{ width: '2000' }, 'REFUSED_SHAPE'], [{ width: 31681 }, 'REFUSED_SHAPE'],
    [{ width: { type: 'dxa', w: 2000 } }, 'REFUSED_SHAPE'], [{ width: '50%' }, 'REFUSED_SHAPE'], [{ width: { type: 'pct', w: 50 } }, 'REFUSED_SHAPE'], [{ widthUnit: 'pt', width: 100 }, 'UNKNOWN_FIELD'],
  ]) {
    const v = verdict([ft('GH257 Name', fields)]);
    assert.equal(v.code, code, `${JSON.stringify(fields)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: a step naming NO axis is REFUSED_SHAPE (the reason lists every axis incl. style after the release-3 merge); style is accepted together with the axes (GH #258 merged); format_table is still a SOLO step', () => {
  const none = verdict([ft('GH257 Name', {})]);
  assert.equal(none.code, 'REFUSED_SHAPE');
  assert.match(none.reason, /borders, shading, cellMargins, indent, width, style/);
  assert.equal(verdict([ft('GH257 Name', { ...WANTED, style: 'Table Grid' })]).ok, true, 'style + the axes in ONE step (the merged union)');
  const solo = verdict([ft('GH257 Name', WANTED), { intent: 'rewrite_paragraph', selector: { text: OUTSIDE }, text: 'X' }]);
  assert.equal(solo.code, 'REFUSED_SHAPE');
  assert.match(solo.reason, /must be the only step/);
});

// ---- the product handler, end to end on the fake transport -----------------------------------------------------------

test('WRITE + READ SYMMETRY: the margins, indent and width land on the named table in the SAME twips; read_tables returns { type:dxa, w } for each margin edge and the width, a bare number for the indent; every other axis and table is untouched', async () => {
  const p = product();
  const r = await p.tool.handler(req([ft('GH257 Name', WANTED)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.operations[0].l1, 'PASS');
  assert.equal(r.operations[0].l3, 'PASS');
  assert.equal(r.wholeTask, 'PASS');
  assert.equal(p.writes() >= 1, true);
  const t = await p.readTables();
  assert.deepEqual(t[0].format.cellMargins, readMargins(MARGINS));
  assert.equal(t[0].format.indent, 100);
  assert.deepEqual(t[0].format.width, dxa(2000));
  assert.deepEqual(t[0].format.borders, { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null }, 'the borders (#256) were not requested and are untouched');
  assert.equal(t[0].format.shading, null);
  assert.deepEqual(t[0].format.style, { id: null, name: null });
  assert.deepEqual(t[1].format.cellMargins, { top: null, right: null, bottom: null, left: null }, 'the twin table is untouched');
  assert.equal(t[1].format.indent, null);
  assert.equal(t[1].format.width, null);
  assert.deepEqual(t[0].rows.map(row => row.map(c => c.text)), [['GH257 Name', 'GH257 Qty'], ['GH257 Alpha', 'GH257 Twin']], 'the cells are untouched');
});

test('INDEPENDENT AXES (L3 own-field preservation): a width-only request leaves the margins and indent of the first request; a borders / shading request leaves all three twips axes', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH257 Name', WANTED)]))).ok, true);
  const second = await p.tool.handler(req([ft('GH257 Name', { width: 3000 })]));
  assert.equal(second.ok, true, JSON.stringify(second));
  let t = await p.readTables();
  assert.deepEqual(t[0].format.width, dxa(3000));
  assert.deepEqual(t[0].format.cellMargins, readMargins(MARGINS), 'the margins survive a width-only request');
  assert.equal(t[0].format.indent, 100, 'the indent survives a width-only request');
  const third = await p.tool.handler(req([ft('GH257 Name', { shading: '#330066', borders: { top: { style: 'single', size: 4, color: '#FF0000' } } })]));
  assert.equal(third.ok, true, JSON.stringify(third));
  t = await p.readTables();
  assert.deepEqual(t[0].format.cellMargins, readMargins(MARGINS));
  assert.equal(t[0].format.indent, 100);
  assert.deepEqual(t[0].format.width, dxa(3000));
});

test('RETRY: the SAME task again is SATISFIED / noOp with ZERO new writes (idempotent, restart-safe)', async () => {
  const p = product();
  const task = req([ft('GH257 Name', WANTED)]);
  assert.equal((await p.tool.handler(task)).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves, 'no write on the retry');
});

test('TABLE IDENTITY: tableAnchor selects the table holding the shared cell text; without it the twips request fails CLOSED with no write', async () => {
  const p = product();
  const amb = await p.tool.handler(req([ft('GH257 Twin', { width: 2500 })]));
  assert.equal(amb.ok, false);
  assert.match(amb.reason, /table-anchor-ambiguous/);
  assert.equal(p.writes(), 0);
  const ok = await p.tool.handler(req([ft('GH257 Twin', { width: 2500, indent: 360 }, 'GH257 Item')]));
  assert.equal(ok.ok, true, JSON.stringify(ok));
  const t = await p.readTables();
  assert.deepEqual(t[1].format.width, dxa(2500));
  assert.equal(t[1].format.indent, 360);
  assert.equal(t[0].format.width, null, 'the other table holding the same cell text is untouched');
});

test('PRE-DISPATCH refusals never open a session: a three-edge cellMargins, a zero width and a unit-carrying width refuse before any driver opens', async () => {
  for (const fields of [{ cellMargins: { top: 1, right: 1, bottom: 1 } }, { width: 0 }, { width: { type: 'dxa', w: 2000 } }]) {
    const p = product();
    const r = await p.tool.handler(req([ft('GH257 Name', fields)]));
    assert.equal(r.code, 'REFUSED_SHAPE', JSON.stringify(fields));
    assert.equal(r.dispatched, false);
    assert.equal(p.opens(), 0);
  }
});
