'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #257b (DOCX P1 WRITE): the pre-dispatch refusal of an explicit ZERO on the format_table margin edges and the indent (GH #257) is lifted. GH #257 refused
 * it because the editor might drop an explicit 0 (read back UNSET = a failed L1 AFTER the write was dispatched); the raw probe of g257v1 ROUND 6 measured the
 * opposite (docs/GH257B-FORMAT-TABLE-ZERO-MARGIN-INDENT-DECISION.md). A zero WIDTH stays refused (never probed or requested). Offline only, on the fake driver
 * (applyTableFormat stores the 0 and the READ returns { type:'dxa', w:0 } / 0, the shape the probe measured); the LIVE acceptance is
 * test/live/gh257b-format-table-zero-margin-indent-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = (...texts) => ({ span: 1, paras: texts.map(text => ({ text, props: { ...PLAIN } })) });
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }];
const tables = () => [
  { rows: [[cell('GH257B Name'), cell('GH257B Qty')], [cell('GH257B Alpha'), cell('GH257B Twin')]] },
  { rows: [[cell('GH257B Item'), cell('GH257B Price')], [cell('GH257B Pear'), cell('GH257B Twin')]] },
];
const ft = (text, fields, tableAnchor) => ({ intent: 'format_table', table: { tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } }, ...fields });
const NONZERO = { cellMargins: { top: 50, right: 60, bottom: 70, left: 80 }, indent: 100, width: 2000 };
const ZERO = { cellMargins: { top: 0, right: 0, bottom: 0, left: 0 }, indent: 0 };
const dxa = w => ({ type: 'dxa', w });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product() {
  const store = createStore(doc(), { doc: { tables: tables() } });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  return { store, tool, readTables, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}

test('SCHEMA: a zero margin edge (each edge alone, with the others positive), all four zero, and a zero indent are accepted pre-dispatch', () => {
  for (const edge of ['top', 'right', 'bottom', 'left']) {
    const r = verdict([ft('GH257B Name', { cellMargins: { top: 5, right: 5, bottom: 5, left: 5, [edge]: 0 } })]);
    assert.equal(r.ok, true, `${edge}=0 -> ${r.code} ${r.reason}`);
  }
  for (const fields of [ZERO, { indent: 0 }, { cellMargins: ZERO.cellMargins }, { ...ZERO, width: 2000 }]) {
    const r = verdict([ft('GH257B Name', fields)]);
    assert.equal(r.ok, true, `${JSON.stringify(fields)} -> ${r.code} ${r.reason}`);
  }
});

test('SCHEMA: what stays refused -- a zero WIDTH, a negative / fractional / string zero-lookalike, null -- is REFUSED_SHAPE pre-dispatch', () => {
  for (const fields of [{ width: 0 }, { indent: -1 }, { indent: 0.5 }, { indent: '0' }, { indent: null }, { indent: -0.0001 },
    { cellMargins: { top: -1, right: 0, bottom: 0, left: 0 } }, { cellMargins: { top: 0, right: 0, bottom: 0, left: 0.5 } }, { cellMargins: { top: '0', right: 0, bottom: 0, left: 0 } }, { cellMargins: { top: 0, right: 0, bottom: 0 } }]) {
    const v = verdict([ft('GH257B Name', fields)]);
    assert.equal(v.code, 'REFUSED_SHAPE', `${JSON.stringify(fields)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('WRITE + READ: zero over a non-zero value (the explicit 0 is a CHANGE) is APPLIED; read_tables returns { dxa, 0 } on all four edges and indent 0 -- not null; the width and the other table are untouched', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH257B Name', NONZERO)]))).ok, true);
  const r = await p.tool.handler(req([ft('GH257B Name', ZERO)]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.operations[0].l1, 'PASS');
  assert.equal(r.wholeTask, 'PASS');
  const t = await p.readTables();
  assert.deepEqual(t[0].format.cellMargins, { top: dxa(0), right: dxa(0), bottom: dxa(0), left: dxa(0) });
  assert.equal(t[0].format.indent, 0);
  assert.deepEqual(t[0].format.width, dxa(2000), 'the width of the first request survives');
  assert.deepEqual(t[1].format.cellMargins, { top: null, right: null, bottom: null, left: null });
  assert.equal(t[1].format.indent, null);
});

test('WRITE + READ: zero on a table that never had a value (unset -> explicit 0) is APPLIED, and reads back 0 -- distinguishable from the unset null', async () => {
  const p = product();
  const r = await p.tool.handler(req([ft('GH257B Twin', ZERO, 'GH257B Item')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  const t = await p.readTables();
  assert.deepEqual(t[1].format.cellMargins, { top: dxa(0), right: dxa(0), bottom: dxa(0), left: dxa(0) });
  assert.equal(t[1].format.indent, 0);
  assert.equal(t[0].format.indent, null, 'the other table is still unset (null), not 0');
});

test('MIXED: one zero edge among positive ones, and an indent-only zero, each land exactly and leave the unnamed axes', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ft('GH257B Name', NONZERO)]))).ok, true);
  const mixed = await p.tool.handler(req([ft('GH257B Name', { cellMargins: { top: 0, right: 60, bottom: 70, left: 80 } })]));
  assert.equal(mixed.ok, true, JSON.stringify(mixed));
  let t = await p.readTables();
  assert.deepEqual(t[0].format.cellMargins, { top: dxa(0), right: dxa(60), bottom: dxa(70), left: dxa(80) });
  assert.equal(t[0].format.indent, 100, 'the indent was not named and is unchanged');
  const ind = await p.tool.handler(req([ft('GH257B Name', { indent: 0 })]));
  assert.equal(ind.ok, true, JSON.stringify(ind));
  t = await p.readTables();
  assert.equal(t[0].format.indent, 0);
  assert.deepEqual(t[0].format.cellMargins.top, dxa(0), 'the margins were not named and are unchanged');
});

test('RETRY: the SAME zero task again is SATISFIED / noOp with ZERO new writes -- a stored 0 is recognised as done, not as unset', async () => {
  const p = product();
  const task = req([ft('GH257B Name', ZERO)]);
  assert.equal((await p.tool.handler(task)).ok, true);
  const saves = p.writes();
  const retry = await p.tool.handler(task);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(retry.noOp, true);
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), saves, 'no write on the retry');
});

test('PRE-DISPATCH refusals still never open a session: a zero width and a negative indent refuse before any driver opens', async () => {
  for (const fields of [{ width: 0 }, { indent: -1 }]) {
    const p = product();
    const r = await p.tool.handler(req([ft('GH257B Name', fields)]));
    assert.equal(r.code, 'REFUSED_SHAPE', JSON.stringify(fields));
    assert.equal(r.dispatched, false);
    assert.equal(p.opens(), 0);
  }
});
