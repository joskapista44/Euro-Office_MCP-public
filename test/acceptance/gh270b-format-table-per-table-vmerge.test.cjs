'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #270 finding (g270v1 G4 / G8): format_table refused a table that carries NO vertical merge because ANOTHER table of the document had one -- table-format.cjs used table.cjs's
 * DOCUMENT-WIDE tableRead, while replace_table_cell and the column family have been per-table reconciled since W06-3 (LIVE-measured READ_DIVERGENCE: a merge-free table's two reads agree regardless
 * of what another table carries). This pins the same narrowing for format_table: only the TARGET table's own vertical merge decides. A table WITH a merge stays FAIL_CLOSED (the merged grid + a table-level
 * format is not measured, so it is not guessed). Offline on the fake; the LIVE is a maintainer's.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, rows, cols) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(`${tag} r${r}c${c}`))) });
/** a 3 x 2 table whose rows 1-2 of column 0 are one vertically merged cell (owner rowspan 2, continuation placeholder) */
const mergedGrid = tag => {
  const t = grid(tag, 3, 2);
  t.rows[1][0] = { ...t.rows[1][0], rowspan: 2 };
  t.rows[2][0] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] };
  return t;
};
const tc = text => ({ tableCell: { text } });
const ft = (text, fields) => ({ intent: 'format_table', table: tc(text), ...fields });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

function product(tables) {
  const store = createStore([{ text: 'Heading', level: 1, props: PLAIN }], { doc: { tables } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  return { store, tool, readTables, writes: () => store.saves };
}

test('PER-TABLE: a table WITHOUT a vertical merge is formatted although another table of the document carries one; the merged table is untouched; the retry writes nothing', async () => {
  const p = product([mergedGrid('VM'), grid('PL', 2, 2)]);
  const mergedBefore = JSON.stringify((await p.readTables())[0]);
  const r = await p.tool.handler(req([ft('PL r0c0', { shading: '#FFFF00' })]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 500));
  const t = await p.readTables();
  assert.ok(t[1].format.shading, 'the plain table carries the shading');
  assert.equal(JSON.stringify(t[0]), mergedBefore, 'the merged table is untouched');
  const saves = p.writes();
  assert.equal((await p.tool.handler(req([ft('PL r0c0', { shading: '#FFFF00' })]))).ok, true);
  assert.equal(p.writes(), saves, 'the retry writes nothing');
});

test('PER-TABLE: the same holds for every axis -- borders and a width on the merge-free table next to a merged one', async () => {
  const p = product([mergedGrid('VM'), grid('PL', 2, 2)]);
  const step = ft('PL r1c1', { borders: { top: { style: 'single', size: 4, color: '#FF0000' } }, width: 9000 });
  const r = await p.tool.handler(req([step]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 500));
});

test('TARGET MERGE (GH #291): format_table on the table that CARRIES the vertical merge is applied -- the merge (rowspan + continuation) and every cell kept, the other table untouched, the retry writes nothing', async () => {
  const p = product([mergedGrid('VM'), grid('PL', 2, 2)]);
  const plainBefore = JSON.stringify(p.store.durableDoc.tables[1]);
  for (const cellText of ['VM r0c0', 'VM r1c1']) {
    const r = await p.tool.handler(req([ft(cellText, { shading: '#FFFF00' })]));
    assert.equal(r.ok, true, `${cellText}: ${JSON.stringify(r).slice(0, 400)}`);
  }
  const vm = p.store.durableDoc.tables[0];
  assert.deepEqual([vm.rows[1][0].rowspan, vm.rows[2][0].vContinue], [2, true]);
  assert.deepEqual(vm.rows.map(r => r.map(c => c.paras.map(x => x.text).join('|'))), [['VM r0c0', 'VM r0c1'], ['VM r1c0', 'VM r1c1'], ['', 'VM r2c1']]);
  assert.equal(JSON.stringify(p.store.durableDoc.tables[1]), plainBefore);
});

test('DOCUMENT WITH ONLY A MERGED TABLE (GH #291): formatted, the merge kept', async () => {
  const p = product([mergedGrid('VM')]);
  const r = await p.tool.handler(req([ft('VM r0c1', { borders: { insideH: { style: 'single', size: 4, color: '#0070C0' } }, width: 7200 })]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 400));
  assert.equal(p.store.durableDoc.tables[0].rows[1][0].rowspan, 2);
});
