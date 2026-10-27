'use strict';

/**
 * GH #137 U2: format_run on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level
 * behaviour is in test/acceptance/xlsx-m7-f04-format-run.test.cjs (X-F04, already adapter-complete and
 * LIVE-measured); this file is the product-tool level, mirroring gh139-u3-insert-drawing-release.test.cjs.
 * The READ mirror is read_runs (GH #137 U1, src/server/read-tools.cjs), already released; the LIVE proof is
 * test/live/release-gh137-u2-format-run-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const RED = '#FF0000';
const book = () => [{ name: 'S', cells: { A1: { v: 'abcdef' }, A2: { v: 'control' }, A3: { f: '=1+1', v: '2' } } }];
const runA = (font = { bold: true, color: RED }) => ({ start: 1, length: 3, font });
const runB = (font = { name: 'Calibri', size: 11 }) => ({ start: 4, length: 3, font });
const run = (over = {}) => ({ intent: 'format_run', id: 's', sheet: 'S', cell: 'A1', runs: [runA(), runB()], ...over });
const statuses = r => (r.operations || []).map(o => o.status);
const cellOf = (store, sheet, addr) => store.durable.find(x => x.name === sheet).cells[addr];

test('GH #137 U2: format_run is a released intent (allowlist + accepted set + schema + constraints), not a sole step, the release gate passes', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('format_run'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('format_run'));
  assert.equal(W.SOLE_STEP.includes('format_run'), false);
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].sort());
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.equal(T.writeReleaseGate({ xlsxAllowlist: [...T.ACCEPTED_XLSX_WRITE_ALLOWLIST].filter(i => i !== 'format_run'), docxAllowlist: [] }).ok, false, 'format_run missing');
});

test('GH #137 U2: the schema + CONSTRAINTS refuse a malformed step before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(run()).ok, true);
  const bad = [
    [run({ runs: undefined }), 'runs missing'],
    [run({ runs: [] }), 'empty runs array'],
    [run({ runs: [{ start: 0, length: 3, font: { bold: true } }] }), 'start below 1'],
    [run({ runs: [{ start: 1, length: 0, font: { bold: true } }] }), 'length below 1'],
    [run({ runs: [{ start: 1.5, length: 3, font: { bold: true } }] }), 'non-integer start'],
    [run({ runs: [{ start: 1, length: 3, font: { bold: true }, extra: 1 }] }), 'a run carries an extra key'],
    [run({ runs: [{ start: 1, length: 3 }] }), 'font missing'],
    [run({ runs: [{ start: 1, length: 3, font: { underline: 'single' } }] }), 'underline is X-F05 scope, not here'],
    [run({ runs: [{ start: 1, length: 3, font: { strikeout: true } }] }), 'strikeout is X-F05 scope, not here'],
    [run({ runs: [{ start: 1, length: 3, font: { color: 'red' } }] }), 'color not a hex string'],
    [run({ runs: [{ start: 1, length: 3, font: { size: 0 } }] }), 'size below 1'],
    [run({ runs: [{ start: 1, length: 3, font: { size: 410 } }] }), 'size above 409'],
    [run({ runs: [{ start: 1, length: 3, font: { name: 'x'.repeat(101) } }] }), 'name over 100 chars'],
    [run({ runs: [{ start: 1, length: 3, font: { bold: 'yes' } }] }), 'bold not a boolean'],
    [run({ cell: 'A1:A2' }), 'cell must be a single cell, not a range'],
    [run({ sheet: '' }), 'an empty sheet name'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #137 U2: the family\'s own plan() catches what the schema cannot -- overlapping runs, an empty font object -- PRE_DISPATCH, 0 writes', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const overlap = await h.tool.handler(req([run({ runs: [{ start: 1, length: 3, font: { bold: true } }, { start: 3, length: 2, font: { italic: true } }] })]));
  assert.equal(overlap.ok, false, J(overlap).slice(0, 300));
  assert.match(J(overlap), /runs must not overlap/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
  const emptyFont = await h.tool.handler(req([run({ runs: [{ start: 1, length: 3, font: {} }] })]));
  assert.equal(emptyFont.ok, false, J(emptyFont).slice(0, 300));
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #137 U2: two runs APPLIED, both hold; retry SATISFIED 0 dispatch; a character outside every run (L3) and the control cell are untouched', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([run()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  const c = cellOf(h.store, 'S', 'A1');
  assert.deepEqual(c.runs[0], { start: 1, length: 3, font: { bold: true, italic: null, name: 'Calibri', size: 11, color: 16711680 } });
  assert.deepEqual(c.runs[1], { start: 4, length: 3, font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0 } });
  assert.equal(cellOf(h.store, 'S', 'A2').v, 'control');
  assert.equal(cellOf(h.store, 'S', 'A2').runs, undefined, 'the untouched control cell is not even given a runs key');
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([run()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #137 U2: a cell over the RUN_MAP_CHAR_CEILING is refused named xlsx-format-run-text-too-long, nothing dispatched', async () => {
  const h = harness({ store: createStore([{ name: 'S', cells: { A1: { v: 'x'.repeat(300) } } }]) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([run()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-format-run-text-too-long/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #137 U2: a run on a formula cell is refused named xlsx-format-run-on-formula-unmeasured, nothing dispatched', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([run({ cell: 'A3' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-format-run-on-formula-unmeasured/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('GH #137 U2: format_run is ALREADY in TIERED_INTENTS (adapter.cjs, pre-existing -- not something this card\'s product wiring adds or needs to add) -- the large-book path is the tiered observation, not a blanket refusal', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(run()).tiered, true);
});
