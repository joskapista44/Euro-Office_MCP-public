'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * Release 3 integration (2/2): ONE format_table step schema from the two chains -- the #256 / #257 / #257b axes (borders, shading, cellMargins, indent, width) AND the #258 `style` name, together in the
 * same step (SOLO stays) -- and the #258 blank-tableAnchor gate standing for the WHOLE step, whichever axes it names. The per-chain acceptance tests (gh256, gh257, gh257b, gh258) pin each side;
 * this file pins the union and the gate. Offline on the fake; the LIVE repeat (g256v2 / g257bv1 / g258v1 on the merged tree) is a maintainer's.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const grid = (tag, rows, cols) => ({ rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => cell(`${tag} r${r}c${c}`))) });
const STYLES = [{ name: 'Table Grid', styleId: 701 }];
const tc = (text, tableAnchor) => ({ tableCell: tableAnchor === undefined ? { text } : { text, tableAnchor } });
const SINGLE = (size, color) => ({ style: 'single', size, color });
const AXES = {
  borders: { borders: { top: SINGLE(4, '#FF0000') } }, shading: { shading: '#330066' }, cellMargins: { cellMargins: { top: 100, right: 100, bottom: 100, left: 100 } },
  indent: { indent: 200 }, width: { width: 9000 }, style: { style: 'Table Grid' },
};
const ft = (text, fields, tableAnchor) => ({ intent: 'format_table', table: tc(text, tableAnchor), ...fields });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product() {
  const store = createStore([{ text: 'Heading', level: 1, props: PLAIN }], { doc: { tables: [grid('R3A', 2, 2), grid('R3B', 2, 2)], tableStyles: STYLES } });
  const createDriver = () => createFakeDocxDriver(store, { saveObserver: true });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readTables = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_tables' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.tables; };
  return { store, tool, readTables, writes: () => store.saves };
}

test('UNION: each of the six axes alone, and every pair of {style} x {another axis}, is ONE accepted step; the published schema carries all six with only intent / table required', () => {
  for (const [k, f] of Object.entries(AXES)) assert.equal(verdict([ft('R3A r0c0', f)]).ok, true, k);
  for (const k of ['borders', 'shading', 'cellMargins', 'indent', 'width']) assert.equal(verdict([ft('R3A r0c0', { ...AXES.style, ...AXES[k] })]).ok, true, `style + ${k}`);
  assert.equal(verdict([ft('R3A r0c0', Object.assign({}, ...Object.values(AXES)))]).ok, true, 'all six together');
  const s = D.inputSchema().properties.task.properties.steps.items.oneOf.find(x => x.properties.intent.const === 'format_table');
  assert.deepEqual(Object.keys(s.properties).sort(), ['borders', 'cellMargins', 'id', 'indent', 'intent', 'shading', 'style', 'table', 'width']);
  assert.deepEqual(s.required, ['intent', 'table']);
  assert.equal(s.additionalProperties, false);
});

test('UNION: a step naming NO axis, an empty borders object, a style that is not a non-empty string, an unknown field -> refused pre-dispatch; format_table stays SOLO', () => {
  assert.equal(verdict([ft('R3A r0c0', {})]).code, 'REFUSED_SHAPE');
  assert.match(verdict([ft('R3A r0c0', {})]).reason, /borders, shading, cellMargins, indent, width, style/);
  assert.equal(verdict([ft('R3A r0c0', { borders: {} })]).code, 'REFUSED_SHAPE');
  for (const style of ['', '   ', 7, null]) assert.equal(verdict([ft('R3A r0c0', { style })]).dispatched, false, JSON.stringify(style));
  assert.equal(verdict([ft('R3A r0c0', { style: '' })]).ok, false);
  assert.equal(verdict([ft('R3A r0c0', { ...AXES.shading, fontSize: 12 })]).code, 'UNKNOWN_FIELD');
  const solo = verdict([ft('R3A r0c0', AXES.shading), { intent: 'rewrite_paragraph', selector: { text: 'Heading' }, text: 'X' }]);
  assert.equal(solo.code, 'REFUSED_SHAPE');
  assert.match(solo.reason, /must be the only step/);
});

test('GATE: a blank (whitespace-only) tableAnchor is REFUSED_SHAPE for EVERY axis form -- not only the style form -- and a real anchor is fine', () => {
  for (const [k, f] of Object.entries(AXES)) {
    const blank = verdict([ft('R3A r0c0', f, '   ')]);
    assert.equal(blank.code, 'REFUSED_SHAPE', `${k} with a blank anchor -> ${blank.code}`);
    assert.match(blank.reason, /tableAnchor must be a non-empty cell text/);
    assert.equal(blank.dispatched, false);
    assert.equal(verdict([ft('R3A r0c0', f, 'R3A r1c1')]).ok, true, `${k} with a real anchor`);
  }
  assert.equal(verdict([ft('R3A r0c0', Object.assign({}, AXES.style, AXES.borders), '\t')]).code, 'REFUSED_SHAPE', 'the merged form is gated too');
});

test('WRITE: style + borders + shading in ONE step applies once on the fake -- the table reads back with the style NAME and the formatting axes; the other table is untouched; the retry writes nothing', async () => {
  const p = product();
  const before = JSON.stringify((await p.readTables())[1]);
  const step = ft('R3A r0c0', { ...AXES.style, ...AXES.borders, ...AXES.shading });
  const r = await p.tool.handler(req([step]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 600));
  const t = await p.readTables();
  assert.equal(t[0].format.style.name, 'Table Grid');
  assert.ok(t[0].format.borders && t[0].format.borders.top, 'the border axis reads back');
  assert.ok(t[0].format.shading, 'the shading axis reads back');
  assert.equal(JSON.stringify(t[1]), before, 'the other table is untouched');
  const saves = p.writes();
  const again = await p.tool.handler(req([step]));
  assert.equal(again.ok, true, JSON.stringify(again).slice(0, 400));
  assert.equal(p.writes(), saves, 'the retry writes nothing');
});
