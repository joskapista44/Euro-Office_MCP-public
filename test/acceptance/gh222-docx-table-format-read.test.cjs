'use strict';

/**
 * GH #222 (DOCX P1 READ): table-level `format` on the product read_tables surface -- an ADDITIVE
 * field next to the existing `rows` (cell text/format/paraFormats, UNCHANGED). Six axes: borders
 * (top/bottom/left/right/insideH/insideV: value/sz/color), shading, cellMargins (top/right/
 * bottom/left), indent, width, style (identity + name). No new driver reader: the source is
 * inventory.cjs's own tableFormatOf() (W04-4), already measured for the write-side
 * families/table-format.cjs verification -- this file adds a product-surface CONSUMER of it, plus
 * the style NAME resolution (id -> name, the reverse of table-format.cjs's own resolveStyleId()).
 * Decision + the six-axis -> source fail:line table: docs/GH222-TABLE-FORMAT-READ-DECISION.md.
 *
 * Offline only: test/fakes/fake-docx-driver.cjs already simulates the per-field UNOBSERVABLE
 * gaps (script.tableFormatUnreadable / tableBorderEdgeUnreadable / tableStylesUnreadable) that
 * the W04-4 write-side offline tests (docx-w04-table-format.test.cjs) already rely on.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });

function harness(doc, script = {}) {
  const s = createStore([{ text: 'Body', props: { ...PLAIN } }], { doc });
  const b = fakeBroker('docx');
  return createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(s, script) });
}

async function readTables(doc, script) {
  const r = await harness(doc, script).handler({ fileId: FILE, operation: 'read_tables' });
  assert.equal(r.ok, true, JSON.stringify(r));
  return r.result.tables;
}

// ---- every axis, fully measured --------------------------------------------------------------------------------------

test('GH #222: all six axes, fully measured -- borders/shading/cellMargins/indent/width pass through verbatim, style resolves id+name (the REAL id shape: a STRING, "701" -- DS 9.3.5 LIVE evidence, docs/evidence/docx-w04-table-format-live-acceptance-rebased-PASS-20261012.json: <w:tblStyle w:val="701"/>)', async () => {
  const format = {
    borders: {
      top: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 255, g: 0, b: 0 } },
      bottom: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 0, g: 255, b: 0 } },
      left: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 0, g: 0, b: 255 } },
      right: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 255, g: 255, b: 0 } },
      insideH: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 255, g: 0, b: 255 } },
      insideV: { value: 'single', sz: 4, space: 0, color: { auto: false, r: 0, g: 255, b: 255 } },
    },
    shading: { val: 'clear', fill: { r: 0, g: 128, b: 0 }, color: null },
    cellMargins: { top: { type: 'dxa', w: 108 }, right: { type: 'dxa', w: 108 }, bottom: { type: 'dxa', w: 108 }, left: { type: 'dxa', w: 108 } },
    indent: 720,
    width: { type: 'dxa', w: 5000 },
    style: '701',
  };
  const tables = await readTables({ tables: [{ rows: [[cell('a'), cell('b')]], format }] });
  assert.deepEqual(tables[0].format, {
    borders: format.borders, shading: format.shading, cellMargins: format.cellMargins, indent: 720, width: format.width,
    style: { id: '701', name: 'Table Grid' },
  });
  // backward compat: the cell-level shape read_tables already returned is untouched
  assert.deepEqual(tables[0].rows[0].map(c => c.text), ['a', 'b']);
});

test('GH #222: the style id ALSO resolves when it arrives as a plain number -- the lookup is String()-on-both-sides (table-format.cjs\'s own convention), not shape-specific', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: { style: 701 } }] });
  assert.deepEqual(tables[0].format.style, { id: 701, name: 'Table Grid' });
});

// ---- absent vs unread ---------------------------------------------------------------------------------------------

test('GH #222: ABSENT -- no table-format read at all is measured null on every field, never UNOBSERVABLE', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]] }] }); // no `format` key at all
  assert.deepEqual(tables[0].format, {
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    shading: null, cellMargins: { top: null, right: null, bottom: null, left: null }, indent: null, width: null,
    style: { id: null, name: null },
  });
});

test('GH #222: the WHOLE table-format axis unreadable -- every field (incl. style id and name) carries the SAME named reason, never a default', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: { style: 701 } }] }, { tableFormatUnreadable: true });
  assert.deepEqual(tables[0].format, {
    borders: 'UNOBSERVABLE:no-such-getter', shading: 'UNOBSERVABLE:no-such-getter', cellMargins: 'UNOBSERVABLE:no-such-getter',
    indent: 'UNOBSERVABLE:no-such-getter', width: 'UNOBSERVABLE:no-such-getter',
    style: { id: 'UNOBSERVABLE:no-such-getter', name: 'UNOBSERVABLE:no-such-getter' },
  });
});

test('GH #222: ONE border edge unreadable -- the other five edges and every other axis stay independently measured, a single gap never widens', async () => {
  const format = { borders: {}, shading: { val: 'clear', fill: { r: 1, g: 2, b: 3 }, color: null }, style: '701' };
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format }] }, { tableBorderEdgeUnreadable: 'top' });
  assert.equal(tables[0].format.borders.top, 'UNOBSERVABLE:no-such-getter');
  for (const edge of ['bottom', 'left', 'right', 'insideH', 'insideV']) assert.equal(tables[0].format.borders[edge], null, edge);
  assert.deepEqual(tables[0].format.shading, { val: 'clear', fill: { r: 1, g: 2, b: 3 }, color: null });
  assert.deepEqual(tables[0].format.style, { id: '701', name: 'Table Grid' });
});

// ---- style identity/name: three distinct, named gaps ------------------------------------------------------------------

test('GH #222: style id measured but the document-wide tableStyles list is unread -- id stays measured, name is a NAMED, DISTINCT gap', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: { style: '701' } }] }, { tableStylesUnreadable: true });
  assert.deepEqual(tables[0].format.style, { id: '701', name: 'UNOBSERVABLE:style-list-unread' });
});

test('GH #222: style id measured but absent from a (readable) tableStyles list -- a distinct named gap, never the list-unread reason', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: { style: '999' } }] });
  assert.deepEqual(tables[0].format.style, { id: '999', name: 'UNOBSERVABLE:style-id-not-in-list' });
});

test('GH #222: no style set, even when the tableStyles list itself is unread -- the measured absence wins (null), never borrowing the lists own gap', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: {} }] }, { tableStylesUnreadable: true });
  assert.deepEqual(tables[0].format.style, { id: null, name: null });
});

// ---- table identity/order stable, each table's format independent ----------------------------------------------------

test('GH #222: table identity/order stable across two tables, each with its own independent format', async () => {
  const tables = await readTables({ tables: [
    { rows: [[cell('T1')]], format: { style: '701', indent: 100 } },
    { rows: [[cell('T2')]], format: { style: '999', indent: 200 } },
  ] });
  assert.equal(tables.length, 2);
  assert.deepEqual([tables[0].rows[0][0].text, tables[1].rows[0][0].text], ['T1', 'T2']);
  assert.equal(tables[0].format.indent, 100);
  assert.equal(tables[0].format.style.name, 'Table Grid');
  assert.equal(tables[1].format.indent, 200);
  assert.equal(tables[1].format.style.name, 'UNOBSERVABLE:style-id-not-in-list');
});

// ---- schema compatibility: the REAL driver/adapter path, not just the fake table-format fixture ------------------------

test('GH #222: schema compatibility -- the REAL driver/adapter path (the shared read-editors.cjs fixture) keeps rows unchanged and adds format additively', async () => {
  const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
  const { docxEditor } = require('../fakes/read-editors.cjs');
  const ed = docxEditor();
  const b = fakeBroker('docx');
  const t = createReadTool('docx', { configSource: b.source, loadPlaywright: ed.loadPlaywright,
    createDriver: (fmt, deps) => docx.createEuroOfficeDocxDriver({ ...deps, pollMs: 5, closeAckTimeoutMs: 200, timeoutMs: 2000 }) });
  const r = await t.handler({ fileId: FILE, operation: 'read_tables' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.tables[0].rows[0].map(c => c.text), ['a', 'b'], 'the cell-level read_tables output is UNCHANGED');
  assert.ok('format' in r.result.tables[0], 'the new table-level format field is additive, never replacing rows');
  assert.deepEqual(r.result.tables[0].format, {
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    shading: null, cellMargins: { top: null, right: null, bottom: null, left: null }, indent: null, width: null,
    style: { id: null, name: null },
  }, 'the fixture seeds no table.format at all -- measured absent, matching the dedicated ABSENT test above');
});

// ---- RED CONTROL: proves the handler actually reads the documents own list, not a hardcoded/pass-by-construction value ----

test('RED CONTROL: the style NAME is resolved from the documents OWN tableStyles list, not the fixtures default', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: { style: 42 } }], tableStyles: [{ name: 'Custom Grid', styleId: 42 }] });
  assert.equal(tables[0].format.style.name, 'Custom Grid');
  assert.notEqual(tables[0].format.style.name, 'Table Grid', 'must not fall back to the fixtures default when a real, different list is seeded');
});

// ---- REGRESSION GUARD: a real STRING style id must NOT be mistaken for an UNOBSERVABLE reason ------------------------
// (GH #222 review, DS 9.3.5 LIVE evidence: the real tblStyle id IS a string, e.g. "701" -- the
// v1 bug keyed the gap-check on `typeof idVal === 'string'`, which would have misread every real
// measured string id as an UNOBSERVABLE:<reason>. If that type-based check is ever reintroduced in
// place of the prefix check, THIS test fails.)

test('REGRESSION GUARD: a measured string style id that is NOT the UNOBSERVABLE: prefix resolves by lookup, never treated as a gap', async () => {
  const tables = await readTables({ tables: [{ rows: [[cell('a')]], format: { style: '701' } }], tableStyles: [{ name: 'Table Grid', styleId: 701 }] });
  assert.deepEqual(tables[0].format.style, { id: '701', name: 'Table Grid' });
  assert.ok(!String(tables[0].format.style.name).startsWith('UNOBSERVABLE:'), 'a real style must never surface as a gap');
});
