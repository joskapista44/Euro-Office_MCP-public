'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #263 (DOCX P1 WRITE): set_page_setup widened to columns + startPageNumber (GH #262 shipped
 * margins + portrait size only). Design and the measured reason for every rule:
 * docs/GH263-PAGE-SETUP-COLUMNS-STARTPAGE-DECISION.md. Offline only: the product handler runs the
 * REAL docx adapter/executor/lifecycle on test/fakes/fake-docx-driver.cjs; the READ side is the
 * product docx_read read_page_setup (GH #223) on the SAME store. The LIVE acceptance is
 * test/live/docx-page-setup-columns-startpage-write-product.live.cjs (not run by the author).
 */

const FILE = '1303627';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const doc = () => [{ text: 'GH263 Heading', level: 1, props: PLAIN }, { text: 'GH263 Body.', props: PLAIN }];
const req = (steps, extra = {}) => ({ fileId: FILE, task: { id: 't', steps, ...extra } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

// fake driver's own starting state (test/fakes/fake-docx-driver.cjs:181/186): columnsCount:1,
// startPageNumber:1 -- writing exactly these is a no-op (the GH #231b/#262 lesson, reapplied).
// {count:2, space:720}/startPageNumber:7 is the W11 LIVE evidence tuple (docs/evidence/
// docx-w11-page-setup-live-acceptance-rebased-PASS-20261012.json) -- non-default, forces a real APPLY.
const COLUMNS = { count: 2, space: 720 };
const START_PAGE_NUMBER = 7;
const DEFAULT_COLUMNS = { count: 1, space: 0 };
const DEFAULT_START_PAGE_NUMBER = 1;

function product(script = {}) {
  const store = createStore(doc());
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true, ...script }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readPageSetup = async () => { const r = await read.handler({ fileId: FILE, operation: 'read_page_setup' }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result.sections; };
  return { store, drivers, tool, readPageSetup, writes: () => store.saves };
}

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: columns {count,space} and startPageNumber are each independently accepted, and together with size/margins', () => {
  const v1 = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns: COLUMNS }]));
  assert.equal(v1.ok, true, JSON.stringify(v1));
  const v2 = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, startPageNumber: START_PAGE_NUMBER }]));
  assert.equal(v2.ok, true, JSON.stringify(v2));
  const v3 = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: { top: 720, right: 720, bottom: 720, left: 720 }, columns: COLUMNS, startPageNumber: START_PAGE_NUMBER }]));
  assert.equal(v3.ok, true, JSON.stringify(v3));
});

test('RED: columns.count outside 1..12, columns.space outside 0..2880, or a partial columns object (missing count or space) is REFUSED_SHAPE pre-dispatch, 0 dispatch', () => {
  const bad = [
    { count: 0, space: 720 }, { count: 13, space: 720 },
    { count: 2, space: -1 }, { count: 2, space: 2881 },
    { count: 2 }, { space: 720 },
  ];
  for (const columns of bad) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(columns));
    assert.equal(v.dispatched, false);
  }
});

test('RED: startPageNumber outside 1..9999, zero, negative or non-integer is REFUSED_SHAPE pre-dispatch, 0 dispatch', () => {
  for (const startPageNumber of [0, -1, 10000, 1.5]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, startPageNumber }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(startPageNumber));
    assert.equal(v.dispatched, false);
  }
});

test('RED: neither size, margins, columns nor startPageNumber named -> REFUSED_SHAPE ("at least one"), 0 dispatch', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 } }]));
  assert.equal(v.code, 'REFUSED_SHAPE');
  assert.match(v.reason, /at least one/);
  assert.equal(v.dispatched, false);
});

// ---- execution (the fake transport) --------------------------------------------------------------------------------------------

test('EXECUTION: writing the DEFAULT columns/startPageNumber (count:1,space:0 / 1) is a no-op -- SATISFIED, zero writes (the GH #231b/#262 lesson, named here)', async () => {
  const p = product();
  const receipt = await p.tool.handler(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns: DEFAULT_COLUMNS, startPageNumber: DEFAULT_START_PAGE_NUMBER }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), 0);
});

test('EXECUTION: non-default columns + startPageNumber together -> APPLIED/FORCESAVE_CONFIRMED; a fresh read_page_setup confirms byte-for-byte; retry is SATISFIED with zero new writes', async () => {
  const p = product();
  const receipt = await p.tool.handler(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns: COLUMNS, startPageNumber: START_PAGE_NUMBER }]));
  assert.equal(receipt.ok, true, JSON.stringify(receipt));
  assert.deepEqual((receipt.operations || []).map(o => o.status), ['APPLIED']);
  assert.equal(receipt.persistence && receipt.persistence.state, 'FORCESAVE_CONFIRMED');
  const writesAfterFirst = p.writes();
  assert.ok(writesAfterFirst > 0, 'expected at least one save after an APPLIED step');

  const sections = await p.readPageSetup();
  const sec = sections.find(s => s.index === 0);
  assert.equal(sec.columnsCount, COLUMNS.count, JSON.stringify(sec));
  assert.deepEqual(sec.columnsSpaces, Array.from({ length: COLUMNS.count - 1 }, () => COLUMNS.space), JSON.stringify(sec));
  assert.equal(sec.startPageNumber, START_PAGE_NUMBER, JSON.stringify(sec));

  const retry = await p.tool.handler(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns: COLUMNS, startPageNumber: START_PAGE_NUMBER }]));
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.deepEqual((retry.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(p.writes(), writesAfterFirst, 'retry must not produce a new write');
});

test('RED CONTROL: a malformed columns value (a string, not an object) is rejected at the type level, not silently coerced', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns: '2x720' }]));
  assert.equal(v.ok, false, JSON.stringify(v));
  assert.equal(v.dispatched, false);
});

module.exports = { FILE, COLUMNS, START_PAGE_NUMBER, DEFAULT_COLUMNS, DEFAULT_START_PAGE_NUMBER };
