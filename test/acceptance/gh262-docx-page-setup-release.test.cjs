'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #262 (DOCX P1 WRITE): set_page_setup served by docx_execute_task, narrowed (at the time this file
 * was written) to margins + portrait size only -- #263 later widened the same schema to accept
 * columns/startPageNumber too (own test file: gh263-docx-page-setup-columns-startpage-release.test.cjs);
 * this file's own acceptance surface stays margins + size. Design and the measured reason for every
 * rule: docs/GH262-PAGE-SETUP-WRITE-DECISION.md. Offline only: the product handler runs the REAL docx
 * adapter/executor/lifecycle on test/fakes/fake-docx-driver.cjs (its page-setup model follows the W11
 * LIVE probes: US Letter portrait defaults, docx-w11-page-setup.test.cjs's own fixture); the READ side
 * is the product docx_read read_page_setup (GH #223) on the SAME store. The LIVE acceptance is
 * test/live/docx-page-setup-write-product.live.cjs (not run by the author).
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const doc = () => [{ text: 'GH262 Heading', level: 1, props: PLAIN }, { text: 'GH262 Body.', props: PLAIN }];
const req = (steps, extra = {}) => ({ fileId: FILE, task: { id: 't', steps, ...extra } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };

// US Letter portrait default (fake driver's own starting state, same as the real DS's own default --
// test/fakes/fake-docx-driver.cjs:172's comment). Written AS-IS this is a no-op SATISFIED (the GH #231b
// lesson: a value equal to the default never dispatches) -- NARROW is the non-default grid point this
// file's execution/retry cases use to force a real APPLY + stored round-trip.
const LETTER = { width: 12240, height: 15840 };
const A4 = { width: 11906, height: 16838 };
const DEFAULT_MARGINS = { top: 1440, right: 1440, bottom: 1440, left: 1440 };
const NARROW_MARGINS = { top: 720, right: 720, bottom: 720, left: 720 };

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

test('SCHEMA: both grid sizes (US Letter, ISO A4) and a full-margins object are each independently accepted', () => {
  for (const size of [LETTER, A4]) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, size }]));
    assert.equal(v.ok, true, `${JSON.stringify(size)} -> ${v.code} ${v.reason}`);
  }
  const v2 = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: NARROW_MARGINS }]));
  assert.equal(v2.ok, true, JSON.stringify(v2));
  const v3 = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, size: LETTER, margins: NARROW_MARGINS }]));
  assert.equal(v3.ok, true, JSON.stringify(v3));
});

test('RED: a size outside the grid (any third width/height pair, including a landscape-shaped mix of the two allowed sizes) is REFUSED_SHAPE pre-dispatch, 0 dispatch', () => {
  const bad = [{ width: 9999, height: 9999 }, { width: LETTER.width, height: A4.height }, { width: A4.width, height: LETTER.height }, { width: LETTER.height, height: LETTER.width }];
  for (const size of bad) {
    const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, size }]));
    assert.equal(v.code, 'REFUSED_SHAPE', JSON.stringify(size));
    assert.equal(v.dispatched, false);
  }
});

test('RED: a margin value outside 0..2880 twips is REFUSED_SHAPE pre-dispatch; a margins object missing any of the four sides is REFUSED_SHAPE', () => {
  const tooFar = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: { top: 2881, right: 720, bottom: 720, left: 720 } }]));
  assert.equal(tooFar.code, 'REFUSED_SHAPE');
  const negative = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: { top: -1, right: 720, bottom: 720, left: 720 } }]));
  assert.equal(negative.code, 'REFUSED_SHAPE');
  const partial = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: { top: 720, right: 720, bottom: 720 } }]));
  assert.equal(partial.code, 'REFUSED_SHAPE');
});

test('RED: neither size nor margins named -> REFUSED_SHAPE ("at least one"), 0 dispatch', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 } }]));
  assert.equal(v.code, 'REFUSED_SHAPE');
  assert.match(v.reason, /at least one/);
});

test('RED: section is required and bounded (0 <= index < of)', () => {
  const noSection = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', margins: NARROW_MARGINS }]));
  assert.equal(noSection.code, 'REFUSED_SHAPE');
  const badSection = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 1, of: 1 }, margins: NARROW_MARGINS }]));
  // adapter-level plan() refuses this (index >= of); product schema itself only bounds index>=0 -- the
  // REAL planner (called from validateDocxWriteRequest's own pre-dispatch plan-alone step) is the second,
  // authoritative layer, same discipline as rewriteSelectorBranch elsewhere in this file.
  assert.equal(badSection.ok, false, JSON.stringify(badSection));
  // columns/startPageNumber WERE UNKNOWN_FIELD here (#263's own scope, not yet built when this test was
  // written) -- #263 widened the schema to accept them; their own RED/GREEN coverage is now
  // gh263-docx-page-setup-columns-startpage-release.test.cjs, not duplicated here.
});

// ---- execution (the fake transport) --------------------------------------------------------------------------------------------

test('EXECUTE: a value equal to the US Letter portrait default (margins 1440 all around, no size named) is a no-op SATISFIED -- the GH #231b lesson, named here rather than surprising a reader', async () => {
  const p = product();
  const r = await p.tool.handler(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: DEFAULT_MARGINS }]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
  assert.deepEqual(r.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(r.persistence.state, 'NOT_REQUIRED');
  assert.equal(p.writes(), 0);
});

test('EXECUTE: NARROW margins + A4 (both non-default) -> APPLIED, FORCESAVE_CONFIRMED; the fresh read_page_setup confirms both, byte-for-byte', async () => {
  const p = product();
  const r = await p.tool.handler(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, size: A4, margins: NARROW_MARGINS }]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED']);
  assert.equal(r.operations[0].l1, 'PASS');
  assert.equal(r.operations[0].l3, 'PASS');
  assert.equal(r.persistence.state, 'FORCESAVE_CONFIRMED');
  const sections = await p.readPageSetup();
  assert.equal(sections[0].pageWidth, A4.width);
  assert.equal(sections[0].pageHeight, A4.height);
  assert.deepEqual(sections[0].margins, NARROW_MARGINS);
});

test('RETRY: the identical NARROW-margins write again -> SATISFIED, 0 writes, no re-dispatch of an already-terminal state', async () => {
  const p = product();
  const step = { intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: NARROW_MARGINS };
  const first = await p.tool.handler(req([step]));
  assert.equal(first.ok, true, JSON.stringify(first).slice(0, 300));
  assert.deepEqual(first.operations.map(o => o.status), ['APPLIED']);
  const writesAfterFirst = p.writes();
  const retry = await p.tool.handler(req([step], { id: 't2' }));
  assert.equal(retry.ok, true, JSON.stringify(retry).slice(0, 300));
  assert.deepEqual(retry.operations.map(o => o.status), ['SATISFIED']);
  assert.equal(retry.persistence.state, 'NOT_REQUIRED');
  assert.equal(p.writes(), writesAfterFirst, 'the retry wrote nothing new');
});

test('MALFORMED_TASK is unreachable for this intent\'s own fields -- RED CONTROL: a non-integer margin value is REFUSED_SHAPE (type-checked by the schema), not silently coerced', () => {
  const v = D.validateDocxWriteRequest(req([{ intent: 'set_page_setup', section: { index: 0, of: 1 }, margins: { top: '720', right: 720, bottom: 720, left: 720 } }]));
  assert.equal(v.code, 'REFUSED_SHAPE');
});
