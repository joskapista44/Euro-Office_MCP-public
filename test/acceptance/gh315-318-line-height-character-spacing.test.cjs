'use strict';

/**
 * GH #315 (exact/atLeast line height) + GH #318 (character spacing) -- RELEASED at the schema layer
 * (docx-write-tools.cjs's FORMAT_TEXT_VALUE, after the live probe below confirmed the unit hypotheses).
 *
 * Unit hypotheses (twips: round(pt*20)) are LIVE-MEASURED -- zero ran test/live/gh315-318-line-height-
 * character-spacing-probe.live.cjs on the disposable the production host DS (internal tracker, agents/an internal branch/
 * gh315-318-probe-20261027/probe1.json): both twips hypotheses hold, and GetSpacingLineRule() exists and
 * round-trips exactly.
 *
 * Most of this file's tests still drive the family/adapter directly (`withSession({driver, ...},
 * docxAdapter, createExecutor(task))`, the test/acceptance/docx-m5-constraints.test.cjs pattern) to prove
 * the FAMILY's own logic (pairing/mutual-exclusion/enum validation, write-body shape, read/compare
 * symmetry) against the fake driver -- that part is unaffected by the schema gate opening. The PRODUCT
 * section near the end goes through the REAL schema (`D.validateDocxWriteRequest`) and the full
 * `D.createDocxWriteTool` handler (gh231-format-text-paragraph-props.test.cjs's own `product()` pattern),
 * to prove the three new fields are actually reachable end-to-end now, not just at the family layer.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const TARGET = 'GH315318 TARGET PARAGRAPH';
const OTHER = 'GH315318 OTHER PARAGRAPH';
const RANGE_HOST = 'GH315318 range host NEEDLE text';
const doc = () => [
  { text: TARGET, props: { ...PLAIN } },
  { text: OTHER, props: { ...PLAIN } },
  { text: RANGE_HOST, props: { ...PLAIN } },
];
const store = () => createStore(doc());

async function run(s, task) {
  let t = 0;
  const driver = createFakeDocxDriver(s, {});
  const result = await withSession({ driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const step = (selector, format) => ({ intent: 'format_text', selector, format });
const planDetail = steps => { const r = planOperations({ id: 't', steps }); return r.ok ? 'OK' : r.detail; };

function readFormattingOf(paragraphs, text) {
  const inv = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  return inv.blocks.find(b => b.text === text).format;
}

async function appliedThenSatisfied(s, task) {
  const first = await run(s, task);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  const second = await run(s, task);
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.dispatch], [0, 0, 0], 'retry after reopening: 0 write / 0 command');
  return first;
}

// ---- SCHEMA (plan-time) ----------------------------------------------------------------------

test('SCHEMA: lineSpacingRule and lineHeight must be given together, neither alone', () => {
  assert.match(planDetail([step({ text: TARGET }, { lineSpacingRule: 'exact' })]), /must be given together/);
  assert.match(planDetail([step({ text: TARGET }, { lineHeight: 24 })]), /must be given together/);
  assert.equal(planDetail([step({ text: TARGET }, { lineSpacingRule: 'exact', lineHeight: 24 })]), 'OK');
  assert.equal(planDetail([step({ text: TARGET }, { lineSpacingRule: 'atLeast', lineHeight: 12 })]), 'OK');
});

test('SCHEMA: lineSpacingRule only accepts exact/atLeast -- not auto (that is the existing lineSpacing field) and not a made-up value', () => {
  assert.match(planDetail([step({ text: TARGET }, { lineSpacingRule: 'auto', lineHeight: 12 })]), /must be one of exact, atLeast/);
  assert.match(planDetail([step({ text: TARGET }, { lineSpacingRule: 'zsiraf', lineHeight: 12 })]), /must be one of exact, atLeast/);
});

test('SCHEMA: lineHeight must be a finite number', () => {
  assert.match(planDetail([step({ text: TARGET }, { lineSpacingRule: 'exact', lineHeight: 'twelve' })]), /lineHeight must be a number/);
});

test('SCHEMA: lineSpacingRule/lineHeight cannot be combined with the existing lineSpacing (same property, different rule)', () => {
  assert.match(planDetail([step({ text: TARGET }, { lineSpacingRule: 'exact', lineHeight: 24, lineSpacing: 1.5 })]), /cannot be combined with lineSpacing/);
});

test('SCHEMA: characterSpacing has no pairing requirement -- a plain number, positive, negative, or zero', () => {
  assert.equal(planDetail([step({ text: TARGET }, { characterSpacing: 1.5 })]), 'OK');
  assert.equal(planDetail([step({ text: TARGET }, { characterSpacing: -2 })]), 'OK');
  assert.equal(planDetail([step({ text: TARGET }, { characterSpacing: 0 })]), 'OK');
});

test('SCHEMA: lineSpacingRule+lineHeight and characterSpacing compose freely in one format_text step', () => {
  assert.equal(planDetail([step({ text: TARGET }, { lineSpacingRule: 'exact', lineHeight: 24, characterSpacing: 1 })]), 'OK');
});

// ---- END-TO-END: whole-paragraph write (GH #315) ---------------------------------------------

test('APPLIED+PRESERVATION+RETRY: lineSpacingRule exact + lineHeight, whole paragraph; read_formatting symmetry', async () => {
  const s = store();
  const first = await appliedThenSatisfied(s, { id: 't', steps: [step({ text: TARGET }, { lineSpacingRule: 'exact', lineHeight: 24 })] });
  const read = readFormattingOf(s.durable, TARGET);
  assert.equal(read.lineSpacingRule, 'exact');
  assert.equal(read.lineHeight, 24);
  // the sibling paragraph was never written -- this harness's inventory projection carries NO
  // `.format` at all for an untouched paragraph (null, not a defaults object); that is a harness
  // characteristic unrelated to this candidate's own fields.
  assert.equal(readFormattingOf(s.durable, OTHER), null);
  assert.equal(first.x.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
});

test('APPLIED+RETRY: lineSpacingRule atLeast + a non-integer-point lineHeight', async () => {
  const s = store();
  await appliedThenSatisfied(s, { id: 't', steps: [step({ text: TARGET }, { lineSpacingRule: 'atLeast', lineHeight: 13.5 })] });
  const read = readFormattingOf(s.durable, TARGET);
  assert.equal(read.lineSpacingRule, 'atLeast');
  assert.equal(read.lineHeight, 13.5);
});

// ---- END-TO-END: run-level write (GH #318) ----------------------------------------------------

test('APPLIED+RETRY: characterSpacing, whole paragraph (run-level write via the paragraph\'s own range)', async () => {
  const s = store();
  await appliedThenSatisfied(s, { id: 't', steps: [step({ text: TARGET }, { characterSpacing: 1.5 })] });
  const read = readFormattingOf(s.durable, TARGET);
  assert.equal(read.characterSpacing, 1.5);
});

test('APPLIED+RETRY: characterSpacing negative (condensed) is a distinct, observable write', async () => {
  const s = store();
  await appliedThenSatisfied(s, { id: 't', steps: [step({ text: TARGET }, { characterSpacing: -2 })] });
  assert.equal(readFormattingOf(s.durable, TARGET).characterSpacing, -2);
});

test('characterSpacing: an untouched run already reads as the observable 0 baseline -- a wanted 0 is SATISFIED on the first try, no write', async () => {
  const { x, m } = await run(store(), { id: 't', steps: [step({ text: TARGET }, { characterSpacing: 0 })] });
  assert.ok(x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(x.appliedOperations));
  assert.equal(m.writes, 0);
});

test('APPLIED: lineSpacingRule+lineHeight and characterSpacing together in one step, both land independently', async () => {
  const s = store();
  await appliedThenSatisfied(s, { id: 't', steps: [step({ text: TARGET }, { lineSpacingRule: 'exact', lineHeight: 30, characterSpacing: -1 })] });
  const read = readFormattingOf(s.durable, TARGET);
  assert.equal(read.lineSpacingRule, 'exact');
  assert.equal(read.lineHeight, 30);
  assert.equal(read.characterSpacing, -1);
});

// ---- RANGE (matchText) write: the SAME paraPr/run block the whole-paragraph path shares --------

test('APPLIED+RETRY: matchText (range) write carries lineSpacingRule/lineHeight to the PARENT paragraph, exactly like a plain paragraph-level key', async () => {
  const s = store();
  const needle = 'NEEDLE';
  await appliedThenSatisfied(s, { id: 't', steps: [{ intent: 'format_text', selector: { text: RANGE_HOST }, matchText: needle, format: { lineSpacingRule: 'exact', lineHeight: 18 } }] });
  const read = readFormattingOf(s.durable, RANGE_HOST);
  assert.equal(read.lineSpacingRule, 'exact');
  assert.equal(read.lineHeight, 18);
});

test('APPLIED+RETRY: matchText (range) write applies characterSpacing only to the matched substring\'s run', async () => {
  const s = store();
  const needle = 'NEEDLE';
  await appliedThenSatisfied(s, { id: 't', steps: [{ intent: 'format_text', selector: { text: RANGE_HOST }, matchText: needle, format: { characterSpacing: 2 } }] });
  const hit = s.durable.find(p => p.text === RANGE_HOST);
  const runs = Array.isArray(hit.runs) ? hit.runs : [{ text: hit.text, props: hit.props || {} }];
  const onNeedle = runs.find(r => r.text === needle);
  const outsideNeedle = runs.find(r => r.text !== needle);
  assert.equal(onNeedle.props.characterSpacing, 2, JSON.stringify(runs));
  assert.notEqual(outsideNeedle && outsideNeedle.props.characterSpacing, 2);
});

// ---- PRODUCT (schema layer, REAL docx-write-tools.cjs handler) --------------------------------

function product() {
  const s = store();
  const createDriver = () => createFakeDocxDriver(s, { saveObserver: true });
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const tool = D.createDocxWriteTool({ configSource, gate: createSessionGate(), runtime, createDriver });
  return { s, tool };
}
const FILE = '1303627';
const req = (format, extra = {}) => ({ fileId: FILE, task: { steps: [{ intent: 'format_text', selector: { text: TARGET }, format, ...extra }] } });

test('PRODUCT SCHEMA: D.validateDocxWriteRequest now accepts the three released fields (the schema gate is OPEN)', () => {
  assert.equal(D.validateDocxWriteRequest(req({ lineSpacingRule: 'exact', lineHeight: 24 })).ok, true);
  assert.equal(D.validateDocxWriteRequest(req({ lineSpacingRule: 'atLeast', lineHeight: 13.5 })).ok, true);
  assert.equal(D.validateDocxWriteRequest(req({ characterSpacing: -2 })).ok, true);
});

test('PRODUCT SCHEMA: the schema itself still rejects an unpaired lineSpacingRule/lineHeight and a bad enum, same detail as the family (defence in depth, same message surfaces at two layers)', () => {
  const a = D.validateDocxWriteRequest(req({ lineSpacingRule: 'exact' }));
  assert.equal(a.ok, false);
  assert.match(a.reason || a.detail || '', /must be given together/);
  const b = D.validateDocxWriteRequest(req({ lineSpacingRule: 'auto', lineHeight: 12 }));
  assert.equal(b.ok, false);
});

test('PRODUCT SCHEMA: an out-of-bound lineHeight/characterSpacing is rejected at the JSON-schema layer itself (before the family ever sees it)', () => {
  assert.equal(D.validateDocxWriteRequest(req({ lineSpacingRule: 'exact', lineHeight: 0 })).ok, false);
  assert.equal(D.validateDocxWriteRequest(req({ characterSpacing: 5000 })).ok, false);
});

test('PRODUCT END-TO-END: lineSpacingRule+lineHeight through the REAL docx_execute_task handler -- APPLIED, then retry SATISFIED, read_formatting symmetry', async () => {
  const { s, tool } = product();
  const first = await tool.handler(req({ lineSpacingRule: 'exact', lineHeight: 24 }));
  assert.equal(first.ok, true, JSON.stringify(first));
  assert.ok((first.operations || []).every(o => o.status === 'APPLIED'), JSON.stringify(first.operations));
  const read = readFormattingOf(s.durable, TARGET);
  assert.equal(read.lineSpacingRule, 'exact');
  assert.equal(read.lineHeight, 24);
  const retry = await tool.handler(req({ lineSpacingRule: 'exact', lineHeight: 24 }));
  assert.ok((retry.operations || []).every(o => o.status === 'SATISFIED'), JSON.stringify(retry.operations));
});

test('PRODUCT END-TO-END: characterSpacing through the REAL docx_execute_task handler -- APPLIED, then retry SATISFIED', async () => {
  const { s, tool } = product();
  const first = await tool.handler(req({ characterSpacing: -2 }));
  assert.equal(first.ok, true, JSON.stringify(first));
  assert.ok((first.operations || []).every(o => o.status === 'APPLIED'));
  assert.equal(readFormattingOf(s.durable, TARGET).characterSpacing, -2);
  const retry = await tool.handler(req({ characterSpacing: -2 }));
  assert.ok((retry.operations || []).every(o => o.status === 'SATISFIED'));
});
