'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const formatting = require('../../src/adapters/docx/families/formatting.cjs');
const { createStore, createFakeDocxDriver, htmlOf, runPayload, formatOf } = require('../fakes/fake-docx-driver.cjs');

/**
 * W-08 (zero internal tracker/#ref, this session): format_text's three NEW keys -- color
 * ('#RRGGBB'), highlight (a named OOXML highlight colour or 'none'), vertAlign (baseline /
 * subscript / superscript) -- on the SAME pipeline docx-w11-partial-format.test.cjs already
 * proves for bold/italic/underline/strike/fontFamily/fontSize/language: plan validation, range
 * (matchText) AND whole-paragraph writes, FIRST applies / SECOND is a 0-write no-op, L1/L3
 * preservation, negative controls.
 */

// color/highlight/vertAlign get a concrete baseline here too, the same modeling choice PLAIN
// already makes for fontFamily/fontSize/language: every run has SOME observable value for every
// text property, matching this suite's existing fixtures (a write is only provable "holds" if the
// baseline was observable to begin with -- the same rule bold/italic/etc. already enforce).
// color's baseline is 'auto' -- the MEASURED default (: GetColor() answers null, a
// defined state, on a run with no explicit colour; the driver reports this as 'auto'), not a
// guessed hex value.
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU', color: 'auto', highlight: 'none', vertAlign: 'baseline' };
const P1 = 'Alpha TARGET omega.';
const P2 = 'Second body paragraph.';
const P5 = 'Mix MIXED end.';
const doc = () => [
  { text: P1, props: { ...PLAIN }, runs: [{ text: 'Alpha ', props: { ...PLAIN } }, { text: 'TARGET', props: { ...PLAIN, color: '#FF0000' } }, { text: ' omega.', props: { ...PLAIN } }] },
  { text: P2, props: { ...PLAIN } },
  { text: P5, props: { ...PLAIN }, runs: [{ text: 'Mix ', props: { ...PLAIN } }, { text: 'MI', props: { ...PLAIN, highlight: 'yellow' } }, { text: 'XED', props: { ...PLAIN } }, { text: ' end.', props: { ...PLAIN } }] },
];

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  return { result, driver, m: result.metrics, x: result.execution };
}
const one = step => ({ id: 'w08', steps: [step] });
const rangeTask = (matchText, format, selector = { text: P1 }) => one({ intent: 'format_text', selector, matchText, format });
const wholeTask = (format, selector = { text: P2 }) => one({ intent: 'format_text', selector, format });
const charsOf = p => runPayload(p).runs.flatMap(r => [...r.text].map(() => r.format));

async function firstAndSecond(store, task) {
  const first = await run(store, task);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), 'L3');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', 'L2');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(store, task);
  assert.equal(second.result.ok, true);
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED));
  assert.equal(second.m.writes, 0);
  assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  return { first, second };
}

// ---- plan validation ----------------------------------------------------------------------------

test('normalizeWanted: color/highlight/vertAlign are validated, canonicalized, and reject bad values', () => {
  assert.deepEqual(formatting.normalizeWanted({ color: '#ff00aa' }).wanted, { color: '#FF00AA' }, 'hex uppercased');
  assert.equal(formatting.normalizeWanted({ color: 'red' }).ok, false, 'not #RRGGBB');
  assert.equal(formatting.normalizeWanted({ color: '#ff00a' }).ok, false, 'wrong length');
  assert.deepEqual(formatting.normalizeWanted({ highlight: 'yellow' }).wanted, { highlight: 'yellow' });
  assert.equal(formatting.normalizeWanted({ highlight: 'orange' }).ok, false, 'not a known OOXML highlight name');
  assert.deepEqual(formatting.normalizeWanted({ vertAlign: 'superscript' }).wanted, { vertAlign: 'superscript' });
  assert.equal(formatting.normalizeWanted({ vertAlign: 'super' }).ok, false, 'not baseline/subscript/superscript');
  assert.equal(formatting.normalizeWanted({ color: '#FF0000', bogus: 1 }).ok, false, 'unknown key still refused');
});

// ---- RANGE (matchText) ---------------------------------------------------------------------------

test('range color: FIRST applies to exactly the range (integral), SECOND is a 0-write no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('TARGET', { color: '#00FF00' }));
  const chars = charsOf(store.durable[0]);
  const at = P1.indexOf('TARGET');
  for (let i = 0; i < P1.length; i += 1) {
    const inside = i >= at && i < at + 'TARGET'.length;
    assert.equal(chars[i].color, inside ? '#00FF00' : 'auto', `colour at ${i}`);
  }
});

test('range color: FROM THE MEASURED "auto" BASELINE to #RRGGBB -> APPLIED, retry SATISFIED (c)', async () => {
  const store = createStore([{ text: P1, props: { ...PLAIN }, runs: [{ text: 'Alpha ', props: { ...PLAIN } }, { text: 'TARGET', props: { ...PLAIN } }, { text: ' omega.', props: { ...PLAIN } }] }]);
  const before = charsOf(store.durable[0]);
  const at = P1.indexOf('TARGET');
  for (let i = at; i < at + 'TARGET'.length; i += 1) assert.equal(before[i].color, 'auto', 'TARGET starts at the measured default');
  await firstAndSecond(store, rangeTask('TARGET', { color: '#9933CC' }));
  const chars = charsOf(store.durable[0]);
  for (let i = 0; i < P1.length; i += 1) assert.equal(chars[i].color, i >= at && i < at + 'TARGET'.length ? '#9933CC' : 'auto', `colour at ${i}`);
});

test('range highlight + vertAlign: both apply together, independent of colour', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('TARGET', { highlight: 'cyan', vertAlign: 'superscript' }));
  const chars = charsOf(store.durable[0]);
  const at = P1.indexOf('TARGET');
  for (let i = at; i < at + 'TARGET'.length; i += 1) {
    assert.equal(chars[i].highlight, 'cyan');
    assert.equal(chars[i].vertAlign, 'superscript');
  }
});

test('range: a MIXED highlight range is a definite not-terminal state -> APPLY, then proven uniform', async () => {
  const store = createStore(doc());
  const inv = formatting.rangeState;
  const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
  const built = normalizeInventory({ html: htmlOf(doc()), paragraphs: doc().map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p) })) });
  const before = inv(built.blocks[2], 'MIXED');
  assert.deepEqual(before.mixed, ['highlight'], 'MI highlight=yellow + XED unset');
  await firstAndSecond(store, rangeTask('MIXED', { highlight: 'green' }, { text: P5 }));
  const chars = charsOf(store.durable[2]);
  const at = P5.indexOf('MIXED');
  for (let i = 0; i < P5.length; i += 1) assert.equal(chars[i].highlight, i >= at && i < at + 5 ? 'green' : 'none', `highlight at ${i}`);
});

// ---- WHOLE PARAGRAPH ------------------------------------------------------------------------------

test('whole-paragraph: color + highlight + vertAlign together persist on the paragraph format', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, wholeTask({ color: '#123456', highlight: 'magenta', vertAlign: 'subscript' }));
  const p = store.durable.find(x => x.text === P2);
  const f = formatOf(p);
  assert.equal(f.color, '#123456');
  assert.equal(f.highlight, 'magenta');
  assert.equal(f.vertAlign, 'subscript');
});

test('whole-paragraph: highlight "none" is a VALID value (removes highlighting), not an empty format', async () => {
  const store = createStore([{ text: P2, props: { ...PLAIN, highlight: 'yellow' } }]);
  await firstAndSecond(store, wholeTask({ highlight: 'none' }));
  assert.equal(formatOf(store.durable.find(x => x.text === P2)).highlight, 'none');
});

// ---- L3: preservation of the OTHER two keys when only one is requested ---------------------------

test('range: requesting only color leaves highlight/vertAlign on the SAME range untouched (L3)', async () => {
  const store = createStore([
    { text: P1, props: { ...PLAIN }, runs: [{ text: 'Alpha ', props: { ...PLAIN } }, { text: 'TARGET', props: { ...PLAIN, highlight: 'red', vertAlign: 'superscript' } }, { text: ' omega.', props: { ...PLAIN } }] },
    { text: P2, props: { ...PLAIN } },
  ]);
  await firstAndSecond(store, rangeTask('TARGET', { color: '#ABCDEF' }));
  const chars = charsOf(store.durable[0]);
  const at = P1.indexOf('TARGET');
  for (let i = at; i < at + 'TARGET'.length; i += 1) {
    assert.equal(chars[i].color, '#ABCDEF', 'requested property set');
    assert.equal(chars[i].highlight, 'red', 'un-requested property preserved');
    assert.equal(chars[i].vertAlign, 'superscript', 'un-requested property preserved');
  }
});

// ---- negative controls: FAIL_CLOSED before any dispatch -------------------------------------------

test('negatives: plan rejects an invalid color/highlight/vertAlign before any dispatch', async () => {
  for (const format of [{ color: 'notahex' }, { highlight: 'rainbow' }, { vertAlign: 'up' }]) {
    const { x, driver } = await run(createStore(doc()), one({ intent: 'format_text', selector: { text: P2 }, format }));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(format));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('negatives: an unobservable colour on a range (non-run paragraph) -> FAIL_CLOSED, 0 dispatch', async () => {
  const d = doc().map(p => (p.text === P1 ? { ...p, nonRunChildren: 1 } : p));
  const { x, driver } = await run(createStore(d), rangeTask('TARGET', { color: '#ABCDEF' }));
  assert.equal(x.appliedOperations[0].status, OPERATION_STATUS.FAILED);
  assert.equal(driver.calls.dispatch, 0);
});

test('lowering: color/highlight/vertAlign write direct range calls, no editor backdoor', () => {
  const body = formatting.rangeCommandBody(P1, 'TARGET', { color: '#AABBCC', highlight: 'blue', vertAlign: 'subscript' });
  assert.match(body, /range\.SetColor\(Api\.HexColor\(/);
  assert.match(body, /range\.SetHighlight\(wanted\.highlight\)/);
  assert.match(body, /range\.SetVertAlign\(wanted\.vertAlign\)/);
  assert.doesNotMatch(body, /frame\.evaluate|window\./, 'no editor backdoor');
});
