'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const formatting = require('../../src/adapters/docx/families/formatting.cjs');
const { createStore, createFakeDocxDriver, htmlOf, runPayload, formatOf } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 W1.1 offline acceptance: format_text on an exact SUBSTRING (`matchText`) through the TARGET
 * path (M1 -> M2 -> M3 -> DOCX adapter), on the fake transport whose run model follows the live
 * probe measurements (runs with offsets and effective formatting; a mixed property is absent from
 * the whole-paragraph read). FIRST applies, SECOND is a 0-write no-op in a new session.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const H = 'Intro heading';
const P1 = 'Alpha TARGET omega.';
const P2 = 'Second body paragraph.';
const P3 = 'Elsewhere TARGET again.';
const P4 = 'DUP and DUP here.';
const P5 = 'Mix MIXED end.';
const doc = () => [
  { text: H, level: 1, props: { ...PLAIN, bold: true } },
  // TARGET differs from its surroundings (italic), so preservation inside and outside is observable
  { text: P1, props: { ...PLAIN }, runs: [{ text: 'Alpha ', props: { ...PLAIN } }, { text: 'TARGET', props: { ...PLAIN, italic: true } }, { text: ' omega.', props: { ...PLAIN } }] },
  { text: P2, props: { ...PLAIN } },
  { text: P3, props: { ...PLAIN } },
  { text: P4, props: { ...PLAIN } },
  { text: P5, props: { ...PLAIN }, runs: [{ text: 'Mix ', props: { ...PLAIN } }, { text: 'MI', props: { ...PLAIN, bold: true } }, { text: 'XED', props: { ...PLAIN } }, { text: ' end.', props: { ...PLAIN } }] },
];

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const one = step => ({ id: 'w11', steps: [step] });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => {
  const e = x.appliedOperations[i].evidence;
  return (e && e.evidence && e.evidence.reason) || (e && e.reason);
};
const rangeTask = (matchText, format, selector = { text: P1 }) => one({ intent: 'format_text', selector, matchText, format });
/** Effective formatting of every character of a durable paragraph (the run model). */
const charsOf = p => runPayload(p).runs.flatMap(r => [...r.text].map(() => r.format));

async function firstAndSecond(store, task) {
  const first = await run(store, task);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.equal(first.m.writes, task.steps.length);
  assert.equal(first.m.commands, task.steps.length);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  const p1 = task.steps.filter(s => s.intent === 'format_text' || s.intent === 'rewrite_paragraph').length;
  assert.equal((first.x.preservationVerification || []).length, p1, 'integral L3 for every P1 operation');
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), 'L3');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', 'L2');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  assert.equal(first.driver.calls.forceSave, 1, 'one barrier, at task end');
  const second = await run(store, task);
  assert.equal(second.result.ok, true);
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED));
  assert.equal(second.m.writes, 0);
  assert.equal(second.m.commands, 0);
  assert.equal(second.driver.calls.forceSave, 0);
  assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

// ---- the observation ----------------------------------------------------------------------------

test('inventory: runs are attached only when they provably cover the live paragraph', () => {
  const d = doc();
  const paragraphs = d.map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p) }));
  const covered = normalizeInventory({ html: htmlOf(d), paragraphs });
  assert.equal(covered.blocks[1].runs.length, 3);
  assert.equal(covered.blocks[1].liveText, P1);
  const withChild = paragraphs.map(p => (p.text === P1 ? { ...p, nonRunChildren: 1 } : p));
  assert.equal(normalizeInventory({ html: htmlOf(d), paragraphs: withChild }).blocks[1].runs, null, 'a non-run child -> no range claim');
  const short = paragraphs.map(p => (p.text === P1 ? { ...p, runs: p.runs.slice(0, 2), runsCoverText: false } : p));
  assert.equal(normalizeInventory({ html: htmlOf(d), paragraphs: short }).blocks[1].runs, null, 'runs not covering the text -> no range claim');
});

test('range state: uniform / mixed / unknown per property, 0 / 1 / >1 occurrences', () => {
  const d = doc();
  const inv = normalizeInventory({ html: htmlOf(d), paragraphs: d.map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p) })) });
  const uni = formatting.rangeState(inv.blocks[1], 'TARGET');
  assert.equal(uni.ok, true);
  assert.equal(uni.format.italic, true);
  assert.deepEqual(uni.mixed, []);
  const mixed = formatting.rangeState(inv.blocks[5], 'MIXED');
  assert.deepEqual(mixed.mixed, ['bold'], 'MI bold + XED plain');
  assert.equal(mixed.format.bold, undefined, 'a mixed property never gets a value');
  assert.equal(formatting.rangeState(inv.blocks[1], 'NOPE').reason, 'range-absent');
  assert.equal(formatting.rangeState(inv.blocks[4], 'DUP').reason, 'range-ambiguous-in-paragraph');
});

// ---- FIRST / SECOND -----------------------------------------------------------------------------

test('partial format: FIRST applies to exactly the range (integral), SECOND is a 0-write no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('TARGET', { bold: true, fontSize: 14 }));
  const chars = charsOf(store.durable[1]);
  const at = P1.indexOf('TARGET');
  for (let i = 0; i < P1.length; i += 1) {
    const inside = i >= at && i < at + 'TARGET'.length;
    assert.equal(chars[i].bold, inside, `bold at ${i}`);
    assert.equal(chars[i].fontSize, inside ? 14 : 11, `size at ${i}`);
    assert.equal(chars[i].italic, inside, `italic (not requested) preserved at ${i}`);
  }
  assert.ok(charsOf(store.durable[3]).every(f => f.bold === false), 'the same text in ANOTHER paragraph is untouched');
});

test('partial format: the same substring in another paragraph is discriminated by the paragraph context', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('TARGET', { underline: true }, { text: P3 }));
  assert.ok(charsOf(store.durable[1]).every(f => f.underline === false), 'P1 TARGET untouched');
});

test('partial format: a MIXED range is a definite not-terminal state -> APPLY, then proven uniform', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('MIXED', { bold: true }, { text: P5 }));
  const chars = charsOf(store.durable[5]);
  const at = P5.indexOf('MIXED');
  for (let i = 0; i < P5.length; i += 1) assert.equal(chars[i].bold, i >= at && i < at + 5, `bold at ${i}`);
});

test('partial format: a paragraph property with matchText goes to the paragraph (source parity)', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('TARGET', { bold: true, alignment: 'center' }));
  assert.equal(formatOf(store.durable[1]).alignment, 'center');
});

// W-09: indentation + shading are paragraph-level (not TEXT_KEYS) on the matchText path too.
test('partial format: indentation + shading with matchText go to the whole paragraph', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, rangeTask('TARGET', { bold: true, indentLeft: 24, shading: 'ffff00' }));
  const f = formatOf(store.durable[1]);
  assert.equal(f.indentLeft, 24);
  assert.equal(f.shading, 'FFFF00');
});

test('REF: rewrite -> partial format of the rewritten paragraph (dependsOn), one session; retry SATISFIED', async () => {
  const store = createStore(doc());
  const task = { id: 'w11-ref', steps: [
    // the rewrite's own retry identity needs the structural selector (text + kind + before), as in Wave 1's mixed chain
    { intent: 'rewrite_paragraph', selector: { text: P2, kind: 'paragraph', before: P1 }, text: 'Rewritten body with KEY words.' },
    { intent: 'format_text', selector: { ref: 's0' }, matchText: 'KEY', format: { bold: true }, dependsOn: ['s0'] },
  ] };
  await firstAndSecond(store, task);
  const p = store.durable.find(x => x.text === 'Rewritten body with KEY words.');
  const chars = charsOf(p);
  const at = p.text.indexOf('KEY');
  for (let i = 0; i < p.text.length; i += 1) assert.equal(chars[i].bold, i >= at && i < at + 3);
});

// ---- negative controls: FAIL_CLOSED before any dispatch -----------------------------------------

test('negatives: absent / duplicate in the paragraph / unobservable / not addressable / stale -> FAIL_CLOSED, 0 dispatch', async () => {
  const cases = [
    [rangeTask('NOPE', { bold: true }), doc(), 'range-absent'],
    [rangeTask('DUP', { bold: true }, { text: P4 }), doc(), 'range-ambiguous-in-paragraph'],
    [rangeTask('TARGET', { bold: true }), doc().map(p => (p.text === P1 ? { ...p, nonRunChildren: 1 } : p)), 'range-format-unobservable'],
    [rangeTask('TARGET', { language: 'de-DE' }), doc().map(p => (p.text === P1 ? { ...p, runs: p.runs.map((r, i) => (i === 1 ? { ...r, props: { ...r.props, language: undefined } } : r)) } : p)), 'range-terminal-property-unobservable'],
    [rangeTask('TARGET', { bold: true }), [...doc(), { text: `${P1} Longer copy.`, props: { ...PLAIN } }], 'format-parent-not-addressable'],
    [rangeTask('TARGET', { bold: true }, { text: P1, index: 1 }), doc(), 'format-target-physical_locator_refused'],
  ];
  for (const [task, d, want] of cases) {
    const { x, driver } = await run(createStore(d), task);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
});

test('negatives: no paragraph context / empty matchText / contained selector -> plan fails before any dispatch', async () => {
  for (const step of [
    { intent: 'format_text', matchText: 'TARGET', format: { bold: true } },
    { intent: 'format_text', selector: { text: P1 }, matchText: '', format: { bold: true } },
    { intent: 'format_text', selector: { text: 'TARGET', contains: true }, matchText: 'TARGET', format: { bold: true } },
  ]) {
    const { x, driver } = await run(createStore(doc()), one(step));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(step));
    assert.equal(driver.calls.dispatch, 0);
  }
});

// ---- negative controls: the write happened but is not the terminal contract -------------------

test('faults at L1: no mutation / part / paragraph-level only / another occurrence -> never APPLIED, no barrier', async () => {
  for (const fault of ['ignore', 'format-partial', 'range-paragraph-only', 'range-other-occurrence']) {
    const { x, result, driver } = await run(createStore(doc()), rangeTask('TARGET', { bold: true, fontSize: 14 }), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, fault);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_NOT_PERSISTED, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('faults at L3: the range holds but something else changed -> NOT APPLIED (integral)', async () => {
  for (const fault of ['range-also-other', 'range-spill', 'range-inside-other-prop']) {
    const { x, driver } = await run(createStore(doc()), rangeTask('TARGET', { bold: true }), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: L1 (the range) held`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
  const changed = await run(createStore(doc()), rangeTask('TARGET', { bold: true }), { fault: 'range-change-text' });
  assert.notEqual(status(changed.x), OPERATION_STATUS.APPLIED, 'text changed by a format -> not applied');
});

test('lowering: the substring write is one callCommand through the common dispatch, paragraph-scoped', () => {
  const body = formatting.rangeCommandBody(P1, 'TARGET', { bold: true });
  assert.match(body, /paragraph\.Search\(matchText, true\)/);
  assert.match(body, /hits\.length !== 1\) return/);
  assert.doesNotMatch(body, /frame\.evaluate|window\./, 'no editor backdoor');
});
