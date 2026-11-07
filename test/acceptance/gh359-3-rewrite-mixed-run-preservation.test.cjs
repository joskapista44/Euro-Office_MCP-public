'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #359 (3/3): rewrite_paragraph on a paragraph whose source runs disagree on bold/italic.
 * Measured LIVE (prod DS cell /26, three run-count shapes):
 * the baseline (target.props) OMITS the property the runs disagree on -- the engine's own
 * multi-run TextPr read returns undefined for a mixed value, same as formatOf() already
 * documents for the whole-paragraph formatting family -- and the rewrite's new, single run
 * DETERMINISTICALLY ends up with the FIRST source run's own value for that property, on all
 * three measured shapes (MULTIRUN: bold true/false -> after bold:true; MULTIRUN_PB: bold
 * false/true -> after bold:false; MULTIRUN3: bold+italic mixed across three runs -> after
 * bold:true, italic:false). Before text.cjs's derivedExpectation fallback, the baseline stayed
 * mixed (missing the key) while the after state had a definite value -- a spurious L3 FAIL
 * (DISPATCHED_OUTCOME_UNKNOWN / INTEGRAL_PRESERVATION_NOT_PROVEN) on a write that, in fact,
 * wrote exactly what the engine deterministically produces.
 */

const BASE = { underline: false, strike: false, fontFamily: 'Arial', fontSize: 11, language: 'hu-HU' };
const OLD = 'RW8Mix_AB';
const NEW = 'RW8Mix_Rewritten';

/** MULTIRUN shape: run1 bold, run2 plain -- only `bold` is mixed. */
function multirunDoc() {
  const run1 = { ...BASE, bold: true, italic: false };
  const run2 = { ...BASE, bold: false, italic: false };
  const topProps = { ...BASE, italic: false }; // `bold` intentionally absent: the measured mixed-key omission
  return [{ text: OLD, runs: [{ text: 'RW8Mix_A', props: run1 }, { text: 'B', props: run2 }], props: topProps }];
}

/** MULTIRUN_PB shape: run1 plain, run2 bold -- the reversed order (first run is now the PLAIN one). */
function multirunPbDoc() {
  const run1 = { ...BASE, bold: false, italic: false };
  const run2 = { ...BASE, bold: true, italic: false };
  const topProps = { ...BASE, italic: false };
  return [{ text: OLD, runs: [{ text: 'RW8Mix_A', props: run1 }, { text: 'B', props: run2 }], props: topProps }];
}

/** MULTIRUN3 shape: three runs, bold AND italic both mixed (run1 bold-only, run2 italic-only, run3 plain). */
function multirun3Doc() {
  const run1 = { ...BASE, bold: true, italic: false };
  const run2 = { ...BASE, bold: false, italic: true };
  const run3 = { ...BASE, bold: false, italic: false };
  const topProps = { ...BASE }; // both `bold` and `italic` absent
  return [{ text: OLD, runs: [{ text: 'RW8Mix_', props: run1 }, { text: 'A', props: run2 }, { text: 'B', props: run3 }], props: topProps }];
}

/** 85b2b2b9 (GH #359 kovetes), MERT (7d827dcf #ref, prod DS cell /26): a VEGYES BETUMERET eset --
 * a run-format fontSize PONTBAN jon, a blokk-props (target.props) FELPONTBAN ugyanarra az ertekre --
 * a fallback elotti verzio a KONVERTALATLAN run-format erteket irta az expected-be, es a tarolt/after
 * target.props a FELPONTOS nyers erteket hordozza, tehat egy VALODI egyezes is hamis L3 FAIL-t adott.
 * run1 fontSize:22 (= 11pt a run-format oldalan), run2 fontSize:28 (= 14pt); topProps.fontSize
 * SZANDEKOSAN hianyzik (a mert mixed-key omission). */
function multirunSizeDoc() {
  const { fontSize: _unused, ...baseNoSize } = BASE;
  const run1 = { ...BASE, fontSize: 22, bold: false, italic: false };
  const run2 = { ...BASE, fontSize: 28, bold: false, italic: false };
  const topProps = { ...baseNoSize, bold: false, italic: false };
  return [{ text: OLD, runs: [{ text: 'RW8Mix_A', props: run1 }, { text: 'B', props: run2 }], props: topProps }];
}

function selectorFrom(paragraphs) {
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const b = primary.blocks.find(x => x.text === OLD);
  return { text: OLD, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  return { result, driver };
}

const rewriteTask = selector => ({ id: 'gh359-3', steps: [{ intent: 'rewrite_paragraph', selector, text: NEW }] });

for (const [name, docFn, expectedBold, expectedItalic] of [
  ['MULTIRUN', multirunDoc, true, false],
  ['MULTIRUN_PB', multirunPbDoc, false, false],
]) {
  test(`GH #359 (3/3) ${name}: the engine's first-run collapse is PROVEN, not a spurious L3 FAIL`, async () => {
    const paragraphs = docFn();
    const { result } = await run(createStore(paragraphs), rewriteTask(selectorFrom(paragraphs)), { fault: 'rewrite-collapses-to-first-run' });
    const e = result.execution.appliedOperations[0];
    assert.equal(e.status, OPERATION_STATUS.APPLIED, `${name}: the write is proven, not DISPATCHED_OUTCOME_UNKNOWN`);
    const l3 = result.execution.preservationVerification[0];
    assert.equal(l3.verdict, 'PASS', `${name}: L3 ${JSON.stringify(l3)}`);
    assert.deepEqual(l3.evidence.evidence.formattingCollapsedToFirstRun, ['bold'], `${name}: only bold was mixed`);
    assert.equal(l3.evidence.evidence.expected.bold, expectedBold, `${name}: the derived expectation is the FIRST run's own value`);
    assert.equal(l3.evidence.evidence.after.bold, expectedBold);
    assert.equal(l3.evidence.evidence.after.italic, expectedItalic);
  });
}

test('GH #359 (3/3) MULTIRUN3: three runs, bold AND italic both mixed -- the first run wins on both', async () => {
  const paragraphs = multirun3Doc();
  const { result } = await run(createStore(paragraphs), rewriteTask(selectorFrom(paragraphs)), { fault: 'rewrite-collapses-to-first-run' });
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.APPLIED);
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS', JSON.stringify(l3));
  assert.deepEqual(l3.evidence.evidence.formattingCollapsedToFirstRun.slice().sort(), ['bold', 'italic']);
  assert.equal(l3.evidence.evidence.expected.bold, true, 'the first run (run1) is bold:true, italic:false');
  assert.equal(l3.evidence.evidence.expected.italic, false);
  assert.equal(l3.evidence.evidence.after.bold, true);
  assert.equal(l3.evidence.evidence.after.italic, false);
});

test('GH #359 (3/3) NEG. KONTROLL: if the engine instead collapsed to the SECOND run, L3 stays FAIL -- the fix does not accept any single value', async () => {
  const paragraphs = multirunDoc();
  const { result } = await run(createStore(paragraphs), rewriteTask(selectorFrom(paragraphs)), { fault: 'rewrite-collapses-to-second-run' });
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'FAIL');
  assert.equal(l3.evidence.evidence.expected.bold, true, 'the expectation stays the FIRST run (true) ...');
  assert.equal(l3.evidence.evidence.after.bold, false, '... but the engine actually gave the second run (false)');
});

test('GH #359 (3/3) regression: a single-run paragraph (no mixing possible) is unaffected -- no formattingCollapsedToFirstRun evidence', async () => {
  const PLAIN = { ...BASE, bold: true, italic: true };
  const paragraphs = [{ text: OLD, props: PLAIN }];
  const { result } = await run(createStore(paragraphs), rewriteTask(selectorFrom(paragraphs)));
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS');
  assert.equal(l3.evidence.evidence.formattingCollapsedToFirstRun, undefined, 'single run: nothing to collapse');
  assert.deepEqual(l3.evidence.evidence.expected, PLAIN);
});

test('GH #359 kovetes (85b2b2b9) MULTIRUN_SIZE: a vegyes betumeret run-format(pt)/blokk-props(felpont) '
  + 'egysegatvaltasa a fallback-ben -- a fix elott egy VALODI egyezes is hamis FAIL lett volna', async () => {
  const paragraphs = multirunSizeDoc();
  const { result } = await run(createStore(paragraphs), rewriteTask(selectorFrom(paragraphs)), { fault: 'rewrite-collapses-to-first-run' });
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.APPLIED, 'the write is proven, not DISPATCHED_OUTCOME_UNKNOWN');
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS', JSON.stringify(l3));
  assert.deepEqual(l3.evidence.evidence.formattingCollapsedToFirstRun, ['fontSize'], 'only fontSize was mixed');
  // expected: the FIRST run's own fontSize (22, the raw half-point unit target.props itself uses),
  // converted from the run-format's point unit (11) -- NOT the unconverted 11.
  assert.equal(l3.evidence.evidence.expected.fontSize, 22);
  assert.equal(l3.evidence.evidence.after.fontSize, 22);
});

test('GH #359 kovetes (85b2b2b9) NEG. KONTROLL: ha az engine a MASODIK run meretet adja, L3 FAIL marad '
  + '-- a javitas nem fogad el akarmilyen erteket, csak a (konvertalt) ELSO run meretet', async () => {
  const paragraphs = multirunSizeDoc();
  const { result } = await run(createStore(paragraphs), rewriteTask(selectorFrom(paragraphs)), { fault: 'rewrite-collapses-to-second-run' });
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'FAIL');
  assert.equal(l3.evidence.evidence.expected.fontSize, 22, 'the expectation stays the FIRST run, converted (22) ...');
  assert.equal(l3.evidence.evidence.after.fontSize, 28, '... but the engine actually gave the second run (28)');
});
