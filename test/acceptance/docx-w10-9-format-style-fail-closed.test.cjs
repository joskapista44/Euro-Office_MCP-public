'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const formatting = require('../../src/adapters/docx/families/formatting.cjs');
const styleDefinition = require('../../src/adapters/docx/families/style-definition.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * W-10-9 (point 2): `format_text {style: <name>}` must FAIL_CLOSED PRE-DISPATCH when
 * `<name>` is not a style that exists in the document -- `classify()` checks
 * `primary.document.styles` before ever reaching dispatch, with a NAMED reason -- the same reason
 * `formatting.cjs`'s commandBody/rangeCommandBody guard (test/static/formatting-command-vm.
 * test.cjs) protects at the lower, write-time layer.
 *
 * CORRECTED (LIVE regression against this file's ORIGINAL claim): `format_text`
 * never creates a style, but `doc.GetAllStyles()` does not see ANY custom style regardless --
 * applying a genuinely EXISTING custom style would have been wrongly refused without
 * `docxAdapter.observationRequestFor` (adapter.cjs) also collecting `formatting.cjs`'s own
 * `styleNamesOf(task)` into the SAME targeted-read request style-definition.cjs already uses
 * (W10-8). The tests below this point exercise that corrected behavior.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU', color: 'auto', highlight: 'none', vertAlign: 'baseline' };
const TARGET_TEXT = 'FormatMe target paragraph.';
const paragraphs = () => [{ text: TARGET_TEXT, props: { ...PLAIN } }];
const store = (styles = {}) => createStore(paragraphs(), { doc: { styles } });

async function run(s, task) {
  let t = 0;
  const driver = createFakeDocxDriver(s, {});
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (format) => ({ id: 'w10-9', steps: [{ intent: 'format_text', selector: { text: TARGET_TEXT }, format }] });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return e && e.evidence && e.evidence.reason; };

test('format_text {style: <existing>}: an existing style name is NOT refused by the pre-dispatch check -- the normal classify/dispatch flow runs', async () => {
  const s = store({ 'Heading 1': { text: {}, paragraph: {} } });
  const { x, m } = await run(s, task({ style: 'Heading 1' }));
  assert.equal(status(x), OPERATION_STATUS.APPLIED, JSON.stringify(x.appliedOperations));
  assert.equal(m.writes, 1);
});

/**
 * W-10-11 (a SECOND LIVE regression, from W10-10): `buildTargetedStyleQuery` writes
 * `out[name] = null` for a requested name it could not find -- `readLive`'s union merges that
 * name into `document.styles` as a KEY, with a `null` VALUE. Since W10-10 made `formatting.cjs`'s
 * own `format.style` names flow into the SAME targeted request, a never-created name NOW gets
 * exactly this null-valued key -- and `styleExists`'s ORIGINAL key-presence-only check (W10-9)
 * wrongly read that as "exists". This is the EXACT inversion zero measured LIVE: APPLY_MISSING
 * went from 0 commands (pre-dispatch FAIL_CLOSED) to 1 command (DISPATCHED_OUTCOME_UNKNOWN).
 */
test('RED/GREEN: format_text {style: <never created>} is STILL FAIL_CLOSED pre-dispatch even once its name is requested via the SAME targeted-read channel W10-10 wires through -- a null-valued key must never read as "exists"', async () => {
  const s = store({ 'Heading 1': { text: {}, paragraph: {} } });
  const { x, m } = await run(s, task({ style: 'ZZNeverCreated2' }));
  assert.equal(status(x), OPERATION_STATUS.FAILED, JSON.stringify(x.appliedOperations));
  assert.equal(reason(x), 'format-style-does-not-exist:ZZNeverCreated2');
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0, 'the pre-dispatch check must refuse BEFORE any command reaches the driver -- this exact count (0) is what zero\'s LIVE regression measured going to 1');
});

test('RED/GREEN: format_text {style: <never created>} FAIL_CLOSED pre-dispatch, 0 writes, 0 commands -- the NAMED reason, never a silent Normal application', async () => {
  const s = store({ 'Heading 1': { text: {}, paragraph: {} } });
  const { x, m } = await run(s, task({ style: 'ZZNeverCreated' }));
  assert.equal(status(x), OPERATION_STATUS.FAILED, JSON.stringify(x.appliedOperations));
  assert.equal(reason(x), 'format-style-does-not-exist:ZZNeverCreated');
  assert.equal(m.writes, 0);
  assert.equal(m.commands, 0, 'the pre-dispatch check must refuse BEFORE any command reaches the driver');
});

test('format_text {style: <name>} is case-insensitive against the document styles map, consistent with this family\'s existing NAME_KEYS convention', async () => {
  const s = store({ 'Heading 1': { text: {}, paragraph: {} } });
  const { x } = await run(s, task({ style: 'heading 1' }));
  assert.notEqual(status(x), OPERATION_STATUS.FAILED, JSON.stringify(x.appliedOperations));
});

test('format_text with NO style field is entirely unaffected by the new pre-dispatch check', async () => {
  const s = store({}); // no styles known at all
  const { x, m } = await run(s, task({ bold: true }));
  assert.equal(status(x), OPERATION_STATUS.APPLIED, JSON.stringify(x.appliedOperations));
  assert.equal(m.writes, 1);
});

test('classify() directly: a missing/unreadable document.styles map fails closed too, the same safe default as every other classification ambiguity in this family', () => {
  const op = { selector: { text: TARGET_TEXT }, format: { style: 'Heading 1' } };
  const primaryNoStyles = { blocks: [{ kind: 'paragraph', text: TARGET_TEXT, format: { ...PLAIN } }], document: {} };
  const c = formatting.classify(primaryNoStyles, op);
  assert.equal(c.state, 'FAIL_CLOSED');
  assert.equal(c.evidence.reason, 'format-style-does-not-exist:Heading 1');
});

// ---- W-10-10: format_text applying a GENUINELY EXISTING custom style ------------
//
// a maintainer's own LIVE diagnostic: a style defined in an EARLIER interaction, invisible to
// doc.GetAllStyles() regardless of WHEN it was created, was wrongly refused by a LATER, format_text-
// ONLY task (no define_style/modify_style step of its own to carry the name through
// style-definition.cjs's styleNamesOf). The fixture below seeds `doc.customStyleNames` (W-10-10's
// own fake-driver extension, test/fakes/fake-docx-driver.cjs) to model EXACTLY that: a style
// that is custom/GetAllStyles-blind from the START of this session, not created by any step in
// the task under test.

const storeWithCustomStyle = (name) => createStore(paragraphs(), { doc: { styles: { [name]: { text: { bold: true }, paragraph: {} } }, customStyleNames: [name] } });

async function runWith(adapter, s, task2) {
  let t = 0;
  const driver = createFakeDocxDriver(s, {});
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, adapter, createExecutor(task2));
  return { x: result.execution, m: result.metrics };
}

test('RED/GREEN: format_text applying a style with NO define_style/modify_style step of its own in the task is APPLIED -- doc.GetAllStyles() alone cannot see it, but observationRequestFor now merges formatting.cjs\'s own requested style names too', async () => {
  const s = storeWithCustomStyle('ZZCustom');
  const { x, m } = await runWith(docxAdapter, s, task({ style: 'ZZCustom' }));
  assert.equal(status(x), OPERATION_STATUS.APPLIED, JSON.stringify(x.appliedOperations));
  assert.equal(m.writes, 1);
});

test('RED/GREEN: reverting formatting.cjs\'s contribution to observationRequestFor reproduces the exact regression -- a genuinely existing custom style is wrongly refused', async () => {
  const s = storeWithCustomStyle('ZZCustom2');
  // a shallow clone whose observationRequestFor ONLY asks style-definition.cjs -- the W10-9 shape,
  // before formatting.cjs's own styleNamesOf(task) was merged in.
  const preFixAdapter = { ...docxAdapter, observationRequestFor(task2) {
    const names = styleDefinition.styleNamesOf(task2);
    return names.length ? { live: { styleNames: names } } : {};
  } };
  const { x, m } = await runWith(preFixAdapter, s, task({ style: 'ZZCustom2' }));
  assert.equal(status(x), OPERATION_STATUS.FAILED, JSON.stringify(x.appliedOperations));
  assert.equal(reason(x), 'format-style-does-not-exist:ZZCustom2', 'reproducing the measured regression: GetAllStyles() alone cannot see a custom style, even one from an earlier session');
  assert.equal(m.writes, 0);
});

test('formatting.styleNamesOf(task): collects format_text{style} names, [] for a task with none', () => {
  assert.deepEqual(formatting.styleNamesOf(task({ style: 'ZZCustom3' })), ['ZZCustom3']);
  assert.deepEqual(formatting.styleNamesOf({ id: 'x', steps: [{ intent: 'format_text', selector: {}, format: { bold: true } }] }), []);
  assert.deepEqual(formatting.styleNamesOf({ id: 'x', steps: [] }), []);
});

test('docxAdapter.observationRequestFor(task): merges style-definition.cjs AND formatting.cjs names, from a MIXED task', () => {
  const t2 = { id: 'x', steps: [
    { intent: 'define_style', name: 'A' },
    { intent: 'format_text', selector: { text: 'x' }, format: { style: 'B' } },
  ] };
  const req = docxAdapter.observationRequestFor(t2);
  assert.deepEqual(new Set(req.live.styleNames), new Set(['A', 'B']));
});
