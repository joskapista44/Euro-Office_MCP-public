'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const rich = require('../../src/adapters/docx/families/rich.cjs');
const notes = require('../../src/adapters/docx/families/notes.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 3 offline acceptance: rich-inline (add_bookmark, add_link) and notes (add_footnote,
 * add_endnote) through the TARGET path (M1 -> M2 -> M3 -> DOCX adapter) on the fake transport whose
 * model follows the live rich probes 1-3: bookmarks with their range text, hyperlinks as non-run
 * children, note references on the target, footnotes projected as an `[N]` anchor + a trailing
 * footnote list (endnotes not projected), and the cursor context a note write must restore.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Courier New', fontSize: 28, language: 'hu-HU' };
const H = 'Intro heading';
const A = 'Alpha paragraph.';
const B = 'Beta paragraph.';
const C = 'Gamma paragraph.';
const L = 'Linked paragraph.';
const Z = 'Closing paragraph.';
const URL = 'https://example.invalid/w3';
const doc = () => [
  { text: H, level: 1, props: { ...PLAIN, bold: true } },
  { text: A, props: { ...PLAIN } },
  { text: B, props: { ...PLAIN } },
  { text: C, props: { ...PLAIN } },
  { text: L, props: { ...RICH } },
  { text: Z, props: { ...PLAIN } },
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
const task = (...steps) => ({ id: 'w3', steps });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => {
  const e = x.appliedOperations[i].evidence;
  return (e && e.evidence && e.evidence.reason) || (e && e.reason);
};
const { INTENTS } = require('../../src/adapters/docx/intents.cjs');
const P1 = new Set(INTENTS.filter(r => r.preservationMode === 'PRIMITIVE_INTEGRAL').map(r => r.intent));

async function firstAndSecond(store, t) {
  const first = await run(store, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => e.evidence)));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => e.status)));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.equal((first.x.preservationVerification || []).length, t.steps.filter(s => P1.has(s.intent)).length, 'integral L3 exactly for the P1 operations');
  assert.ok((first.x.preservationVerification || []).every(v => v.verdict === 'PASS'), 'L3');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', 'L2');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  assert.equal(first.driver.calls.forceSave, 1, 'one barrier, at task end');
  const second = await run(store, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(second.m.writes, 0);
  assert.equal(second.m.commands, 0);
  assert.equal(second.driver.calls.forceSave, 0);
  assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

// ---- inventory: the footnote projection is representation, not text ----------------------------

test('inventory: a footnote reference anchor and the trailing footnote list do not become text', () => {
  const html = `<body><p><span>${A}</span><a style="mso-footnote-id:ftn1" href="#_ftn1" name="_ftnref1"><span class="MsoFootnoteReference"><span>[1]</span></span></a></p><p>${Z}</p>`
    + '<div style="mso-element:footnote-list"><br clear="all"/><hr/><div style="mso-element:footnote" id="ftn1"><p class="MsoFootnoteText"><a style="mso-footnote-id:ftn1" href="#_ftnref1" name="_ftn1"><span>[1]</span></a><span>&nbsp;</span><span>the note</span></p></div></div></body>';
  const inv = normalizeInventory({ html, paragraphs: [] });
  assert.deepEqual(inv.blocks.map(b => b.text), [A, Z]);
  // NEGATIVE CONTROL: without the stripping the same projection yields the marker and a note paragraph
  const raw = normalizeInventory({ html: html.replace(/mso-footnote-id/g, 'x-id').replace('mso-element:footnote-list', 'x-list'), paragraphs: [] });
  assert.deepEqual(raw.blocks.map(b => b.text), [`${A}[1]`, Z, '[1] the note']);
});

// ---- FIRST / SECOND per intent ------------------------------------------------------------------

test('add_bookmark: FIRST applies, SECOND is a 0-write no-op; the range is the target', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'add_bookmark', selector: { text: A }, name: 'w3_mark' }));
  assert.deepEqual(store.durableDoc.bookmarks, [{ name: 'w3_mark', text: A }]);
});

test('add_link: FIRST applies with the seven text properties preserved (P1), SECOND is a no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'add_link', selector: { text: L }, url: URL, tooltip: 'tip' }));
  const p = store.durable.find(x => x.text === L);
  assert.deepEqual(p.links, [{ url: URL, tip: 'tip' }]);
  assert.deepEqual(p.props, RICH);
});

test('add_link without a tooltip: the url is the screen tip (the source\'s durable state), FIRST / SECOND', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'add_link', selector: { text: L }, url: URL }));
  assert.deepEqual(store.durable.find(x => x.text === L).links, [{ url: URL, tip: URL }]);
});

test('add_footnote / add_endnote: FIRST / SECOND, the reference sits on the target', async () => {
  for (const [intent, kind] of [['add_footnote', 'footnote'], ['add_endnote', 'endnote']]) {
    const store = createStore(doc());
    await firstAndSecond(store, task({ intent, selector: { text: B }, text: `the ${kind}` }));
    assert.deepEqual(store.durableDoc[`${kind}s`], [`the ${kind}`]);
    assert.equal(store.durable.find(x => x.text === B).refs[kind], 1);
  }
});

test('one session, all four intents and two footnotes: every note restores the main body for the next write', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'add_footnote', selector: { text: A }, text: 'first footnote' },
    { intent: 'add_footnote', selector: { text: B }, text: 'second footnote' },
    { intent: 'add_endnote', selector: { text: C }, text: 'an endnote' },
    { intent: 'add_bookmark', selector: { text: A }, name: 'after_notes' },
    { intent: 'add_link', selector: { text: L }, url: URL },
  ));
  assert.deepEqual(store.durableDoc.footnotes, ['first footnote', 'second footnote']);
  assert.deepEqual(store.durableDoc.bookmarks, [{ name: 'after_notes', text: A }], 'footnote-then-bookmark on one paragraph: the bookmark is on the target');
});

test('composability: rewrite -> footnote on its result (ref); a footnoted paragraph stays targetable by text', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'rewrite_paragraph', selector: { text: A, kind: 'paragraph', before: H }, text: 'Alpha rewritten.' },
    { intent: 'add_footnote', selector: { ref: 's0' }, text: 'on the rewrite', dependsOn: ['s0'] },
    { intent: 'format_text', selector: { text: 'Alpha rewritten.' }, format: { bold: true } },
    { intent: 'add_bookmark', selector: { text: 'Alpha rewritten.' }, name: 'rewritten' },
  ));
  assert.equal(store.durable.find(x => x.text === 'Alpha rewritten.').refs.footnote, 1);
});

// ---- negatives: FAIL_CLOSED before any dispatch -------------------------------------------------

test('negatives: absent / ambiguous / table cell / unread / existing name / other link / stray note / cursor in a note -> FAIL_CLOSED, 0 dispatch', async () => {
  const withDoc = extra => createStore(doc(), { doc: extra });
  const cases = [
    [task({ intent: 'add_bookmark', selector: { text: 'Nope.' }, name: 'x' }), createStore(doc()), {}, 'add_bookmark-target-not-found'],
    [task({ intent: 'add_link', selector: { text: A }, url: URL }), createStore([...doc(), { text: A, props: { ...PLAIN } }]), {}, 'add_link-target-ambiguous'],
    [task({ intent: 'add_bookmark', selector: { text: A }, name: 'b' }), createStore(doc()), { docUnreadable: true }, 'bookmarks-unreadable'],
    [task({ intent: 'add_bookmark', selector: { text: A }, name: 'b' }), withDoc({ bookmarks: [{ name: 'b', text: B }] }), {}, 'bookmark-name-exists'],
    [task({ intent: 'add_link', selector: { text: A }, url: URL }), createStore(doc()), { richUnreadable: true }, 'hyperlinks-unreadable'],
    [task({ intent: 'add_link', selector: { text: A }, url: URL }), createStore(doc().map(p => (p.text === A ? { ...p, links: [{ url: 'https://other.invalid', tip: '' }] } : p))), {}, 'target-already-linked'],
    [task({ intent: 'add_footnote', selector: { text: A }, text: 'stray' }), withDoc({ footnotes: ['stray'] }), {}, 'footnote-text-without-reference-on-target'],
    [task({ intent: 'add_endnote', selector: { text: A }, text: 'dup' }), createStore(doc().map(p => (p.text === A ? { ...p, refs: { footnote: 0, endnote: 2 } } : p)), { doc: { endnotes: ['dup', 'dup'] } }), {}, 'endnote-duplicated'],
    [task({ intent: 'add_footnote', selector: { text: A }, text: 'n' }), createStore(doc()), { docUnreadable: true }, 'footnotes-unreadable'],
    [task({ intent: 'add_footnote', selector: { text: A }, text: 'n' }), createStore(doc()), { startInNote: true }, 'editor-not-in-main-body'],
    [task({ intent: 'add_bookmark', selector: { text: A }, name: 'b' }), createStore(doc()), { startInNote: true }, 'editor-not-in-main-body'],
  ];
  for (const [t, store, script, want] of cases) {
    const { x, driver } = await run(store, t, script);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
});

test('negatives: table-cell collision -> the target is not unique', () => {
  const primary = { blocks: [{ kind: 'paragraph', text: A }, { kind: 'table', rows: [[{ text: A }]] }], document: null };
  assert.deepEqual(rich.inlineTarget(primary, A).reason, 'target-ambiguous');
});

test('negatives: bad parameters and physical selectors -> plan fails before any dispatch', async () => {
  for (const step of [
    { intent: 'add_bookmark', selector: { text: A }, name: '1bad' },
    { intent: 'add_bookmark', selector: { text: A }, name: 'x'.repeat(41) },
    { intent: 'add_link', selector: { text: A }, url: '  ' },
    { intent: 'add_link', selector: { text: A }, url: URL, tooltip: 3 },
    { intent: 'add_footnote', selector: { text: A }, text: ' ' },
    { intent: 'add_endnote', selector: { text: A, index: 2 }, text: 'n' },
  ]) {
    const { x, driver } = await run(createStore(doc()), task(step));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(step));
    assert.equal(driver.calls.dispatch, 0);
  }
});

// ---- the write happened but is not the terminal contract ----------------------------------------

test('faults at L1: no write / wrong target / duplicate / wrong url / note elsewhere / cursor left in the note -> never APPLIED, no barrier', async () => {
  const cases = [
    ['ignore', { intent: 'add_bookmark', selector: { text: A }, name: 'b' }],
    ['bookmark-wrong-target', { intent: 'add_bookmark', selector: { text: A }, name: 'b' }],
    ['bookmark-twice', { intent: 'add_bookmark', selector: { text: A }, name: 'b' }],
    ['ignore', { intent: 'add_link', selector: { text: L }, url: URL }],
    ['link-wrong-url', { intent: 'add_link', selector: { text: L }, url: URL }],
    ['ignore', { intent: 'add_footnote', selector: { text: A }, text: 'n' }],
    ['note-twice', { intent: 'add_footnote', selector: { text: A }, text: 'n' }],
    ['note-wrong-paragraph', { intent: 'add_endnote', selector: { text: A }, text: 'n' }],
    ['note-stuck', { intent: 'add_footnote', selector: { text: A }, text: 'n' }],
  ];
  for (const [fault, step] of cases) {
    const { x, result, driver } = await run(createStore(doc()), task(step), { fault });
    const tag = `${fault} ${step.intent}`;
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, tag);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', tag);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, tag);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST, tag);
    assert.equal(driver.calls.forceSave, 0, tag);
  }
});

test('fault at L3: the link is there but the Hyperlink style was not undone -> NOT APPLIED (P1)', async () => {
  const { x } = await run(createStore(doc()), task({ intent: 'add_link', selector: { text: L }, url: URL }), { fault: 'link-no-restore' });
  assert.equal(x.primitiveVerification[0].verdict, 'PASS');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
});

test('fault at L3: a link that swallows the target\'s existing note reference -> NOT APPLIED (probe 4)', async () => {
  const { x } = await run(createStore(doc()), task(
    { intent: 'add_endnote', selector: { text: L }, text: 'an endnote' },
    { intent: 'add_link', selector: { text: L }, url: URL },
  ), { fault: 'link-swallows-reference', faultOn: [2] });
  assert.equal(status(x, 0), OPERATION_STATUS.APPLIED);
  assert.equal(x.primitiveVerification[1].verdict, 'PASS', 'the link itself is there');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'but the reference on the target is gone');
  assert.equal(x.appliedOperations[1].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
});

test('note -> link on one paragraph: both APPLIED, L2 holds, retry 0/0 (the w3_mixed LIVE failure shape)', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'add_endnote', selector: { text: L }, text: 'an endnote' },
    { intent: 'add_link', selector: { text: L }, url: URL, tooltip: 't' },
  ));
});

test('owner review 3: link L3 round trip -- another paragraph linked / an unrelated footnote lost / a bookmark moved -> NOT APPLIED', async () => {
  for (const fault of ['link-touch-other', 'link-drop-other-footnote', 'link-move-bookmark']) {
    const store = createStore(doc().map(p => (p.text === A ? { ...p, refs: { footnote: 1, endnote: 0 } } : p)), { doc: { footnotes: ['kept note'], bookmarks: [{ name: 'kept', text: B }] } });
    const { x } = await run(store, task({ intent: 'add_link', selector: { text: L }, url: URL, tooltip: 't' }), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the link itself is there`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
  }
});

test('owner review 1: the right url with a WRONG tooltip is not SATISFIED -> FAIL_CLOSED, 0 dispatch', async () => {
  const store = createStore(doc().map(p => (p.text === L ? { ...p, links: [{ url: URL, tip: 'wrong tip' }] } : p)));
  const { x, driver } = await run(store, task({ intent: 'add_link', selector: { text: L }, url: URL, tooltip: 'requested tip' }));
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'target-already-linked');
  assert.equal(driver.calls.dispatch, 0);
  // POSITIVE CONTROL: the same link with the requested tooltip is SATISFIED
  const ok = await run(createStore(doc().map(p => (p.text === L ? { ...p, links: [{ url: URL, tip: 'requested tip' }] } : p))), task({ intent: 'add_link', selector: { text: L }, url: URL, tooltip: 'requested tip' }));
  assert.equal(status(ok.x), OPERATION_STATUS.SATISFIED);
});

test('owner review 3 regression: normal text + footnote reference -> add_link on the text; the reference stays a note reference', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'add_footnote', selector: { text: L }, text: 'kept footnote' },
    { intent: 'add_link', selector: { text: L }, url: URL, tooltip: 't' },
  ));
  const p = store.durable.find(x => x.text === L);
  assert.equal(p.refs.footnote, 1);
  assert.deepEqual(p.links, [{ url: URL, tip: 't' }]);
});

test('owner review 4: mixed regressions footnote -> bookmark and footnote -> endnote, one session each, retry 0/0', async () => {
  const s1 = createStore(doc());
  await firstAndSecond(s1, task({ intent: 'add_footnote', selector: { text: A }, text: 'fn' }, { intent: 'add_bookmark', selector: { text: A }, name: 'after_fn' }));
  assert.deepEqual(s1.durableDoc.bookmarks, [{ name: 'after_fn', text: A }]);
  const s2 = createStore(doc());
  await firstAndSecond(s2, task({ intent: 'add_footnote', selector: { text: A }, text: 'fn' }, { intent: 'add_endnote', selector: { text: B }, text: 'en' }));
  assert.deepEqual([s2.durableDoc.footnotes, s2.durableDoc.endnotes], [['fn'], ['en']]);
});

test('the second note of a task after a note that left the cursor inside it: FAIL_CLOSED, never a blind write', async () => {
  const { x, driver } = await run(createStore(doc()), task(
    { intent: 'add_footnote', selector: { text: A }, text: 'one' },
    { intent: 'add_footnote', selector: { text: B }, text: 'two' },
  ), { fault: 'note-stuck', faultOn: [1] });
  assert.notEqual(status(x, 0), OPERATION_STATUS.APPLIED, 'the first note did not restore the main body');
  assert.equal(driver.calls.dispatch, 1, 'no second dispatch into the note context');
});

// ---- lowering ------------------------------------------------------------------------------------

test('lowering: one callCommand per operation, public API only, no editor backdoor', () => {
  const bm = rich.commandBody({ action: 'bookmark', targetText: A, name: 'b' });
  const ln = rich.commandBody({ action: 'link', targetText: A, url: URL, tooltip: '' });
  const nt = notes.commandBody({ kind: 'footnote', targetText: A, text: 'n' });
  for (const body of [bm, ln, nt]) {
    assert.match(body, /GetCurrentFootEndnote\(\) !== null\) return/, 'the main-body precondition inside the command');
    assert.doesNotMatch(body, /frame\.evaluate|window\.|SetDocPosType|private_|asc_|getSelectedElements/);
    assert.doesNotThrow(() => new Function(body));
  }
  assert.match(ln, /again\[0\]\.SetTextPr\(before\)/, 'the link restores the text properties');
  assert.match(ln, /hits\[0\]\.AddHyperlink\(m\.url, m\.tooltip\)/, 'the link covers the exact target RANGE');
  assert.doesNotMatch(ln, /GetParagraph\(0\)|paragraph\.AddHyperlink/, 'never the paragraph-wide link (probe 4)');
  assert.match(nt, /finally \{ doc\.MoveCursorToStart\(\); \}/, 'the note returns the cursor to the main body');
  assert.match(bm, /GetAllBookmarksNames\(\) \|\| \[\]\)\.indexOf\(m\.name\) >= 0\) return/, 'an existing name -> no write');
});
