'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const rich = require('../../src/adapters/docx/families/rich.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #266 -- add_link / add_bookmark with `matchText`, adapter level on the fake whose model follows the rich probes: the selector names the PARENT unit by its exact text, the link / bookmark
 * covers matchText, which must occur exactly once in it. Without matchText the whole-unit behaviour (test/acceptance/docx-w3-rich.test.cjs) is unchanged. The fake's sub-range model is NOT a
 * LIVE measurement: the harness test/live/gh266-docx-link-bookmark-product.live.cjs measures it.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const A = 'Alpha beta gamma delta.';
const B = 'Second paragraph here.';
const URL1 = 'https://example.invalid/one';
const URL2 = 'https://example.invalid/two';
const doc = () => [{ text: A, props: { ...PLAIN } }, { text: B, props: { ...PLAIN } }];

async function run(store, steps, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor({ id: 'g266', steps }));
  return { result, driver, x: result.execution, m: result.metrics };
}
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const link = (match, url = URL1, extra = {}) => ({ intent: 'add_link', selector: { text: A }, matchText: match, url, tooltip: 'tip', ...extra });
const mark = (match, name = 'g266_mark') => ({ intent: 'add_bookmark', selector: { text: A }, matchText: match, name });

test('add_link matchText: FIRST applies on the sub-range only (L1/L2/L3 PASS), SECOND is a 0-write no-op', async () => {
  const store = createStore(doc());
  const first = await run(store, [link('beta')]);
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED, JSON.stringify(first.x.appliedOperations.map(e => e.evidence)));
  assert.equal(first.x.primitiveVerification[0].verdict, 'PASS');
  assert.equal(first.x.preservationVerification[0].verdict, 'PASS');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  assert.deepEqual(store.durable.find(p => p.text === A).links, [{ url: URL1, tip: 'tip', text: 'beta' }]);
  assert.equal(store.durable.find(p => p.text === B).links, undefined, 'another paragraph is untouched');
  const second = await run(store, [link('beta')]);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.equal(second.m.writes, 0);
  assert.equal(second.m.commands, 0);
});

test('add_bookmark matchText: FIRST applies, the bookmark range IS matchText (not the paragraph), SECOND is a 0-write no-op', async () => {
  const store = createStore(doc());
  const first = await run(store, [mark('gamma')]);
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED, JSON.stringify(first.x.appliedOperations.map(e => e.evidence)));
  assert.deepEqual(store.durableDoc.bookmarks.map(b => ({ name: b.name, text: b.text })), [{ name: 'g266_mark', text: 'gamma' }]);
  const second = await run(store, [mark('gamma')]);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.equal(second.m.writes, 0);
});

test('two links on two DIFFERENT words of one paragraph both apply and both survive (each step keeps the other)', async () => {
  const store = createStore(doc());
  assert.equal(status((await run(store, [link('Alpha')])).x), OPERATION_STATUS.APPLIED);
  const r = await run(store, [link('delta', URL2)]);
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => e.evidence)));
  assert.equal(r.x.preservationVerification[0].verdict, 'PASS');
  assert.deepEqual(store.durable.find(p => p.text === A).links.map(l => [l.url, l.text]), [[URL1, 'Alpha'], [URL2, 'delta']]);
});

test('negatives: matchText absent / ambiguous / blank / overlapping an existing link / different url on the same range -> FAIL_CLOSED, 0 dispatch', async () => {
  const linked = (url, text) => createStore(doc().map(p => (p.text === A ? { ...p, links: [{ url, tip: 'tip', text }] } : p)));
  const cases = [
    [[link('zzz')], createStore(doc()), 'add_link-match-not-found'],
    [[link('a')], createStore(doc()), 'add_link-match-ambiguous'],
    [[mark('zzz')], createStore(doc()), 'add_bookmark-match-not-found'],
    [[mark('a')], createStore(doc()), 'add_bookmark-match-ambiguous'],
    [[link('beta gamma')], linked(URL1, 'beta'), 'target-already-linked'],
    [[link('beta')], linked(URL2, 'beta'), 'target-already-linked'],
    [[link('beta')], linked(URL1, A), 'target-already-linked'],
  ];
  for (const [steps, store, want] of cases) {
    const { x, driver } = await run(store, steps);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
  assert.equal(rich.plan({ intent: 'add_link', selector: { text: A }, matchText: '  ', url: URL1 }, 0).ok, false, 'a blank matchText never plans');
  assert.equal(rich.plan({ intent: 'add_link', selector: { text: A }, matchText: 5, url: URL1 }, 0).ok, false);
});

test('faults: a link / bookmark over the WHOLE parent instead of matchText, a lost neighbour link, a wrong url -> never APPLIED', async () => {
  const linkedAlpha = () => createStore(doc().map(p => (p.text === A ? { ...p, links: [{ url: URL2, tip: 'tip', text: 'Alpha' }] } : p)));
  const cases = [
    ['link-whole-parent', createStore(doc()), link('beta'), 'primitive'],
    ['bookmark-whole-parent', createStore(doc()), mark('beta'), 'primitive'],
    ['link-wrong-url', createStore(doc()), link('beta'), 'primitive'],
    ['link-drop-other-link', linkedAlpha(), link('delta'), 'preservation'],
    ['link-no-restore', createStore(doc()), link('beta'), 'preservation'],
  ];
  for (const [fault, store, step, layer] of cases) {
    const { x } = await run(store, [step], { fault });
    assert.notEqual(status(x), OPERATION_STATUS.APPLIED, fault);
    if (layer === 'primitive') assert.equal(x.primitiveVerification[0].verdict, 'FAIL', fault);
    else assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
  }
  const { x } = await run(linkedStore(), [link('delta')], { fault: 'link-drop-other-link' });
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  function linkedStore() { return createStore(doc().map(p => (p.text === A ? { ...p, links: [{ url: URL2, tip: 'tip', text: 'Alpha' }] } : p))); }
});

test('lowering: matchText bodies search INSIDE the one parent paragraph, public API only; the whole-unit body has no paragraph search', () => {
  const ranged = rich.commandBody({ action: 'link', targetText: A, matchText: 'beta', url: URL1, tooltip: 'tip' });
  const whole = rich.commandBody({ action: 'link', targetText: A, url: URL1, tooltip: 'tip' });
  assert.match(ranged, /parents\.length !== 1\) return \[\]/);
  assert.match(ranged, /paragraph\.Search\(m\.matchText, true\)/);
  assert.match(ranged, /hits\[0\]\.AddHyperlink\(m\.url, m\.tooltip\)/, 'the link is on the RANGE, never the paragraph');
  assert.doesNotMatch(ranged, /paragraph\.AddHyperlink|frame\.evaluate|window\.|private_|asc_/);
  assert.doesNotMatch(whole, /GetParagraph|paragraph\.Search/);
  assert.doesNotThrow(() => new Function(ranged));
  assert.doesNotThrow(() => new Function(rich.commandBody({ action: 'bookmark', targetText: A, matchText: 'beta', name: 'n' })));
});

// ---- the pre-dispatch gate: a bookmark range must not enclose a hyperlink (g266v1 ROUND4B) --------------------------------------------------------------------
// g266v1 LIVE (2026-10-19): a whole-unit bookmark added to a paragraph that ALREADY carried a hyperlink LANDED in the file and then every rich read of the document (bookmarks, hyperlinks)
// was unreadable, even in a fresh process. The reverse order (a bookmark first, then a link over the same unit) is LIVE-green since M4-W3 (w3_mixed), so only the bookmark side is gated.

const linkedStore = links => createStore(doc().map(p => (p.text === A ? { ...p, links } : p)));
const noDispatch = (r, want) => { assert.equal(status(r.x), OPERATION_STATUS.FAILED); assert.equal(reason(r.x), want); assert.equal(r.driver.calls.dispatch, 0); assert.equal(r.m.writes, 0); };

test('GATE: a whole-unit bookmark on a unit that already holds a hyperlink (whole-unit or sub-range) -> FAIL_CLOSED bookmark-target-holds-a-hyperlink before any dispatch, the document unchanged', async () => {
  for (const links of [[{ url: URL1, tip: 'tip' }], [{ url: URL1, tip: 'tip', text: 'beta' }]]) {
    const store = linkedStore(links);
    const before = JSON.stringify(store.durable.map(p => [p.text, p.links || null]));
    noDispatch(await run(store, [{ intent: 'add_bookmark', selector: { text: A }, name: 'g266_gate' }]), 'bookmark-target-holds-a-hyperlink');
    assert.equal(JSON.stringify(store.durable.map(p => [p.text, p.links || null])), before);
    assert.deepEqual(store.durableDoc.bookmarks, []);
  }
});

test('GATE: a matchText bookmark is refused only when its range ENCLOSES a link text; a range clear of the link is written', async () => {
  const links = [{ url: URL1, tip: 'tip', text: 'beta' }];
  noDispatch(await run(linkedStore(links), [mark('Alpha beta')]), 'bookmark-target-holds-a-hyperlink');
  noDispatch(await run(linkedStore(links), [mark('beta gamma')]), 'bookmark-target-holds-a-hyperlink');
  const ok = await run(linkedStore(links), [mark('gamma')]);
  assert.equal(status(ok.x), OPERATION_STATUS.APPLIED, JSON.stringify(ok.x.appliedOperations.map(e => e.evidence)));
});

// GH #266 gate extension (g266v2 ROUND9: a bookmark INSIDE a link -> DISPATCHED_OUTCOME_UNKNOWN, reads unproven): ANY intersection of the bookmark range with a link refuses, by geometry.
test('GATE (geometry): a matchText bookmark INSIDE a link, PARTIALLY overlapping one, or EQUAL to one is refused before any dispatch with a named reason; a disjoint range is written', async () => {
  const cases = [
    ['inside a whole-unit link', [{ url: URL1, tip: 'tip' }], 'beta', 'bookmark-target-inside-a-hyperlink'],
    ['inside a sub-range link (strictly larger)', [{ url: URL1, tip: 'tip', text: 'Alpha beta gamma' }], 'beta', 'bookmark-target-inside-a-hyperlink'],
    ['partial overlap, bookmark runs past the link end', [{ url: URL1, tip: 'tip', text: 'Alpha beta' }], 'beta gamma', 'bookmark-target-overlaps-a-hyperlink'],
    ['partial overlap, bookmark starts before the link', [{ url: URL1, tip: 'tip', text: 'beta gamma' }], 'Alpha beta', 'bookmark-target-overlaps-a-hyperlink'],
    ['equal range', [{ url: URL1, tip: 'tip', text: 'beta' }], 'beta', 'bookmark-target-holds-a-hyperlink'],
    ['encloses (unchanged reason)', [{ url: URL1, tip: 'tip', text: 'beta' }], 'Alpha beta gamma', 'bookmark-target-holds-a-hyperlink'],
  ];
  for (const [label, links, match, want] of cases) {
    const store = linkedStore(links);
    const before = JSON.stringify(store.durable.map(p => [p.text, p.links || null]));
    noDispatch(await run(store, [mark(match)]), want);
    assert.equal(JSON.stringify(store.durable.map(p => [p.text, p.links || null])), before, label);
    assert.deepEqual(store.durableDoc.bookmarks, [], label);
  }
  for (const [links, match] of [[[{ url: URL1, tip: 'tip', text: 'beta' }], 'gamma'], [[{ url: URL1, tip: 'tip', text: 'beta' }], 'Alpha'], [[{ url: URL1, tip: 'tip', text: 'delta' }], 'beta gamma']]) {
    const ok = await run(linkedStore(links), [mark(match)]);
    assert.equal(status(ok.x), OPERATION_STATUS.APPLIED, `${match}: ${JSON.stringify(ok.x.appliedOperations.map(e => e.evidence))}`);
  }
});

test('GATE (geometry): the link text occurring TWICE in the unit is tested at EVERY occurrence (one of them intersecting the range refuses)', async () => {
  const T = 'beta one beta two';
  const store = createStore([{ text: T, props: { ...PLAIN } }].map(p => ({ ...p, links: [{ url: URL1, tip: 'tip', text: 'beta' }] })));
  const r = await run(store, [{ intent: 'add_bookmark', selector: { text: T }, matchText: 'beta two', name: 'g266_twice' }]);
  noDispatch(r, 'bookmark-target-holds-a-hyperlink');
  const clear = await run(createStore([{ text: T, props: { ...PLAIN }, links: [{ url: URL1, tip: 'tip', text: 'beta' }] }]), [{ intent: 'add_bookmark', selector: { text: T }, matchText: ' one ', name: 'g266_clear' }]);
  assert.equal(status(clear.x), OPERATION_STATUS.APPLIED, JSON.stringify(clear.x.appliedOperations.map(e => e.evidence)));
});

test('GATE: the reverse order stays released -- a bookmark first, then a whole-unit link over the same unit (the LIVE-green M4-W3 w3_mixed order) applies; the retry of the SAME bookmark after the link is SATISFIED, not refused', async () => {
  const store = createStore(doc());
  const m = await run(store, [{ intent: 'add_bookmark', selector: { text: A }, name: 'g266_first' }]);
  assert.equal(status(m.x), OPERATION_STATUS.APPLIED);
  const l = await run(store, [{ intent: 'add_link', selector: { text: A }, url: URL1, tooltip: 'tip' }]);
  assert.equal(status(l.x), OPERATION_STATUS.APPLIED, JSON.stringify(l.x.appliedOperations.map(e => e.evidence)));
  const again = await run(store, [{ intent: 'add_bookmark', selector: { text: A }, name: 'g266_first' }]);
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED, 'the bookmark that already IS there is a no-op, never a refusal');
});
