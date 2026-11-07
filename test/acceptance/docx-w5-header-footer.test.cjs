'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const hf = require('../../src/adapters/docx/families/header-footer.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const { check } = require('../../src/server/write-tools.cjs');

/**
 * M4 Wave 5 offline acceptance: header-footer (edit_header, edit_footer) through the TARGET path on
 * the fake transport whose model follows header/footer probes 1-2 (2026-09-25): per section, per
 * kind and variant a part of its own or none; a write replaces the whole content by ONE plain
 * paragraph; a created part exists only on its section; the cursor stays in the body.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const COURIER = { bold: true, italic: true, underline: true, strike: true, fontFamily: 'Courier New', fontSize: 28, language: 'hu-HU' };
// M5 gate 6: `part` is what the full replacement itself yields (one plain paragraph each, no style) --
// replaceable without loss; `fmt` carries formatting the replacement would drop (refused before dispatch)
const part = (...texts) => ({ paras: texts.map(t => ({ text: t, props: { ...PLAIN, fontFamily: 'Arial', fontSize: 22 }, alignment: 'left' })) });
const fmt = (...texts) => ({ paras: texts.map((t, i) => ({ text: t, props: i === 0 ? { ...COURIER } : { ...PLAIN }, alignment: i === 0 ? 'center' : 'left' })) });
const section = (h, f) => ({ header: { default: h || null }, footer: { default: f || null } });
const paragraphs = () => [{ text: 'Body heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Body first.', props: { ...PLAIN } }, { text: 'Body second.', props: { ...PLAIN } }];
const store = (sections, paras = paragraphs()) => createStore(paras, { doc: sections ? { sections } : {} });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w5', steps });
const eh = (text, index = 0, of = 1, extra = {}) => ({ intent: 'edit_header', section: { index, of }, text, ...extra });
const ef = (text, index = 0, of = 1, extra = {}) => ({ intent: 'edit_footer', section: { index, of }, text, ...extra });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const partOf = (s, i, k, v = 'default') => { const p = s.durableDoc.sections[i][k][v]; return p ? p.paras.map(x => x.text) : null; };
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.sections]);

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per operation');
  assert.deepEqual([first.driver.calls.open, first.driver.calls.close], [1, 1]);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.primitiveVerification));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(s, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome || second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

test('one section, no parts: edit_header / edit_footer create the part with one paragraph; the other kind stays absent; SECOND 0/0', async () => {
  for (const [st, k, other] of [[eh('Agent header'), 'header', 'footer'], [ef('Agent footer'), 'footer', 'header']]) {
    const s = store();
    await firstAndSecond(s, task(st));
    assert.deepEqual(partOf(s, 0, k), [st.text]);
    assert.equal(partOf(s, 0, other), null, `${k}: the ${other} is not created`);
    assert.deepEqual(s.durable.map(p => p.text), paragraphs().map(p => p.text), 'the body unchanged');
  }
});

test('M5 gate 6 POSITIVE: an existing part the replacement cannot lose anything from -> ONE plain paragraph; the formatted footer and the body unchanged', async () => {
  const s = store([section(part('HA plain'), fmt('FA fmt'))]);
  await firstAndSecond(s, task(eh('HA NEW')));
  assert.deepEqual(partOf(s, 0, 'header'), ['HA NEW']);
  assert.deepEqual(partOf(s, 0, 'footer'), ['FA fmt']);
  assert.equal(s.durableDoc.sections[0].footer.default.paras[0].props.fontFamily, 'Courier New', 'the footer keeps its own format');
});

test('M5 gate 6 NEGATIVE: an existing part whose formatting the replacement would drop -> refused BEFORE dispatch (0 dispatch / 0 write), the lost property named', async () => {
  const one = (props, extra = {}) => ({ paras: [{ text: 'H', props: { ...PLAIN, fontFamily: 'Arial', fontSize: 22, ...props }, alignment: 'left', ...extra }] });
  const L = 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-replacement-would-drop-';
  const cases = [
    [fmt('HA fmt', 'HA two'), `${L}paragraphs`],
    [part('HA', 'HA two'), `${L}paragraphs`, 'two PLAIN paragraphs: the replacement leaves one'],
    [one({}, { style: 'Header' }), `${L}style:Header`, 'the built-in Header style alone -- a plain-text part is NOT "unformatted"'],
    [one({}, { alignment: 'center' }), `${L}alignment:center`],
    [one({ bold: true }), `${L}run-bold`], [one({ italic: true }), `${L}run-italic`],
    [one({ underline: true }), `${L}run-underline`], [one({ strike: true }), `${L}run-strike`],
    [one({ fontFamily: 'Courier New' }), `${L}run-font-family`], [one({ fontSize: 28 }), `${L}run-font-size`],
    // GH #316 review R1 (a maintainer, LIVE): color is now readable on a header/footer run -- an explicitly-coloured
    // run the caller does not name would be silently decolourised by the plain-run replacement
    [one({ color: '#FF0000' }), `${L}run-color`],
    // GH #289 (LIVE probe 2): a PAGE field / an image reads as runs with no text -- the run classes alone said "plain"
    [one({}, { inline: ['pgNum'] }), `${L}inline:pgNum`, 'a PAGE field in an otherwise plain one-paragraph part'],
    [one({}, { inline: ['image', 'paraDrawing'] }), `${L}inline:image,paraDrawing`, 'an image in an otherwise plain one-paragraph part'],
    [one({}, { inline: null }), 'header-footer-inline-content-unreadable', 'the inline-content read failed -> refused, never assumed plain'],
  ];
  for (const [p, want, why] of cases) {
    const s = store([section(p)]);
    const before = snapshot(s);
    const { x, driver, m } = await run(s, task(eh('HA NEW')));
    assert.equal(status(x), OPERATION_STATUS.FAILED, why || want);
    assert.equal(reason(x), want, why || want);
    assert.deepEqual([driver.calls.dispatch, m.writes], [0, 0], want);
    assert.equal(snapshot(s), before, `${want}: nothing written`);
  }
  // the document default unread -> every EXISTING-part replacement refused; a MISSING part is still created
  const u = await run(store([section(part('H'))]), task(eh('x')), { defaultTextUnreadable: true });
  assert.equal(reason(u.x), 'header-footer-default-text-format-unreadable');
  assert.equal(u.driver.calls.dispatch, 0);
  const c = store([section(null, fmt('F'))]);
  await firstAndSecond(c, task(eh('Created')));
  assert.deepEqual(partOf(c, 0, 'header'), ['Created'], 'the CREATE scope is separate: nothing to lose');
  // CONTROL: a part that already holds the text is SATISFIED (no write, nothing lost) even when styled
  const d = await run(store([section(one({}, { style: 'Header' }))]), task(eh('H')));
  assert.equal(status(d.x), OPERATION_STATUS.SATISFIED);
});

test('SECTION_TARGETING: three sections with their own parts -- section 1 header written, every other part unchanged', async () => {
  const sections = () => [section(part('HA'), part('FA')), section(part('HB'), part('FB')), section(part('HC'), part('FC'))];
  const s = store(sections());
  await firstAndSecond(s, task(eh('HB NEW', 1, 3)));
  assert.deepEqual([0, 1, 2].map(i => [partOf(s, i, 'header'), partOf(s, i, 'footer')]), [[['HA'], ['FA']], [['HB NEW'], ['FB']], [['HC'], ['FC']]]);
  // the LAST section with an own part (nothing inherits from it) and the FIRST (every later one has its own)
  const s2 = store(sections());
  await firstAndSecond(s2, task(ef('FC NEW', 2, 3), eh('HA NEW', 0, 3)));
  assert.deepEqual([partOf(s2, 2, 'footer'), partOf(s2, 0, 'header'), partOf(s2, 1, 'header')], [['FC NEW'], ['HA NEW'], ['HB']]);
});

test('HEADER_FOOTER_COMPOSITION: header -> footer and footer -> header in one session; each semantic, SECOND 2x SATISFIED 0/0', async () => {
  for (const steps of [[eh('H one'), ef('F one')], [ef('F two'), eh('H two')]]) {
    const s = store([section(part('old H'), part('old F'))]);
    await firstAndSecond(s, task(...steps));
    assert.deepEqual([partOf(s, 0, 'header'), partOf(s, 0, 'footer')], [[steps.find(x => x.intent === 'edit_header').text], [steps.find(x => x.intent === 'edit_footer').text]]);
  }
});

test('MIXED_BODY_HEADER_FOOTER: rewrite -> header -> footer -> format in ONE session; the body ops resolve from the fresh state; retry 4x SATISFIED 0/0', async () => {
  const s = store([section(part('old H'), part('old F'))]);
  const t = task(
    { intent: 'rewrite_paragraph', selector: { text: 'Body first.', kind: 'paragraph', before: 'Body heading' }, text: 'Body first rewritten.' },
    eh('Mixed header'),
    ef('Mixed footer'),
    { intent: 'format_text', selector: { text: 'Body second.' }, format: { bold: true } },
  );
  await firstAndSecond(s, t);
  assert.deepEqual(s.durable.map(p => p.text), ['Body heading', 'Body first rewritten.', 'Body second.']);
  assert.equal(s.durable[2].props.bold, true);
  assert.deepEqual([partOf(s, 0, 'header'), partOf(s, 0, 'footer')], [['Mixed header'], ['Mixed footer']]);
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch, nothing written', async () => {
  const two = () => [section(part('HA'), part('FA')), section(part('HB'), part('FB'))];
  const failed = [
    // M5: a variant whose flag the read cannot report is refused (the internal reader absent)
    [task(eh('x', 0, 1, { variant: 'even' })), [section(part('H'))], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:variant-even-flag-unobservable', { flagsUnreadable: true }],
    [task(ef('x', 0, 1, { variant: 'title' })), [section(null, part('F'))], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:variant-title-flag-unobservable', { flagsUnreadable: true }],
    // an even part that ALREADY holds the text is not SATISFIED either while its visibility flag is unobservable
    [task(eh('x', 0, 1, { variant: 'even' })), [{ header: { default: part('H'), even: part('x') }, footer: { default: null } }], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:variant-even-flag-unobservable', { flagsUnreadable: true }],
    [task(eh('x', 0, 2)), [section(part('H'))], 'FAIL_CLOSED:section-descriptor-mismatch'],
    [task(eh('x', 0, 1)), two(), 'FAIL_CLOSED:section-descriptor-mismatch'],
    // section 0 written while section 1 has no header of its own: section 1 would change too
    [task(eh('x', 0, 2)), [section(part('HA'), part('FA')), section(null, part('FB'))], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:later-section-inherits-this-part:1'],
    // section 1 has no footer of its own: creating one breaks the link to section 0's
    [task(ef('x', 1, 2)), [section(part('HA'), part('FA')), section(part('HB'), null)], 'FAIL_CLOSED_UNSUPPORTED_SHAPE:linked-to-previous'],
    [task(eh('x')), [section({ ...part('H'), nonPlain: true })], 'header-footer-not-plain-paragraphs-of-runs'],
  ];
  for (const [t, sections, want, script] of failed) {
    const s = store(sections);
    const before = snapshot(s);
    const { x, driver } = await run(s, t, script);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
    assert.equal(snapshot(s), before, `${want}: nothing written`);
  }
  // unreadable sections -> FAIL_CLOSED, never a guess
  const u = await run(store(), task(eh('x')), { sectionsUnreadable: true });
  assert.equal(reason(u.x), 'sections-unreadable');
  assert.equal(u.driver.calls.dispatch, 0);
  for (const st of [
    { intent: 'edit_header', text: 'x' }, // no section: never implicit
    eh('x', 1, 1), eh('x', -1, 1), eh('x', 0, 0), { intent: 'edit_header', section: { index: 0 }, text: 'x' },
    { intent: 'edit_header', section: { index: 0, of: 1, name: 'A' }, text: 'x' },
    eh('x', 0, 1, { variant: 'odd' }), eh(''), eh('two\nlines'),
    // GH #317: a tab is no longer rejected here -- a literal tab in text is now ALLOWED (only
    // `\r`/`\n` are a line break); see 'plan: a literal tab in text is ALLOWED' elsewhere.
    eh('x', 0, 1, { selector: { text: 'Body first.' } }), // a text selector is not a header target
  ]) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('faults at L1: ignored / written into the other kind / another section / another variant / a leftover empty paragraph / appended -> never APPLIED', async () => {
  for (const fault of ['ignore', 'hf-into-other-kind', 'hf-wrong-section', 'hf-wrong-variant', 'hf-leftover-paragraph', 'hf-append']) {
    const sections = [section(part('HA'), part('FA')), section(part('HB'), part('FB'))];
    const { x, driver } = await run(store(sections), task(eh('HA NEW', 0, 2)), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('faults at L3 (P1): the body changed / another section changed / the other kind changed / the cursor left the body -> NOT APPLIED', async () => {
  for (const fault of ['hf-body-change', 'hf-other-section', 'hf-other-kind-changed', 'hf-cursor-stuck']) {
    const sections = [section(part('HA'), part('FA')), section(part('HB'), part('FB'))];
    const { x } = await run(store(sections), task(eh('HA NEW', 0, 2)), { fault });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', `${fault}: the part itself is right`);
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN, fault);
  }
});

test('retry: a correct terminal part is SATISFIED without a write; the old text is never searched; an unresolved earlier run -> FAIL_CLOSED', async () => {
  const s = store([section(part('Final header'))]);
  const { x, m } = await run(s, task(eh('Final header')));
  assert.equal(status(x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  const r = await run(store([section(part('Final header'))]), { ...task(eh('Final header')), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(reason(r.x), 'preservation-baseline-unavailable');
  assert.equal(r.driver.calls.dispatch, 0);
});

// GH #316: optional `format` on edit_header/edit_footer -- a subset of format_text's own keys
// (bold/italic/underline/strike/fontFamily/fontSize/color) + alignment, applied to the ONE
// paragraph/run the intent writes. The fake's applyHeaderFooter merges `format` into the written
// paragraph's `props`/`alignment` (test/fakes/fake-docx-driver.cjs), the SAME shape its read side
// (partRead -> runs()/fmt()) already produces for every other header/footer test here.
test('GH #316: edit_header/edit_footer format -- CREATE applies bold/color/fontFamily/alignment, reads back, retry SATISFIED 0/0', async () => {
  const wanted = { bold: true, color: '#FF00AA', fontFamily: 'Courier New', fontSize: 28, alignment: 'center' };
  const s = store();
  await firstAndSecond(s, task(eh('HA fmt', 0, 1, { format: wanted })));
  const p = s.durableDoc.sections[0].header.default.paras[0];
  assert.equal(p.text, 'HA fmt');
  assert.equal(p.props.bold, true);
  assert.equal(p.props.color, '#FF00AA');
  assert.equal(p.props.fontFamily, 'Courier New');
  // the RAW store is half-points (the SAME convention as every other fontSize write in this fake) --
  // the paragraph-level `format.fontSize` read projection halves it back to the requested 28
  assert.equal(p.props.fontSize, 56);
  assert.equal(p.alignment, 'center');
  // footer, same shape, independently
  const s2 = store();
  await firstAndSecond(s2, task(ef('FA fmt', 0, 1, { format: wanted })));
  const fp = s2.durableDoc.sections[0].footer.default.paras[0];
  assert.equal(fp.props.bold, true);
  assert.equal(fp.props.color, '#FF00AA');
});

test('GH #316: SAME text, DIFFERENT format -> NOT satisfied (re-applies the format, not just a no-op); matching text+format IS satisfied', async () => {
  const s = store([section(part('HA'))]);
  // first write: text changes AND format is requested -> APPLIED (not satisfied by the text match alone)
  const first = await run(s, task(eh('HA', 0, 1, { format: { bold: true } })));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED, 'same text, but format now requested -- must still write');
  assert.equal(s.durableDoc.sections[0].header.default.paras[0].props.bold, true);
  // retry: SAME text AND SAME format -> SATISFIED, zero writes
  const retry = await run(s, task(eh('HA', 0, 1, { format: { bold: true } })));
  assert.equal(status(retry.x), OPERATION_STATUS.SATISFIED);
  assert.equal(retry.m.writes, 0);
  // a DIFFERENT format on the same text -> APPLIED again (formatMatches correctly says "not done")
  const changed = await run(s, task(eh('HA', 0, 1, { format: { bold: false, italic: true } })));
  assert.equal(status(changed.x), OPERATION_STATUS.APPLIED, 'text unchanged, format changed -- formatMatches must not call this done');
  assert.equal(s.durableDoc.sections[0].header.default.paras[0].props.italic, true);
});

test('GH #316 review R1 (a maintainer, LIVE): naming EVERY non-default property the existing part holds (including color) escapes replacementLoss and applies; naming only SOME of them still refuses', async () => {
  const colored = { paras: [{ text: 'HA', props: { ...PLAIN, fontFamily: 'Courier New', fontSize: 28, bold: true, color: '#FF00AA' }, alignment: 'center' }] };
  // PARTIAL re-specification (bold/italic/alignment only) -- fontFamily/fontSize/color are non-default and
  // unnamed -> refused (the SAME shape a maintainer's LIVE ROUND3a measured: run-font-family, the first unnamed
  // non-default property the guard's own key order reaches)
  const sA = store([section(colored)]);
  const beforeA = snapshot(sA);
  const partial = await run(sA, task(eh('HA', 0, 1, { format: { bold: false, italic: true, alignment: 'right' } })));
  assert.equal(status(partial.x), OPERATION_STATUS.FAILED);
  assert.equal(reason(partial.x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-replacement-would-drop-run-font-family');
  assert.equal(snapshot(sA), beforeA, 'partial re-specification: nothing written');
  // FULL re-specification (every non-default property named, even re-asserting fontFamily/fontSize/color to
  // their CURRENT values) -> APPLIED; the new bold/italic/alignment land, color survives because it was named
  const sB = store([section(colored)]);
  const full = await run(sB, task(eh('HA', 0, 1, { format: { bold: false, italic: true, alignment: 'right', fontFamily: 'Courier New', fontSize: 28, color: '#FF00AA' } })));
  assert.equal(status(full.x), OPERATION_STATUS.APPLIED, JSON.stringify(full.x.appliedOperations.map(e => [e.status, e.evidence])));
  const p = sB.durableDoc.sections[0].header.default.paras[0];
  assert.equal(p.props.bold, false);
  assert.equal(p.props.italic, true);
  assert.equal(p.props.color, '#FF00AA', 'color named -> survives the replacement');
  assert.equal(p.alignment, 'right');
});

test('GH #316: paragraph mode + format -- the named paragraph gets both the new text and the format, every other paragraph untouched', async () => {
  const two = { paras: [{ text: 'P0', props: { ...PLAIN }, alignment: 'left' }, { text: 'P1', props: { ...PLAIN }, alignment: 'left' }] };
  const s = store([{ header: { default: two }, footer: { default: null } }]);
  await firstAndSecond(s, task(eh('P0 NEW', 0, 1, { paragraph: { index: 0, of: 2 }, format: { bold: true, alignment: 'right' } })));
  const paras = s.durableDoc.sections[0].header.default.paras;
  assert.equal(paras[0].text, 'P0 NEW');
  assert.equal(paras[0].props.bold, true);
  assert.equal(paras[0].alignment, 'right');
  assert.equal(paras[1].text, 'P1', 'the other paragraph untouched');
  assert.equal(!!paras[1].props.bold, false, 'the other paragraph\'s format untouched');
});

test('GH #316: format validation -- unknown keys and invalid values refused BEFORE dispatch (plan-failed), 0 dispatch', async () => {
  const bad = [
    eh('x', 0, 1, { format: { highlight: 'yellow' } }), // not in HF_FORMAT_KEYS (format_text-only key)
    eh('x', 0, 1, { format: { language: 'hu-HU' } }), // not in HF_FORMAT_KEYS
    eh('x', 0, 1, { format: { color: 'FF00AA' } }), // missing leading '#'
    eh('x', 0, 1, { format: { alignment: 'diagonal' } }),
    eh('x', 0, 1, { format: { fontSize: 'large' } }),
    eh('x', 0, 1, { format: {} }), // empty format
    eh('x', 0, 1, { format: 'bold' }), // not an object
    eh('x', 0, 1, { format: ['bold'] }), // an array
  ];
  for (const st of bad) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0, JSON.stringify(st));
  }
  assert.equal(hf.HF_FORMAT_KEYS.includes('highlight'), false);
  assert.equal(hf.HF_FORMAT_KEYS.includes('language'), false);
});

// GH #317: `tabs` added to HF_FORMAT_KEYS -- the SAME W-09 shape format_text already ships for
// body paragraphs (points, align restricted to formatting.cjs's own TAB_ALIGNS), PARAGRAPH-level
// like `alignment` (never a run property). The fake's applyFmt/fmt (test/fakes/fake-docx-driver.cjs)
// store/read it on the paragraph object directly, the same shape header-footer.cjs's own
// formatMatches/replacementLoss now expect.
test('GH #317: edit_header tabs -- CREATE applies, reads back paragraph-level (not a run property), retry SATISFIED 0/0', async () => {
  const wanted = { tabs: [{ pos: 288, align: 'right' }] };
  const s = store();
  await firstAndSecond(s, task(eh('HA tabs', 0, 1, { format: wanted })));
  const p = s.durableDoc.sections[0].header.default.paras[0];
  assert.equal(p.text, 'HA tabs');
  assert.deepEqual(p.tabs, [{ pos: 288, align: 'right' }]);
  assert.equal('tabs' in p.props, false, 'tabs is paragraph-level, never merged into the run props');
});

test('GH #317: SAME text, DIFFERENT tabs -> NOT satisfied (re-applies); matching text+tabs IS satisfied', async () => {
  const s = store([section(part('HA'))]);
  const first = await run(s, task(eh('HA', 0, 1, { format: { tabs: [{ pos: 144, align: 'left' }] } })));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED, 'same text, but tabs now requested -- must still write');
  assert.deepEqual(s.durableDoc.sections[0].header.default.paras[0].tabs, [{ pos: 144, align: 'left' }]);
  const retry = await run(s, task(eh('HA', 0, 1, { format: { tabs: [{ pos: 144, align: 'left' }] } })));
  assert.equal(status(retry.x), OPERATION_STATUS.SATISFIED);
  assert.equal(retry.m.writes, 0);
  const changed = await run(s, task(eh('HA', 0, 1, { format: { tabs: [{ pos: 288, align: 'right' }] } })));
  assert.equal(status(changed.x), OPERATION_STATUS.APPLIED, 'tabs changed -- formatMatches must not call this done');
  assert.deepEqual(s.durableDoc.sections[0].header.default.paras[0].tabs, [{ pos: 288, align: 'right' }]);
});

test('GH #317: an existing part\'s own tabs, unnamed, refuse the full replacement (silent loss); naming tabs escapes it, the SAME shape color/fontFamily already use', () => {
  const tabbed = { paras: [{ text: 'HA', props: { ...PLAIN }, alignment: 'left', tabs: [{ pos: 200, align: 'left' }] }] };
  return (async () => {
    const sA = store([section(tabbed)]);
    const beforeA = snapshot(sA);
    const partial = await run(sA, task(eh('HA', 0, 1, { format: { bold: true } })));
    assert.equal(status(partial.x), OPERATION_STATUS.FAILED);
    assert.equal(reason(partial.x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-replacement-would-drop-tabs');
    assert.equal(snapshot(sA), beforeA, 'nothing written');
    const sB = store([section(tabbed)]);
    const full = await run(sB, task(eh('HA', 0, 1, { format: { bold: true, tabs: [{ pos: 200, align: 'left' }] } })));
    assert.equal(status(full.x), OPERATION_STATUS.APPLIED, JSON.stringify(full.x.appliedOperations.map(e => [e.status, e.evidence])));
    assert.equal(sB.durableDoc.sections[0].header.default.paras[0].props.bold, true);
  })();
});

test('GH #317: paragraph mode + tabs format -- the named paragraph gets the tabs, every other paragraph untouched', async () => {
  const two = { paras: [{ text: 'P0', props: { ...PLAIN }, alignment: 'left' }, { text: 'P1', props: { ...PLAIN }, alignment: 'left' }] };
  const s = store([{ header: { default: two }, footer: { default: null } }]);
  await firstAndSecond(s, task(eh('P0 NEW', 0, 1, { paragraph: { index: 0, of: 2 }, format: { tabs: [{ pos: 288, align: 'right' }] } })));
  const paras = s.durableDoc.sections[0].header.default.paras;
  assert.deepEqual(paras[0].tabs, [{ pos: 288, align: 'right' }]);
  assert.equal(paras[1].text, 'P1', 'the other paragraph untouched');
  assert.deepEqual(paras[1].tabs, undefined, 'the other paragraph\'s tabs untouched');
});

test('GH #317: tabs validation -- a malformed entry is refused BEFORE dispatch (plan-failed), 0 dispatch', async () => {
  const bad = [
    eh('x', 0, 1, { format: { tabs: [{ pos: -1, align: 'left' }] } }), // negative pos
    eh('x', 0, 1, { format: { tabs: [{ pos: 100, align: 'decimal' }] } }), // decimal refused (format_text's own rule)
    eh('x', 0, 1, { format: { tabs: 'left' } }), // not an array
    eh('x', 0, 1, { format: { tabs: [{ align: 'left' }] } }), // missing pos
  ];
  for (const st of bad) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0, JSON.stringify(st));
  }
  assert.equal(hf.HF_FORMAT_KEYS.includes('tabs'), true);
});

// GH #317 "lines" (internal tracker, step 3): edit_header/edit_footer gain an optional `lines: string[]`
// instead of `text` -- a full-part write of N paragraphs, each own text, the SAME format (tabs
// included). Mutually exclusive with `text` and with `paragraph` (lines IS a full-replacement
// mode, never a single-existing-paragraph edit). The fake's applyHeaderFooter (test/fakes/
// fake-docx-driver.cjs) models the write's END STATE (N paragraphs, each own applyFmt), not the
// RemoveAllElements+Push calls themselves -- the SAME level of abstraction `fresh()` already uses.
const ehLines = (lines, index = 0, of = 1, extra = {}) => ({ intent: 'edit_header', section: { index, of }, lines, ...extra });
const efLines = (lines, index = 0, of = 1, extra = {}) => ({ intent: 'edit_footer', section: { index, of }, lines, ...extra });

test('GH #317 lines: CREATE on a fresh header writes N paragraphs, in order, SAME tabs format on every one; retry SATISFIED 0/0', async () => {
  const s = store();
  const lines = ['L0', 'L1', 'L2'];
  const wanted = { tabs: [{ pos: 288, align: 'right' }] };
  await firstAndSecond(s, task(ehLines(lines, 0, 1, { format: wanted })));
  const paras = s.durableDoc.sections[0].header.default.paras;
  assert.deepEqual(paras.map(x => x.text), lines);
  assert.ok(paras.every(x => Array.isArray(x.tabs) && x.tabs.length === 1 && x.tabs[0].pos === 288 && x.tabs[0].align === 'right'), JSON.stringify(paras));
});

test('GH #317 lines: a DIFFERENT number of lines than the existing PLAIN paragraph count still applies (no loss, nothing to acknowledge)', async () => {
  const s = store([section(part('ONE'))]); // existing: 1 plain paragraph
  await firstAndSecond(s, task(ehLines(['A', 'B', 'C'])));
  assert.deepEqual(s.durableDoc.sections[0].header.default.paras.map(x => x.text), ['A', 'B', 'C']);
});

test('GH #317 lines: text and lines together -> plan-failed, 0 dispatch; lines and paragraph together -> plan-failed; pageNumber and lines together -> plan-failed', async () => {
  const cases = [
    { intent: 'edit_header', section: { index: 0, of: 1 }, text: 'x', lines: ['a', 'b'] },
    { intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['a', 'b'], paragraph: { index: 0, of: 2 } },
    { intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['a', 'b'], pageNumber: true },
  ];
  for (const st of cases) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0, JSON.stringify(st));
  }
});

test('GH #317 lines: malformed lines (empty array, non-string element, a newline inside a line, neither text nor lines) -> plan-failed before dispatch', async () => {
  const bad = [
    { intent: 'edit_header', section: { index: 0, of: 1 }, lines: [] },
    { intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['ok', 42] },
    { intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['ok', ''] },
    // GH #317: a tab inside a line is no longer malformed -- now ALLOWED; see
    // 'plan: a literal tab in each line of lines is ALLOWED' elsewhere.
    { intent: 'edit_header', section: { index: 0, of: 1 }, lines: ['ok', 'a\nb'] },
    { intent: 'edit_header', section: { index: 0, of: 1 } }, // neither text nor lines
  ];
  for (const st of bad) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0, JSON.stringify(st));
  }
});

test('GH #317 lines: an UNNAMED formatted EXISTING paragraph -- ANY of them, not just the first -- refuses the full replacement; naming the property escapes it', async () => {
  // two existing paragraphs: the SECOND one carries bold the caller's format does not name
  const mixed = { paras: [{ text: 'P0', props: { ...PLAIN }, alignment: 'left' }, { text: 'P1', props: { ...PLAIN, bold: true }, alignment: 'left' }] };
  const sA = store([section(mixed)]);
  const beforeA = snapshot(sA);
  const partial = await run(sA, task(ehLines(['A', 'B', 'C'])));
  assert.equal(status(partial.x), OPERATION_STATUS.FAILED);
  assert.equal(reason(partial.x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-replacement-would-drop-run-bold');
  assert.equal(snapshot(sA), beforeA, 'nothing written');
  const sB = store([section(mixed)]);
  const full = await run(sB, task(ehLines(['A', 'B', 'C'], 0, 1, { format: { bold: true } })));
  assert.equal(status(full.x), OPERATION_STATUS.APPLIED, JSON.stringify(full.x.appliedOperations.map(e => [e.status, e.evidence])));
});

test('GH #317 lines: SAME lines, DIFFERENT tabs -> NOT satisfied (re-applies, checked on EVERY paragraph); matching lines+tabs IS satisfied', async () => {
  const s = store();
  const lines = ['L0', 'L1'];
  const first = await run(s, task(ehLines(lines, 0, 1, { format: { tabs: [{ pos: 144, align: 'left' }] } })));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED);
  const retry = await run(s, task(ehLines(lines, 0, 1, { format: { tabs: [{ pos: 144, align: 'left' }] } })));
  assert.equal(status(retry.x), OPERATION_STATUS.SATISFIED);
  assert.equal(retry.m.writes, 0);
  const changed = await run(s, task(ehLines(lines, 0, 1, { format: { tabs: [{ pos: 288, align: 'right' }] } })));
  assert.equal(status(changed.x), OPERATION_STATUS.APPLIED, 'tabs changed on both paragraphs -- formatMatches must not call this done');
});

test('GH #317 lines: edit_footer works the SAME way as edit_header; the other kind stays absent', async () => {
  const s = store();
  await firstAndSecond(s, task(efLines(['F0', 'F1'])));
  assert.deepEqual(partOf(s, 0, 'footer'), ['F0', 'F1']);
  assert.equal(partOf(s, 0, 'header'), null);
});

test('schema (GH #317): edit_header/edit_footer accept lines instead of text; text is no longer schema-required (plan() enforces exactly-one, not the schema)', () => {
  const schemaProblem = (intent, s) => check(D.DOCX_STEP_SCHEMAS_ALL[intent], s, 'step');
  for (const intent of ['edit_header', 'edit_footer']) {
    assert.equal(schemaProblem(intent, { intent, section: { index: 0, of: 1 }, lines: ['A', 'B'] }), null);
    assert.equal(schemaProblem(intent, { intent, section: { index: 0, of: 1 }, text: 'A' }), null, 'text alone still validates');
    assert.equal(schemaProblem(intent, { intent, section: { index: 0, of: 1 } }), null, 'neither text nor lines is SCHEMA-valid -- the family\'s own plan() is the one that refuses it (tested above, via the full run() pipeline)');
    assert.ok(schemaProblem(intent, { intent, section: { index: 0, of: 1 }, lines: [] }), 'lines needs at least one item');
  }
});

test('lowering (lines): meta carries `lines`, never `text`; the body generalizes RemoveAllElements + Push per further line', () => {
  const op = planOperations(task(ehLines(['A', 'B', 'C'], 0, 2))).operations[0];
  const low = hf.lower(op, { kind: 'section_part', sectionIndex: 0, sectionCount: 2, part: 'header', variant: 'default', create: false });
  assert.deepEqual(low.meta.lines, ['A', 'B', 'C']);
  assert.equal('text' in low.meta, false);
  assert.match(low.body, /Array\.isArray\(m\.lines\)/);
  assert.match(low.body, /hf\.Push\(np\)/);
  assert.doesNotMatch(low.body, /evaluate|asc_|private_|Set_CurHdrFtr|Create_SectionHdrFtr|SetDocPosType|MoveCursor|\.Document\b|\.Search\(/);
});

test('lowering: public section API only (no frame.evaluate / internal / cursor call); the command re-checks count, existence, inheritance and content', () => {
  const op = planOperations(task(eh('H', 0, 2))).operations[0];
  const low = hf.lower(op, { kind: 'section_part', sectionIndex: 0, sectionCount: 2, part: 'header', variant: 'default', create: false });
  assert.deepEqual([low.meta.sectionIndex, low.meta.sectionCount, low.meta.part, low.meta.variant, low.meta.create, low.meta.text], [0, 2, 'header', 'default', false, 'H']);
  assert.match(low.body, /secs\.length !== m\.sectionCount/);
  assert.match(low.body, /!!existing === m\.create/);
  // GH #289 H-07a: the inheriting sections re-read and compared with the named set (was: any later section without a part refused)
  assert.match(low.body, /JSON\.stringify\(aff\) !== JSON\.stringify\(m\.inheritedBy\)/);
  assert.match(low.body, /hf\.RemoveAllElements\(\)/);
  assert.doesNotMatch(low.body, /evaluate|asc_|private_|Set_CurHdrFtr|Create_SectionHdrFtr|SetDocPosType|MoveCursor|\.Document\b|\.Search\(/);
  assert.equal(op.preservationMode, 'PRIMITIVE_INTEGRAL');
});
