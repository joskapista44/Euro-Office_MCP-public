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
    eh('x', 0, 1, { variant: 'odd' }), eh(''), eh('two\nlines'), eh('tab\there'),
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
