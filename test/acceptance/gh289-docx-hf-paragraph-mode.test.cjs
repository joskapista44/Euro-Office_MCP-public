'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const hf = require('../../src/adapters/docx/families/header-footer.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #289 H-01 offline acceptance: edit_header / edit_footer `paragraph: { index, of }` -- the text of ONE paragraph
 * of an existing part changes, its formatting and every run without text (a field, an image) kept
 * (docs/GH289-HF-DOCUMENTED-LIMITS-DECISION.md; LIVE basis: docs/evidence/gh289-hf-shapes-public-api-20261021.json).
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const COURIER = { bold: true, italic: false, underline: false, strike: false, fontFamily: 'Courier New', fontSize: 28, language: 'hu-HU' };
const styled = (text, extra = {}) => ({ text, props: { ...COURIER }, alignment: 'center', style: 'Header', ...extra });
const plain = (text, extra = {}) => ({ text, props: { ...PLAIN }, alignment: 'left', ...extra });
const section = (h, f) => ({ header: { default: h || null }, footer: { default: f || null } });
const store = sections => createStore([{ text: 'Body first.', props: { ...PLAIN } }], { doc: { sections } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'gh289', steps });
const ph = (text, index, of, extra = {}) => ({ intent: 'edit_header', section: { index: 0, of: 1 }, paragraph: { index, of }, text, ...extra });
const pf = (text, index, of, extra = {}) => ({ intent: 'edit_footer', section: { index: 0, of: 1 }, paragraph: { index, of }, text, ...extra });
const status = x => x.appliedOperations[0].status;
const reason = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const parts = s => s.durableDoc.sections[0];
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.sections]);

test('a styled two-paragraph header: paragraph 0 of 2 changes, its style, alignment and run format kept, paragraph 1 untouched; retry 0/0', async () => {
  const s = store([section({ paras: [styled('HA styled one'), plain('HA plain two')] }, { paras: [plain('Footer stays')] })]);
  const first = await run(s, task(ph('HA styled NEW', 0, 2)));
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(status(first.x), OPERATION_STATUS.APPLIED);
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  const h = parts(s).header.default.paras;
  assert.deepEqual([h[0].text, h[0].style, h[0].alignment, h[0].props.fontFamily, h[0].props.bold], ['HA styled NEW', 'Header', 'center', 'Courier New', true]);
  assert.equal(h[1].text, 'HA plain two');
  assert.equal(parts(s).footer.default.paras[0].text, 'Footer stays');
  const second = await run(s, task(ph('HA styled NEW', 0, 2)));
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([second.m.writes, second.m.commands], [0, 0]);
});

test('"Page <PAGE>": the literal run changes, the field run (no text) and the inline content kept', async () => {
  const s = store([section(null, { paras: [plain('', { runs: [{ text: 'Page ' }, { text: '' }], inline: ['pgNum'], style: 'Footer' })] })]);
  const r = await run(s, task(pf('Oldal', 0, 1)));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  const p = parts(s).footer.default.paras[0];
  assert.deepEqual([p.runs.map(x => x.text), p.inline, p.style], [['Oldal', ''], ['pgNum'], 'Footer']);
});

test('refusals before dispatch: count mismatch, no part, two text runs, a table in the part, an unread inline list, a later section inheriting', async () => {
  const two = { paras: [plain('', { runs: [{ text: 'Bold ', props: { ...COURIER } }, { text: 'plain' }] })] };
  const cases = [
    [store([section({ paras: [styled('H')] })]), ph('x', 0, 2), 'FAIL_CLOSED:paragraph-descriptor-mismatch'],
    [store([section(null)]), ph('x', 0, 1), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-paragraph-mode-needs-an-existing-part'],
    [store([section(two)]), ph('x', 0, 1), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:header-footer-paragraph-text-runs:2'],
    [store([section({ nonPlain: true, paras: [styled('H')] })]), ph('x', 0, 1), 'header-footer-not-plain-paragraphs-of-runs'],
    [store([section({ paras: [styled('H', { inline: null })] })]), ph('x', 0, 1), 'header-footer-inline-content-unreadable'],
  ];
  for (const [s, st, want] of cases) {
    const before = snapshot(s);
    const { x, driver, m } = await run(s, task(st));
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.deepEqual([driver.calls.dispatch, m.writes], [0, 0], want);
    assert.equal(snapshot(s), before, `${want}: nothing written`);
  }
  const s2 = createStore([{ text: 'Body.', props: { ...PLAIN } }], { doc: { sections: [section({ paras: [styled('H0')] }), section(null)] } });
  const r2 = await run(s2, task({ ...ph('x', 0, 1), section: { index: 0, of: 2 } }));
  assert.equal(reason(r2.x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:later-section-inherits-this-part:1');
  assert.equal(r2.driver.calls.dispatch, 0);
});

test('L3 is the gate: a write that drops the format, the inline content, or changes the OTHER paragraph is never APPLIED', async () => {
  for (const fault of ['hf-para-drop-format', 'hf-para-drop-inline', 'hf-para-other-paragraph']) {
    const s = store([section({ paras: [styled('HA one', { inline: ['pgNum'] }), plain('HA two')] })]);
    const { x } = await run(s, task(ph('HA one NEW', 0, 2)), { fault });
    assert.notEqual(status(x), OPERATION_STATUS.APPLIED, fault);
  }
});

test('H-05 (LIVE, gh289-290-remaining-probe-20261023.json): a header holding a paragraph, a table and a paragraph -- the paragraphs OUTSIDE the table change in place (the table not counted), the table kept; retry 0/0; the full replacement stays refused', async () => {
  for (const [label, index, want] of [['before the table', 0, ['HT before NEW', 'HT after']], ['after the table', 1, ['HT before', 'HT after NEW']]]) {
    const s = store([section({ paras: [styled('HT before'), plain('HT after')], tables: [{ at: 1, rows: [['HT cell 0', 'HT cell 1']] }] })]);
    const text = `${want[index]}`;
    const first = await run(s, task(ph(text, index, 2)));
    assert.equal(status(first.x), OPERATION_STATUS.APPLIED, `${label}: ${JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence]))}`);
    assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
    assert.deepEqual(parts(s).header.default.paras.map(x => x.text), want, label);
    assert.deepEqual(parts(s).header.default.tables[0].rows, [['HT cell 0', 'HT cell 1']], `${label}: the table kept`);
    const second = await run(s, task(ph(text, index, 2)));
    assert.equal(status(second.x), OPERATION_STATUS.SATISFIED, label);
    assert.deepEqual([second.m.writes, second.m.commands], [0, 0], label);
  }
  // L3: a write that also touches the table is never APPLIED
  const t = store([section({ paras: [styled('HT before'), plain('HT after')], tables: [{ at: 1, rows: [['HT cell 0', 'HT cell 1']] }] })]);
  const { x } = await run(t, task(ph('HT after NEW', 1, 2)), { fault: 'hf-para-touch-table' });
  assert.notEqual(status(x), OPERATION_STATUS.APPLIED);
  // the count is the part's own paragraphs: naming the element count (3) is a descriptor mismatch; the full replacement is refused
  for (const [st, want] of [[ph('x', 0, 3), 'FAIL_CLOSED:paragraph-descriptor-mismatch'], [{ intent: 'edit_header', section: { index: 0, of: 1 }, text: 'x' }, 'header-footer-not-plain-paragraphs-of-runs']]) {
    const s = store([section({ paras: [styled('HT before'), plain('HT after')], tables: [{ at: 1, rows: [['HT cell 0', 'HT cell 1']] }] })]);
    const before = snapshot(s);
    const r = await run(s, task(st));
    assert.equal(reason(r.x), want, want);
    assert.equal(r.driver.calls.dispatch, 0, want);
    assert.equal(snapshot(s), before, want);
  }
});

test('plan: a malformed paragraph descriptor is refused at planning', () => {
  for (const q of [{ index: 1, of: 1 }, { index: -1, of: 1 }, { index: 0 }, { index: 0, of: 1, extra: 1 }, 'p0']) {
    assert.equal(hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, paragraph: q, text: 'x' }, 0).ok, false, JSON.stringify(q));
  }
  assert.equal(hf.plan({ intent: 'edit_header', section: { index: 0, of: 1 }, paragraph: { index: 0, of: 1 }, text: 'x' }, 0).ok, true);
});

// ---- H-07a: a part that LATER sections inherit -- written only with the affected set named (inheritedBy) ----------

const multi = sections => createStore([{ text: 'Body.', props: { ...PLAIN } }], { doc: { sections } });
const eh3 = (text, extra = {}) => ({ intent: 'edit_header', section: { index: 0, of: 3 }, text, ...extra });
const evidenceOf = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence) || e || {}; };

test('H-07a: S1 and S2 inherit S0\'s header -- refused without inheritedBy (the set named), APPLIED with [1, 2], they keep no part of their own', async () => {
  const sections = () => [section({ paras: [plain('H0')] }), section(null), section(null)];
  const s1 = multi(sections());
  const r1 = await run(s1, task(eh3('H0 NEW')));
  assert.equal(reason(r1.x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:later-section-inherits-this-part:1,2');
  assert.deepEqual(evidenceOf(r1.x).inheritingSections, [1, 2]);
  assert.equal(r1.driver.calls.dispatch, 0);
  const s2 = multi(sections());
  const r2 = await run(s2, task(eh3('H0 NEW', { inheritedBy: [1, 2] })));
  assert.equal(status(r2.x), OPERATION_STATUS.APPLIED, JSON.stringify(r2.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual(s2.durableDoc.sections.map(x => x.header.default && x.header.default.paras[0].text), ['H0 NEW', null, null]);
  const again = await run(s2, task(eh3('H0 NEW', { inheritedBy: [1, 2] })));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
});

test('H-07a: a stale or wrong acknowledgement is refused, naming the real set', async () => {
  for (const [sections, named] of [[[section({ paras: [plain('H0')] }), section(null), section(null)], [1]], [[section({ paras: [plain('H0')] }), section({ paras: [plain('H1')] }), section({ paras: [plain('H2')] })], [1]]]) {
    const s = multi(sections);
    const before = snapshot(s);
    const r = await run(s, task(eh3('H0 NEW', { inheritedBy: named })));
    assert.ok(/^FAIL_CLOSED:inherited-by-mismatch:(1,2|none)$/.test(reason(r.x)), `${reason(r.x)} ${JSON.stringify(named)}`);
    assert.equal(r.driver.calls.dispatch, 0);
    assert.equal(snapshot(s), before);
  }
});

test('H-07a precision: S2 without a part inherits S1 (which has its own), not S0 -- a write to S0 needs no acknowledgement', async () => {
  const s = multi([section({ paras: [plain('H0')] }), section({ paras: [plain('H1')] }), section(null)]);
  const r = await run(s, task(eh3('H0 NEW')));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual(s.durableDoc.sections.map(x => x.header.default && x.header.default.paras[0].text), ['H0 NEW', 'H1', null]);
});

test('H-07a with paragraph mode: the styled first paragraph of an inherited header, its format kept', async () => {
  const s = multi([section({ paras: [styled('H0 styled'), plain('H0 two')] }), section(null), section(null)]);
  const r = await run(s, task(eh3('H0 styled NEW', { paragraph: { index: 0, of: 2 }, inheritedBy: [1, 2] })));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  const h = s.durableDoc.sections[0].header.default.paras[0];
  assert.deepEqual([h.text, h.style, h.alignment], ['H0 styled NEW', 'Header', 'center']);
});

test('H-07a plan: inheritedBy must be an ascending list of later section indexes', () => {
  for (const v of [[0], [3], [2, 1], [1, 1], ['1'], 1]) {
    assert.equal(hf.plan({ intent: 'edit_header', section: { index: 0, of: 3 }, inheritedBy: v, text: 'x' }, 0).ok, false, JSON.stringify(v));
  }
  assert.equal(hf.plan({ intent: 'edit_header', section: { index: 0, of: 3 }, inheritedBy: [1, 2], text: 'x' }, 0).ok, true);
});

// ---- H-04b: pageNumber -- a PAGE field after the text (full mode) ------------------------------------------------

test('H-04b: a new footer "Oldal <PAGE>" -- APPLIED, the paragraph holds the text and one PAGE field; retry 0/0; a write that loses the field is not APPLIED', async () => {
  const s = store([section(null, null)]);
  const st = { intent: 'edit_footer', section: { index: 0, of: 1 }, text: 'Oldal ', pageNumber: true };
  const r = await run(s, task(st));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations.map(e => [e.status, e.evidence])));
  const p = parts(s).footer.default.paras[0];
  assert.deepEqual([p.runs.map(x => x.text), p.inline], [['Oldal ', ''], ['pgNum']]);
  const again = await run(s, task(st));
  assert.equal(status(again.x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([again.m.writes, again.m.commands], [0, 0]);
  const lost = await run(store([section(null, null)]), task(st), { fault: 'hf-page-number-missing' });
  assert.notEqual(status(lost.x), OPERATION_STATUS.APPLIED);
});

test('H-04b plan: pageNumber is a boolean and a full-part write only', () => {
  const base = { intent: 'edit_footer', section: { index: 0, of: 1 }, text: 'x' };
  assert.equal(hf.plan({ ...base, pageNumber: 'yes' }, 0).ok, false);
  assert.equal(hf.plan({ ...base, pageNumber: true, paragraph: { index: 0, of: 1 } }, 0).ok, false);
  assert.equal(hf.plan({ ...base, pageNumber: true }, 0).ok, true);
});
