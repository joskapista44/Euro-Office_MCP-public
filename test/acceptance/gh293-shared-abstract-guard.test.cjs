'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
// GH #293 L7: the set_numbering / set_list_type co-member guards read the word /5 Bridge's whole affected set -- the released paths run on /5
const L7_V5 = '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5';
const { createReadTool } = require('../../src/server/read-tools.cjs');

/**
 * GH #293 N5 (LIVE, docs/evidence/gh293-n5-word-shared-abstractnum-sole-member-B-changed-20261022.json): Word's shared definition -- list A
 * (numId 1) and list B (numId 2) on ONE abstractNum. With A1 the sole member of numId A, the released numId-only guard let
 * set_numbering { code: 'a.' } through; it changed list B too (decimal -> lowerLetter, saved) and was caught only after the fact (L3).
 * The public ApiNumbering.ToJSON abstractNum key tells the lists apart BEFORE the write (docs/evidence/gh293-n5-word-shared-abstractnum-
 * public-recognition-20261022.json): the refuse guard now counts other-numId items on the same abstractId, and fails closed when the
 * target's abstractId is unread. read_numbering does NOT report it (session-local: a save + reopen renumbers it); the refusal names the affected items.
 */
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const item = (text, numId) => ({ text, props: PLAIN, list: { type: 'numbered', level: 0, numId } });
const FMT = [{ numFmt: 'decimal', lvlText: '%1.', suff: 'tab', lvlJc: 'left', start: 1 }, null, null, null, null, null, null, null, null];
function store(shared) {
  return createStore([{ text: 'GH293 intro', props: PLAIN }, item('GH293 A1', 'n1'), { text: 'GH293 between', props: PLAIN }, item('GH293 B1', 'n2'), item('GH293 B2', 'n2')],
    { doc: { numberingFormats: { n1: FMT, n2: FMT }, ...(shared ? { abstractOf: { n1: 'abs-0', n2: 'abs-0' } } : {}) } });
}
async function run(s, step, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, { runtime: L7_V5, ...script });
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor({ id: 'g293a', steps: [step] }));
  return { x: result.execution, driver };
}
const status = x => x.appliedOperations[0].status;
const reason = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const sn = () => ({ intent: 'set_numbering', selector: { text: 'GH293 A1' }, sharedDefinition: 'refuse', code: 'a.' });
const lt = () => ({ intent: 'set_list_type', selector: { text: 'GH293 A1' }, sharedDefinition: 'refuse', type: 'bullet' });

test('GH #293 N5: set_numbering / set_list_type on the SOLE member of numId A, while list B (another numId) sits on the same abstractNum -> FAILED before dispatch, B named', async () => {
  for (const [step, prefix] of [[sn(), 'set-numbering'], [lt(), 'list']]) {
    const { x, driver } = await run(store(true), step);
    assert.equal(status(x), OPERATION_STATUS.FAILED, step.intent);
    assert.equal(reason(x), `${prefix}-shared-abstract-definition:abs-0:2`, step.intent);
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('GH #293 N5: the same write with list B on its OWN abstractNum is not refused by this rule (negative control); an unread abstractId fails closed', async () => {
  const own = await run(store(false), sn());
  assert.notEqual(reason(own.x) || '', 'set-numbering-shared-abstract-definition:abs-0:2');
  assert.equal(status(own.x), OPERATION_STATUS.APPLIED, JSON.stringify(own.x.appliedOperations[0]));
  for (const [step, want] of [[sn(), 'set-numbering-abstract-definition-unreadable'], [lt(), 'list-abstract-definition-unreadable']]) {
    const { x, driver } = await run(store(true), step, { abstractUnreadable: true });
    assert.equal(status(x), OPERATION_STATUS.FAILED);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('GH #293 N5: read_numbering does NOT report the session-local abstractId (a reopen renumbers it: a false change in every comparison)', async () => {
  const s = store(true);
  const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(s, {}) });
  const r = await read.handler({ fileId: '1', operation: 'read_numbering' });
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 300));
  const ids = Object.fromEntries(r.result.items.filter(i => i.numbering && i.numbering.present).map(i => [i.text, [i.numbering.numId, 'abstractId' in i.numbering]]));
  assert.deepEqual(ids, { 'GH293 A1': ['n1', false], 'GH293 B1': ['n2', false], 'GH293 B2': ['n2', false] });
});
