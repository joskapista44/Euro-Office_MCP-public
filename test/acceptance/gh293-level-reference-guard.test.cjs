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

/**
 * GH #293 N6 (LIVE, docs/evidence/gh293-cross-level-label-probe-GAP-20261022.json): a deeper level whose label quotes the target level
 * ("%1.%2") renders the target level's counter in the target level's format. On release 16, set_numbering { code: 'a.' } on the
 * SOLE level-0 item was APPLIED and silently changed the level-1 labels "1.1" -> "a.1". The refuse guard (the product's
 * policy) now also refuses a format / start change of a level that a deeper item of the same numId quotes; a suffix change
 * (not part of the quoted label) and a deeper level that does not quote it stay writable. set_list_type gets the same rule.
 */
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const item = (text, numId, level, type = 'numbered') => ({ text, props: PLAIN, list: { type, level, numId } });
const P = 'GH293 parent'; const C1 = 'GH293 child 1'; const C2 = 'GH293 child 2';
function store(childLvlText) {
  return createStore([item(P, 'n1', 0), item(C1, 'n1', 1), item(C2, 'n1', 1)], { doc: { numberingFormats: { n1: [
    { numFmt: 'decimal', lvlText: '%1.', suff: 'tab', lvlJc: 'left', start: 1 }, { numFmt: 'decimal', lvlText: childLvlText, suff: 'tab', lvlJc: 'left', start: 1 }, null, null, null, null, null, null, null] } } });
}
async function run(s, step) {
  let t = 0;
  const driver = createFakeDocxDriver(s, { runtime: L7_V5 });
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor({ id: 'g293', steps: [step] }));
  return { x: result.execution, driver };
}
const status = x => x.appliedOperations[0].status;
const reason = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const sn = fields => ({ intent: 'set_numbering', selector: { text: P }, sharedDefinition: 'refuse', ...fields });

test('GH #293 N6: set_numbering code / start on a level that deeper items quote ("%1.%2") -> FAILED before dispatch, the quoting items named', async () => {
  for (const fields of [{ code: 'a.' }, { start: 3 }]) {
    const { x, driver } = await run(store('%1.%2'), sn(fields));
    assert.equal(status(x), OPERATION_STATUS.FAILED, JSON.stringify(fields));
    assert.equal(reason(x), 'set-numbering-level-referenced-by-other-levels:n1:level0:2', JSON.stringify(fields));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('GH #293 N6: a suffix change of that level, and a code change when the deeper level does not quote it ("%2."), are APPLIED', async () => {
  assert.equal(status((await run(store('%1.%2'), sn({ suffix: 'space' }))).x), OPERATION_STATUS.APPLIED);
  assert.equal(status((await run(store('%2.'), sn({ code: 'a.' }))).x), OPERATION_STATUS.APPLIED);
});

test('GH #293 N6: an unreadable lvlText on a deeper level counts as quoting (fail closed)', async () => {
  const { x } = await run(store(null), sn({ code: 'a.' }));
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.match(reason(x), /set-numbering-level-referenced-by-other-levels/);
});

test('GH #293 N6: set_list_type (refuse) on a level that deeper items quote -> FAILED before dispatch; not quoted -> APPLIED', async () => {
  const quoted = await run(store('%1.%2'), { intent: 'set_list_type', selector: { text: P }, type: 'bullet', sharedDefinition: 'refuse' });
  assert.equal(status(quoted.x), OPERATION_STATUS.FAILED);
  assert.equal(reason(quoted.x), 'list-level-referenced-by-other-levels:n1:level0:2');
  assert.equal(quoted.driver.calls.dispatch, 0);
  assert.equal(status((await run(store('%2.'), { intent: 'set_list_type', selector: { text: P }, type: 'bullet', sharedDefinition: 'refuse' })).x), OPERATION_STATUS.APPLIED);
});
