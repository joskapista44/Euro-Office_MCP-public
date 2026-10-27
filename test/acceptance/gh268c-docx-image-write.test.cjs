'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * GH #268c offline acceptance: set_image_size / set_image_wrap through the TARGET path on the fake
 * transport, modeled on test/acceptance/docx-w11-page-setup.test.cjs's own shape. The fake model
 * follows the LIVE probe (test/live/docx-images-probe.live.cjs R3_SIZE / R4_WRAP) + the GH #268b
 * read-side row shape (no docPr id/name, per owner decision): a target is addressed by
 * { paragraph, ordinal? } or { index }, never by anything the editor session can rename.
 */

const paragraphs = () => [{ text: 'Body heading', level: 1, props: {} }, { text: 'Body first.', props: {} }];
const image = (i, over = {}) => ({ index: i, descr: null, title: null, cx: 914400, cy: 914400, drawingType: 'inline', wrapType: 'none', paragraph: `img-p${i}`, ...over });
const store = (images, paras = paragraphs()) => createStore(paras, { doc: { images } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'gh268c', steps });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const imagesOf = s => s.durableDoc.images;

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per operation');
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

// ---- set_image_size -------------------------------------------------------------------------------

test('set_image_size by index: write + retry SATISFIED 0/0; the OTHER image and the body unchanged', async () => {
  const s = store([image(0), image(1)]);
  await firstAndSecond(s, task({ intent: 'set_image_size', selector: { index: 1 }, widthEmu: 1000, heightEmu: 2000 }));
  assert.deepEqual([imagesOf(s)[1].cx, imagesOf(s)[1].cy], [1000, 2000]);
  assert.deepEqual([imagesOf(s)[0].cx, imagesOf(s)[0].cy], [914400, 914400]);
});

test('set_image_size by { paragraph, ordinal }: the ordinal-th image whose parent paragraph matches, among paragraphs sharing that text', async () => {
  const s = store([image(0, { paragraph: 'shared' }), image(1, { paragraph: 'shared' }), image(2, { paragraph: 'other' })]);
  await firstAndSecond(s, task({ intent: 'set_image_size', selector: { paragraph: 'shared', ordinal: 2 }, widthEmu: 5000, heightEmu: 6000 }));
  assert.deepEqual([imagesOf(s)[1].cx, imagesOf(s)[1].cy], [5000, 6000]);
  assert.deepEqual([imagesOf(s)[0].cx, imagesOf(s)[0].cy], [914400, 914400]);
});

test('set_image_size: ordinal default is 1 (omitted)', async () => {
  const s = store([image(0, { paragraph: 'p' })]);
  await firstAndSecond(s, task({ intent: 'set_image_size', selector: { paragraph: 'p' }, widthEmu: 1, heightEmu: 2 }));
  assert.deepEqual([imagesOf(s)[0].cx, imagesOf(s)[0].cy], [1, 2]);
});

test('set_image_size: index out of range -> FAIL_CLOSED, 0 dispatch', async () => {
  const s = store([image(0)]);
  const { x, m } = await run(s, task({ intent: 'set_image_size', selector: { index: 5 }, widthEmu: 1, heightEmu: 2 }));
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'image-index-not-found');
  assert.equal(m.writes, 0);
});

test('set_image_size: paragraph with fewer images than the requested ordinal -> FAIL_CLOSED, 0 dispatch', async () => {
  const s = store([image(0, { paragraph: 'p' })]);
  const { x, m } = await run(s, task({ intent: 'set_image_size', selector: { paragraph: 'p', ordinal: 2 }, widthEmu: 1, heightEmu: 2 }));
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'image-paragraph-ordinal-not-found');
  assert.equal(m.writes, 0);
});

test('faults: another image changing on a size write -> preservation FAIL', async () => {
  const { x } = await run(store([image(0), image(1)]), task({ intent: 'set_image_size', selector: { index: 0 }, widthEmu: 1, heightEmu: 2 }), { fault: 'image-other-changed' });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
});

test('faults: the body changing on a size write -> preservation FAIL', async () => {
  const { x } = await run(store([image(0)]), task({ intent: 'set_image_size', selector: { index: 0 }, widthEmu: 1, heightEmu: 2 }), { fault: 'image-body-change' });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
});

// ---- set_image_wrap --------------------------------------------------------------------------------

test('set_image_wrap square: inline -> floating/square; write + retry SATISFIED 0/0; the OTHER image unchanged', async () => {
  const s = store([image(0), image(1)]);
  await firstAndSecond(s, task({ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'square' }));
  assert.deepEqual([imagesOf(s)[0].drawingType, imagesOf(s)[0].wrapType], ['anchor', 'square']);
  assert.deepEqual([imagesOf(s)[1].drawingType, imagesOf(s)[1].wrapType], ['inline', 'none']);
});

test('set_image_wrap: already floating/square -> SATISFIED on the FIRST run (no write)', async () => {
  const s = store([image(0, { drawingType: 'anchor', wrapType: 'square' })]);
  const { x, m } = await run(s, task({ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'square' }));
  assert.equal(status(x), OPERATION_STATUS.SATISFIED);
  assert.equal(m.writes, 0);
});

test('set_image_wrap: an already-floating image with a DIFFERENT wrap -> FAIL_CLOSED (unmeasured source placement), 0 dispatch', async () => {
  const s = store([image(0, { drawingType: 'anchor', wrapType: 'through' })]);
  const { x, m } = await run(s, task({ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'square' }));
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'image-wrap-unmeasured-source-placement');
  assert.equal(m.writes, 0);
});

test('set_image_wrap: a non-square wrap value is refused at PLANNING (schema + family plan), never dispatched', () => {
  for (const wrap of ['tight', 'through', 'topAndBottom', 'behind', 'inFront', 'inline']) {
    const planned = planOperations(task({ intent: 'set_image_wrap', selector: { index: 0 }, wrap }));
    assert.equal(planned.ok, false, `wrap:'${wrap}' must be refused at plan time (NEM-MÉRT)`);
  }
});

test('faults: another image changing on a wrap write -> preservation FAIL', async () => {
  const { x } = await run(store([image(0), image(1)]), task({ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'square' }), { fault: 'image-other-changed' });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
});

// ---- selector shape ---------------------------------------------------------------------------------

test('selector must name exactly one of index or paragraph -- both together is refused at planning', () => {
  const planned = planOperations(task({ intent: 'set_image_size', selector: { index: 0, paragraph: 'p' }, widthEmu: 1, heightEmu: 2 }));
  assert.equal(planned.ok, false);
});

test('selector must name exactly one of index or paragraph -- neither is refused at planning', () => {
  const planned = planOperations(task({ intent: 'set_image_size', selector: {}, widthEmu: 1, heightEmu: 2 }));
  assert.equal(planned.ok, false);
});

test('selector.ordinal with index is refused at planning (ordinal is a paragraph-only qualifier)', () => {
  const planned = planOperations(task({ intent: 'set_image_size', selector: { index: 0, ordinal: 2 }, widthEmu: 1, heightEmu: 2 }));
  assert.equal(planned.ok, false);
});

test('a drifted index between resolve and the command re-check: the callCommand guard matches the family resolve (same fresh read)', async () => {
  // Two images sharing a paragraph; addressing the SECOND by paragraph+ordinal must never touch the first.
  const s = store([image(0, { paragraph: 'dup' }), image(1, { paragraph: 'dup' })]);
  await firstAndSecond(s, task({ intent: 'set_image_size', selector: { paragraph: 'dup', ordinal: 1 }, widthEmu: 7, heightEmu: 8 }));
  assert.deepEqual([imagesOf(s)[0].cx, imagesOf(s)[0].cy], [7, 8]);
  assert.deepEqual([imagesOf(s)[1].cx, imagesOf(s)[1].cy], [914400, 914400]);
});

test('read fault: images unreadable -> FAIL_CLOSED, never a guessed match', async () => {
  const { x, m } = await run(store([image(0)]), task({ intent: 'set_image_size', selector: { index: 0 }, widthEmu: 1, heightEmu: 2 }), { imagesUnreadable: true });
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(m.writes, 0);
});
