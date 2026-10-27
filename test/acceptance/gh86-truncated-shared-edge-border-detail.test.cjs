'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #86: above the 5000-cell read bound a shared-edge border mode (outline / inside / all / remove) has no bounded
 * variant (X-L13/8) and refuses xlsx-observation-truncated. The refusal must name the bound and -- for the drawing
 * modes -- the supported path (border.mode 'cells', merge-aware). The reason code itself stays unchanged.
 */

const book = () => [{ name: 'Data', cells: { A1: { v: 'a' }, B2: { v: 1 } } }];
const task = spec => ({ id: 'gh86', steps: [{ intent: 'format_range', sheet: 'Data', range: 'A1:B2', format: spec }] });

async function run(spec) {
  let n = 0;
  const store = createStore(book());
  const driver = createFakeXlsxDriver(store, { readTruncated: true });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const t = task(spec);
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, store };
}

const evidenceOf = result => JSON.stringify(result);

for (const mode of ['all', 'inside', 'outline']) {
  test(`GH #86: a truncated read with border mode '${mode}' names the bound and border.mode 'cells'`, async () => {
    const { result, driver } = await run({ border: { mode, style: 'thin', color: '#000000' } });
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    const s = evidenceOf(result);
    assert.match(s, /xlsx-observation-truncated/);
    assert.match(s, /45000-cell read bound/);
    assert.match(s, /border\.mode 'cells'/);
  });
}

test("GH #86: a truncated read with border mode 'remove' names the bound, offers no drawing mode", async () => {
  const { result } = await run({ border: { mode: 'remove', sides: ['top'] } });
  const s = evidenceOf(result);
  assert.match(s, /xlsx-observation-truncated/);
  assert.match(s, /45000-cell read bound/);
  assert.doesNotMatch(s, /border\.mode 'cells'/);
});

test('GH #86 negative control: a truncated read on a non-border format carries no border hint', async () => {
  const { result } = await run({ wrap: true });
  const s = evidenceOf(result);
  assert.match(s, /xlsx-observation-truncated/);
  assert.doesNotMatch(s, /border\.mode/);
});
