'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { observeFresh } = require('../../../src/core/observation/observe.cjs');
const { POLICY } = require('../../../src/core/observation/enrichment.cjs');
const { SessionHandle, readPackage } = require('../../../src/core/session/handle.cjs');
const { bindDriver } = require('../../../src/core/session/internal.cjs');
const { currentGeneration } = require('../../../src/core/mutation-epoch.cjs');
const { Metrics, collectMetrics } = require('../../../src/core/metrics/collect.cjs');
const { createFakeSession } = require('../../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../../fakes/fake-adapter.cjs');

/**
 * M4 common package transport (owner M4 brief point 10). Physical package I/O is counted at the
 * one core boundary that performs it; an enrichment INVOCATION is never taken for a package read.
 */

function session(sessionScript = {}) {
  let t = 0;
  const metrics = new Metrics({ now: () => (t += 1) });
  const handle = new SessionHandle({ id: 'pkg', fileId: 'f', adapterKind: 'fake', metrics });
  const driver = createFakeSession(sessionScript);
  bindDriver(handle, driver);
  return { handle, driver, metrics };
}

function readingAdapter(reads) {
  const a = createFakeAdapter();
  a.enrichObservation = async (_h, _primary, { readPackage: rp }) => {
    for (let i = 0; i < reads; i += 1) {
      const part = await rp({ part: `part-${i}` });
      if (!part.ok) return { ok: false, detail: part.detail };
    }
    return { ok: true, enrichment: { reads } };
  };
  return a;
}

test('an enrichment hook that reads the package twice -> packageReads 2, enrichmentInvocations 1', async () => {
  const { handle, driver, metrics } = session();
  const obs = await observeFresh(handle, readingAdapter(2), { enrichment: POLICY.REQUIRED });
  assert.equal(obs.ok, true);
  const m = collectMetrics(metrics);
  assert.equal(driver.calls.readPackage, 2);
  assert.equal(m.packageReads, 2);
  assert.equal(m.packageInteractions, 2);
  assert.equal(m.enrichmentInvocations, 1);
  assert.equal(m.totalExternalInteractions, m.editorInteractions + 2);

  // NEGATIVE CONTROL: a hook that is invoked but reads nothing is 1 invocation and 0 physical reads.
  const s2 = session();
  await observeFresh(s2.handle, readingAdapter(0), { enrichment: POLICY.REQUIRED });
  const m2 = collectMetrics(s2.metrics);
  assert.equal(m2.enrichmentInvocations, 1);
  assert.equal(m2.packageReads, 0);
});

test('a package read is read-only: it never bumps the mutation generation', async () => {
  const { handle } = session();
  const before = currentGeneration(handle);
  await readPackage(handle, { part: 'word/document.xml' });
  assert.equal(currentGeneration(handle), before);
});

test('no transport -> explicit refusal, nothing counted', async () => {
  const { handle, driver, metrics } = session();
  delete driver.readPackage;
  const r = await readPackage(handle, { part: 'x' });
  assert.equal(r.ok, false);
  assert.equal(collectMetrics(metrics).packageReads, 0);
});

test('a REQUIRED enrichment whose package read fails fails the observation closed', async () => {
  const { handle } = session({ readPackageFails: true });
  const obs = await observeFresh(handle, readingAdapter(1), { enrichment: POLICY.REQUIRED });
  assert.equal(obs.ok, false);
  assert.equal(obs.outcome, 'observation-required-enrichment-failed');
});
