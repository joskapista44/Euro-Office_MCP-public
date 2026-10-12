'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { POLICY, ENRICHMENT_STATE, runEnrichment, isFailClosed } =
  require('../../../src/core/observation/enrichment.cjs');
const { Metrics } = require('../../../src/core/metrics/collect.cjs');
const { AUTHORITY } = require('../../../src/core/authority.cjs');

const handle = () => ({ id: 'h', metrics: new Metrics({ now: () => 1 }) });
const adapterWith = hook => ({ kind: 'fake', enrichObservation: hook });

test('policy NONE never calls the hook', async () => {
  let called = 0;
  const e = await runEnrichment(handle(), adapterWith(async () => { called++; return { ok: true }; }), {}, { policy: POLICY.NONE });
  assert.equal(e.state, ENRICHMENT_STATE.SKIPPED);
  assert.equal(called, 0);
  assert.equal(isFailClosed(e), false);
});

test('an adapter without the hook is NOT_IMPLEMENTED under OPTIONAL — this is XLSX', async () => {
  const e = await runEnrichment(handle(), { kind: 'xlsx', enrichObservation: null }, {}, { policy: POLICY.OPTIONAL });
  assert.equal(e.state, ENRICHMENT_STATE.NOT_IMPLEMENTED);
  assert.equal(isFailClosed(e), false, 'having nothing to add is not a failure');
});

test('the same adapter under REQUIRED is FAILED, and the reason names the capability', async () => {
  const e = await runEnrichment(handle(), { kind: 'xlsx', enrichObservation: null }, {}, { policy: POLICY.REQUIRED });
  assert.equal(e.state, ENRICHMENT_STATE.FAILED);
  assert.equal(isFailClosed(e), true);
  assert.match(e.reason, /requires it/);
});

test('OPTIONAL failure is DEGRADED, and distinguishable from NOT_IMPLEMENTED', async () => {
  const thrown = await runEnrichment(handle(), adapterWith(async () => { throw new Error('boom'); }), {}, { policy: POLICY.OPTIONAL });
  assert.equal(thrown.state, ENRICHMENT_STATE.DEGRADED);
  assert.equal(thrown.threw, true);
  assert.equal(thrown.detail, 'boom');
  assert.equal(isFailClosed(thrown), false);

  const refused = await runEnrichment(handle(), adapterWith(async () => ({ ok: false, detail: 'no' })), {}, { policy: POLICY.OPTIONAL });
  assert.equal(refused.state, ENRICHMENT_STATE.DEGRADED);
  assert.equal(refused.threw, false);

  // The distinction owner requirement 6 asks for: a broken package read must not look like an
  // adapter that simply had nothing to say.
  assert.notEqual(thrown.state, ENRICHMENT_STATE.NOT_IMPLEMENTED);
});

test('REQUIRED failure is FAILED for both a refusal and a throw', async () => {
  for (const hook of [async () => { throw new Error('x'); }, async () => ({ ok: false })]) {
    const e = await runEnrichment(handle(), adapterWith(hook), {}, { policy: POLICY.REQUIRED });
    assert.equal(e.state, ENRICHMENT_STATE.FAILED);
    assert.equal(isFailClosed(e), true);
  }
});

test('a hook returning null means "nothing to add here", not a failure', async () => {
  const e = await runEnrichment(handle(), adapterWith(async () => null), {}, { policy: POLICY.OPTIONAL });
  assert.equal(e.state, ENRICHMENT_STATE.NOT_IMPLEMENTED);
  assert.equal(e.reason, 'hook returned null');
});

test('enrichment never claims more than LIVE_READ and is marked non-live', async () => {
  const e = await runEnrichment(handle(), adapterWith(async () => ({ ok: true, enrichment: { a: 1 } })), {}, {});
  assert.equal(e.authority, AUTHORITY.LIVE_READ);
  assert.equal(e.isLiveEvidence, false, 'so a verifier cannot mistake it for what the editor showed');
  assert.deepEqual(e.data, { a: 1 });
});

test('the enrichment invocation is counted even when the hook throws — and it is not package I/O', async () => {
  const h = handle();
  await runEnrichment(h, adapterWith(async () => { throw new Error('late'); }), {}, { policy: POLICY.OPTIONAL });
  assert.equal(h.metrics.enrichmentInvocations, 1, 'the hook WAS invoked');
  const { collectMetrics } = require('../../../src/core/metrics/collect.cjs');
  const s = collectMetrics(h.metrics);
  assert.equal(s.packageReads, 0, 'an invocation is not a physical package read');
  assert.equal(s.packageInteractions, 0);
});

test('an unknown policy is a CALLER error, not a boundary failure', async () => {
  // It must not silently become OPTIONAL: a typo in a capability declaration would then quietly
  // downgrade a REQUIRED enrichment.
  await assert.rejects(
    () => runEnrichment(handle(), adapterWith(async () => ({ ok: true })), {}, { policy: 'MAYBE' }),
    /unknown enrichment policy/
  );
});

test('OPTIONAL is the default', async () => {
  const e = await runEnrichment(handle(), { kind: 'xlsx', enrichObservation: null }, {}, {});
  assert.equal(e.policy, POLICY.OPTIONAL);
});
