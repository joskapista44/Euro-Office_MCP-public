'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { Metrics, collectMetrics } = require('../../../src/core/metrics/collect.cjs');
const { BOUNDARIES } = require('../../../src/core/metrics/clock.cjs');

function marked() {
  let t = 0;
  const m = new Metrics({ now: () => t });
  t = 0; m.clock.mark('openEntry');
  t = 5; m.clock.mark('handleReturned');
  t = 9; m.clock.mark('ready');
  t = 30; m.clock.mark('taskDone');
  t = 40; m.clock.mark('closed');
  m.clock
    .span(BOUNDARIES.SESSION_OPEN, 'openEntry', 'handleReturned')
    .span(BOUNDARIES.READY, 'handleReturned', 'ready')
    .span(BOUNDARIES.TASK_ELAPSED, 'ready', 'taskDone')
    .span(BOUNDARIES.TOTAL_ELAPSED, 'openEntry', 'closed');
  return m;
}

test('the two clock boundaries keep their source meanings', () => {
  const s = collectMetrics(marked());
  assert.equal(s.sessionOpenMs, 5);
  assert.equal(s.readyMs, 4);
  assert.equal(s.taskElapsedMs, 21, 'taskElapsedMs EXCLUDES the open (XLSX meaning)');
  assert.equal(s.totalElapsedMs, 40, 'totalElapsedMs INCLUDES the open (DOCX meaning)');
  assert.ok(s.taskElapsedMs < s.totalElapsedMs, 'the two spans are not interchangeable');
});

test('an unmeasured span is null, never 0', () => {
  const s = collectMetrics(new Metrics({ now: () => 0 }));
  for (const b of Object.values(BOUNDARIES)) {
    assert.equal(s[b], null, `${b} was never measured and must not read as an instant duration`);
  }
});

/**
 * CONTRACT CHANGE, not a behaviour regression. This test was written in M1 against the schema of
 * M0 12.1, where editorInteractions included verificationReads. The M2 final review established
 * that the sum DOUBLE-COUNTS a verifier that reads through readLive(): the read is already in
 * freshReads. Physical I/O and semantic orchestration are now separate accounts.
 *
 * This is the ONLY one of M1's 76 tests the change touches, and it asserts the same property --
 * that the total is the sum of its parts -- against the corrected parts.
 */
test('totalExternalInteractions = editorInteractions + packageInteractions (physical I/O only)', () => {
  const m = marked();
  m.count('freshReads', 3).count('verificationInvocations', 4).count('commands', 5)
    .count('persistencePolls', 6).count('enrichmentInvocations', 7);
  const s = collectMetrics(m);

  assert.equal(s.editorInteractions, 3 + 5 + 6, 'PHYSICAL interactions only: no verifier invocations');
  // Enrichment-hook invocations are semantic orchestration, not package I/O. M2 has no common
  // package transport, so physical package I/O is 0 however many hooks ran.
  assert.equal(s.packageReads, 0);
  assert.equal(s.packageInteractions, 0, 'enrichment invocations are not physical package I/O');
  assert.equal(s.totalExternalInteractions, s.editorInteractions + s.packageInteractions);
  assert.equal(s.totalExternalInteractions, 14);

  // Semantic orchestration is reported, and deliberately outside every physical total.
  assert.equal(s.verificationInvocations, 4);
  assert.equal(s.enrichmentInvocations, 7);
  assert.ok(
    s.totalExternalInteractions < 3 + 4 + 5 + 6 + 7,
    'counting either semantic counter as physical I/O would inflate the total'
  );

  // The retired name is gone rather than redefined: undefined is loud, a wrong number is not.
  assert.equal(s.totalInteractions, undefined);
  assert.equal(s.verificationReads, undefined);
});

test('operationsFailed is derived: planned - applied - satisfied', () => {
  const m = marked();
  m.count('operationsPlanned', 10).count('operationsApplied', 6).count('operationsSatisfied', 3);
  assert.equal(collectMetrics(m).operationsFailed, 1);
});

test('an unknown counter name fails loudly rather than being silently created', () => {
  assert.throws(() => new Metrics().count('zsirafReads'), /unknown metric counter/);
});

test('re-marking a clock boundary is a caller error', () => {
  const m = new Metrics({ now: () => 1 });
  m.clock.mark('openEntry');
  assert.throws(() => m.clock.mark('openEntry'), /already taken/);
});
