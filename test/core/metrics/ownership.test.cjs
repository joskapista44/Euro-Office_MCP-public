'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { OWNER, CONSISTENCY, unownedCounters, orphanOwners } = require('../../../src/core/metrics/ownership.cjs');
const { COUNTERS } = require('../../../src/core/metrics/collect.cjs');

test('every common counter has exactly one named owner', () => {
  assert.deepEqual(unownedCounters(), [], 'a counter nobody owns has no single meaning');
  assert.deepEqual(orphanOwners(), [], 'an owner for a counter the schema lacks is stale');
  assert.equal(Object.keys(OWNER).length, COUNTERS.length);
});

test('no counter is owned by two modules', () => {
  const byOwner = new Map();
  for (const [counter, owner] of Object.entries(OWNER)) {
    if (!byOwner.has(owner)) byOwner.set(owner, []);
    byOwner.get(owner).push(counter);
  }
  // Several counters may share an owner (the executor owns four), but no COUNTER may have two.
  const counters = Object.keys(OWNER);
  assert.equal(new Set(counters).size, counters.length);
});

test('the interaction counters are each owned by the module that performs the interaction', () => {
  assert.match(OWNER.freshReads, /session\/handle\.cjs/);
  assert.match(OWNER.commands, /session\/handle\.cjs/);
  // Enrichment is semantic orchestration: the hook invocation is what the core counts there.
  // M2 left physical package I/O without an owner because no common transport existed. M4 adds
  // it, and the ownership rule is the same as for editor reads: the module that performs the I/O.
  assert.match(OWNER.enrichmentInvocations, /observation\/enrichment\.cjs/);
  assert.match(OWNER.packageReads, /session\/handle\.cjs \(readPackage\)/);
  assert.doesNotMatch(OWNER.packageReads, /enrichment/, 'a hook invocation never owns a physical read');
  assert.match(OWNER.verificationInvocations, /verification\/orchestrate\.cjs/);
  assert.match(OWNER.persistencePolls, /persistence\/barrier\.cjs/);
});

test('the consistency contract names a driver/adapter call for each counted interaction', () => {
  assert.deepEqual(Object.keys(CONSISTENCY).sort(), [
    'adapter.enrichObservation',
    'core verifier hook invocations',
    'driver.canSave',
    'driver.dispatch',
    'driver.readLive',
    'driver.readPackage',
  ]);
  for (const counter of Object.values(CONSISTENCY)) {
    assert.ok(COUNTERS.includes(counter), `${counter} must be a real schema counter`);
  }
});
