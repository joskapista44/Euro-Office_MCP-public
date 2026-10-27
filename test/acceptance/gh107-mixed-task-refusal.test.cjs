'use strict';

/**
 * GH #107 (a) (the reporter's NAD tasks, 1.0.1-4c014d5ca25e): a task mixing format_range with set_validation above
 * 5000 used cells was refused 'family-not-migrated-to-the-tiered-observation' -- although each step alone is observed
 * (set_validation tiered, format_range through its bounded variant). The tiered observation is TASK-level: one step
 * outside TIERED_INTENTS makes the whole task non-tiered, and the bounded variant needs every step to have one. (a) made the
 * refusal name the blockers; (b) (docs/GH107B-MIXED-TIERED-BOUNDED-DESIGN-20261014.md) observes the mix through the bounded
 * variant, so the two mixed-task cases below flip. A task with a step that has NEITHER path (filter_range) is still refused
 * with the (a) note, and a single-step refusal keeps its exact old text.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { tierOf } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { liveScope, xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');

// 10 000 used cells: above the 5000 full-read bound, below the tiered ceiling
const sheets = [{ name: 'S', used: 'A1:AG1500' }];
const fmt = { intent: 'format_range', id: 'f1', sheet: 'S', range: 'C18:C67', format: { fill: '#FFFF00' } };
const val = { intent: 'set_validation', id: 'v1', sheet: 'S', range: 'D18:D67', type: 'list', items: ['a', 'b'], alertStyle: 'stop', replace: true };
const wr = { intent: 'write_range', id: 'w1', sheet: 'S', range: 'L1', values: [['x']] };
const fr = { intent: 'filter_range', id: 'fr1', sheet: 'S', range: 'A1:B2', field: 1, value: 'x' };
const tier = steps => {
  assert.equal(xlsxAdapter.planOperations({ id: 't', steps }).ok, true, 'the task plans (an unplannable one is refused for another reason)');
  return tierOf(sheets, liveScope({ id: 't', steps }));
};

test('each step alone is observed above 5000 (the control: the mix is the cause, not a step)', () => {
  assert.deepEqual([tier([fmt]).tier, tier([fmt]).bounded], ['tiered', true]);
  assert.equal(tier([val]).tier, 'tiered');
});

test('GH #107 (b): format_range + set_validation is OBSERVED (tiered + bounded), in either order', () => {
  for (const steps of [[fmt, val], [val, fmt]]) {
    const t = tier(steps);
    assert.deepEqual([t.tier, t.bounded], ['tiered', true]);
    assert.equal(t.reason, undefined);
  }
});

test('GH #107 (b): write_range + format_range + set_validation is observed too', () => {
  const t = tier([wr, fmt, val]);
  assert.deepEqual([t.tier, t.bounded], ['tiered', true]);
});

test('a mix with a step that is not observed alone does not advise a split', () => {
  const t = tier([wr, fr]);
  assert.equal(t.tier, 'refused');
  assert.equal(t.reason, 'family-not-migrated-to-the-tiered-observation; mixed task -- not tiered: fr1 (filter_range); no bounded variant: fr1 (filter_range)');
});

test('a single-step refusal keeps its exact old text (NEG: the note is only for a multi-step task)', () => {
  assert.equal(tier([fr]).reason, 'family-not-migrated-to-the-tiered-observation');
});

test('an eligible mix is untouched: write_range + set_validation is tiered, no note in the request', () => {
  const s = liveScope({ id: 't', steps: [wr, val] });
  assert.equal(s.mixedTask, undefined);
  assert.equal(tierOf(sheets, s).tier, 'tiered');
});
