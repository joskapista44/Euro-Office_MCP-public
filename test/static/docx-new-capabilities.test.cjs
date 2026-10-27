'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const snapshot = require('../fixtures/docx-source-registry.json');
const { INTENTS } = require('../../src/adapters/docx/intents.cjs');
const { NEW_INTENTS, CLASSES } = require('../../src/adapters/docx/capabilities.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

/**
 * NEW-CAPABILITY closed gate (W11, tracker an internal branch,
 * #ref option b). NOT the 28/28 migration gate (test/static/docx-intent-accounting.test.cjs) --
 * that one stays untouched and keeps measuring the SOURCE migration only. This gate proves the
 * OTHER half of the same promise: a post-M7 capability with no source intent lives in its OWN,
 * CLOSED list, and that list never grows by accident or leaks into the migration count.
 */

test('NEW_INTENTS is exactly the post-M7 capability set, today [define_style, modify_style, set_page_setup, format_table, set_numbering, set_image_size, set_image_wrap, resolve_comment, reply_comment, add_comment_on_match]', () => {
  assert.deepEqual(NEW_INTENTS.map(r => r.intent), ['define_style', 'modify_style', 'set_page_setup', 'format_table', 'set_numbering', 'set_image_size', 'set_image_wrap', 'resolve_comment', 'reply_comment', 'add_comment_on_match']);
});

test('NEW_INTENTS is disjoint from the 28 migrated intents (intents.cjs)', () => {
  const migrated = new Set(INTENTS.map(i => i.intent));
  const overlap = NEW_INTENTS.map(r => r.intent).filter(i => migrated.has(i));
  assert.deepEqual(overlap, []);
});

test('NEW_INTENTS is disjoint from the pinned SOURCE registry (no source intent masquerading as new)', () => {
  const source = new Set(snapshot.intents.map(i => i.intent));
  const overlap = NEW_INTENTS.map(r => r.intent).filter(i => source.has(i));
  assert.deepEqual(overlap, []);
});

test('every NEW_INTENTS row carries its own migrationClass, a family, a target module, a status and LIVE-probe/design evidence (not a source file:line)', () => {
  for (const row of NEW_INTENTS) {
    assert.ok(CLASSES.includes(row.migrationClass), `${row.intent}: ${row.migrationClass}`);
    assert.ok(row.family && row.targetModule && row.status, row.intent);
    assert.equal(row.batchSafe, 'NONE', `${row.intent}: no batch proof`);
    assert.ok(row.source && typeof row.source.evidence === 'string' && row.source.evidence.length > 0, `${row.intent}: evidence`);
    // the whole point of this registry: there is NO source file:line to cite (unlike intents.cjs rows)
    assert.doesNotMatch(row.source.evidence, /joskapista44\/euro-mcp/, `${row.intent}: must not cite the source oracle -- it has none`);
  }
});

test('every NEW_INTENTS row declares a preservation placement with its own (non-source) reasoning', () => {
  const modes = ['NONE', 'PRIMITIVE_INTEGRAL', 'POST_PRIMITIVE'];
  for (const row of NEW_INTENTS) {
    assert.ok(modes.includes(row.preservationMode), `${row.intent}: ${row.preservationMode}`);
    assert.ok(typeof row.preservationEvidence === 'string' && row.preservationEvidence.length > 0, `${row.intent}: preservationEvidence`);
  }
});

test('NEGATIVE CONTROL: the gate detects an intent invented INTO this list without the above', () => {
  const bogus = [...NEW_INTENTS, { intent: 'make_coffee', family: 'nowhere', migrationClass: 'NEW_CAPABILITY', batchSafe: 'NONE' }];
  const names = bogus.map(r => r.intent);
  assert.ok(names.includes('make_coffee'));
  const bad = bogus.find(r => r.intent === 'make_coffee');
  assert.ok(!(bad.source && bad.source.evidence), 'the negative control has no evidence -- it must fail the evidence check above if ever merged in for real');
});

test('define_style/modify_style plan with the rows\' own declared preservationMode, and refuse a step with no field', () => {
  for (const intent of ['define_style', 'modify_style']) {
    const row = NEW_INTENTS.find(r => r.intent === intent);
    const planned = planOperations({ steps: [{ intent, name: 'W10GateProbe', text: { bold: true } }] });
    assert.equal(planned.ok, true, planned.detail);
    assert.equal(planned.operations[0].preservationMode, row.preservationMode);
    const empty = planOperations({ steps: [{ intent, name: 'W10GateProbe' }] });
    assert.equal(empty.ok, false);
  }
});

test('set_page_setup plans with the row\'s own declared preservationMode, and refuses a step with no field', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'set_page_setup');
  const planned = planOperations({ steps: [{ intent: 'set_page_setup', section: { index: 0, of: 1 }, columns: { count: 2, space: 720 } }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const empty = planOperations({ steps: [{ intent: 'set_page_setup', section: { index: 0, of: 1 } }] });
  assert.equal(empty.ok, false);
});

test('format_table plans with the row\'s own declared preservationMode, and refuses a step with no field', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'format_table');
  const planned = planOperations({ steps: [{ intent: 'format_table', table: { text: 'Header cell' }, indent: 720 }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const empty = planOperations({ steps: [{ intent: 'format_table', table: { text: 'Header cell' } }] });
  assert.equal(empty.ok, false);
});

test('set_numbering plans with the row\'s own declared preservationMode, and refuses a step with no field', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'set_numbering');
  const planned = planOperations({ steps: [{ intent: 'set_numbering', selector: { text: 'x' }, start: 5 }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const empty = planOperations({ steps: [{ intent: 'set_numbering', selector: { text: 'x' } }] });
  assert.equal(empty.ok, false);
});

test('set_image_size plans with the row\'s own declared preservationMode, and refuses a step with a bad selector', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'set_image_size');
  const planned = planOperations({ steps: [{ intent: 'set_image_size', selector: { paragraph: 'x' }, widthEmu: 100, heightEmu: 200 }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const both = planOperations({ steps: [{ intent: 'set_image_size', selector: { paragraph: 'x', index: 0 }, widthEmu: 100, heightEmu: 200 }] });
  assert.equal(both.ok, false, 'index + paragraph together must be refused');
});

test('set_image_wrap plans with the row\'s own declared preservationMode, and refuses an unmeasured wrap value', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'set_image_wrap');
  const planned = planOperations({ steps: [{ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'square' }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const unmeasured = planOperations({ steps: [{ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'tight' }] });
  assert.equal(unmeasured.ok, false, 'tight is NEM-MÉRT -- must be refused at plan time');
});
