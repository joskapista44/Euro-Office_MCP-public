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

test('NEW_INTENTS is exactly the post-M7 capability set, today [define_style, modify_style, set_page_setup, format_table, set_numbering, set_image_size, set_image_wrap, resolve_comment, reply_comment, add_comment_on_match, insert_table, insert_section_break, insert_image]', () => {
  assert.deepEqual(NEW_INTENTS.map(r => r.intent), ['define_style', 'modify_style', 'set_page_setup', 'format_table', 'set_numbering', 'set_image_size', 'set_image_wrap', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'insert_table', 'insert_section_break', 'insert_image', 'insert_paragraphs', 'insert_shape', 'insert_chart', 'delete_image', 'replace_image']);
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

test('insert_table plans with the row\'s own declared preservationMode, and refuses an out-of-bound rows/cols', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'insert_table');
  const planned = planOperations({ steps: [{ intent: 'insert_table', selector: { afterParagraph: 'x' }, rows: 2, cols: 2 }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const zero = planOperations({ steps: [{ intent: 'insert_table', selector: { afterParagraph: 'x' }, rows: 0, cols: 2 }] });
  assert.equal(zero.ok, false, 'rows: 0 is the measured CreateTable null boundary -- must be refused at plan time');
});

test('insert_section_break plans with the row\'s own declared preservationMode, and refuses an unmeasured type', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'insert_section_break');
  const planned = planOperations({ steps: [{ intent: 'insert_section_break', selector: { afterParagraph: 'x' }, type: 'nextPage' }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  const unmeasured = planOperations({ steps: [{ intent: 'insert_section_break', selector: { afterParagraph: 'x' }, type: 'evenPage' }] });
  assert.equal(unmeasured.ok, false, 'evenPage/oddPage/nextColumn were not measured -- must be refused at plan time');
});

test('insert_image plans with the row\'s own declared preservationMode, and refuses a malformed source and an out-of-range EMU', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'insert_image');
  const source = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
  const planned = planOperations({ steps: [{ intent: 'insert_image', selector: { afterParagraph: 'x' }, source, widthEmu: 914400, heightEmu: 914400 }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  assert.equal(planned.operations[0].position, 'after', 'position defaults to \'after\'');
  const bad = planOperations({ steps: [{ intent: 'insert_image', selector: { afterParagraph: 'x' }, source: 'data:image/gif;base64,abcd', widthEmu: 914400, heightEmu: 914400 }] });
  assert.equal(bad.ok, false, 'a GIF data URI is not PNG/JPEG -- must be refused at plan time');
  const zero = planOperations({ steps: [{ intent: 'insert_image', selector: { afterParagraph: 'x' }, source, widthEmu: 0, heightEmu: 914400 }] });
  assert.equal(zero.ok, false, 'widthEmu: 0 is not a positive integer -- must be refused at plan time');
});

test('insert_paragraphs plans with the row\'s own declared preservationMode, and refuses an out-of-range count and an unknown format key', () => {
  const row = NEW_INTENTS.find(r => r.intent === 'insert_paragraphs');
  const planned = planOperations({ steps: [{ intent: 'insert_paragraphs', selector: { afterParagraph: 'x' }, paragraphs: [{ text: 'a' }, { text: 'b', style: 'Normal', format: { bold: true } }] }] });
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planned.operations[0].preservationMode, row.preservationMode);
  assert.equal(planned.operations[0].paragraphs.length, 2);
  const tooMany = planOperations({ steps: [{ intent: 'insert_paragraphs', selector: { afterParagraph: 'x' }, paragraphs: Array.from({ length: 201 }, (_, i) => ({ text: `p${i}` })) }] });
  assert.equal(tooMany.ok, false, '201 paragraphs is over the measured 200 ceiling -- must be refused at plan time');
  const badFormat = planOperations({ steps: [{ intent: 'insert_paragraphs', selector: { afterParagraph: 'x' }, paragraphs: [{ text: 'a', format: { highlight: 'yellow' } }] }] });
  assert.equal(badFormat.ok, false, 'highlight is not in the accepted format subset -- must be refused at plan time');
});
