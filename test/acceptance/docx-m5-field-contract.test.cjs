'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { STEP_KEYS } = require('../../src/adapters/docx/field-contract.cjs');
const { INTENTS } = require('../../src/adapters/docx/intents.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M5 fail-closed field contract (owner decision after owner channel, point 1). A field the
 * target does not read and enforce refuses the WHOLE task at planning -- 0 dispatch, 0 write --
 * instead of being silently dropped. Before the guard, every task below was APPLIED with 1 write
 * (measured offline, M5 card #ref: a mismatching precondition and a constraints.section on a
 * nonexistent heading both executed).
 */

const RICH = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const P1 = 'First body paragraph.';
const doc = () => [{ text: 'Intro heading', level: 1, props: { ...RICH, bold: true } }, { text: P1, props: { ...RICH } }, { text: 'Second body paragraph.', props: { ...RICH } }];

async function run(task) {
  let t = 0;
  const store = createStore(doc());
  const before = JSON.stringify(store.durable);
  const driver = createFakeDocxDriver(store, {});
  const result = await withSession({ driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, docxAdapter, createExecutor(task));
  return { result, driver, unchanged: JSON.stringify(store.durable) === before };
}

const FMT = { intent: 'format_text', selector: { text: P1 }, format: { underline: true } };

async function refusedBeforeDispatch(task, detailPattern) {
  const planned = planOperations(task);
  assert.equal(planned.ok, false, 'the planner refuses the task');
  assert.match(planned.detail, detailPattern);
  const { result, driver, unchanged } = await run(task);
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0, '0 dispatch');
  assert.equal(result.metrics.writes, 0, '0 write');
  assert.equal(unchanged, true, 'the document is unchanged');
}

test('control: the same format_text task WITHOUT the extra fields is applied (the guard is not a blanket refusal)', async () => {
  const { result, driver } = await run({ id: 't', steps: [FMT] });
  assert.equal(result.editOk, true);
  assert.equal(driver.calls.dispatch, 1);
});

// M5 migration (owner channel): precondition, constraints and sectionHeading + paragraphOrdinal
// are ENFORCED now (docx-m5-constraints.test.cjs has the full positive/negative set). What stays here
// is the original point of this file: the same task shapes that were silently APPLIED before M5 are
// now refused before any dispatch -- by the enforcement, no longer by the transitional guard.

async function refusedAtExecution(task, ambiguity) {
  assert.equal(planOperations(task).ok, true, 'the field is accepted at planning (migrated)');
  const { result, driver, unchanged } = await run(task);
  assert.equal(result.ok, false);
  assert.equal(result.execution.appliedOperations[0].evidence.ambiguity, ambiguity);
  assert.equal(driver.calls.dispatch, 0, '0 dispatch');
  assert.equal(result.metrics.writes, 0, '0 write');
  assert.equal(unchanged, true, 'the document is unchanged');
}

test('C1 selector.precondition mismatch -> enforced: refused before dispatch (was: APPLIED, 1 write)', async () => {
  await refusedAtExecution({ id: 't', steps: [{ ...FMT, selector: { text: P1, precondition: { text: 'SOMETHING ELSE' } } }] }, 'PRECONDITION_TEXT_MISMATCH');
});

test('C2 task.constraints.section on a nonexistent heading -> enforced: refused before dispatch (was: APPLIED, 1 write)', async () => {
  await refusedAtExecution({ id: 't', constraints: { section: 'Nonexistent heading' }, steps: [FMT] }, 'CONSTRAINT_SECTION_NOT_FOUND');
});

test('the source selector sectionHeading + paragraphOrdinal -> resolved (migrated); on intents not migrated still refused, named', async () => {
  assert.equal(planOperations({ id: 't', steps: [{ ...FMT, selector: { sectionHeading: 'Intro heading', paragraphOrdinal: 1 } }] }).ok, true);
  // M5 gate 3: migrated on insert_paragraph_before/_after too; delete stays refused, named
  assert.equal(planOperations({ id: 't', steps: [{ intent: 'insert_paragraph_after', selector: { sectionHeading: 'Intro heading', paragraphOrdinal: 1 }, text: 'x' }] }).ok, true);
  await refusedBeforeDispatch({ id: 't', steps: [{ intent: 'delete_paragraph', selector: { sectionHeading: 'Intro heading', paragraphOrdinal: 1 } }] }, /field-not-enforced-yet: .*selector\.sectionHeading/);
});

test('the source reference syntax (targetRef / destinationRef / step.ref) -> refused, named', async () => {
  await refusedBeforeDispatch({ id: 't', steps: [{ intent: 'format_text', targetRef: 'a', format: { bold: true } }] }, /field-not-enforced-yet: .*step\.targetRef/);
  await refusedBeforeDispatch({ id: 't', steps: [{ ...FMT, ref: 'a' }] }, /field-not-enforced-yet: .*step\.ref/);
  await refusedBeforeDispatch({ id: 't', steps: [{ intent: 'move_paragraph_after', selector: { text: P1 }, destinationRef: 'a' }] }, /field-not-enforced-yet: .*step\.destinationRef/);
});

test('an unknown step field or task field -> refused (never dropped)', async () => {
  await refusedBeforeDispatch({ id: 't', steps: [{ ...FMT, zzqUnknown: 1 }] }, /field-unsupported: step 0 \(format_text\)\.zzqUnknown/);
  await refusedBeforeDispatch({ id: 't', zzqUnknown: 1, steps: [FMT] }, /field-unsupported: task\.zzqUnknown/);
});

test('a field valid for ANOTHER intent is refused where the plan would drop it', async () => {
  // set_list_type does not read level; rewrite_paragraph does not carry dependsOn
  assert.match(planOperations({ steps: [{ intent: 'set_list_type', selector: { text: P1 }, type: 'bullet', level: 2 }] }).detail, /field-unsupported: step 0 \(set_list_type\)\.level/);
  assert.match(planOperations({ steps: [{ intent: 'rewrite_paragraph', selector: { text: P1 }, text: 'x', dependsOn: [] }] }).detail, /field-unsupported: step 0 \(rewrite_paragraph\)\.dependsOn/);
  // the source split terminal assertions (the target plan does not read them)
  assert.match(planOperations({ steps: [{ intent: 'split_table_cell', selector: { text: 'a' }, splitRows: 1, splitCols: 2, terminalTableAnchor: 'b', terminalRowCount: 3 }] }).detail, /split_table_cell\)\.terminalRowCount/);
  assert.match(planOperations({ steps: [{ intent: 'split_table_cell', selector: { text: 'a' }, splitRows: 1, splitCols: 2, terminalTableAnchor: 'b', terminalTargetRowCellCount: 3 }] }).detail, /step\.terminalTargetRowCellCount/);
});

test('a selector key outside the resolver branch it picks would be ignored -> refused', async () => {
  const detail = sel => planOperations({ steps: [{ ...FMT, selector: sel }] }).detail;
  assert.match(detail({ ref: 's0', text: P1 }), /text would be ignored beside the ref selector/);
  assert.match(detail({ fingerprint: 'f', text: P1 }), /text would be ignored beside the fingerprint selector/);
  assert.match(detail({ heading: 'Intro heading', text: P1 }), /text would be ignored beside the heading selector/);
  assert.match(detail({ text: P1, level: 1 }), /level would be ignored beside the text selector/);
  // the keys the resolver REFUSES itself stay on its own named path (not a silent drop)
  assert.equal(planOperations({ steps: [{ ...FMT, selector: { text: P1, index: 3 } }] }).ok, true);
});

test('every registry intent has an accepted field set, and every set carries intent + id', () => {
  const registry = INTENTS.map(r => r.intent).sort();
  assert.deepEqual(Object.keys(STEP_KEYS).sort(), registry);
  for (const [intent, keys] of Object.entries(STEP_KEYS)) assert.ok(keys.includes('intent') && keys.includes('id'), intent);
});
