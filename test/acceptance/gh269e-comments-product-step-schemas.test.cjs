'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const D = require('../../src/server/docx-write-tools.cjs');
const { check } = require('../../src/server/write-tools.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');

/**
 * GH #269e: resolve_comment / reply_comment / add_comment_on_match are routed (the adapter, GH #269c/#269d), coded in DOCX_STEP_SCHEMAS_ALL / DOCX_CONSTRAINTS_ALL and (g269ev1 PASS) RELEASED:
 * NOT in DOCX_WRITE_ALLOWLIST (r3-release-trim.test.cjs covers that half, and the release itself is a later, LIVE-PASS-gated commit). Without these entries, adding an intent to the
 * allowlist would throw "no entry for the allowlisted intent" (releasedOnly). This file covers the schema, the constraint, and the agreement with the adapter's own contract.
 * Direct against the _ALL tables -- the only way to reach them before release.
 */

const resolveStep = (over = {}) => ({ intent: 'resolve_comment', text: 'Please check', author: 'Alice', quote: 'Body first.', ...over });
const replyStep = (over = {}) => ({ intent: 'reply_comment', target: { text: 'Please check', author: 'Alice', quote: 'Body first.' }, text: 'Done', ...over });
const addStep = (over = {}) => ({ intent: 'add_comment_on_match', matchText: 'Body first.', text: 'New note', ...over });
const schemaProblem = (intent, step) => check(D.DOCX_STEP_SCHEMAS_ALL[intent], step, 'step');
const constraintProblem = (intent, step) => D.DOCX_CONSTRAINTS_ALL[intent](step);

test('SCHEMAS: all three exist in the _ALL tables AND are published (released-only tables), identical objects', () => {
  for (const i of ['resolve_comment', 'reply_comment', 'add_comment_on_match']) {
    assert.ok(D.DOCX_STEP_SCHEMAS_ALL[i] && D.DOCX_CONSTRAINTS_ALL[i], i);
    assert.equal(D.DOCX_STEP_SCHEMAS[i], D.DOCX_STEP_SCHEMAS_ALL[i], `${i} published`);
    assert.equal(D.DOCX_CONSTRAINTS[i], D.DOCX_CONSTRAINTS_ALL[i], `${i} published constraint`);
  }
});

test('resolve_comment: a well-formed step passes (resolved optional, quote may be empty); the identity triple is REQUIRED; unknown fields (incl. solved) are refused', () => {
  assert.equal(schemaProblem('resolve_comment', resolveStep()), null);
  assert.equal(schemaProblem('resolve_comment', resolveStep({ resolved: false })), null);
  assert.equal(schemaProblem('resolve_comment', resolveStep({ quote: '' })), null);
  for (const k of ['text', 'author', 'quote']) { const s = resolveStep(); delete s[k]; assert.ok(schemaProblem('resolve_comment', s), `${k} is required`); }
  assert.ok(schemaProblem('resolve_comment', resolveStep({ solved: true })), 'the adapter flag is `resolved`, not `solved`');
  assert.ok(schemaProblem('resolve_comment', resolveStep({ resolved: 'yes' })));
  assert.ok(schemaProblem('resolve_comment', resolveStep({ extra: 1 })));
  assert.ok(schemaProblem('resolve_comment', resolveStep({ text: '' })));
});

test('reply_comment: a well-formed step passes; NO author field on the reply (UNKNOWN_FIELD), a target needs all three keys and allows no others', () => {
  assert.equal(schemaProblem('reply_comment', replyStep()), null);
  assert.ok(schemaProblem('reply_comment', replyStep({ author: 'Mallory' })), 'the reply author is the session identity - no field');
  assert.ok(schemaProblem('reply_comment', replyStep({ userId: 'u' })));
  for (const k of ['text', 'author', 'quote']) { const t = { ...replyStep().target }; delete t[k]; assert.ok(schemaProblem('reply_comment', replyStep({ target: t })), `target.${k} is required`); }
  assert.ok(schemaProblem('reply_comment', replyStep({ target: { ...replyStep().target, extra: 1 } })));
  assert.ok(schemaProblem('reply_comment', replyStep({ target: 'Please check' })));
  const noText = replyStep(); delete noText.text;
  assert.ok(schemaProblem('reply_comment', noText));
  assert.ok(schemaProblem('reply_comment', replyStep({ text: '' })));
});

test('add_comment_on_match: a well-formed step passes (occurrence optional, 1-based integer); NO author field; unknown fields refused', () => {
  assert.equal(schemaProblem('add_comment_on_match', addStep()), null);
  assert.equal(schemaProblem('add_comment_on_match', addStep({ occurrence: 2 })), null);
  assert.ok(schemaProblem('add_comment_on_match', addStep({ author: 'Mallory' })), 'no author field');
  assert.ok(schemaProblem('add_comment_on_match', addStep({ occurrence: 0 })));
  assert.ok(schemaProblem('add_comment_on_match', addStep({ occurrence: 1.5 })));
  assert.ok(schemaProblem('add_comment_on_match', addStep({ selector: { text: 'x' } })));
  for (const k of ['matchText', 'text']) { const s = addStep(); delete s[k]; assert.ok(schemaProblem('add_comment_on_match', s), `${k} is required`); }
});

test('CONSTRAINTS: a blank text / author / matchText is refused with its own reason; a good step passes', () => {
  assert.equal(constraintProblem('resolve_comment', resolveStep()), null);
  assert.match(constraintProblem('resolve_comment', resolveStep({ text: '  ' })), /text must not be blank/);
  assert.match(constraintProblem('resolve_comment', resolveStep({ author: ' ' })), /author must not be blank/);
  assert.equal(constraintProblem('reply_comment', replyStep()), null);
  assert.match(constraintProblem('reply_comment', replyStep({ text: ' ' })), /text must not be blank/);
  assert.match(constraintProblem('reply_comment', replyStep({ target: { ...replyStep().target, text: ' ' } })), /target\.text/);
  assert.match(constraintProblem('reply_comment', replyStep({ target: { ...replyStep().target, author: ' ' } })), /target\.author/);
  assert.equal(constraintProblem('add_comment_on_match', addStep()), null);
  assert.match(constraintProblem('add_comment_on_match', addStep({ matchText: ' ' })), /matchText must not be blank/);
  assert.match(constraintProblem('add_comment_on_match', addStep({ text: ' ' })), /text must not be blank/);
});

test('AGREEMENT with the adapter: every well-formed product step is accepted by the adapter planner; an adapter-side refusal is not hidden by the schema', () => {
  for (const step of [resolveStep(), resolveStep({ resolved: false }), replyStep(), addStep(), addStep({ occurrence: 2 })]) {
    assert.equal(planOperations({ id: 't', steps: [step] }).ok, true, JSON.stringify(step));
  }
  assert.equal(typeof docxAdapter.kind, 'string');
});

test('RELEASED (g269ev1 PASS), positive control of the former withheld check: the handler no longer refuses a WELL-FORMED step for the allowlist -- it reaches the RUNTIME stage exactly like an already-released intent (add_footnote)', async () => {
  const tool = D.createDocxWriteTool({ createDriver: () => { throw new Error('must not be needed: the runtime stage stops first in this environment'); } });
  const ref = await tool.handler({ fileId: '1', task: { steps: [{ intent: 'add_footnote', selector: { text: 'p' }, text: 'n' }] } });
  assert.equal(ref.stage, 'RUNTIME', 'the released reference intent reaches the runtime stage here');
  for (const step of [resolveStep(), replyStep(), addStep()]) {
    assert.notEqual((D.validateDocxWriteRequest({ fileId: '1', task: { steps: [step] } }) || {}).code, 'REFUSED_NOT_ALLOWLISTED', step.intent);
    const r = await tool.handler({ fileId: '1', task: { steps: [step] } });
    assert.notEqual(r.code, 'REFUSED_NOT_ALLOWLISTED', step.intent);
    assert.equal(r.stage, ref.stage, `${step.intent}: past validation, same stage as a released intent`);
    assert.equal(r.code, ref.code, step.intent);
  }
});

test('the schema still refuses at the handler, BEFORE any driver: an author on reply_comment / add_comment_on_match, solved on resolve_comment', async () => {
  let reached = 0;
  const tool = D.createDocxWriteTool({ createDriver: () => { reached += 1; throw new Error('must not be reached'); } });
  for (const step of [replyStep({ author: 'Mallory' }), addStep({ author: 'Mallory' }), resolveStep({ solved: true })]) {
    const r = await tool.handler({ fileId: '1', task: { steps: [step] } });
    assert.equal(r.dispatched, false, step.intent);
    assert.notEqual(r.code, 'REFUSED_NOT_ALLOWLISTED', step.intent);
  }
  assert.equal(reached, 0);
});
