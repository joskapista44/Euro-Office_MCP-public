'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const D = require('../../src/server/docx-write-tools.cjs');
const { check } = require('../../src/server/write-tools.cjs');

/**
 * GH #268d: set_image_size / set_image_wrap are routed and coded (DOCX_STEP_SCHEMAS_ALL /
 * DOCX_CONSTRAINTS_ALL). This file covers the schema itself, tested directly against
 * DOCX_STEP_SCHEMAS_ALL (independent of whether the intents are allowlisted). The adapter-level
 * behaviour (plan/classify/resolve/write/verify) is #268c's own coverage (test/acceptance/
 * gh268c-docx-image-write.test.cjs, LIVE-PASSed through withSession directly, g268cv1).
 *
 * FLIP (this unit, after the withheld-gate LIVE PASS g268dv1 and the handler-acceptance LIVE
 * MEASURED g268dv2): both intents moved from WITHHELD into DOCX_WRITE_ALLOWLIST (r3-release-trim.
 * test.cjs). The last two tests below now prove the OPPOSITE of the pre-FLIP state: the allowlist
 * gate no longer refuses either intent, and the handler's pre-dispatch validation reaches past it
 * toward the driver (the driver-throws probe below proves dispatch is ATTEMPTED, not that it
 * succeeds -- the handler-level APPLIED/FILE_CHECKS/retry proof is test/live's own
 * docx-image-write-handler-acceptance.live.cjs).
 */

const sizeStep = (over = {}) => ({ intent: 'set_image_size', selector: { index: 0 }, widthEmu: 100, heightEmu: 200, ...over });
const wrapStep = (over = {}) => ({ intent: 'set_image_wrap', selector: { index: 0 }, wrap: 'square', ...over });
const problem = (intent, step) => check(D.DOCX_STEP_SCHEMAS_ALL[intent], step, 'step');

test('set_image_size schema: a well-formed step (by index, and by {paragraph, ordinal}) passes', () => {
  assert.equal(problem('set_image_size', sizeStep()), null);
  assert.equal(problem('set_image_size', sizeStep({ selector: { paragraph: 'p' } })), null);
  assert.equal(problem('set_image_size', sizeStep({ selector: { paragraph: 'p', ordinal: 2 } })), null);
});

test('set_image_size schema: RED -- unknown field, bad types, out-of-range EMU', () => {
  assert.ok(problem('set_image_size', { ...sizeStep(), extra: 1 }));
  assert.ok(problem('set_image_size', sizeStep({ widthEmu: 'x' })));
  assert.ok(problem('set_image_size', sizeStep({ widthEmu: 0 })), 'widthEmu must be >= 1');
  assert.ok(problem('set_image_size', sizeStep({ heightEmu: -5 })));
  assert.ok(problem('set_image_size', { ...sizeStep(), selector: { index: 0, bogus: 1 } }), 'additionalProperties:false on the selector itself rejects an unknown key');
});

test('set_image_size schema: index + paragraph together, or ordinal with no paragraph/index, PASS the bare schema (additionalProperties:false has no XOR) -- the family\'s own plan() is the actual authority for that shape, exercised in gh268c-docx-image-write.test.cjs', () => {
  assert.equal(problem('set_image_size', sizeStep({ selector: { index: 0, paragraph: 'p' } })), null);
  assert.equal(problem('set_image_size', sizeStep({ selector: { ordinal: 2 } })), null);
});

test('set_image_wrap schema: well-formed passes, and ONLY \'square\' is in the enum', () => {
  assert.equal(problem('set_image_wrap', wrapStep()), null);
  for (const wrap of ['tight', 'through', 'topAndBottom', 'behind', 'inFront', 'inline', 'none']) {
    assert.ok(problem('set_image_wrap', wrapStep({ wrap })), `wrap:'${wrap}' must fail the schema -- NEM-MÉRT, not assumed`);
  }
});

test('set_image_wrap schema: RED -- unknown field, missing required wrap', () => {
  assert.ok(problem('set_image_wrap', { ...wrapStep(), extra: 1 }));
  const noWrap = { intent: 'set_image_wrap', selector: { index: 0 } };
  assert.ok(problem('set_image_wrap', noWrap));
});

test('DOCX_CONSTRAINTS_ALL defers both intents to the family\'s own plan() (no product-schema-level constraint of its own)', () => {
  assert.equal(D.DOCX_CONSTRAINTS_ALL.set_image_size(sizeStep()), null);
  assert.equal(D.DOCX_CONSTRAINTS_ALL.set_image_wrap(wrapStep()), null);
});

test('FLIP POS. CONTROL: the pre-dispatch allowlist check no longer refuses either intent -- validateDocxWriteRequest with a WELL-FORMED step is NOT REFUSED_NOT_ALLOWLISTED', () => {
  for (const step of [sizeStep(), wrapStep()]) {
    const v = D.validateDocxWriteRequest({ fileId: '1', task: { steps: [step] } });
    assert.notEqual(v.code, 'REFUSED_NOT_ALLOWLISTED', step.intent);
  }
});

test('FLIP POS. CONTROL: createDocxWriteTool().handler no longer stops at the allowlist gate -- it reaches past validateDocxWriteRequest toward the driver (here: BROWSER_RUNTIME_UNAVAILABLE from the unconfigured fixture, never the allowlist refusal)', async () => {
  const tool = D.createDocxWriteTool({ createDriver: () => { throw new Error('must not be called before BROWSER_RUNTIME_UNAVAILABLE'); } });
  for (const step of [sizeStep(), wrapStep()]) {
    const r = await tool.handler({ fileId: '1', task: { steps: [step] } });
    assert.notEqual(r.code, 'REFUSED_NOT_ALLOWLISTED', step.intent);
  }
});
