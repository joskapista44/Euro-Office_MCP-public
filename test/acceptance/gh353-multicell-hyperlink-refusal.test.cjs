'use strict';

/**
 * GH #353 (H02): a multi-cell range on set_hyperlink / remove_hyperlink is a NAMED, pre-dispatch refusal
 * (`xlsx-hyperlink-range-unsupported`), not the schema's anonymous CELL-pattern error. Offline: the reason it is
 * refused is a LIVE-measured engine fact (docs/DS-9.3.5-BRIDGE18-L2-HYPERLINK-WRITE-PLAN.md section 4:
 * ws.SetHyperlink(<range>, ...) is a SILENT NO-OP -- no throw, no link, nothing in read_hyperlinks or the stored file).
 * remove_hyperlink on a range (ApiRange.ClearHyperlinks) was never measured, so it is refused the same way in v1.
 * Per-cell expansion is new behaviour and is NOT part of this refusal.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/hyperlinks-write.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const CODE = 'xlsx-hyperlink-range-unsupported';
const setStep = (extra = {}) => ({ intent: 'set_hyperlink', id: 's0', sheet: 'Sheet1', cell: 'B2', address: 'https://example.com/t', ...extra });
const removeStep = (extra = {}) => ({ intent: 'remove_hyperlink', id: 's0', sheet: 'Sheet1', cell: 'B2', ...extra });
const RANGES = ['A1:B2', 'B2:B9', 'A1:A1'];

test('GH #353: the PRODUCT surface (validateWriteRequest) refuses a range for both intents with the NAMED code, not the anonymous schema-pattern error', () => {
  for (const cell of RANGES) {
    for (const step of [setStep({ cell }), removeStep({ cell })]) {
      const v = W.validateWriteRequest(req([step]));
      assert.equal(v.ok, false, `${step.intent} ${cell}`);
      const text = `${v.reason || ''} ${v.detail || ''}`;
      assert.ok(text.includes(CODE), `${step.intent} ${cell}: expected ${CODE}, got: ${text}`);
    }
  }
  assert.equal(W.validateWriteRequest(req([setStep()])).ok, true, 'a single cell still validates');
  assert.equal(W.validateWriteRequest(req([removeStep()])).ok, true);
});

test('GH #353: set_hyperlink / remove_hyperlink on a multi-cell range -> NAMED family refusal (code + the measured reason), nothing planned', () => {
  for (const cell of RANGES) {
    for (const step of [setStep({ cell }), removeStep({ cell })]) {
      const p = fam.plan(step, 0);
      assert.equal(p.ok, false, `${step.intent} ${cell}`);
      assert.ok(p.detail.includes(CODE), `${step.intent} ${cell}: the refusal must carry ${CODE}, got: ${p.detail}`);
      assert.ok(/silent no-op/i.test(p.detail), `${step.intent} ${cell}: the refusal states the measured reason, got: ${p.detail}`);
      const t = planOperations({ id: 'gh353', steps: [step] });
      assert.equal(t.ok, false);
      assert.ok(t.detail.includes(CODE), t.detail);
    }
  }
});

test('GH #353: a malformed cell is NOT mislabelled as a range refusal (the code names the one measured reason only)', () => {
  for (const cell of ['', '1A', 'A', 'A0']) {
    const p = fam.plan(setStep({ cell }), 0);
    assert.equal(p.ok, false, `"${cell}"`);
    assert.ok(!p.detail.includes(CODE), `"${cell}": ${p.detail}`);
  }
});

test('GH #353: the single-cell paths are unchanged (set + remove still plan)', () => {
  assert.equal(fam.plan(setStep(), 0).ok, true);
  assert.equal(fam.plan(setStep({ address: undefined, location: 'Sheet2!A1' }), 0).ok, true);
  assert.equal(fam.plan(removeStep(), 0).ok, true);
});
