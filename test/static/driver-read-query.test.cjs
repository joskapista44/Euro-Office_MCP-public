'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { PROPERTY_QUERY } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * The driver's LIVE read runs a constant editor query. A read must never mutate, so the query may
 * call only getters and the Search it needs to find the exact range.
 */
const MUTATING = /\.(Add\w*|Delete|Remove\w*|Insert\w*|Set\w*|Merge\w*|Split\w*|Push|Paste\w*|Replace\w*|Move\w*|Apply\w*)\s*\(/;

test('the property query calls no mutating editor method', () => {
  assert.doesNotMatch(PROPERTY_QUERY, MUTATING);
  const calls = [...PROPERTY_QUERY.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(m => m[1]);
  assert.ok(calls.length >= 5, 'positive control: the extractor sees the query calls');
  // W-08: IsAutoColor() is a pure boolean query (ApiColor), the same read-only class
  // as Get*/Search -- it never matches MUTATING above, so allowing exactly this one name (not every Is*) does not weaken the gate.
  // W-09 (shading) reads through the same single exception.
  // W-09-3: ToJSON() is a pure serialization query (ParaPr) -- the only reader for
  // tabs/per-side border/pagination flags, called with no arguments, never matches MUTATING above.
  // Same narrow, named exception as IsAutoColor; not every *Json/*JSON-shaped name, just this one.
  for (const c of calls) assert.match(c, /^(Get\w*|IsAutoColor|ToJSON|Search)$/, `unexpected editor call in a read: ${c}`);
});

test('NEGATIVE CONTROL: the pattern catches a mutation smuggled into a read', () => {
  assert.match("const r = doc.Search('x')[0]; r.Delete();", MUTATING);
  assert.match('p.SetBold(true);', MUTATING);
});
