'use strict';

/**
 * Static gate (Loki / llama.cpp, MEASURED 2026-11-01 on the grammar engine, one request per tool): xlsx_execute_task
 * was the only tool refused ("Failed to initialize samplers: failed to parse grammar"), and the failing nodes were the
 * step alternatives add_comment / edit_comment / reply_to_comment / set_hyperlink. Field-level bisection: a string
 * `maxLength: 2000` inside those object alternatives -> 400, 1999 -> 200, no maxLength -> 200. The exact llama.cpp
 * mechanism is NOT measured, so this gate pins the measured fields only (it does not claim a general threshold).
 * The 2000 ceiling itself stays: each family's plan() refuses a longer value, named (LIVE-measured at exactly 2000).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const comments = require('../../src/adapters/xlsx/families/comments.cjs');
const links = require('../../src/adapters/xlsx/families/hyperlinks-write.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const config = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker' } };
const MEASURED = [['add_comment', 'text'], ['edit_comment', 'text'], ['reply_to_comment', 'text'],
  ['set_hyperlink', 'address'], ['set_hyperlink', 'location'], ['set_hyperlink', 'tooltip']];

test('the measured grammar-breaking fields carry no schema maxLength in the SERVED xlsx_execute_task schema', () => {
  const { tools } = buildServerTools({ config, configIdentity: 'fixture' });
  const steps = JSON.stringify(tools.xlsx_execute_task.inputSchema);
  assert.ok(steps.includes('"add_comment"') && steps.includes('"set_hyperlink"'), 'positive control: the walk reaches the step alternatives');
  for (const [intent, field] of MEASURED) {
    const prop = W.STEP_SCHEMAS[intent].properties[field];
    assert.ok(prop, `${intent}.${field} exists`);
    assert.equal(prop.maxLength, undefined, `${intent}.${field} has no schema maxLength`);
    assert.equal(prop.type, 'string');
  }
});

test('the 2000 ceiling is still enforced by the families, at exactly 2000', () => {
  assert.equal(comments.MAX_TEXT, 2000);
  assert.deepEqual([links.MAX_ADDRESS, links.MAX_LOCATION, links.MAX_TOOLTIP], [2000, 2000, 2000]);
  const step = text => ({ intent: 'add_comment', sheet: 'S', cell: 'A1', text });
  assert.equal(comments.plan(step('x'.repeat(2000)), 0).ok, true, '2000 characters pass (the LIVE-proven length)');
  const over = comments.plan(step('x'.repeat(2001)), 0);
  assert.equal(over.ok, false);
  assert.match(over.detail, /over the 2000 ceiling/);
  for (const intent of ['edit_comment', 'reply_to_comment']) {
    assert.equal(comments.plan({ ...step('x'.repeat(2001)), intent }, 0).ok, false, `${intent} over the ceiling`);
  }
});

test('the server refuses an over-length comment pre-dispatch, named (no schema maxLength needed)', () => {
  const v = n => W.validateWriteRequest(req([{ intent: 'add_comment', sheet: 'S', cell: 'A1', text: 'x'.repeat(n) }]));
  assert.equal(v(2000).ok, true, JSON.stringify(v(2000)).slice(0, 300));
  const r = v(2001);
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r), /2000 ceiling/);
});
