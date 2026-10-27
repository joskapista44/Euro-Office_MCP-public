'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { richOutcome } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { readOnce, FORMATS } = require('../../src/server/read-tools.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');

/**
 * readLive no longer hands the product a SILENT null: when the page-side property query fails ({ error } from its outer catch, or a failed call) or returns a part-less result, the payload
 * carries `richUnobservable` (the NAMED reason), the inventory keeps it, and the read says WHY it is incomplete. The payload stays ok:true so the HTML-only reads keep working.
 * Before: paragraphs and document null, complete:false and NO reason (g266v1 ROUND4B / gbolv1: every rich field and the bookmark list null).
 */

const OK = { ok: true, result: { paragraphs: [{ text: 'a' }], document: { bookmarks: [] } } };

test('richOutcome: a query that read both parts has NO richUnobservable (the payload is unchanged for a healthy read)', () => {
  const o = richOutcome(OK);
  assert.equal(o.richUnobservable, undefined); assert.equal(o.paragraphs.length, 1); assert.deepEqual(o.document, { bookmarks: [] });
});

test('richOutcome: the page-side { error } is NAMED with the page error; a failed call is NAMED with its error / outcome; a part-less result names the missing part', () => {
  assert.deepEqual(richOutcome({ ok: true, result: { error: 'TypeError: x is not a function' } }), { paragraphs: null, document: null, richUnobservable: 'UNOBSERVABLE:property-query: TypeError: x is not a function' });
  assert.equal(richOutcome({ ok: false, outcome: 'call-threw', error: 'boom' }).richUnobservable, 'UNOBSERVABLE:property-query-call: boom');
  assert.equal(richOutcome({ ok: false, outcome: 'command-callback-timeout' }).richUnobservable, 'UNOBSERVABLE:property-query-call: command-callback-timeout');
  assert.equal(richOutcome(undefined).richUnobservable, 'UNOBSERVABLE:property-query-call: no result');
  assert.equal(richOutcome({ ok: true, result: { document: {} } }).richUnobservable, 'UNOBSERVABLE:property-query: no paragraphs array');
  assert.equal(richOutcome({ ok: true, result: { paragraphs: [] } }).richUnobservable, 'UNOBSERVABLE:property-query: no document object');
  assert.equal(richOutcome({ ok: true, result: { error: 'x'.repeat(500) } }).richUnobservable.length, 'UNOBSERVABLE:property-query: '.length + 160, 'the message is clipped');
});

const payload = extra => ({ html: '<p>a</p>', paragraphs: [{ text: 'a', props: null, hyperlinks: [], noteRefs: { footnote: 0, endnote: 0 } }], document: { bookmarks: [] }, ...extra });
const driverOf = p => ({ readLive: async () => ({ ok: true, authority: 'LIVE_READ', payload: p }) });
const read = (p, operation = 'read_formatting') => readOnce(FORMATS.docx, driverOf(p), { fileId: 'f', operation }, 'docx');

test('inventory: richUnobservable is kept when the payload carries it and ABSENT otherwise (no new key on a healthy read)', () => {
  assert.equal(normalizeInventory(payload({ richUnobservable: 'UNOBSERVABLE:property-query: boom' })).richUnobservable, 'UNOBSERVABLE:property-query: boom');
  assert.equal('richUnobservable' in normalizeInventory(payload()), false);
});

test('read: a healthy payload is complete with NO incomplete reason (unchanged)', async () => {
  const r = await read(payload());
  assert.equal(r.ok, true); assert.equal(r.complete, true); assert.equal('incomplete' in r, false);
});

test('read: the BLIND payload (paragraphs / document null + the named reason) is incomplete AND says why -- for a rich read AND for read_text; before the change the reason was absent', async () => {
  const blind = { html: '<p>a</p>', paragraphs: null, document: null, richUnobservable: 'UNOBSERVABLE:property-query: TypeError: x is not a function' };
  for (const op of ['read_formatting', 'read_bookmarks', 'read_text']) {
    const r = await read(blind, op);
    assert.equal(r.ok, true, op); assert.equal(r.complete, false, op);
    assert.deepEqual(r.incomplete, ['rich: UNOBSERVABLE:property-query: TypeError: x is not a function'], op);
  }
  const legacy = await read({ html: '<p>a</p>', paragraphs: null, document: null });
  assert.equal(legacy.complete, false); assert.equal('incomplete' in legacy, false, 'a payload WITHOUT the named reason still has no reason (the pre-change state this test pins as the thing that was wrong)');
});

test('read: paragraphs read but the document part missing is NOT complete any more (it used to be "complete" with every document-level field silently null)', async () => {
  const r = await read({ html: '<p>a</p>', paragraphs: [{ text: 'a' }], document: null, richUnobservable: 'UNOBSERVABLE:property-query: no document object' }, 'read_bookmarks');
  assert.equal(r.complete, false); assert.deepEqual(r.incomplete, ['rich: UNOBSERVABLE:property-query: no document object']);
  assert.equal(r.result.bookmarks, null, 'the field is still null (absent) -- the difference is that the read now says why');
});
