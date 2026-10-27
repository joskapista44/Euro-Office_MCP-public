'use strict';

/**
 * GH #138 U4: the remaining 8 print.cjs intents on the PRODUCT surface (the WRITE tool with the fake XLSX
 * editor) -- fit, header/footer text, manual page breaks, first page number, print centering, scale, paper
 * size, header/footer flags. The family-level behaviour is in test/acceptance/xlsx-m7-p05-scale-paper*,
 * xlsx-m7-f-gates.test.cjs and the family's own LIVE probes (print.cjs's header comment cites them); this
 * file is the product-tool level, mirroring gh138-u3-page-layout-release.test.cjs. The READ mirror
 * (read_page_setup's `print` field) was already released; the LIVE proof is
 * test/live/release-gh138-u4-print-family-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const INTENTS = ['set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering', 'set_print_scale', 'set_paper_size', 'set_header_footer_flag'];
const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const statuses = r => (r.operations || []).map(o => o.status);
// B10-WRITE is narrowed to the EXPLICIT list XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES (runtime-identity.cjs) --
// unlike most families, the fake driver's DEFAULT runtime is NOT enough for a print WRITE to dispatch.
// DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): the historical 9.3.4.60 /7..16 chain (widened
// progressively, a maintainer's review #ref, test/static/b10-write-handler-identity.test.cjs) is retired -- the
// ACCEPTED runtime is now the single 9.3.5.0/16 identity, used below for both the 8 APPLIED tests and the
// "PROD runtime" test (the two no longer differ -- there is only one accepted entry).
const RT935 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const printHarness = (store) => harness({ store, script: { runtime: RT935 } });

test('GH #138 U4: all 8 print intents are released (allowlist + accepted set + schema + constraints), normal steps, the release gate passes, and B10 no longer depends on them', () => {
  for (const intent of INTENTS) {
    assert.ok(W.XLSX_WRITE_ALLOWLIST.includes(intent), intent);
    assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes(intent), intent);
    assert.ok(!W.SOLE_STEP.includes(intent), intent);
    assert.equal(Object.prototype.hasOwnProperty.call(W.TEMPORARY_BOUNDARY, intent), false, intent);
  }
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
});

test('GH #138 U4: set_print_fit APPLIED on the accepted 9.3.5.0/16 runtime -- zero\'s regression: an unwidened XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES would FAIL_CLOSED here, named xlsx-print-bridge-write-runtime-not-proven', async () => {
  const h = harness({ store: createStore(book()), script: { runtime: RT935 } });
  const r = await h.tool.handler(req([{ intent: 'set_print_fit', id: 's', sheet: 'P', fitToWidth: 80, fitToHeight: 90 }]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.doesNotMatch(J(r), /xlsx-print-bridge-write-runtime-not-proven/);
});

test('GH #138 U4: the schema + constraints refuse a malformed step before any dispatch, per intent', () => {
  const v = s => W.validateWriteRequest(req([s]));
  const bad = [
    [{ intent: 'set_print_fit', sheet: 'P' }, 'no field at all'],
    [{ intent: 'set_print_fit', sheet: 'P', fitToWidth: 50 }, 'fitToHeight missing (both required together)'],
    [{ intent: 'set_print_fit', sheet: 'P', fitToWidth: 101, fitToHeight: 50 }, 'fitToWidth over 100'],
    [{ intent: 'set_header_footer', sheet: 'P', slot: 'centerHeader', value: 'x' }, 'an unreleased slot (centerHeader is not in HF_SLOTS)'],
    [{ intent: 'set_header_footer', sheet: 'P', slot: 'oddHeader', value: '' }, 'an empty string (not a valid clear -- must be null)'],
    [{ intent: 'set_header_footer', sheet: 'P', slot: 'oddHeader' }, 'value missing'],
    [{ intent: 'set_page_break', sheet: 'P', axis: 'diagonal', at: 5 }, 'an unknown axis'],
    [{ intent: 'set_page_break', sheet: 'P', axis: 'row', at: 5, remove: false }, 'remove must be true when given, never false'],
    [{ intent: 'set_page_break', sheet: 'P', axis: 'row' }, 'at missing'],
    [{ intent: 'set_first_page_number', sheet: 'P', value: 0 }, 'value must be >= 1'],
    [{ intent: 'set_first_page_number', sheet: 'P', value: 32768 }, 'value over the 32767 cap'],
    [{ intent: 'set_print_centering', sheet: 'P' }, 'neither horizontal nor vertical'],
    [{ intent: 'set_print_centering', sheet: 'P', horizontal: true, vertical: false }, 'both given -- exactly one required'],
    [{ intent: 'set_print_centering', sheet: 'P', horizontal: 'yes' }, 'a non-boolean axis value'],
    [{ intent: 'set_print_scale', sheet: 'P', scale: 9 }, 'scale under the 10% floor'],
    [{ intent: 'set_print_scale', sheet: 'P', scale: 401 }, 'scale over the 400% cap'],
    [{ intent: 'set_paper_size', sheet: 'P', paper: 'A4 Oversize' }, 'a name not in the editor menu'],
    [{ intent: 'set_header_footer_flag', sheet: 'P', flag: 'centerHeaderFlag', value: true }, 'an unreleased flag (only differentFirst / differentOddEven)'],
    [{ intent: 'set_header_footer_flag', sheet: 'P', flag: 'differentFirst', value: 'yes' }, 'value not a boolean'],
    [{ intent: 'set_print_fit', sheet: 'P', fitToWidth: 50, fitToHeight: 50, wide: 1 }, 'a print-family field this intent does not support'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
  // and the accepted shape of each, so the bad list above is contrasted against a real pass, not a typo
  const good = [
    { intent: 'set_print_fit', sheet: 'P', fitToWidth: 80, fitToHeight: 80 },
    { intent: 'set_header_footer', sheet: 'P', slot: 'oddFooter', value: '&P of &N' },
    { intent: 'set_header_footer', sheet: 'P', slot: 'firstHeader', value: null },
    { intent: 'set_page_break', sheet: 'P', axis: 'row', at: 11 },
    { intent: 'set_page_break', sheet: 'P', axis: 'column', at: 'E', remove: true },
    { intent: 'set_first_page_number', sheet: 'P', value: 5 },
    { intent: 'set_first_page_number', sheet: 'P', value: null },
    { intent: 'set_print_centering', sheet: 'P', vertical: true },
    { intent: 'set_print_scale', sheet: 'P', scale: 75 },
    { intent: 'set_paper_size', sheet: 'P', paper: 'US Letter' },
    { intent: 'set_header_footer_flag', sheet: 'P', flag: 'differentOddEven', value: true },
  ];
  for (const s of good) assert.equal(v(s).ok, true, J(s));
});

test('GH #138 U4: set_print_fit APPLIED (fit ON, the other sheet untouched), retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_print_fit', id: 's', sheet: 'P', fitToWidth: 80, fitToHeight: 90 }]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  const p = h.store.durable.find(x => x.name === 'P').print;
  assert.equal(p.fitW, 80); assert.equal(p.fitH, 90); assert.equal(p.fitFlag, true);
  assert.equal(h.store.durable.find(x => x.name === 'O').print, undefined, 'the other sheet carries no print state at all');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_print_fit', id: 's', sheet: 'P', fitToWidth: 80, fitToHeight: 90 }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #138 U4: set_header_footer (a non-oddHeader slot, the F2-widened read) APPLIED, clear (null) APPLIED, retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_header_footer', id: 's', sheet: 'P', slot: 'firstFooter', value: 'Page &P' }]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  assert.equal(h.store.durable.find(x => x.name === 'P').print.firstFooter, 'Page &P');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_header_footer', id: 's', sheet: 'P', slot: 'firstFooter', value: 'Page &P' }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
  const cleared = await h.tool.handler(req([{ intent: 'set_header_footer', id: 's', sheet: 'P', slot: 'firstFooter', value: null }]));
  assert.deepEqual(statuses(cleared), ['APPLIED']);
  assert.equal(h.store.durable.find(x => x.name === 'P').print.firstFooter, null);
});

test('GH #138 U4: set_page_break (row add, column add, row remove) APPLIED, retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const add = await h.tool.handler(req([{ intent: 'set_page_break', id: 's', sheet: 'P', axis: 'row', at: 11 }]));
  assert.equal(add.ok, true, J(add).slice(0, 500));
  assert.deepEqual(statuses(add), ['APPLIED']);
  assert.deepEqual(h.store.durable.find(x => x.name === 'P').print.rowBreaks, [{ id: 10, man: true }], 'the family\'s own id-is-at-minus-1 convention, W8b measured id 10 for at 11');
  const col = await h.tool.handler(req([{ intent: 'set_page_break', id: 's2', sheet: 'P', axis: 'column', at: 'E' }]));
  assert.deepEqual(statuses(col), ['APPLIED']);
  assert.deepEqual(h.store.durable.find(x => x.name === 'P').print.colBreaks, [{ id: 4, man: true }], 'F2 Q11: a column break at E is id 4');
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_page_break', id: 's', sheet: 'P', axis: 'row', at: 11 }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
  const removed = await h.tool.handler(req([{ intent: 'set_page_break', id: 's3', sheet: 'P', axis: 'row', at: 11, remove: true }]));
  assert.deepEqual(statuses(removed), ['APPLIED']);
  assert.deepEqual(h.store.durable.find(x => x.name === 'P').print.rowBreaks, []);
});

test('GH #138 U4: set_first_page_number set and clear, both APPLIED, retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_first_page_number', id: 's', sheet: 'P', value: 7 }]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 500));
  assert.equal(h.store.durable.find(x => x.name === 'P').print.firstPage, 7);
  assert.equal(h.store.durable.find(x => x.name === 'P').print.useFirstPage, true);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_first_page_number', id: 's', sheet: 'P', value: 7 }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
  const cleared = await h.tool.handler(req([{ intent: 'set_first_page_number', id: 's2', sheet: 'P', value: null }]));
  assert.deepEqual(statuses(cleared), ['APPLIED']);
  assert.equal(h.store.durable.find(x => x.name === 'P').print.firstPage, null);
});

test('GH #138 U4: set_print_centering (vertical, the F2-widened vCenter read) APPLIED, retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_print_centering', id: 's', sheet: 'P', vertical: true }]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 500));
  assert.equal(h.store.durable.find(x => x.name === 'P').print.vCenter, true);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_print_centering', id: 's', sheet: 'P', vertical: true }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('GH #138 U4: set_print_scale APPLIED (fit forced OFF, scale 100 rule does not apply here -- OD-A/OD-C), retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_print_scale', id: 's', sheet: 'P', scale: 75 }]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 500));
  const p = h.store.durable.find(x => x.name === 'P').print;
  assert.equal(p.scale, 75); assert.equal(p.fitW, 0); assert.equal(p.fitH, 0); assert.equal(p.fitFlag, false);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_print_scale', id: 's', sheet: 'P', scale: 75 }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('GH #138 U4: set_paper_size APPLIED (the editor\'s exact mm, owner OD-C/OD-A), retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_paper_size', id: 's', sheet: 'P', paper: 'US Letter' }]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 500));
  const p = h.store.durable.find(x => x.name === 'P').print;
  assert.equal(p.width, 215.9); assert.equal(p.height, 279.4);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_paper_size', id: 's', sheet: 'P', paper: 'US Letter' }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('GH #138 U4: set_header_footer_flag (always F2-widened) APPLIED, retry SATISFIED 0 dispatch', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_header_footer_flag', id: 's', sheet: 'P', flag: 'differentOddEven', value: true }]));
  assert.deepEqual(statuses(r), ['APPLIED'], J(r).slice(0, 500));
  assert.equal(h.store.durable.find(x => x.name === 'P').print.diffOddEven, true);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([{ intent: 'set_header_footer_flag', id: 's', sheet: 'P', flag: 'differentOddEven', value: true }]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('GH #138 U4: NEG -- a sheet that does not exist is refused named (xlsx-sheet-not-found), 0 dispatch, for a representative intent', async () => {
  const h = printHarness(createStore(book()));
  const r = await h.tool.handler(req([{ intent: 'set_print_fit', id: 's', sheet: 'Nope', fitToWidth: 50, fitToHeight: 50 }]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.equal(h.calls().dispatch, 0);
  assert.match(J(r), /xlsx-sheet-not-found/);
});

test('GH #138 U4: NEG -- the OWN family refusal surfaces (set_header_footer_flag always needs the F2-widened read; an unwidened build refuses it named, before dispatch is even reached)', async () => {
  const h = harness({ store: createStore(book()), script: { runtime: RT935, printWide: false } });
  // print.cjs's own needsWide()/stepGap(): set_header_footer_flag ALWAYS needs the F2-widened read
  // (diffFirst/diffOddEven); on an unwidened build this is xlsx-print-unobservable, not silently accepted.
  const r = await h.tool.handler(req([{ intent: 'set_header_footer_flag', id: 's', sheet: 'P', flag: 'differentFirst', value: true }]));
  assert.equal(r.ok, false, J(r).slice(0, 400));
  assert.match(J(r), /xlsx-print-unobservable/);
  assert.equal(h.calls().dispatch, 0);
});
