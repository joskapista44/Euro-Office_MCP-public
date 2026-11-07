'use strict';

/**
 * GH #296 (offline): delete_image / replace_image in the `image` family (src/adapters/docx/families/image.cjs)
 * and the driver's publicReceipt rule. plan/classify/resolve/verify run against fabricated `normalizeInventory`
 * observations; the REAL callCommand body runs via vm against a fake editor that behaves like the measured one
 * (test/live/docx-gh296-image-replace-delete-probe.live.cjs, run p296-r1: Delete() and ReplaceDrawing() return
 * true, the fingerprint the sandbox re-checks is the ToJSON extent + drawingType + the parent paragraph text).
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const img = require('../../src/adapters/docx/families/image.cjs');
const { receiptDispatchResult } = require('../../src/drivers/eurooffice/docx-driver.cjs');

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const JPEG = 'data:image/jpeg;base64,/9j/4AAQSkZJRg==';
const E = 914400;

// the p296-r1 SEED shape: four header/footer images (two with empty paragraph text) + a body inline + a body floating
const row = (index, part, paragraph, cx, drawingType = 'inline', wrapType = 'through') => ({ index, part, paragraphIndex: part === 'body' ? 4 + (drawingType === 'anchor' ? 1 : 0) : 0, cx, cy: E, drawingType, wrapType, paragraph });
const SEED = [
  row(0, 'header:0:title', 'ZQW296 HDR TITLE', 3 * E),
  row(1, 'header:0:default', 'ZQW296 HDR DEF', 1 * E),
  row(2, 'header:0:even', '', 4 * E),
  row(3, 'footer:0:default', '', 2 * E),
  row(4, 'body', 'ZQW296 BODY INLINE', 5 * E),
  row(5, 'body', 'ZQW296 BODY FLOAT', 6 * E, 'anchor', 'square'),
];
const HTML = '<p>ZQW296 BODY INLINE</p><p>ZQW296 BODY FLOAT</p>';
const reindex = rows => rows.map((r, i) => ({ ...r, index: i }));
const obs = (images, html = HTML) => ({ primary: normalizeInventory({ html, document: { images } }) });

const planned = step => {
  const p = img.plan(step, 0);
  assert.equal(p.ok, true, p.detail);
  return p.operation;
};

// ---- plan -----------------------------------------------------------------------------------------

test('plan: delete_image / replace_image take the image selector plus an optional part; replace takes source and an optional size pair', () => {
  const d = planned({ intent: 'delete_image', selector: { index: 2, part: 'header:0:even' } });
  assert.deepEqual(d.selector, { index: 2, part: 'header:0:even' });
  assert.equal(d.preservationMode, 'PRIMITIVE_INTEGRAL');
  const r = planned({ intent: 'replace_image', selector: { paragraph: 'ZQW296 BODY FLOAT' }, source: JPEG });
  assert.equal(r.source, JPEG);
  assert.equal('widthEmu' in r, false, 'no size given = keep the current size');
  const sized = planned({ intent: 'replace_image', selector: { index: 0, part: 'header:0:title' }, source: PNG, widthEmu: 7 * E, heightEmu: 2 * E });
  assert.deepEqual([sized.widthEmu, sized.heightEmu], [7 * E, 2 * E]);
});

test('plan: refuses a bad part, part on the older intents, an external / non-PNG-JPEG source, and a lone width or height', () => {
  const no = step => assert.equal(img.plan(step, 0).ok, false, JSON.stringify(step));
  no({ intent: 'delete_image', selector: { index: 0, part: 'header:0:odd' } });
  no({ intent: 'delete_image', selector: { index: 0, part: 'section:0:default' } });
  no({ intent: 'set_image_size', selector: { index: 0, part: 'body' }, widthEmu: E, heightEmu: E });
  no({ intent: 'delete_image', selector: { index: 0, paragraph: 'x' } });
  for (const source of ['https://example.com/x.png', 'data:image/gif;base64,abcd', '', undefined]) no({ intent: 'replace_image', selector: { index: 0 }, source });
  no({ intent: 'replace_image', selector: { index: 0 }, source: PNG, widthEmu: E });
  no({ intent: 'replace_image', selector: { index: 0 }, source: PNG, widthEmu: E, heightEmu: 0 });
});

// ---- classify / resolve ---------------------------------------------------------------------------

test('classify/resolve: the selector resolves on the fresh read; a part mismatch fails closed instead of re-targeting; never SATISFIED', () => {
  const o = obs(SEED);
  const d = planned({ intent: 'delete_image', selector: { index: 2, part: 'header:0:even' } });
  assert.equal(img.classify(o.primary, d).state, 'APPLY');
  const t = img.resolve(o.primary, d.selector, { operation: d });
  assert.equal(t.ok, true);
  assert.deepEqual(t.target, { kind: 'image', imageIndex: 2, expectedParagraph: '', expectedWidthEmu: 4 * E, expectedHeightEmu: E, expectedDrawingType: 'inline' });
  const wrongPart = planned({ intent: 'delete_image', selector: { index: 2, part: 'header:0:default' } });
  assert.equal(img.classify(o.primary, wrongPart).state, 'FAIL_CLOSED');
  assert.equal(img.resolve(o.primary, wrongPart.selector, { operation: wrongPart }).ambiguity, 'IMAGE_PART_MISMATCH');
  const floatT = img.resolve(o.primary, null, { operation: planned({ intent: 'replace_image', selector: { paragraph: 'ZQW296 BODY FLOAT' }, source: PNG }) });
  assert.equal(floatT.target.expectedDrawingType, 'anchor');
  // a landed delete retried: the index now names a different image -> the part assertion fails closed
  const afterDelete = obs(reindex(SEED.filter(r => r.index !== 2)));
  assert.equal(img.classify(afterDelete.primary, d).state, 'FAIL_CLOSED');
});

test('classify: an unreadable fingerprint (part enumeration gap) fails closed for delete/replace', () => {
  const gap = SEED.map(r => ({ ...r, part: 'UNOBSERVABLE:part-docpr-id-set-mismatch' }));
  const d = planned({ intent: 'delete_image', selector: { index: 4 } });
  assert.equal(img.classify(obs(gap).primary, d).state, 'FAIL_CLOSED');
});

// ---- the callCommand body (vm, fake editor) --------------------------------------------------------

function editor(rows) {
  const calls = [];
  const draw = r => ({
    GetParentParagraph: () => ({ GetText: () => `${r.paragraph}\r\n` }),
    ToJSON: () => JSON.stringify({ drawingType: r.drawingType, extent: { cx: r.cx, cy: r.cy } }),
    Delete: () => { calls.push(['Delete', r.index]); return true; },
  });
  const imgs = rows.map(draw);
  const doc = {
    GetAllImages: () => imgs,
    ReplaceDrawing: (old, neu, keep) => { calls.push(['ReplaceDrawing', imgs.indexOf(old), neu, keep]); return true; },
  };
  const Api = { GetDocument: () => doc, CreateImage: (src, w, h) => ({ src, w, h }) };
  return { Api, calls };
}

function runBody(op, rows, mutate) {
  const o = obs(SEED);
  const t = img.resolve(o.primary, op.selector, { operation: op });
  assert.equal(t.ok, true);
  const lowered = img.lower(op, t.target);
  assert.equal(lowered.meta.publicReceipt, true);
  const e = editor(mutate ? mutate(rows) : rows);
  const sandbox = { Api: e.Api };
  vm.createContext(sandbox);
  const ret = vm.runInContext(`(function(){ ${lowered.body} })()`, sandbox);
  return JSON.parse(JSON.stringify({ ret, calls: e.calls })); // the vm realm's objects have another prototype
}

test('body: delete re-checks the fingerprint, deletes, and returns the receipt', () => {
  const r = runBody(planned({ intent: 'delete_image', selector: { index: 3, part: 'footer:0:default' } }), SEED);
  assert.deepEqual(r.ret, { ok: true });
  assert.deepEqual(r.calls, [['Delete', 3]]);
});

test('body: replace keeps the old size when none is given, uses the new one when it is, always ReplaceDrawing(..., true)', () => {
  const keep = runBody(planned({ intent: 'replace_image', selector: { index: 3 }, source: JPEG }), SEED);
  assert.deepEqual(keep.ret, { ok: true });
  assert.deepEqual(keep.calls, [['ReplaceDrawing', 3, { src: JPEG, w: 2 * E, h: E }, true]]);
  const sized = runBody(planned({ intent: 'replace_image', selector: { index: 0, part: 'header:0:title' }, source: PNG, widthEmu: 7 * E, heightEmu: 2 * E }), SEED);
  assert.deepEqual(sized.calls, [['ReplaceDrawing', 0, { src: PNG, w: 7 * E, h: 2 * E }, true]]);
});

test('body NEG: a drifted image at the index (other size / other placement / other paragraph) is NOT touched and the receipt is refused', () => {
  const op = planned({ intent: 'delete_image', selector: { index: 2 } });
  for (const drift of [r => ({ ...r, cx: r.cx + 1 }), r => ({ ...r, drawingType: 'anchor' }), r => ({ ...r, paragraph: 'other' })]) {
    const r = runBody(op, SEED, rows => rows.map(x => (x.index === 2 ? drift(x) : x)));
    assert.equal(r.ret.ok, false);
    assert.deepEqual(r.calls, []);
  }
});

test('driver: the publicReceipt rule is ok only on a { ok: true } body result', () => {
  assert.equal(receiptDispatchResult({ ok: true, result: { ok: true } }).ok, true);
  assert.equal(receiptDispatchResult({ ok: true, result: { ok: false, reason: 'fingerprint-mismatch' } }).ok, false);
  assert.equal(receiptDispatchResult({ ok: true, result: undefined }).ok, false, 'a body that returned nothing is not a receipt');
  assert.equal(receiptDispatchResult({ ok: false, outcome: 'timeout' }).ok, false);
});

// ---- verify ---------------------------------------------------------------------------------------

test('verify delete: count -1 and every other image accounted for; NEG: an extra image vanishing, or nothing deleted, fails', () => {
  const op = planned({ intent: 'delete_image', selector: { index: 1, part: 'header:0:default' } });
  const before = obs(SEED);
  const good = obs(reindex(SEED.filter(r => r.index !== 1)));
  assert.equal(img.verifyPrimitive(op, before, good).ok, true);
  assert.equal(img.verifyPreservation(op, before, good).ok, true);
  assert.equal(img.verifyPrimitive(op, before, obs(SEED)).ok, false, 'nothing deleted');
  assert.equal(img.verifyPrimitive(op, before, obs(reindex(SEED.filter(r => r.index !== 1 && r.index !== 2)))).ok, false, 'two gone');
  assert.equal(img.verifyPrimitive(op, before, obs(reindex(SEED.filter(r => r.index !== 2)))).ok, false, 'the wrong one gone');
});

test('verify replace: the replacement at the same part + paragraph index + placement at the expected size; NEG: moved, resized, or another image changed', () => {
  const op = planned({ intent: 'replace_image', selector: { index: 0, part: 'header:0:title' }, source: PNG, widthEmu: 7 * E, heightEmu: 2 * E });
  const before = obs(SEED);
  const swap = f => obs(SEED.map(r => (r.index === 0 ? f(r) : r)));
  assert.equal(img.verifyPrimitive(op, before, swap(r => ({ ...r, cx: 7 * E, cy: 2 * E - 1 }))).ok, true, '±1 EMU drift accepted');
  assert.equal(img.verifyPreservation(op, before, swap(r => ({ ...r, cx: 7 * E, cy: 2 * E }))).ok, true);
  assert.equal(img.verifyPrimitive(op, before, obs(SEED)).ok, false, 'not resized = not the replacement asked for');
  assert.equal(img.verifyPrimitive(op, before, swap(r => ({ ...r, cx: 7 * E, cy: 2 * E, part: 'header:0:default' }))).ok, false, 'moved to another part');
  const otherChanged = obs(SEED.map(r => (r.index === 0 ? { ...r, cx: 7 * E, cy: 2 * E } : (r.index === 5 ? { ...r, wrapType: 'tight' } : r))));
  assert.equal(img.verifyPrimitive(op, before, otherChanged).ok, false, 'another image changed');
  const keep = planned({ intent: 'replace_image', selector: { index: 5 }, source: JPEG });
  assert.equal(img.verifyPrimitive(keep, before, obs(SEED)).ok, true, 'keep-size replace: same key at the same place');
});

test('verify preservation NEG: a body text change fails the edit intents', () => {
  const op = planned({ intent: 'delete_image', selector: { index: 2 } });
  const after = obs(reindex(SEED.filter(r => r.index !== 2)), '<p>ZQW296 BODY INLINE changed</p><p>ZQW296 BODY FLOAT</p>');
  assert.equal(img.verifyPrimitive(op, obs(SEED), after).ok, true);
  assert.equal(img.verifyPreservation(op, obs(SEED), after).ok, false);
});
