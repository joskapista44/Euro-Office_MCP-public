'use strict';

/**
 * GH #309 part 1 (offline): the `insert_shape` family (src/adapters/docx/families/shape-insert.cjs), the driver's page-side shape read (docx-shapes-read.cjs) and the inventory's
 * `document.shapes` rows. plan/classify/resolve/verify against fabricated `normalizeInventory` observations; the read and the callCommand body run against SCRIPTED ApiDocument stand-ins
 * -- no editor, no VM. What the real editor does is covered by the LIVE probes (docx-gh309-shape-probe / -shape-read-probe, zero PASS) and the LIVE product harness
 * (docx-gh309-insert-shape-product.live.cjs).
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const si = require('../../src/adapters/docx/families/shape-insert.cjs');
const { canonicalBlock } = require('../../src/adapters/docx/families/table-insert.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { readShapesInto, SHAPES_READ_CALL } = require('../../src/drivers/eurooffice/docx-shapes-read.cjs');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

const ANCHOR = 'GH309 anchor';
const SIBLING = 'GH309 sibling';
const OTHER = 'GH309 other';
const TEXT = 'GH309 shape text';
const MM = 36000;
const W = 60 * MM;
const H = 20 * MM;
const SHAPE = { type: 'rect', widthMm: 60, heightMm: 20, fill: 'ffe699', stroke: 'C00000', text: TEXT };

// a driver-shape row (docx-shapes-read.cjs row shape) and an observation built through the REAL normalizeInventory
// the placement fields an INLINE shape reads back through ToJSON (LIVE: drawingType 'inline', wrapType 'through', default positions)
const INLINE = Object.freeze({ drawingType: 'inline', wrapType: 'through', behindDoc: false, hRel: 'column', hAlign: null, hOffset: 0, vRel: 'paragraph', vAlign: null, vOffset: 0 });
const row = (previous, over = {}) => ({ index: 0, text: TEXT, cx: W, cy: H, paragraph: '', previous, ...INLINE, ...over });
function obsOf(html, shapes = [], images = []) {
  return { primary: normalizeInventory({ html, document: { shapes, images } }) };
}
const step = (over = {}) => ({ intent: 'insert_shape', selector: { afterParagraph: ANCHOR }, shape: { ...SHAPE }, ...over });
const op = (over = {}) => si.plan(step(over), 0).operation;

// ---- plan -----------------------------------------------------------------------------------------

test('plan: accepts { afterParagraph } + the full shape object; converts mm to EMU (x36000) and upper-cases the colours', () => {
  const p = si.plan(step(), 0);
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operation.preservationMode, 'PRIMITIVE_INTEGRAL');
  assert.equal(p.operation.preset, 'rect');
  assert.equal(p.operation.widthEmu, W);
  assert.equal(p.operation.heightEmu, H);
  assert.equal(p.operation.fill, 'FFE699');
  assert.equal(p.operation.stroke, 'C00000');
  assert.equal(p.operation.text, TEXT);
  assert.equal(si.plan(step({ shape: { ...SHAPE, type: 'flowChartProcess' } }), 0).ok, true);
  for (const t of ['roundRect', 'ellipse', 'triangle', 'trapezoid', 'flowChartDecision', 'flowChartTerminator', 'rightArrow', 'chevron', 'homePlate']) {
    const q = si.plan(step({ shape: { ...SHAPE, type: t } }), 0);
    assert.equal(q.ok, true, t);
    assert.equal(q.operation.preset, t);
  }
});

test('plan: refuses a wrong-shaped selector, a missing / extra / wrong-typed shape field, an UNMEASURED preset, a non-integer or out-of-range size, a bad colour, and an empty / multi-line / over-long text', () => {
  const bad = (over, why) => assert.equal(si.plan(step(over), 0).ok, false, why);
  bad({ selector: {} }, 'empty selector');
  bad({ selector: { afterParagraph: '' } }, 'blank anchor');
  bad({ selector: { afterParagraph: ANCHOR, extra: 1 } }, 'extra selector key');
  bad({ shape: undefined }, 'no shape');
  bad({ shape: 'rect' }, 'shape is not an object');
  bad({ shape: [SHAPE] }, 'shape is an array');
  for (const k of Object.keys(SHAPE)) { const s = { ...SHAPE }; delete s[k]; bad({ shape: s }, `missing ${k}`); }
  bad({ shape: { ...SHAPE, extra: 1 } }, 'extra shape key');
  for (const t of ['star5', 'cloud', 'RECT', 'roundrect', '', undefined]) bad({ shape: { ...SHAPE, type: t } }, `type ${t}`);
  for (const v of [0, -1, 1.5, 401, '60', undefined, NaN]) { bad({ shape: { ...SHAPE, widthMm: v } }, `widthMm ${v}`); bad({ shape: { ...SHAPE, heightMm: v } }, `heightMm ${v}`); }
  for (const v of ['red', 'FFE69', 'FFE6999', '#FFE699', 'GGGGGG', 123456, undefined]) { bad({ shape: { ...SHAPE, fill: v } }, `fill ${v}`); bad({ shape: { ...SHAPE, stroke: v } }, `stroke ${v}`); }
  for (const t of ['', '   ', 'a\nb', 'a\rb', 'a\tb', 'a\u0000b', 5, 'x'.repeat(501)]) bad({ shape: { ...SHAPE, text: t } }, `text ${JSON.stringify(t).slice(0, 12)}`);
  assert.equal(si.plan(step({ shape: { ...SHAPE, text: 'x'.repeat(500) } }), 0).ok, true, '500 chars is the cap, not over it');
});

test('resolve / preDispatchCheck: round-trip through the real planOperations pipeline', () => {
  const planned = planOperations({ steps: [step()] });
  assert.equal(planned.ok, true, planned.detail);
  const o = planned.operations[0];
  const r = si.resolve(obsOf(`<p>${ANCHOR}</p>`).primary, o.selector);
  assert.equal(r.ok, true);
  assert.equal(r.target.anchorText, ANCHOR);
  assert.equal(si.preDispatchCheck(o, r.target).ok, true);
  assert.equal(si.preDispatchCheck(o, { kind: 'insert_shape_anchor', anchorText: 'wrong' }).ok, false);
  assert.equal(si.resolve(obsOf(`<p>${OTHER}</p>`).primary, o.selector).ok, false, 'a missing anchor does not resolve');
});

// ---- classify ---------------------------------------------------------------------------------------

test('classify: APPLY when nothing follows the anchor, the next block is not empty, or no matching shape row exists', () => {
  assert.equal(si.classify(obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p>`).primary, op()).state, 'APPLY', 'anchor is last');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p>${SIBLING}</p>`).primary, op()).state, 'APPLY', 'next block not empty');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p></p>`).primary, op()).state, 'APPLY', 'an EMPTY paragraph alone is NOT a shape -- the very false positive the shape list exists to prevent');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR, { text: 'other text' })]).primary, op()).state, 'APPLY', 'wrong own text');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR, { cx: 1 })]).primary, op()).state, 'APPLY', 'wrong size');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [row(OTHER)]).primary, op()).state, 'APPLY', 'a shape after a DIFFERENT paragraph is not this one');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR, { paragraph: 'has text' })]).primary, op()).state, 'APPLY', 'the shape paragraph must be empty');
});

test('classify: SATISFIED when the block right after the anchor is an empty paragraph AND a shape row there has the requested text and size (+-1 EMU)', () => {
  const html = `<p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  assert.equal(si.classify(obsOf(html, [row(ANCHOR)]).primary, op()).state, 'SATISFIED');
  assert.equal(si.classify(obsOf(html, [row(ANCHOR, { cx: W - 1, cy: H + 1 })]).primary, op()).state, 'SATISFIED', '+-1 EMU drift tolerated');
  assert.equal(si.classify(obsOf(html, [row(ANCHOR, { cx: W - 2 })]).primary, op()).state, 'APPLY', '+-2 is not');
});

test('classify: FAIL_CLOSED when the anchor is missing / ambiguous / a non-paragraph, and when the shape list is UNOBSERVABLE (a retry could duplicate -- never guessed as "none")', () => {
  assert.equal(si.classify(obsOf(`<p>${OTHER}</p>`).primary, op()).state, 'FAIL_CLOSED');
  assert.equal(si.classify(obsOf(`<p>${ANCHOR}</p><p>${ANCHOR}</p>`).primary, op()).state, 'FAIL_CLOSED');
  for (const unread of ['UNOBSERVABLE:no-GetAllShapes', 'UNOBSERVABLE:unread', undefined]) {
    const primary = normalizeInventory({ html: `<p>${ANCHOR}</p>`, document: { shapes: unread } }); // a string / absent -> not an array
    const c = si.classify(primary, op());
    assert.equal(c.state, 'FAIL_CLOSED', String(unread));
    assert.equal(c.evidence.reason, 'insert-shape-shapes-unobservable');
  }
});

// ---- the inventory ----------------------------------------------------------------------------------

test('normalizeInventory: document.shapes rows are frozen primitives; absent -> null, a failed field stays UNOBSERVABLE; the whole list unread is a named string', () => {
  const p = normalizeInventory({ html: `<p>${ANCHOR}</p>`, document: { shapes: [row(ANCHOR), { index: 1, text: 'UNOBSERVABLE:x', cx: 'UNOBSERVABLE:y', cy: 'UNOBSERVABLE:y', paragraph: null, previous: undefined }] } });
  const rows = si.shapeRows(p);
  assert.equal(rows.length, 2);
  assert.deepEqual({ ...rows[0] }, { index: 0, text: TEXT, widthEmu: W, heightEmu: H, paragraph: '', previous: ANCHOR, ...INLINE });
  assert.equal(rows[1].text, 'UNOBSERVABLE:x');
  assert.equal(rows[1].widthEmu, 'UNOBSERVABLE:y');
  assert.equal(rows[1].paragraph, null);
  assert.equal(rows[1].previous, null);
  assert.ok(Object.isFrozen(rows) && Object.isFrozen(rows[0]));
  assert.equal(si.shapeRows(normalizeInventory({ html: '<p>x</p>', document: { shapes: 'UNOBSERVABLE:no-GetAllShapes' } })), null);
  assert.equal(normalizeInventory({ html: '<p>x</p>', document: { shapes: 'UNOBSERVABLE:no-GetAllShapes' } }).document.shapes, 'UNOBSERVABLE:no-GetAllShapes');
  assert.equal(normalizeInventory({ html: '<p>x</p>', document: {} }).document.shapes, 'UNOBSERVABLE:unread');
});

test('NO OTHER INTENT MOVES: the shapes list never reaches a block -- blocks and canonicalBlock are byte-identical with and without shapes', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  const a = normalizeInventory({ html, document: { shapes: [row(ANCHOR)] } });
  const b = normalizeInventory({ html, document: { shapes: [] } });
  assert.equal(JSON.stringify(a.blocks.map(x => canonicalBlock(x))), JSON.stringify(b.blocks.map(x => canonicalBlock(x))));
  assert.equal(JSON.stringify(a.blocks), JSON.stringify(b.blocks));
});

// ---- verify ---------------------------------------------------------------------------------------

test('verifyPrimitive: ok only when the after-state has the empty paragraph right after the anchor AND the matching shape row; an empty paragraph alone fails', () => {
  const before = obsOf(`<p>${ANCHOR}</p>`);
  const good = si.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR)]));
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
  assert.equal(si.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p><p></p>`)).ok, false, 'empty paragraph, no shape');
  assert.equal(si.verifyPrimitive(op(), before, obsOf(`<p>${ANCHOR}</p>`, [row(ANCHOR)])).ok, false, 'a shape row but no new block');
  const unread = { primary: normalizeInventory({ html: `<p>${ANCHOR}</p><p></p>`, document: { shapes: 'UNOBSERVABLE:unread' } }) };
  assert.equal(si.verifyPrimitive(op(), before, unread).ok, false, 'unreadable shape list in the after-state');
});

test('verifyPreservation: OK when exactly one new shape + exactly the one new empty paragraph appear and everything else (blocks, images, other shapes) is unchanged', () => {
  const before = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`);
  const after = obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`, [row(ANCHOR)]);
  const good = si.verifyPreservation(op(), before, after);
  assert.equal(good.ok, true, JSON.stringify(good.evidence));
});

test('verifyPreservation: an EXISTING shape that followed the anchor keeps its identity (its `previous` moved, its own key did not) -- not a false loss', () => {
  const existing = { index: 0, text: 'old shape', cx: 5 * MM, cy: 5 * MM, paragraph: '', previous: ANCHOR };
  const before = obsOf(`<p>${ANCHOR}</p><p></p>`, [existing]);
  const after = obsOf(`<p>${ANCHOR}</p><p></p><p></p>`, [{ ...row(ANCHOR), index: 0 }, { ...existing, index: 1, previous: '' }]);
  const r = si.verifyPreservation(op(), before, after);
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

test('verifyPreservation: FAILS on a lost / mutated existing shape, no new shape, TWO new shapes, a changed image, a missing / non-empty new block, or any block that shifted', () => {
  const html = `<p>${OTHER}</p><p>${ANCHOR}</p><p>${SIBLING}</p>`;
  const withNew = `<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  const existing = { index: 0, text: 'old', cx: 1, cy: 1, paragraph: '', previous: OTHER };
  const before = obsOf(html, [existing]);
  assert.equal(si.verifyPreservation(op(), before, obsOf(withNew, [row(ANCHOR)])).ok, false, 'the existing shape vanished');
  assert.equal(si.verifyPreservation(op(), before, obsOf(withNew, [{ ...existing, text: 'changed' }, row(ANCHOR)])).ok, false, 'the existing shape was mutated');
  assert.equal(si.verifyPreservation(op(), before, obsOf(withNew, [existing])).ok, false, 'no new shape');
  assert.equal(si.verifyPreservation(op(), before, obsOf(withNew, [existing, row(ANCHOR), row(ANCHOR, { text: 'second' })])).ok, false, 'two new shapes');
  const img = { index: 0, drawingType: 'inline', cx: 1, cy: 1, paragraph: OTHER };
  assert.equal(si.verifyPreservation(op(), obsOf(html, [existing], [img]), obsOf(withNew, [existing, row(ANCHOR)], [])).ok, false, 'an image was lost');
  assert.equal(si.verifyPreservation(op(), before, obsOf(html, [existing, row(ANCHOR)])).ok, false, 'no new block');
  assert.equal(si.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p>stray</p><p>${SIBLING}</p>`, [existing, row(ANCHOR)])).ok, false, 'the new block carries text');
  assert.equal(si.verifyPreservation(op(), before, obsOf(`<p>${OTHER} CHANGED</p><p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`, [existing, row(ANCHOR)])).ok, false, 'a block before the anchor changed');
  assert.equal(si.verifyPreservation(op(), before, obsOf(`<p>${OTHER}</p><p>${ANCHOR}</p><p></p><p>${SIBLING} CHANGED</p>`, [existing, row(ANCHOR)])).ok, false, 'a block after the new one changed');
  const unread = { primary: normalizeInventory({ html: withNew, document: { shapes: 'UNOBSERVABLE:unread' } }) };
  assert.equal(si.verifyPreservation(op(), before, unread).ok, false, 'the after shape list is unreadable');
});

test('verifyPreservation: the ANCHOR block alone tolerates the DS mso-border-*-alt:none rewrite (the #322 lesson); the same rewrite on any OTHER block still fails', () => {
  const styled = style => `<p style="${style}">${ANCHOR}</p>`;
  const before = obsOf(`${styled('border-left:none')}<p>${SIBLING}</p>`);
  const afterAnchorRewritten = obsOf(`${styled('mso-border-left-alt:none')}<p></p><p>${SIBLING}</p>`, [row(ANCHOR)]);
  const okAnchor = si.verifyPreservation(op(), before, afterAnchorRewritten);
  assert.equal(okAnchor.ok, true, JSON.stringify(okAnchor.evidence));
  const beforeOther = obsOf(`${styled('border-left:none')}<p style="border-left:none">${SIBLING}</p>`);
  const afterOtherRewritten = obsOf(`${styled('border-left:none')}<p></p><p style="mso-border-left-alt:none">${SIBLING}</p>`, [row(ANCHOR)]);
  assert.equal(si.verifyPreservation(op(), beforeOther, afterOtherRewritten).ok, false, 'a non-anchor block must still compare byte-exact');
});

test('wholeTaskChecks: one named check, true only on the matching final shape; producedText: a shape\'s text is not a body unit a later step can ref', () => {
  const ok = si.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR)]));
  assert.equal(ok.length, 1);
  assert.equal(ok[0].ok, true);
  assert.equal(si.wholeTaskChecks(op(), obsOf(`<p>${ANCHOR}</p><p></p>`))[0].ok, false);
  assert.equal(si.wholeTaskChecks(op(), obsOf(`<p>${OTHER}</p>`))[0].ok, false, 'an unresolvable anchor is false, never a throw');
  assert.equal(si.producedText({}), null);
});

// ---- the callCommand body ---------------------------------------------------------------------------

/** a scripted ApiDocument + Api that RECORDS what the body does, in order. */
function scripted(texts, over = {}) {
  const calls = [];
  const mkPara = text => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, AddDrawing(s) { calls.push('AddDrawing'); this.drawing = s; }, text });
  const elements = texts.map(mkPara);
  const doc = { GetElementsCount: () => elements.length, GetElement: i => elements[i], AddElement(i, p) { calls.push(`AddElement@${i}`); elements.splice(i, 0, p); } };
  const Api = {
    GetDocument: () => doc,
    CreateParagraph: () => mkPara(''),
    CreateRGBColor: (r, g, b) => { calls.push(`rgb(${r},${g},${b})`); return { r, g, b }; },
    CreateSolidFill: c => ({ fill: c }),
    CreateStroke: (w, f) => { calls.push(`stroke(${w})`); return { w, f }; },
    CreateShape: (prst, w, h, fill, stroke) => {
      calls.push(`CreateShape(${prst},${w},${h})`);
      const para = { AddText: t => calls.push(`text(${t})`) };
      return { prst, w, h, fill, stroke, GetDocContent: () => ({ GetElement: () => para }) };
    },
    ...over,
  };
  return { Api, doc, elements, calls };
}
const run = (body, Api) => new Function('Api', body)(Api);
const lowered = o => si.lower(o, { kind: 'insert_shape_anchor', anchorText: ANCHOR });

test('commandBody: builds the shape + its text FIRST, then a fresh paragraph carries it and lands at anchorIndex + 1 -- the insert is the LAST step', () => {
  const s = scripted([OTHER, ANCHOR, SIBLING]);
  run(lowered(op()).body, s.Api);
  assert.deepEqual(s.calls, ['rgb(255,230,153)', 'rgb(192,0,0)', 'stroke(19050)', `CreateShape(rect,${W},${H})`, `text(${TEXT})`, 'AddDrawing', 'AddElement@2']);
  assert.equal(s.elements.length, 4);
  assert.equal(s.elements[2].drawing.prst, 'rect');
  assert.equal(s.elements[1].text, ANCHOR, 'the anchor itself is untouched');
});

test('commandBody: any guard miss leaves the document UNCHANGED -- missing / ambiguous anchor, a missing API, a shape with no text content, a throwing step', () => {
  const body = lowered(op()).body;
  const untouched = (s, why) => { run(body, s.Api); assert.equal(s.elements.length, s.n || 3, why); assert.ok(!s.calls.includes('AddElement@2') && !s.calls.some(c => c.startsWith('AddElement')), why); };
  untouched(scripted([OTHER, 'not the anchor', SIBLING]), 'no anchor');
  const dup = scripted([ANCHOR, SIBLING, ANCHOR]); run(body, dup.Api); assert.equal(dup.elements.length, 3, 'ambiguous anchor refused'); assert.ok(!dup.calls.some(c => c.startsWith('AddElement')));
  for (const missing of ['CreateShape', 'CreateSolidFill', 'CreateRGBColor', 'CreateStroke']) { const s = scripted([OTHER, ANCHOR, SIBLING], { [missing]: undefined }); untouched(s, `missing ${missing}`); }
  const noDocContent = scripted([OTHER, ANCHOR, SIBLING], { CreateShape: () => ({}) }); untouched(noDocContent, 'no GetDocContent');
  const noPara = scripted([OTHER, ANCHOR, SIBLING], { CreateShape: () => ({ GetDocContent: () => ({ GetElement: () => null }) }) }); untouched(noPara, 'no first paragraph');
  const noAddText = scripted([OTHER, ANCHOR, SIBLING], { CreateShape: () => ({ GetDocContent: () => ({ GetElement: () => ({}) }) }) }); untouched(noAddText, 'no AddText');
  const throwing = scripted([OTHER, ANCHOR, SIBLING], { CreateShape: () => { throw new Error('boom'); } }); assert.doesNotThrow(() => run(body, throwing.Api)); untouched(throwing, 'CreateShape throws');
  const throwText = scripted([OTHER, ANCHOR, SIBLING], { CreateShape: () => ({ GetDocContent: () => ({ GetElement: () => ({ AddText() { throw new Error('t'); } }) }) }) }); untouched(throwText, 'AddText throws BEFORE anything touches the document');
  const noAddEl = scripted([OTHER, ANCHOR, SIBLING]); noAddEl.doc.AddElement = undefined; run(body, noAddEl.Api); assert.equal(noAddEl.elements.length, 3);
});

test('commandBody: the body is valid JavaScript and carries the measured stroke width, never a regex-class / template hazard', () => {
  const body = lowered(op()).body;
  assert.doesNotThrow(() => new Function('Api', body));
  assert.ok(body.includes('19050'));
  assert.equal(si.STROKE_WIDTH_EMU, 19050);
});

// ---- the driver's page-side read ------------------------------------------------------------------------

const INLINE_JSON = { drawingType: 'inline', wrapType: 'through', behindDoc: false, positionH: { relativeFrom: 'column', posOffset: 0, percent: false }, positionV: { relativeFrom: 'paragraph', posOffset: 0, percent: false } };
const shapeObj = (text, w, h, cls = 'shape', json = INLINE_JSON) => ({ GetClassType: () => cls, GetDocContent: () => ({ GetElement: () => ({ GetText: () => `${text}\r\n` }) }), GetWidth: () => w, GetHeight: () => h, ToJSON: () => JSON.stringify(json) });
const para = (text, drawings = []) => ({ GetClassType: () => 'paragraph', GetText: () => `${text}\r\n`, GetAllDrawingObjects: () => drawings });
const docOf = (elements, over = {}) => ({ GetAllShapes: () => [], GetElementsCount: () => elements.length, GetElement: i => elements[i], ...over });
const readOf = doc => { const info = {}; readShapesInto(doc, info); return info.shapes; };

test('readShapesInto: one primitive row per SHAPE (class "shape" only), with the shape\'s own text, EMU size, its own (empty) paragraph and the text of the element right before it', () => {
  const rows = readOf(docOf([para(OTHER), para(ANCHOR), para('', [shapeObj(TEXT, W, H)]), para(SIBLING), para('', [shapeObj('two', 1, 2), shapeObj('img', 9, 9, 'image')])]));
  assert.deepEqual(rows, [
    { index: 0, text: TEXT, cx: W, cy: H, paragraph: '', previous: ANCHOR, ...INLINE },
    { index: 1, text: 'two', cx: 1, cy: 2, paragraph: '', previous: SIBLING, ...INLINE },
  ]);
  assert.ok(rows.every(r => Object.values(r).every(v => v === null || ['string', 'number', 'boolean'].includes(typeof v))), 'every field is a primitive');
});

test('readShapesInto: `previous` is null when the preceding element is not a paragraph (a table) or the shape paragraph is first; no shapes -> an empty list, not a gap', () => {
  const table = { GetClassType: () => 'table' };
  assert.equal(readOf(docOf([table, para('', [shapeObj('a', 1, 1)])]))[0].previous, null);
  assert.equal(readOf(docOf([para('', [shapeObj('a', 1, 1)])]))[0].previous, null);
  assert.deepEqual(readOf(docOf([para(OTHER)])), []);
});

test('readShapesInto: a failed read is NAMED, never a guessed value -- a throwing own-text / size, a missing or throwing GetAllShapes, a non-numeric size', () => {
  const badText = { ...shapeObj('x', 1, 1), GetDocContent: () => { throw new Error('no content'); } };
  const badSize = { ...shapeObj('x', 1, 1), GetWidth: () => { throw new Error('no width'); } };
  const rows = readOf(docOf([para(ANCHOR), para('', [badText, badSize, shapeObj('z', 'wide', null)])]));
  assert.equal(rows[0].text, 'UNOBSERVABLE:no content'); assert.equal(rows[0].cx, 1, 'a text failure does not hide the size');
  assert.equal(rows[1].cx, 'UNOBSERVABLE:no width'); assert.equal(rows[1].cy, 'UNOBSERVABLE:no width');
  assert.equal(rows[2].cx, null); assert.equal(rows[2].cy, null);
  assert.equal(readOf({ GetElementsCount: () => 0 }), 'UNOBSERVABLE:no-GetAllShapes');
  assert.equal(readOf(docOf([], { GetElementsCount: () => { throw new Error('gone'); } })), 'UNOBSERVABLE:gone');
});

test('the page body: SHAPES_READ_CALL is self-contained (compiles and runs on its own), holds no backtick / dollar-brace, and sits in the driver\'s readLive body; the body still compiles', () => {
  const info = {};
  new Function('doc', 'docInfo', SHAPES_READ_CALL)(docOf([para(ANCHOR), para('', [shapeObj(TEXT, W, H)])]), info);
  assert.equal(info.shapes[0].text, TEXT);
  assert.ok(!/`|\$\{/.test(readShapesInto.toString()), 'the injected source must hold no backtick / dollar-brace (it lives inside a template literal)');
  const body = propertyQuery({});
  assert.ok(body.includes('GetAllShapes') && body.includes('docInfo.shapes'), 'the driver body carries the projection');
  assert.doesNotThrow(() => new Function(body), 'the whole readLive body must still be valid JavaScript');
});

test('NEG control of the read: a doc whose shape paragraph has NO text is still read (the per-paragraph read would skip it -- this walk must not)', () => {
  const rows = readOf(docOf([para(ANCHOR), para('', [shapeObj(TEXT, W, H)])]));
  assert.equal(rows.length, 1);
  assert.equal(rows[0].paragraph, '');
});

// ---- GH #309 part 3: floating placement ------------------------------------------------------------------

const PL = { wrap: 'square', horizontal: { relativeTo: 'margin', align: 'right' }, vertical: { relativeTo: 'paragraph', offsetMm: 5 } };
const PL_PAGE = { wrap: 'behind', horizontal: { relativeTo: 'page', offsetMm: 25 }, vertical: { relativeTo: 'page', offsetMm: 60 } };
const SQUARE_READ = Object.freeze({ drawingType: 'anchor', wrapType: 'square', behindDoc: false, hRel: 'margin', hAlign: 'right', hOffset: null, vRel: 'paragraph', vAlign: null, vOffset: 5 * MM });
const BEHIND_READ = Object.freeze({ drawingType: 'anchor', wrapType: 'none', behindDoc: true, hRel: 'page', hAlign: null, hOffset: 25 * MM, vRel: 'page', vAlign: null, vOffset: 60 * MM });
const fop = (pl = PL) => op({ shape: { ...SHAPE, placement: pl } });

test('plan (placement): the four measured wraps x the measured references are accepted and normalised to the expected read-back (EMU)', () => {
  assert.deepEqual({ ...fop().placement }, { wrap: 'square', wrapType: 'square', behindDoc: false, hRel: 'margin', hAlign: 'right', hOffsetEmu: null, vRel: 'paragraph', vOffsetEmu: 5 * MM });
  assert.deepEqual({ ...fop(PL_PAGE).placement }, { wrap: 'behind', wrapType: 'none', behindDoc: true, hRel: 'page', hAlign: null, hOffsetEmu: 25 * MM, vRel: 'page', vOffsetEmu: 60 * MM });
  for (const [wrap, wrapType, behindDoc] of [['topAndBottom', 'top_and_bottom', false], ['inFront', 'none', false]]) {
    const p = fop({ ...PL, wrap }).placement;
    assert.equal(p.wrapType, wrapType); assert.equal(p.behindDoc, behindDoc);
  }
  assert.equal(op().placement, null, 'no placement -> inline (part 1 unchanged)');
});

test('plan (placement): unmeasured / malformed placements are refused before dispatch', () => {
  const bad = (pl, why) => { const p = si.plan(step({ shape: { ...SHAPE, placement: pl } }), 0); assert.equal(p.ok, false, why); assert.match(p.detail, /placement/, why); };
  for (const w of ['tight', 'through', 'inline', 'Square', '', undefined]) bad({ ...PL, wrap: w }, `wrap ${w}`);
  bad({ ...PL, extra: 1 }, 'extra placement key');
  bad({ wrap: 'square', horizontal: PL.horizontal }, 'missing vertical');
  bad(null, 'null placement'); bad('square', 'string placement'); bad([PL], 'array placement');
  bad({ ...PL, horizontal: { relativeTo: 'margin', align: 'inside' } }, 'unmeasured align');
  bad({ ...PL, horizontal: { relativeTo: 'margin', offsetMm: 5 } }, 'margin + offset not measured');
  bad({ ...PL, horizontal: { relativeTo: 'page', align: 'right' } }, 'page + align not measured');
  bad({ ...PL, horizontal: { relativeTo: 'column', offsetMm: 5 } }, 'column reference not measured');
  bad({ ...PL, horizontal: { relativeTo: 'page', offsetMm: 5, align: 'left' } }, 'both align and offset');
  for (const v of [-1, 1.5, 401, '5', undefined]) bad({ ...PL, vertical: { relativeTo: 'page', offsetMm: v } }, `vertical offset ${v}`);
  bad({ ...PL, vertical: { relativeTo: 'margin', offsetMm: 5 } }, 'vertical margin not measured');
  bad({ ...PL, vertical: { relativeTo: 'page', align: 'top' } }, 'vertical align not measured');
  assert.equal(si.plan(step({ shape: { ...SHAPE, placement: { ...PL, vertical: { relativeTo: 'page', offsetMm: 400 } } } }), 0).ok, true, '400 mm is the cap, not over it');
});

test('classify (placement): SATISFIED only when the read-back placement matches; inline and floating never satisfy each other', () => {
  const html = `<p>${ANCHOR}</p><p></p><p>${SIBLING}</p>`;
  assert.equal(si.classify(obsOf(html, [row(ANCHOR, SQUARE_READ)]).primary, fop()).state, 'SATISFIED');
  assert.equal(si.classify(obsOf(html, [row(ANCHOR, BEHIND_READ)]).primary, fop(PL_PAGE)).state, 'SATISFIED');
  assert.equal(si.classify(obsOf(html, [row(ANCHOR)]).primary, fop()).state, 'APPLY', 'an inline shape does not satisfy a floating request');
  assert.equal(si.classify(obsOf(html, [row(ANCHOR, SQUARE_READ)]).primary, op()).state, 'APPLY', 'a floating shape does not satisfy an inline request');
  for (const [k, v] of [['wrapType', 'none'], ['behindDoc', true], ['hAlign', 'left'], ['hRel', 'page'], ['vOffset', 0], ['vRel', 'page'], ['vAlign', 'top']]) {
    assert.equal(si.classify(obsOf(html, [row(ANCHOR, { ...SQUARE_READ, [k]: v })]).primary, fop()).state, 'APPLY', `${k} differs`);
  }
  assert.equal(si.classify(obsOf(html, [row(ANCHOR, { ...BEHIND_READ, hOffset: 25 * MM + 1 })]).primary, fop(PL_PAGE)).state, 'APPLY', 'an offset compares EMU-exact');
});

test('classify (placement): a shape with this text and size after the anchor whose placement is UNREADABLE fails closed (a retry could duplicate it)', () => {
  const html = `<p>${ANCHOR}</p><p></p>`;
  const unread = row(ANCHOR, { drawingType: 'UNOBSERVABLE:x', wrapType: 'UNOBSERVABLE:x', behindDoc: 'UNOBSERVABLE:x', hRel: 'UNOBSERVABLE:x', hAlign: 'UNOBSERVABLE:x', hOffset: 'UNOBSERVABLE:x', vRel: 'UNOBSERVABLE:x', vAlign: 'UNOBSERVABLE:x', vOffset: 'UNOBSERVABLE:x' });
  for (const o of [op(), fop()]) {
    const c = si.classify(obsOf(html, [unread]).primary, o);
    assert.equal(c.state, 'FAIL_CLOSED'); assert.equal(c.evidence.reason, 'insert-shape-placement-unobservable');
  }
  assert.equal(si.classify(obsOf(html, [row(OTHER, { drawingType: 'UNOBSERVABLE:x' })]).primary, op()).state, 'APPLY', 'an unreadable shape ELSEWHERE does not block');
});

test('verify (placement): verifyPrimitive needs the placement; preservation sees an existing shape whose placement changed as mutated', () => {
  const before = obsOf(`<p>${ANCHOR}</p>`);
  assert.equal(si.verifyPrimitive(fop(), before, obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR, SQUARE_READ)])).ok, true);
  assert.equal(si.verifyPrimitive(fop(), before, obsOf(`<p>${ANCHOR}</p><p></p>`, [row(ANCHOR)])).ok, false, 'landed inline instead');
  const existing = { index: 0, text: 'old', cx: 1, cy: 1, paragraph: '', previous: OTHER, ...INLINE };
  const html = `<p>${OTHER}</p><p></p><p>${ANCHOR}</p>`;
  const withNew = `<p>${OTHER}</p><p></p><p>${ANCHOR}</p><p></p>`;
  assert.equal(si.verifyPreservation(fop(), obsOf(html, [existing]), obsOf(withNew, [existing, row(ANCHOR, SQUARE_READ)])).ok, true);
  assert.equal(si.verifyPreservation(fop(), obsOf(html, [existing]), obsOf(withNew, [{ ...existing, ...SQUARE_READ }, row(ANCHOR, SQUARE_READ)])).ok, false, 'an existing shape turned floating');
  assert.match(si.wholeTaskChecks(fop(), obsOf(withNew, [existing, row(ANCHOR, SQUARE_READ)]))[0].name, /floating \(square\)/);
});

function scriptedFloating(over = {}) {
  const s = scripted([OTHER, ANCHOR, SIBLING], {
    CreateShape: (prst, w, h) => {
      s.calls.push(`CreateShape(${prst},${w},${h})`);
      const para = { AddText: t => s.calls.push(`text(${t})`) };
      return {
        prst, GetDocContent: () => ({ GetElement: () => para }),
        SetWrappingStyle: w2 => { s.calls.push(`wrap(${w2})`); return true; },
        SetHorAlign: (r, a) => { s.calls.push(`hAlign(${r},${a})`); return true; },
        SetHorPosition: (r, o) => { s.calls.push(`hPos(${r},${o})`); return true; },
        SetVerPosition: (r, o) => { s.calls.push(`vPos(${r},${o})`); return true; },
        ...over,
      };
    },
  });
  return s;
}

test('commandBody (placement): wrap + horizontal + vertical are set on the shape BEFORE it is added; an inline op calls none of them', () => {
  const a = scriptedFloating(); run(lowered(fop()).body, a.Api);
  assert.deepEqual(a.calls.slice(-5), ['wrap(square)', 'hAlign(margin,right)', `vPos(paragraph,${5 * MM})`, 'AddDrawing', 'AddElement@2']);
  const b = scriptedFloating(); run(lowered(fop(PL_PAGE)).body, b.Api);
  assert.deepEqual(b.calls.slice(-5), ['wrap(behind)', `hPos(page,${25 * MM})`, `vPos(page,${60 * MM})`, 'AddDrawing', 'AddElement@2']);
  const c = scriptedFloating(); run(lowered(op()).body, c.Api);
  assert.ok(!c.calls.some(x => /^(wrap|hAlign|hPos|vPos)\(/.test(x)), 'inline: no placement setter');
  assert.equal(c.elements.length, 4);
});

test('commandBody (placement): a missing or refusing setter leaves the document UNCHANGED', () => {
  const cases = [
    [{ SetWrappingStyle: undefined }, fop()], [{ SetVerPosition: undefined }, fop()], [{ SetHorAlign: undefined }, fop()], [{ SetHorPosition: undefined }, fop(PL_PAGE)],
    [{ SetWrappingStyle: () => false }, fop()], [{ SetHorAlign: () => false }, fop()], [{ SetHorPosition: () => false }, fop(PL_PAGE)], [{ SetVerPosition: () => false }, fop()],
    [{ SetVerPosition: () => { throw new Error('boom'); } }, fop()],
  ];
  for (const [over, o] of cases) {
    const s = scriptedFloating(over);
    assert.doesNotThrow(() => run(lowered(o).body, s.Api));
    assert.equal(s.elements.length, 3, JSON.stringify(Object.keys(over)));
    assert.ok(!s.calls.some(x => x.startsWith('AddElement')));
  }
  const ok = scriptedFloating({ SetHorPosition: undefined }); run(lowered(fop()).body, ok.Api);
  assert.equal(ok.elements.length, 4, 'an align placement does not need SetHorPosition');
});

test('readShapesInto (placement): the ToJSON projection -- a floating shape reads drawingType anchor + wrap + positions; a throwing ToJSON is NAMED on every placement field', () => {
  const j = { drawingType: 'anchor', wrapType: 'square', behindDoc: false, positionH: { align: 'right', relativeFrom: 'margin', percent: false }, positionV: { relativeFrom: 'paragraph', posOffset: 5 * MM, percent: false } };
  const [r] = readOf(docOf([para(ANCHOR), para('', [shapeObj(TEXT, W, H, 'shape', j)])]));
  assert.deepEqual({ drawingType: r.drawingType, wrapType: r.wrapType, behindDoc: r.behindDoc, hRel: r.hRel, hAlign: r.hAlign, hOffset: r.hOffset, vRel: r.vRel, vAlign: r.vAlign, vOffset: r.vOffset }, SQUARE_READ);
  const throwing = { ...shapeObj(TEXT, W, H), ToJSON: () => { throw new Error('no json'); } };
  const [t] = readOf(docOf([para(ANCHOR), para('', [throwing])]));
  for (const k of Object.keys(SQUARE_READ)) assert.equal(t[k], 'UNOBSERVABLE:no json', k);
  assert.equal(t.text, TEXT, 'a ToJSON failure does not hide the text');
});
