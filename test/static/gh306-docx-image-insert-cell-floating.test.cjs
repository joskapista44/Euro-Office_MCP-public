'use strict';

/**
 * GH #306 part 4 (offline): insert_image FLOATING placement (body target, position 'after') and the TABLE-CELL target
 * (src/adapters/docx/families/image-insert.cjs). plan/classify/resolve/verify against fabricated `normalizeInventory`
 * observations; the callCommand bodies are LIVE-measured (test/live/docx-gh306-cell-floating-image-probe.live.cjs and the
 * product harness), here only their guards and setter order are asserted on the generated string.
 */
const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const ii = require('../../src/adapters/docx/families/image-insert.cjs');

const ANCHOR = 'GH306F anchor';
const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const W = 914400;
const H = 457200;
const MM = 36000;

const SQUARE_RIGHT = { wrap: 'square', horizontal: { relativeTo: 'margin', align: 'right' }, vertical: { relativeTo: 'paragraph', offsetMm: 0 } };
const BEHIND_PAGE = { wrap: 'behind', horizontal: { relativeTo: 'page', offsetMm: 25 }, vertical: { relativeTo: 'page', offsetMm: 60 } };

/** a driver image row carrying the ToJSON placement projection (docx-images-read.cjs) */
const floatRow = (paragraph, o = {}) => ({ index: 0, part: 'body', paragraphIndex: 2, drawingType: 'anchor', wrapType: 'square', behindDoc: false, hRel: 'margin', hAlign: 'right', hOffset: null, vRel: 'paragraph', vAlign: null, vOffset: 0, cx: W, cy: H, paragraph, ...o });
const inlineRow = (paragraph, o = {}) => ({ index: 0, part: 'body', paragraphIndex: 2, drawingType: 'inline', wrapType: 'through', behindDoc: false, hRel: 'column', hAlign: null, hOffset: 0, vRel: 'paragraph', vAlign: null, vOffset: 0, cx: W, cy: H, paragraph, ...o });

const obsOf = (html, images = [], tables) => ({ primary: normalizeInventory({ html, document: { images, ...(tables ? { tables } : {}) } }) });

const floatOp = (placement = SQUARE_RIGHT, over = {}) => {
  const p = ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H, placement, ...over }, 0);
  assert.equal(p.ok, true, p.detail);
  return p.operation;
};

// ---- floating: plan -------------------------------------------------------------------------------

test('floating plan: accepts the measured placement shapes (insert_shape\'s own planPlacement) and freezes the expected read-back', () => {
  assert.deepEqual({ ...floatOp().placement }, { wrap: 'square', wrapType: 'square', behindDoc: false, hRel: 'margin', hAlign: 'right', hOffsetEmu: null, vRel: 'paragraph', vOffsetEmu: 0 });
  assert.deepEqual({ ...floatOp(BEHIND_PAGE).placement }, { wrap: 'behind', wrapType: 'none', behindDoc: true, hRel: 'page', hAlign: null, hOffsetEmu: 25 * MM, vRel: 'page', vOffsetEmu: 60 * MM });
  assert.equal(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H }, 0).operation.placement, null, 'no placement -> inline');
});

test('floating plan: refuses an unmeasured wrap/position, placement with position \'end\', and placement on a header/footer or table-cell target', () => {
  const bad = pl => ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H, placement: pl }, 0).ok;
  assert.equal(bad({ ...SQUARE_RIGHT, wrap: 'tight' }), false, 'tight is not offered');
  assert.equal(bad({ ...SQUARE_RIGHT, horizontal: { relativeTo: 'page', align: 'left' } }), false);
  assert.equal(bad({ ...SQUARE_RIGHT, vertical: { relativeTo: 'margin', offsetMm: 0 } }), false);
  assert.equal(bad({ wrap: 'square' }), false);
  const end = ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H, position: 'end', placement: SQUARE_RIGHT }, 0);
  assert.equal(end.ok, false);
  assert.match(end.detail, /only supported with position 'after'/);
  const hf = ii.plan({ intent: 'insert_image', selector: { section: { index: 0, of: 1 }, part: 'header' }, source: PNG, widthEmu: W, heightEmu: H, placement: SQUARE_RIGHT }, 0);
  assert.equal(hf.ok, false);
  const cell = ii.plan({ intent: 'insert_image', selector: { tableCell: { text: 'c' } }, source: PNG, widthEmu: W, heightEmu: H, placement: SQUARE_RIGHT }, 0);
  assert.equal(cell.ok, false);
});

// ---- floating: classify / verify -------------------------------------------------------------------

test('floating classify: SATISFIED only by an image with exactly the requested wrap / behindDoc / position; an inline image never satisfies it, and vice versa', () => {
  const html = `<p>${ANCHOR}</p><p></p>`;
  assert.equal(ii.classify(obsOf(html, [floatRow('')]).primary, floatOp()).state, 'SATISFIED');
  assert.equal(ii.classify(obsOf(html, [inlineRow('')]).primary, floatOp()).state, 'APPLY', 'inline is not the requested floating image');
  assert.equal(ii.classify(obsOf(html, [floatRow('', { hAlign: 'left' })]).primary, floatOp()).state, 'APPLY', 'another alignment');
  assert.equal(ii.classify(obsOf(html, [floatRow('', { wrapType: 'top_and_bottom' })]).primary, floatOp()).state, 'APPLY', 'another wrap');
  assert.equal(ii.classify(obsOf(html, [floatRow('', { vOffset: 1 })]).primary, floatOp()).state, 'APPLY', 'another vertical offset');
  const behind = floatRow('', { wrapType: 'none', behindDoc: true, hRel: 'page', hAlign: null, hOffset: 25 * MM, vRel: 'page', vOffset: 60 * MM });
  assert.equal(ii.classify(obsOf(html, [behind]).primary, floatOp(BEHIND_PAGE)).state, 'SATISFIED');
  assert.equal(ii.classify(obsOf(html, [{ ...behind, behindDoc: false }]).primary, floatOp(BEHIND_PAGE)).state, 'APPLY', 'inFront is not behind');
  const inlineOp = ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H }, 0).operation;
  assert.equal(ii.classify(obsOf(html, [floatRow('')]).primary, inlineOp).state, 'APPLY', 'a floating image does not satisfy an inline request');
  assert.equal(ii.classify(obsOf(html, [inlineRow('')]).primary, inlineOp).state, 'SATISFIED');
});

test('floating: an UNOBSERVABLE placement field never matches (a failed ToJSON names every placement field)', () => {
  const html = `<p>${ANCHOR}</p><p></p>`;
  const failed = { index: 0, part: 'body', paragraphIndex: 2, json: 'UNOBSERVABLE:boom', paragraph: '' };
  assert.equal(ii.classify(obsOf(html, [failed]).primary, floatOp()).state, 'APPLY');
  const row = obsOf(html, [failed]).primary.document.images[0];
  for (const k of ['behindDoc', 'hRel', 'hAlign', 'hOffset', 'vRel', 'vAlign', 'vOffset']) assert.equal(row[k], 'UNOBSERVABLE:boom', k);
});

test('floating verifyPrimitive / wholeTaskChecks: ok only when the new empty paragraph carries the requested floating image', () => {
  const before = obsOf(`<p>${ANCHOR}</p>`);
  assert.equal(ii.verifyPrimitive(floatOp(), before, obsOf(`<p>${ANCHOR}</p><p></p>`, [floatRow('')])).ok, true);
  assert.equal(ii.verifyPrimitive(floatOp(), before, obsOf(`<p>${ANCHOR}</p><p></p>`, [inlineRow('')])).ok, false, 'landed inline instead');
  const checks = ii.wholeTaskChecks(floatOp(), obsOf(`<p>${ANCHOR}</p><p></p>`, [floatRow('')]));
  assert.equal(checks[0].ok, true);
  assert.match(checks[0].name, /floating, wrap square/);
});

test('floating lower: the command sets wrap, then the horizontal, then the vertical position on the image BEFORE AddDrawing, and refuses on a missing setter', () => {
  const l = ii.lower(floatOp(BEHIND_PAGE), { kind: 'insert_image_anchor', anchorText: ANCHOR });
  assert.deepEqual({ ...l.meta.placement }, { ...floatOp(BEHIND_PAGE).placement });
  const b = l.body;
  const at = s => b.indexOf(s);
  assert.ok(at('img.SetWrappingStyle(p.wrap)') > 0 && at('img.SetWrappingStyle(p.wrap)') < at('img.SetHorPosition(p.hRel, p.hOffsetEmu)'));
  assert.ok(at('img.SetVerPosition(p.vRel, p.vOffsetEmu)') < at('np.AddDrawing(img)'));
  assert.match(b, /typeof img\.SetWrappingStyle !== 'function'/);
  assert.doesNotThrow(() => new Function(b));
  const inline = ii.lower(ii.plan({ intent: 'insert_image', selector: { afterParagraph: ANCHOR }, source: PNG, widthEmu: W, heightEmu: H }, 0).operation, { kind: 'insert_image_anchor', anchorText: ANCHOR });
  assert.equal(inline.meta.placement, null);
});

// ---- table cell -----------------------------------------------------------------------------------

const TABLE_HTML = '<table><tr><td><p>c00</p></td><td><p>c01</p></td></tr><tr><td><p>c10</p></td><td><p>c11</p></td></tr></table>';
const cellOf = (text, o = {}) => ({ text: `${text}\t`, paragraphs: 1, simple: true, format: {}, paraFormats: [{ text: `${text}\t`, format: {} }], inline: [], ...o });
const tablesOf = (over = {}) => [{ rows: [[cellOf('c00'), cellOf('c01')], [cellOf('c10'), over.c11 || cellOf('c11')]] }];
const cellImage = (paragraph = 'c11', o = {}) => ({ index: 0, part: 'body', paragraphIndex: null, drawingType: 'inline', wrapType: 'through', cx: W, cy: H, paragraph, ...o });
const WITH_IMAGE = ['blipFill', 'image', 'paraDrawing'];
const cellObs = (images = [], over = {}, prefix = `<p>${ANCHOR}</p>`) => obsOf(`${prefix}${TABLE_HTML}`, images, tablesOf(over));
const cellOp = (tableCell = { text: 'c11' }) => {
  const p = ii.plan({ intent: 'insert_image', selector: { tableCell }, source: PNG, widthEmu: W, heightEmu: H }, 0);
  assert.equal(p.ok, true, p.detail);
  return p.operation;
};

test('cell plan: { tableCell: { text, tableAnchor? } }; refuses an empty text, extra keys, a position and a placement', () => {
  assert.equal(cellOp().cellText, 'c11');
  assert.equal(cellOp().tableAnchor, null);
  assert.equal(cellOp({ text: 'c11', tableAnchor: 'c00' }).tableAnchor, 'c00');
  const bad = (sel, extra = {}) => ii.plan({ intent: 'insert_image', selector: sel, source: PNG, widthEmu: W, heightEmu: H, ...extra }, 0).ok;
  assert.equal(bad({ tableCell: { text: '' } }), false);
  assert.equal(bad({ tableCell: { text: 'c11', row: 1 } }), false);
  assert.equal(bad({ tableCell: { text: 'c11', tableAnchor: '' } }), false);
  assert.equal(bad({ tableCell: 'c11' }), false);
  assert.equal(bad({ tableCell: { text: 'c11' }, afterParagraph: ANCHOR }), false, 'never combined with the body target');
  assert.equal(bad({ tableCell: { text: 'c11' } }, { position: 'end' }), false);
});

test('cell classify: APPLY on a plain cell; SATISFIED only when a matching inline image row sits in a cell paragraph with that text AND the cell reads an image marker', () => {
  assert.equal(ii.classify(cellObs().primary, cellOp()).state, 'APPLY');
  assert.equal(ii.classify(cellObs([cellImage()], { c11: cellOf('c11', { inline: WITH_IMAGE }) }).primary, cellOp()).state, 'SATISFIED');
  assert.equal(ii.classify(cellObs([cellImage()]).primary, cellOp()).state, 'APPLY', 'no image marker in the cell');
  assert.equal(ii.classify(cellObs([cellImage('c11', { paragraphIndex: 3 })], { c11: cellOf('c11', { inline: WITH_IMAGE }) }).primary, cellOp()).state, 'APPLY', 'a body paragraph image with that text is not in a cell');
  assert.equal(ii.classify(cellObs([cellImage('c11', { cx: 1 })], { c11: cellOf('c11', { inline: WITH_IMAGE }) }).primary, cellOp()).state, 'APPLY', 'another size');
});

test('cell classify: FAIL_CLOSED on a missing / ambiguous cell, a non-plain cell, an unreadable cell content or unreadable tables', () => {
  const reason = (obs, op = cellOp()) => ii.classify(obs.primary, op).evidence.reason;
  assert.equal(reason(cellObs(), cellOp({ text: 'nope' })), 'insert-image-table-cell-not-found');
  assert.equal(reason(obsOf(`${TABLE_HTML}`, [], [{ rows: [[cellOf('c11'), cellOf('c01')], [cellOf('c10'), cellOf('c11')]] }])), 'insert-image-table-cell-ambiguous');
  assert.equal(reason(cellObs([], { c11: cellOf('c11', { simple: false }) })), 'insert-image-table-cell-not-plain-paragraphs-of-runs');
  assert.equal(reason(cellObs([], { c11: cellOf('c11', { inline: null }) })), 'insert-image-table-cell-content-unreadable');
  assert.equal(reason(obsOf(`<p>${ANCHOR}</p>${TABLE_HTML}`)), 'tables-unreadable');
  assert.equal(reason(cellObs(), cellOp({ text: 'c11', tableAnchor: 'nope' })), 'table-anchor-not-found');
});

test('cell resolve / preDispatchCheck: the resolved target names the same cell text and table anchor', () => {
  const r = ii.resolve(cellObs().primary, { tableCell: { text: 'c11' } });
  assert.equal(r.ok, true);
  assert.equal(r.target.kind, 'insert_image_table_cell');
  assert.equal(r.target.at, '0:1:1');
  assert.deepEqual(ii.preDispatchCheck(cellOp(), r.target), { ok: true });
  assert.equal(ii.preDispatchCheck(cellOp({ text: 'c11', tableAnchor: 'c00' }), r.target).ok, false, 'a different anchor scope');
  assert.equal(ii.resolve(cellObs().primary, { tableCell: { text: 'nope' } }).ambiguity, 'INSERT_IMAGE_TABLE_CELL_NOT_FOUND');
});

test('cell verifyPrimitive: exactly one more matching row in the cell and the image marker; verifyPreservation: everything else unchanged', () => {
  const before = cellObs();
  const after = cellObs([cellImage()], { c11: cellOf('c11', { inline: WITH_IMAGE }) });
  assert.equal(ii.verifyPrimitive(cellOp(), before, after).ok, true);
  assert.equal(ii.verifyPrimitive(cellOp(), before, cellObs([cellImage()])).ok, false, 'no marker');
  assert.equal(ii.verifyPrimitive(cellOp(), before, cellObs([cellImage(), cellImage()], { c11: cellOf('c11', { inline: WITH_IMAGE }) })).ok, false, 'two new rows');
  assert.equal(ii.verifyPreservation(cellOp(), before, after).ok, true);
  const otherCellChanged = obsOf(`<p>${ANCHOR}</p>${TABLE_HTML}`, [cellImage()], [{ rows: [[cellOf('c00', { format: { bold: true } }), cellOf('c01')], [cellOf('c10'), cellOf('c11', { inline: WITH_IMAGE })]] }]);
  assert.equal(ii.verifyPreservation(cellOp(), before, otherCellChanged).ok, false, 'another cell changed');
  const targetFormatChanged = cellObs([cellImage()], { c11: cellOf('c11', { inline: WITH_IMAGE, format: { italic: true } }) });
  assert.equal(ii.verifyPreservation(cellOp(), before, targetFormatChanged).ok, false, 'the target cell changed beyond its inline list');
  const bodyChanged = cellObs([cellImage()], { c11: cellOf('c11', { inline: WITH_IMAGE }) }, `<p>${ANCHOR} edited</p>`);
  assert.equal(ii.verifyPreservation(cellOp(), before, bodyChanged).ok, false, 'a body block changed');
  const lost = cellObs([inlineRow('x')]);
  assert.equal(ii.verifyPreservation(cellOp(), lost, after).ok, false, 'an existing image was lost');
});

test('cell wholeTaskChecks / lower: one named check; the command appends to the cell content\'s last paragraph after the plain-runs guard', () => {
  const checks = ii.wholeTaskChecks(cellOp(), cellObs([cellImage()], { c11: cellOf('c11', { inline: WITH_IMAGE }) }));
  assert.equal(checks.length, 1);
  assert.equal(checks[0].ok, true);
  assert.match(checks[0].name, /table cell "c11"/);
  const r = ii.resolve(cellObs().primary, { tableCell: { text: 'c11', tableAnchor: 'c00' } });
  const l = ii.lower(cellOp({ text: 'c11', tableAnchor: 'c00' }), r.target);
  assert.equal(l.meta.action, 'insertCell');
  assert.equal(l.meta.tableAnchor, 'c00');
  const b = l.body;
  assert.ok(b.indexOf("ch.GetClassType() !== 'run'") < b.indexOf('last.AddDrawing(img)'));
  assert.match(b, /content\.GetElement\(n - 1\)/);
  assert.doesNotThrow(() => new Function(b));
});
