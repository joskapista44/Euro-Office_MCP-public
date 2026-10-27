'use strict';

/**
 * GH #224 (DOCX P2 READ): the DEDICATED `read_styles` operation on the product `docx_read` surface -- the style
 * DEFINITIONS the canonical inventory already carries (inventory.cjs `document.styles`, keyed by NAME). A plain
 * re-shape: no new driver read, no extra query option. Decision: docs/GH224-STYLE-DEFINITIONS-READ-DECISION.md.
 *
 * Offline only: test/fakes/fake-docx-driver.cjs already models the real driver's styles read (its styleOf unit
 * conversions, script.styleUnobservableFields, script.stylesUnreadable and doc.customStyleNames -- the
 * GetAllStyles()-blindness for a custom style, W-10 run3), so every case below runs the REAL read-tools pipeline.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool, readOnce, FORMATS } = require('../../src/server/read-tools.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

const PARAS = [{ text: 'Body', props: {} }];
const DOC = {
  styles: {
    Normal: { text: { bold: false, size: 22, color: '000000' }, paragraph: { align: 'left', indentLeft: 0, spacingAfter: 160 } },
    'Heading 1': { text: { bold: true, size: 32, color: '2F5496' }, paragraph: { align: 'inherited', indentLeft: 'inherited', spacingAfter: 'inherited' } },
    Quote: { text: { bold: 'inherited', size: 'inherited', color: 'auto' }, paragraph: { align: 'center', indentLeft: 720, spacingAfter: 'inherited' } },
  },
};

function harness({ doc = DOC, script = {}, wrap } = {}) {
  const s = createStore(PARAS, { doc });
  const b = fakeBroker('docx');
  const driver = () => { const d = createFakeDocxDriver(s, script); return wrap ? wrap(d) : d; };
  return { store: s, tool: createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver: driver }) };
}
async function readStyles(opts) {
  const r = await harness(opts).tool.handler({ fileId: FILE, operation: 'read_styles' });
  assert.equal(r.ok, true, JSON.stringify(r));
  return r.result;
}
const byName = (res, name) => res.styles.find(x => x.name === name);

test('GH #224: built-in styles come back by name with identity + text + paragraph fields, units converted by the driver (half-points -> pt, twips -> pt), "inherited" kept as the word', async () => {
  const res = await readStyles();
  assert.deepEqual(res.styles.map(x => x.name), ['Normal', 'Heading 1', 'Quote']);
  assert.deepEqual(byName(res, 'Normal'), { name: 'Normal', identity: { name: 'Normal', basedOn: null, next: null, link: null, styleType: 'paragraphStyle', basedOnName: null, nextName: null, linkName: null },
    text: { bold: false, size: 11, color: '000000', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' }, paragraph: { align: 'left', indentLeft: 0, spacingAfter: 8, spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' } }); // GH #292: the further keys, unset here
  assert.deepEqual(byName(res, 'Heading 1').text, { bold: true, size: 16, color: '2F5496', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' });
  assert.deepEqual(byName(res, 'Heading 1').paragraph, { align: 'inherited', indentLeft: 'inherited', spacingAfter: 'inherited', spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' });
  assert.deepEqual(byName(res, 'Quote').paragraph, { align: 'center', indentLeft: 36, spacingAfter: 'inherited', spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' });
  assert.equal(byName(res, 'Quote').text.color, 'auto');
});

test('GH #224: the result carries exactly {styles, scope} and per style exactly {name, identity, text, paragraph} -- no effective-appearance field is invented', async () => {
  const res = await readStyles();
  assert.deepEqual(Object.keys(res).sort(), ['scope', 'styles']);
  for (const st of res.styles) assert.deepEqual(Object.keys(st).sort(), ['identity', 'name', 'paragraph', 'text']);
  assert.match(res.scope, /GetAllStyles/);
  assert.match(res.scope, /not effective rendered appearance/);
  assert.equal(/effective|rendered/i.test(JSON.stringify(res.styles)), false, 'no per-style field claims an effective/rendered value');
});

test('GH #224: a custom style GetAllStyles() is blind to is NOT listed (the named boundary) -- and the listed ones are unaffected', async () => {
  const doc = { ...DOC, styles: { ...DOC.styles, MyCustom: { text: { bold: true, size: 24, color: 'FF0000' }, paragraph: { align: 'left', indentLeft: 0, spacingAfter: 0 } } }, customStyleNames: ['MyCustom'] };
  const res = await readStyles({ doc });
  assert.equal(byName(res, 'MyCustom'), undefined, 'the enumeration does not list it; the result makes no completeness claim (scope says so)');
  assert.deepEqual(res.styles.map(x => x.name), ['Normal', 'Heading 1', 'Quote']);
});

test('GH #224: ONE unreadable getter is a named per-field UNOBSERVABLE; the other fields of the SAME style stay measured', async () => {
  const res = await readStyles({ script: { styleUnobservableFields: ['text.color'] } });
  const n = byName(res, 'Normal');
  assert.equal(n.text.color, 'UNOBSERVABLE:no-such-getter');
  assert.equal(n.text.size, 11, 'size on the same style is still measured');
  assert.equal(n.paragraph.align, 'left');
});

test('GH #224: the whole style map unread (stylesUnreadable) -> styles:null (a gap), NEVER an empty list that would read as "no styles"; a measured-empty map is [] and distinct', async () => {
  const unread = await readStyles({ script: { stylesUnreadable: true } });
  assert.equal(unread.styles, null);
  const empty = await readStyles({ doc: { styles: {} } });
  assert.deepEqual(empty.styles, []);
});

test('GH #224: a listed name whose definition could not be read back is a NAMED gap entry, not dropped and not a guessed definition', async () => {
  const wrap = d => ({ ...d, async readLive(h, req) { const r = await d.readLive(h, req); r.payload.document.styles.Quote = null; return r; } });
  const res = await readStyles({ wrap });
  assert.deepEqual(byName(res, 'Quote'), { name: 'Quote', gap: 'unread' });
  assert.equal(byName(res, 'Normal').identity.name, 'Normal', 'the other styles are untouched');
});

test('GH #224: an unreadable ToJSON on one style (driver value passed through) keeps identity:null and the per-field UNOBSERVABLE strings -- no invented identity', async () => {
  const u = 'UNOBSERVABLE:no-readable-ToJSON';
  const wrap = d => ({ ...d, async readLive(h, req) { const r = await d.readLive(h, req); r.payload.document.styles.Quote = { name: 'Quote', identity: null, text: { bold: u, size: u, color: u }, paragraph: { align: u, indentLeft: u, spacingAfter: u } }; return r; } });
  const q = byName(await readStyles({ wrap }), 'Quote');
  assert.equal(q.identity, null);
  assert.deepEqual(q.text, { bold: u, size: u, color: u });
});

test('GH #224: READ only -- no dispatch, no save, the store is byte-identical after the read; read_formatting does not leak the style map', async () => {
  const h = harness();
  const before = JSON.stringify(h.store);
  const r = await h.tool.handler({ fileId: FILE, operation: 'read_styles' });
  assert.equal(r.ok, true);
  assert.equal(JSON.stringify(h.store), before);
  assert.equal(h.store.saves, 0);
  const f = await h.tool.handler({ fileId: FILE, operation: 'read_formatting' });
  assert.equal(f.ok, true);
  assert.equal(/"identity"|"basedOn"|Heading 1/.test(JSON.stringify(f.result)), false);
});

test('GH #224: reopen -- two independent reads of the same store agree', async () => {
  const h = harness();
  const a = await h.tool.handler({ fileId: FILE, operation: 'read_styles' });
  const b = await h.tool.handler({ fileId: FILE, operation: 'read_styles' });
  assert.deepEqual(a.result, b.result);
});

test('GH #224: the real readOnce pipeline projects a hand-built payload (no fake driver in the path) -- the same names/fields, no reshaping beyond the inventory', async () => {
  const driver = { async readLive() { return { ok: true, payload: { html: '<p>x</p>', paragraphs: [{ text: 'x' }], document: { styles: { Normal: { name: 'Normal', identity: { name: 'Normal', basedOn: null, styleType: 'paragraph' }, text: { bold: false, size: 11, color: '000000' }, paragraph: { align: 'left', indentLeft: 0, spacingAfter: 8 } } } } } }; } };
  const r = await readOnce(FORMATS.docx, driver, { fileId: FILE, operation: 'read_styles' }, 'docx');
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.styles.map(x => x.name), ['Normal']);
});
