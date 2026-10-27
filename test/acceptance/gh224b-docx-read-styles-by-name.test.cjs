'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool, validateReadRequest, inputSchema } = require('../../src/server/read-tools.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

/**
 * GH #224b (DOCX P2 READ): read_styles gains an optional `names` argument -- a NAME-TARGETED read of exactly the requested styles through the driver's buildTargetedStyleQuery (doc.GetStyle(name)),
 * which sees a style created with CreateStyle that doc.GetAllStyles() is blind to (GH #224 CUSTOM_STYLE_LISTED=false on DS 9.3.5). Decision: docs/GH224B-READ-STYLES-BY-NAME-DECISION.md. Without `names`
 * the operation is UNCHANGED (the GetAllStyles enumeration, the custom style not listed). Offline only: test/fakes/fake-docx-driver.cjs models the GetAllStyles blindness (doc.customStyleNames: a custom
 * style is enumerated ONLY when the request names it) and the targeted query's `null` for a name GetStyle could not return; the LIVE acceptance is test/live/gh224b-read-styles-by-name.live.cjs (not run by the author).
 */

const PARAS = [{ text: 'Body', props: {} }];
const CUSTOM = { text: { bold: true, size: 32, color: '2255AA' }, paragraph: { align: 'center', indentLeft: 360, spacingAfter: 240 } }; // RAW units: 16pt / 18pt / 12pt
const DOC = {
  styles: {
    Normal: { text: { bold: false, size: 22, color: '000000' }, paragraph: { align: 'left', indentLeft: 0, spacingAfter: 160 } },
    'Heading 1': { text: { bold: true, size: 32, color: '2F5496' }, paragraph: { align: 'inherited', indentLeft: 'inherited', spacingAfter: 'inherited' } },
    'GH224b Custom': CUSTOM,
  },
  customStyleNames: ['GH224b Custom'],
};

function harness({ doc = DOC, script = {} } = {}) {
  const s = createStore(PARAS, { doc });
  const b = fakeBroker('docx');
  let reads = 0;
  const createDriver = () => { const d = createFakeDocxDriver(s, script); const orig = d.readLive.bind(d); d.readLive = async (h, r) => { reads += 1; d.lastRequest = r; return orig(h, r); }; return d; };
  return { store: s, reads: () => reads, tool: createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver }) };
}
async function readStyles(args = {}, opts) {
  const h = harness(opts);
  const r = await h.tool.handler({ fileId: FILE, operation: 'read_styles', ...args });
  return { r, h };
}
const ok = async (args, opts) => { const { r } = await readStyles(args, opts); assert.equal(r.ok, true, JSON.stringify(r)); return r.result; };
const check = args => validateReadRequest('docx', { fileId: FILE, operation: 'read_styles', ...args });

// ---- the argument (pre-driver) ------------------------------------------------------------------------------------------------

test('SCHEMA: `names` is offered in the published inputSchema, optional, an array of 1..32 distinct 1..128-character names', () => {
  const schema = inputSchema('docx');
  assert.equal(schema.properties.names.type, 'array');
  assert.equal(schema.properties.names.minItems, 1);
  assert.equal(schema.properties.names.maxItems, 32);
  assert.equal(schema.properties.names.uniqueItems, true);
  assert.deepEqual(schema.properties.names.items, { type: 'string', minLength: 1, maxLength: 128 });
  assert.deepEqual(schema.required, ['fileId', 'operation'], '`names` is not required');
  assert.equal(check({}).ok, true, 'no names: the old call shape still validates');
});

test('VALIDATION: a valid names list is accepted and carried; an empty list, a duplicate, a non-array, a non-string, an empty / over-long / control-character name, more than 32 names are INVALID_ARGS pre-driver', () => {
  assert.deepEqual(check({ names: ['Normal', 'Heading 1'] }).args, { names: ['Normal', 'Heading 1'] });
  assert.equal(check({ names: ['x'.repeat(128)] }).ok, true);
  assert.equal(check({ names: Array.from({ length: 32 }, (_, i) => `S${i}`) }).ok, true);
  for (const names of [[], ['A', 'A'], 'Normal', { 0: 'Normal' }, [1], [null], [''], ['x'.repeat(129)], ['bad\nname'], ['bad\u0000name'], Array.from({ length: 33 }, (_, i) => `S${i}`), null]) {
    const v = check({ names });
    assert.equal(v.ok, false, JSON.stringify(names));
    assert.equal(v.code, 'INVALID_ARGS', `${JSON.stringify(names)} -> ${v.code}`);
    assert.equal(v.dispatched, false);
  }
});

test('VALIDATION: `names` belongs to read_styles ONLY -- every other docx read operation still refuses it as an unknown field', () => {
  for (const operation of ['read_text', 'read_formatting', 'read_tables', 'read_numbering', 'read_page_setup']) {
    const v = validateReadRequest('docx', { fileId: FILE, operation, names: ['Normal'] });
    assert.equal(v.code, 'UNKNOWN_FIELD', operation);
  }
});

// ---- the read ---------------------------------------------------------------------------------------------------------------------

test('WITHOUT names the operation is UNCHANGED: the GetAllStyles enumeration, the custom style NOT listed, and the driver gets NO styleNames', async () => {
  const h = harness();
  const r = await h.tool.handler({ fileId: FILE, operation: 'read_styles' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.styles.map(x => x.name), ['Normal', 'Heading 1']);
  assert.match(r.result.scope, /GetAllStyles\(\) enumeration/);
});

test('WITH names: a CUSTOM style comes back by name with the six measured fields in POINTS (the very style the enumeration cannot see), and exactly the requested names, in REQUEST order', async () => {
  const res = await ok({ names: ['GH224b Custom'] });
  assert.deepEqual(res.styles.map(x => x.name), ['GH224b Custom']);
  assert.deepEqual(res.styles[0], { name: 'GH224b Custom', identity: { name: 'GH224b Custom', basedOn: null, next: null, link: null, styleType: 'paragraphStyle', basedOnName: null, nextName: null, linkName: null },
    text: { bold: true, size: 16, color: '2255AA', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' }, paragraph: { align: 'center', indentLeft: 18, spacingAfter: 12, spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' } }); // GH #292: the further keys, unset
  assert.deepEqual(Object.keys(res).sort(), ['scope', 'styles']);
  assert.match(res.scope, /name-targeted/);
  const mixed = await ok({ names: ['Heading 1', 'GH224b Custom', 'Normal'] });
  assert.deepEqual(mixed.styles.map(x => x.name), ['Heading 1', 'GH224b Custom', 'Normal'], 'the request order, not the enumeration order');
});

test('WITH names: built-in styles read the SAME as in the enumeration (units, "inherited" kept as the word) -- the built-in read is unchanged', async () => {
  const all = await ok({});
  const named = await ok({ names: ['Normal', 'Heading 1'] });
  for (const st of named.styles) assert.deepEqual(st, all.styles.find(x => x.name === st.name), st.name);
});

test('WITH names: a name that does not exist is { name, gap: "not-found" } -- named, in its place in the order, never dropped and never a guessed definition; the others are unaffected', async () => {
  const res = await ok({ names: ['GH224b Custom', 'GH224b NoSuchStyle', 'Normal'] });
  assert.deepEqual(res.styles.map(x => x.name), ['GH224b Custom', 'GH224b NoSuchStyle', 'Normal']);
  assert.deepEqual(res.styles[1], { name: 'GH224b NoSuchStyle', gap: 'not-found' });
  assert.equal(res.styles[0].text.size, 16);
  assert.equal(res.styles[2].text.size, 11);
  assert.match(res.scope, /"not-found" = GetStyle\(name\) returned no style of that name/);
});

test('WITH names: the driver gets the names as styleNames ONLY for this operation (the other docx reads keep paying nothing), and exactly one live read happens', async () => {
  const seen = [];
  const mk = () => {
    const s = createStore(PARAS, { doc: DOC });
    const b = fakeBroker('docx');
    const createDriver = () => { const d = createFakeDocxDriver(s, {}); const orig = d.readLive.bind(d); d.readLive = async (h, r) => { seen.push(r); return orig(h, r); }; return d; };
    return createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver });
  };
  const r = await mk().handler({ fileId: FILE, operation: 'read_styles', names: ['GH224b Custom'] });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(seen.length, 1);
  assert.deepEqual(seen[0].styleNames, ['GH224b Custom']);
  seen.length = 0;
  assert.equal((await mk().handler({ fileId: FILE, operation: 'read_styles' })).ok, true);
  assert.equal(seen[0].styleNames, undefined, 'no names: no styleNames in the driver request');
  seen.length = 0;
  assert.equal((await mk().handler({ fileId: FILE, operation: 'read_text' })).ok, true);
  assert.equal(seen[0].styleNames, undefined, 'another operation never asks for the targeted style read');
});

test('WITH names: one unreadable getter is still a per-field UNOBSERVABLE on the named style (the driver\'s own value passes through); an unreadable style map is { styles: null } -- a gap, not "not found"', async () => {
  const res = await ok({ names: ['GH224b Custom'] }, { script: { styleUnobservableFields: ['text.color'] } });
  assert.equal(res.styles[0].text.color, 'UNOBSERVABLE:no-such-getter');
  assert.equal(res.styles[0].text.size, 16, 'the other fields of the SAME style stay measured');
  const gone = await ok({ names: ['GH224b Custom'] }, { script: { stylesUnreadable: true } });
  assert.equal(gone.styles, null);
});

test('PRE-DRIVER refusals never open a driver: a bad names list is refused before any read', async () => {
  const h = harness();
  const r = await h.tool.handler({ fileId: FILE, operation: 'read_styles', names: ['A', 'A'] });
  assert.equal(r.ok, false);
  assert.equal(r.code, 'INVALID_ARGS');
  assert.equal(h.reads(), 0);
});
