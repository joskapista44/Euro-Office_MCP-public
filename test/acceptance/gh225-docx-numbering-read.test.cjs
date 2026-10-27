'use strict';

/**
 * GH #225 (DOCX P1 READ): the DEDICATED `read_numbering` operation on the product `docx_read`
 * surface, PARTIAL -> ACCEPTED. A plain field-select over the canonical inventory's own
 * per-item `numbering` (inventory.cjs:161-183, the SAME shape families/numbering.cjs's own
 * write-side verification already reads via families/list.cjs's listState) -- no new driver
 * reader; the ONLY new wiring is read-tools.cjs's FORMATS.docx.read requesting the extra
 * `numberingFormat: true` ToJSON() pass, and ONLY for this one operation.
 * Decision: docs/GH225-NUMBERING-READ-DECISION.md.
 *
 * Offline only: test/fakes/fake-docx-driver.cjs's own `numberingOf`/`levelFormatsOf`/
 * `formatOfNumbering` (fake-docx-driver.cjs:969-991) already simulate every gap this card's
 * acceptance list asks for (script.numberingUnreadable / levelFormatsUnreadable /
 * numberingFormatUnreadable / numberingFormatFullyUnreadable) -- the SAME fixture
 * docx-rewrite-slice.test.cjs's and families/numbering's own offline tests already trust.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

function harness(paragraphs, { doc = {}, script = {} } = {}) {
  const s = createStore(paragraphs, { doc });
  const b = fakeBroker('docx');
  return createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(s, script) });
}

async function readNumbering(paragraphs, opts) {
  const r = await harness(paragraphs, opts).handler({ fileId: FILE, operation: 'read_numbering' });
  assert.equal(r.ok, true, JSON.stringify(r));
  return r.result.items;
}

const fmt = (numFmt, lvlText, suff, lvlJc, start) => ({ numFmt, lvlText, suff, lvlJc, ...(start === undefined ? {} : { start }) });

// ---- bullet + numbered + multilevel, present:true, every field measured -------------------------------------------

test('GH #225: bullet list, numbered list and a two-level numbered list, each fully measured -- type derived from the SAME block.ordered rule list.cjs itself uses', async () => {
  const paras = [
    { text: 'Heading', level: 1, props: {} },
    { text: 'Bullet A', props: {}, list: { type: 'bullet', level: 0, numId: 'n1' } },
    { text: 'Bullet B', props: {}, list: { type: 'bullet', level: 0, numId: 'n1' } },
    { text: 'Numbered A', props: {}, list: { type: 'numbered', level: 0, numId: 'n2' } },
    { text: 'Numbered A sub', props: {}, list: { type: 'numbered', level: 1, numId: 'n2' } },
    { text: 'Plain paragraph', props: {} },
  ];
  const doc = { numberingFormats: {
    n1: Object.assign(Array(9).fill(null), { 0: fmt('bullet', '-', 'tab', 'left', 1) }),
    n2: Object.assign(Array(9).fill(null), { 0: fmt('decimal', '%1.', 'tab', 'left', 1), 1: fmt('lowerLetter', '%2.', 'tab', 'left', 1) }),
  } };
  const items = await readNumbering(paras, { doc });

  assert.deepEqual(items.find(i => i.text === 'Heading'), { text: 'Heading', kind: 'heading', numbering: { present: false } });
  assert.deepEqual(items.find(i => i.text === 'Plain paragraph'), { text: 'Plain paragraph', kind: 'paragraph', numbering: { present: false } });

  const bA = items.find(i => i.text === 'Bullet A');
  assert.deepEqual(bA, { text: 'Bullet A', kind: 'list_item', numbering: {
    present: true, type: 'bullet', numId: 'n1', level: 0, levelFormats: [5, 5, 5, 5, 5, 5, 5, 5, 5],
    format: fmt('bullet', '-', 'tab', 'left', 1), // GH #293 N5: no abstractId -- session-local, never a read field
  } });
  const bB = items.find(i => i.text === 'Bullet B');
  assert.equal(bB.numbering.numId, bA.numbering.numId, 'SHARED_NUMID_OBSERVABLE: both items name the SAME numId');

  const nA = items.find(i => i.text === 'Numbered A');
  assert.deepEqual(nA.numbering, { present: true, type: 'numbered', numId: 'n2', level: 0, levelFormats: [13, 13, 13, 13, 13, 13, 13, 13, 13], format: fmt('decimal', '%1.', 'tab', 'left', 1) });
  const nSub = items.find(i => i.text === 'Numbered A sub');
  assert.deepEqual(nSub.numbering, { present: true, type: 'numbered', numId: 'n2', level: 1, levelFormats: [13, 13, 13, 13, 13, 13, 13, 13, 13], format: fmt('lowerLetter', '%2.', 'tab', 'left', 1) });
  // distinct numIds: n1 (bullet) and n2 (numbered) never share identity
  assert.notEqual(bA.numbering.numId, nA.numbering.numId);
});

// ---- the default-start-omitted rule (W07-7): absence means 1, a MEASURED value, never UNOBSERVABLE ----------------

test('GH #225: start OMITTED from the stored level -- reads as the measured integer 1, not a gap (the SAME driver rule docx-driver.cjs:268-275 implements)', async () => {
  const paras = [{ text: 'Item', props: {}, list: { type: 'numbered', level: 0, numId: 'n1' } }];
  const doc = { numberingFormats: { n1: Object.assign(Array(9).fill(null), { 0: { numFmt: 'decimal', lvlText: '%1.', suff: 'tab', lvlJc: 'left' } }) } };
  const [item] = await readNumbering(paras, { doc });
  assert.equal(item.numbering.format.start, 1);
  assert.equal(typeof item.numbering.format.start, 'number');
});

// ---- genuinely unobservable: the default-start convenience OFF, the key never arrives --------------------------

test('GH #225: with the default-start convenience OFF, an omitted start is named UNOBSERVABLE:start-not-set -- never guessed as 1', async () => {
  const paras = [{ text: 'Item', props: {}, list: { type: 'numbered', level: 0, numId: 'n1' } }];
  const doc = { numberingFormats: { n1: Object.assign(Array(9).fill(null), { 0: { numFmt: 'decimal', lvlText: '%1.', suff: 'tab', lvlJc: 'left' } }) } };
  const [item] = await readNumbering(paras, { doc, script: { numberingFormatFullyUnreadable: true } });
  assert.equal(item.numbering.format.start, 'UNOBSERVABLE:start-not-set');
  assert.equal(item.numbering.format.numFmt, 'decimal', 'the OTHER fields on the SAME item stay independently measured');
});

// ---- the WHOLE format object never read (ToJSON itself failed) -- every field names the SAME reason, level/numId untouched ----

test('GH #225: the ToJSON format read fails entirely -- the WHOLE format is one named gap (not five), but level/numId/type/levelFormats (a DIFFERENT gate) stay measured', async () => {
  const paras = [{ text: 'Item', props: {}, list: { type: 'bullet', level: 0, numId: 'n1' } }];
  const [item] = await readNumbering(paras, { script: { numberingFormatUnreadable: true } });
  // the whole ToJSON() call never ran (undefined, not a wrong-shaped object) -- ONE reason for the
  // whole axis, distinct from the per-field case above (where the call ran but ONE key was absent).
  assert.equal(item.numbering.format, 'UNOBSERVABLE:numbering-format-not-read');
  assert.equal(item.numbering.present, true);
  assert.equal(item.numbering.type, 'bullet');
  assert.equal(item.numbering.numId, 'n1');
  assert.deepEqual(item.numbering.levelFormats, [5, 5, 5, 5, 5, 5, 5, 5, 5], 'levelFormats is a SEPARATE Bridge/internal-gate read -- one gate failing never hides the other');
});

test('GH #225: levelFormats unreadable (its OWN gate) -- format and identity stay independently measured', async () => {
  const paras = [{ text: 'Item', props: {}, list: { type: 'numbered', level: 0, numId: 'n1' } }];
  const doc = { numberingFormats: { n1: Object.assign(Array(9).fill(null), { 0: fmt('decimal', '%1.', 'tab', 'left', 3) }) } };
  const [item] = await readNumbering(paras, { doc, script: { levelFormatsUnreadable: true } });
  assert.equal(item.numbering.levelFormats, null);
  assert.deepEqual(item.numbering.format, fmt('decimal', '%1.', 'tab', 'left', 3));
  assert.equal(item.numbering.numId, 'n1');
});

// ---- the whole axis unread (the bridge call itself never attached) -- named 'unread', never "not numbered" --------

test('GH #225: the numbering axis never attached at all (unread) -- distinct from a measured present:false (not numbered)', async () => {
  const paras = [
    { text: 'Item', props: {}, list: { type: 'bullet', level: 0, numId: 'n1' } },
    { text: 'Plain', props: {} },
  ];
  const items = await readNumbering(paras, { script: { numberingUnreadable: true } });
  assert.deepEqual(items.find(i => i.text === 'Item'), { text: 'Item', kind: 'list_item', numbering: null, numberingGap: 'unread' });
  assert.deepEqual(items.find(i => i.text === 'Plain'), { text: 'Plain', kind: 'paragraph', numbering: null, numberingGap: 'unread' });
});

// ---- RED: a non-list-item paragraph whose numbering inconsistently reports present:true is named, never trusted ----

/**
 * test/fakes/fake-docx-driver.cjs's OWN `numberingOf` cannot produce this shape (it derives
 * present:false unconditionally from the absence of `p.list` -- the SAME guard a correct driver
 * has, which is exactly why this state is normally unreachable). To exercise read-tools.cjs's
 * fail-closed handling of it anyway, this ONE test drives the REAL pipeline (readOnce ->
 * normalizeInventory -> the real read_numbering projection, read-tools.cjs's own exported
 * `readOnce`/`FORMATS`) with a hand-built driver that returns the inconsistent payload directly --
 * real code under test, a minimal fake transport, no reimplemented classification logic.
 */
test('RED: a plain paragraph whose own numbering read says present:true (a state a correct driver never produces) is named inconsistent, never passed through as a list item', async () => {
  const { readOnce, FORMATS } = require('../../src/server/read-tools.cjs');
  const html = '<p>Odd</p>';
  const driver = { async readLive() { return { ok: true, payload: { html, paragraphs: [{ text: 'Odd', numbering: { present: true, level: 0, numId: 'zz' } }], document: {} } }; } };
  const r = await readOnce(FORMATS.docx, driver, { fileId: FILE, operation: 'read_numbering' }, 'docx');
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result.items, [{ text: 'Odd', kind: 'paragraph', numbering: null, numberingGap: 'inconsistent' }]);
});

// ---- reopen: a fresh session re-reads the SAME, now-durable state -------------------------------------------------

test('GH #225: reopen -- a second, independent read on the same (unwritten) store sees the identical numbering state', async () => {
  const paras = [{ text: 'Item', props: {}, list: { type: 'numbered', level: 2, numId: 'n9' } }];
  const doc = { numberingFormats: { n9: Object.assign(Array(9).fill(null), { 2: fmt('upperRoman', '%3.', 'space', 'left', 1) }) } };
  const s = createStore(paras, { doc });
  const b = fakeBroker('docx');
  const tool = createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(s, {}) });
  const r1 = await tool.handler({ fileId: FILE, operation: 'read_numbering' });
  const r2 = await tool.handler({ fileId: FILE, operation: 'read_numbering' });
  assert.deepEqual(r1.result, r2.result, 'two independent sessions (fresh open each time) agree -- no write happened in between');
  assert.equal(r1.result.items[0].numbering.level, 2);
  assert.equal(r1.result.items[0].numbering.numId, 'n9');
});
