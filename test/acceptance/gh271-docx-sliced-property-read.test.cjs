'use strict';

/**
 * GH #271 -- the sliced DOCX property read (docx-driver.cjs slicedPropertyRead): a LARGE document is read in
 * time-budgeted callCommand slices so the editor's websocket heartbeat runs between them (LIVE g271c: one 41-175 s
 * call dropped the socket and lost the write). Offline, against an in-memory Word Api with a deterministic clock:
 *   - the default body is byte-identical to what it was (small documents keep the single call);
 *   - the merged sliced read equals the single-call read, field for field;
 *   - RED: the document changing between slices, or a slice that does not advance, ends the read NAMED, never stitched.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');

// an in-memory Word Api: `paras` top-level paragraphs, `tables` 2x2 tables between them; unknown Get*/Is* answer null
function wordApi(state) {
  const getters = o => new Proxy(o, { get: (t, k) => (k in t ? t[k] : (typeof k === 'string' && /^(Get|Is)/.test(k) ? () => null : undefined)) });
  const pr = () => getters({});
  const rng = () => getters({ GetTextPr: pr });
  const runOf = text => getters({ GetClassType: () => 'run', GetText: () => text, GetRange: rng });
  const paraOf = text => getters({ GetClassType: () => 'paragraph', GetText: () => text, GetElementsCount: () => 1, GetElement: () => runOf(text), GetRange: rng, GetParaPr: pr, GetTextPr: pr });
  const tableOf = tag => {
    const cell = (r, c) => getters({ GetContent: () => getters({ GetElementsCount: () => 1, GetElement: () => paraOf(`${tag} r${r}c${c}`), GetText: () => `${tag} r${r}c${c}` }) });
    return getters({ GetClassType: () => 'table', GetRowsCount: () => 2, GetRow: r => getters({ GetCellsCount: () => 2, GetCell: c => cell(r, c) }) });
  };
  const elements = () => state.items.map(x => (x.table ? tableOf(x.table) : paraOf(x.text)));
  const doc = getters({
    GetElementsCount: () => state.items.length, GetElement: i => elements()[i] || null,
    GetText: () => state.items.map(x => x.table || x.text).join('\n'),
    Search: text => (state.items.filter(x => x.text === text).length ? [getters({ GetText: () => text, GetTextPr: pr })] : []),
    GetAllTables: () => [], GetSections: () => [], GetAllBookmarksNames: () => [], GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [], GetAllStyles: () => [], GetAllImages: () => [], GetAllComments: () => [],
  });
  return { GetDocument: () => doc };
}
const doc = (paras, tables) => {
  const items = [];
  for (let i = 0; i < paras; i += 1) { items.push({ text: `Qs paragraph ${i}` }); if (i % 5 === 4 && tables > 0) { items.push({ table: `QsT${tables}` }); tables -= 1; } }
  return { items };
};

// a deterministic clock: every Date.now() call advances 1 ms, so a small budget cuts slices at predictable points
function withTickingClock(f) {
  const real = Date.now; let t = 1e12;
  Date.now = () => (t += 1);
  return f().finally(() => { Date.now = real; });
}
const runner = state => async body => {
  try { return { ok: true, outcome: 'command-completed', result: new Function('Api', body)(wordApi(state)) }; }
  catch (err) { return { ok: false, outcome: 'call-threw', error: String(err) }; }
};

test('the default body is the same string; a slice request is validated', () => {
  assert.equal(docx.PROPERTY_QUERY, docx.propertyQuery());
  assert.equal(docx.propertyQuery().includes('__slice'), false);
  assert.throws(() => docx.propertyQuery({ slice: { part: 'bogus', from: 0, budgetMs: 10 } }), /invalid slice/);
  assert.throws(() => docx.propertyQuery({ slice: { part: 'paragraphs', from: -1, budgetMs: 10 } }), /invalid slice/);
});

test('the sliced read equals the single-call read on a document that needs several slices of both parts', () => withTickingClock(async () => {
  const state = doc(40, 6);
  const single = await runner(state)(docx.propertyQuery());
  assert.equal(single.ok, true);
  assert.equal(typeof single.result.error, 'undefined', single.result.error);
  let calls = 0;
  const counting = async body => { calls += 1; return runner(state)(body); };
  const sliced = await docx.slicedPropertyRead(counting, {}, 25);
  assert.equal(typeof sliced.result.error, 'undefined', sliced.result.error);
  assert.ok(sliced.slices > 2, `expected several slices, got ${sliced.slices}`);
  assert.equal(calls, sliced.slices);
  assert.equal(single.result.paragraphs.length, 40);
  assert.equal(single.result.document.tables.length, 6);
  assert.deepEqual(sliced.result, single.result);
}));

test('a read that fits one budget is one slice per part', () => withTickingClock(async () => {
  const state = doc(3, 1);
  const sliced = await docx.slicedPropertyRead(runner(state), {}, 1e9);
  assert.equal(sliced.slices, 2);
  assert.deepEqual(sliced.result, (await runner(state)(docx.propertyQuery())).result);
}));

test('RED: the document changing between slices ends the read NAMED, never a stitched result', () => withTickingClock(async () => {
  const state = doc(40, 6);
  let calls = 0;
  const mutating = async body => { calls += 1; if (calls === 2) state.items.push({ text: 'Qs inserted meanwhile' }); return runner(state)(body); };
  const r = await docx.slicedPropertyRead(mutating, {}, 25);
  assert.match(r.result.error, /the document changed between slices/);
  assert.match(docx.richOutcome(r).richUnobservable, /^UNOBSERVABLE:property-query: sliced read: the document changed/);
}));

test('RED: a slice that does not advance, a failed slice and a missing fingerprint are each named', async () => {
  const stuck = async () => ({ ok: true, outcome: 'command-completed', result: { paragraphs: [], searchFailures: [], next: 0, fp: 'x' } });
  assert.match((await docx.slicedPropertyRead(stuck, {}, 10)).result.error, /paragraphs slice 0 did not advance/);
  const failed = { ok: false, outcome: 'command-callback-timeout' };
  assert.deepEqual(await docx.slicedPropertyRead(async () => failed, {}, 10), failed);
  const noFp = async () => ({ ok: true, outcome: 'command-completed', result: { paragraphs: [], searchFailures: [], next: null } });
  assert.match((await docx.slicedPropertyRead(noFp, {}, 10)).result.error, /reported no fingerprint/);
});

test('the slicing threshold counts the HTML projection\'s paragraph tags', () => {
  assert.equal(docx.paragraphTagsOf('<p>a</p><p class="x">b</p><pre>c</pre><td><p>d</p></td>'), 3);
  assert.equal(docx.paragraphTagsOf(''), 0);
  assert.ok(docx.PROPERTY_SLICE_MIN_PARAGRAPH_TAGS > 0 && docx.PROPERTY_SLICE_BUDGET_MS > 0);
});
