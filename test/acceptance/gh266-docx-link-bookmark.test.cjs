'use strict';

const test = require('node:test');
const { testFor } = require('../helpers/trim-skip.cjs'); // release 3 trim: the add_link / add_bookmark cases run per intent, as each is (re-)released
const testLink = testFor('add_link');
const testBookmark = testFor('add_bookmark');
const assert = require('node:assert/strict');

const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const has = i => W.DOCX_WRITE_ALLOWLIST.includes(i);
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

/**
 * GH #266 (DOCX P1 WRITE; the release itself is an OWNER DECISION, alongside #228..#265): add_link and add_bookmark served by docx_execute_task, whole unit or a `matchText` piece of it
 * (the format_text pattern). Design and the measured reason for every rule: docs/GH266-DOCX-LINK-BOOKMARK-DECISION.md. Offline only: the product handler runs the REAL docx adapter /
 * executor / lifecycle on test/fakes/fake-docx-driver.cjs. The LIVE acceptance is test/live/gh266-docx-link-bookmark-product.live.cjs (not run by the author).
 *
 * READ<->WRITE symmetry (card rule): a link reads back through the product read_formatting (the paragraph's `hyperlinks`: url, screen tip, displayed text) and a bookmark through read_bookmarks
 * (name + the text of its range) -- both exist, so no separate READ unit is needed; the READ-BACK tests below pin that the product READ returns exactly what the WRITE put there.
 */

const FILE = '1303626';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const PARENT = 'Alpha beta gamma delta.';
const OTHER = 'Second paragraph here.';
const URL1 = 'https://example.invalid/one';
const doc = () => [{ text: 'Heading', level: 1, props: PLAIN }, { text: PARENT, props: PLAIN }, { text: OTHER, props: PLAIN }];
const al = (extra = {}) => ({ intent: 'add_link', selector: { text: PARENT }, href: URL1, ...extra });
const ab = (extra = {}) => ({ intent: 'add_bookmark', selector: { text: PARENT }, name: 'g266_mark', ...extra });
const req = steps => ({ fileId: FILE, task: { steps } });
const cfg = { kind: 'BROKER', async getEditorConfig() { return { document: { key: 'KEY-1' } }; } };
const verdict = steps => D.validateDocxWriteRequest(req(steps));

function product(paragraphs = doc(), docExtra = {}) {
  const store = createStore(paragraphs, { doc: docExtra });
  const drivers = [];
  const createDriver = () => { const d = createFakeDocxDriver(store, { saveObserver: true }); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const tool = D.createDocxWriteTool({ configSource: cfg, gate: createSessionGate(), runtime, createDriver });
  const read = createReadTool('docx', { configSource: cfg, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(store, {}) });
  const readOp = async operation => { const r = await read.handler({ fileId: FILE, operation }); assert.equal(r.ok, true, JSON.stringify(r)); return r.result; };
  return { store, tool, readOp, writes: () => store.saves, opens: () => drivers.reduce((n, d) => n + d.calls.open, 0) };
}
const snapshot = s => JSON.stringify([s.durable.map(p => [p.text, p.links || null]), s.durableDoc.bookmarks]);
const linksOf = async (p, text) => {
  const r = await p.readOp('read_formatting');
  const para = (r.units || []).find(x => x.text === text);
  assert.ok(para, `read_formatting has no paragraph ${text}: ${JSON.stringify(r).slice(0, 300)}`);
  return para.hyperlinks;
};

// ---- the step schema (pre-dispatch) --------------------------------------------------------------------------------------------

test('SCHEMA: add_link / add_bookmark are in both allowlists and the published schema; http(s) and mailto links, matchText, tooltip and a step id are accepted', () => {
  for (const i of ['add_link', 'add_bookmark'].filter(has)) {
    assert.ok(W.DOCX_WRITE_ALLOWLIST.includes(i), i);
    assert.ok(T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes(i), i);
    assert.ok(D.inputSchema().properties.task.properties.steps.items.oneOf.some(s => s.properties.intent.const === i), i);
  }
  for (const step of [al(), al({ matchText: 'beta' }), al({ tooltip: 'tip', id: 'l1' }), al({ href: 'http://example.invalid' }), al({ href: 'mailto:someone@example.invalid' }), al({ href: 'https://example.invalid/a?b=c#d' }),
    ab(), ab({ matchText: 'gamma' }), ab({ name: '_x9', id: 'b1' }), ab({ name: 'a'.repeat(40) })].filter(st => has(st.intent))) {
    const r = verdict([step]);
    assert.equal(r.ok, true, `${JSON.stringify(step)} -> ${r.code} ${r.reason}`);
  }
});

test('SCHEMA: every other URL scheme, a relative / bare / upper-case / whitespace / control-character URL, a bad bookmark name, a blank text, an extra selector key and an unknown field are refused pre-dispatch', () => {
  const refused = [
    [al({ href: 'javascript:alert(1)' }), 'REFUSED_SHAPE'], [al({ href: 'file:///etc/passwd' }), 'REFUSED_SHAPE'], [al({ href: 'data:text/html,x' }), 'REFUSED_SHAPE'], [al({ href: 'ftp://example.invalid/x' }), 'REFUSED_SHAPE'],
    [al({ href: '/relative/path' }), 'REFUSED_SHAPE'], [al({ href: 'example.invalid' }), 'REFUSED_SHAPE'], [al({ href: 'www.example.invalid' }), 'REFUSED_SHAPE'], [al({ href: 'HTTPS://EXAMPLE.INVALID' }), 'REFUSED_SHAPE'],
    [al({ href: 'https://' }), 'REFUSED_SHAPE'], [al({ href: 'https:///x' }), 'REFUSED_SHAPE'], [al({ href: 'https://exa mple.invalid' }), 'REFUSED_SHAPE'], [al({ href: 'https://example.invalid/\u0001' }), 'REFUSED_SHAPE'],
    [al({ href: `https://example.invalid/${'a'.repeat(2048)}` }), 'REFUSED_SHAPE'], [al({ href: 'mailto:' }), 'REFUSED_SHAPE'], [al({ tooltip: '' }), 'REFUSED_SHAPE'], [al({ tooltip: 'x'.repeat(256) }), 'REFUSED_SHAPE'],
    [{ intent: 'add_link', selector: { text: PARENT } }, 'REFUSED_SHAPE'], [al({ url: URL1 }), 'ROUTE_OVERRIDE'], [al({ selector: { text: '   ' } }), 'REFUSED_SHAPE'], [al({ matchText: '   ' }), 'REFUSED_SHAPE'], [al({ matchText: '' }), 'REFUSED_SHAPE'],
    [al({ selector: { ref: 's0' } }), 'UNKNOWN_FIELD'], [al({ selector: { text: PARENT, contains: true } }), 'UNKNOWN_FIELD'], [al({ target: 'x' }), 'UNKNOWN_FIELD'], [al({ style: 'Hyperlink' }), 'UNKNOWN_FIELD'],
    [ab({ name: '1abc' }), 'REFUSED_SHAPE'], [ab({ name: 'has space' }), 'REFUSED_SHAPE'], [ab({ name: 'a'.repeat(41) }), 'REFUSED_SHAPE'], [ab({ name: '' }), 'REFUSED_SHAPE'], [ab({ name: 'ékezet' }), 'REFUSED_SHAPE'],
    [{ intent: 'add_bookmark', selector: { text: PARENT } }, 'REFUSED_SHAPE'], [ab({ href: URL1 }), 'UNKNOWN_FIELD'], [ab({ matchText: '  ' }), 'REFUSED_SHAPE'],
  ];
  for (const [step, code] of refused.filter(([st]) => has(st.intent))) {
    const v = verdict([step]);
    assert.equal(v.code, code, `${JSON.stringify(step).slice(0, 120)} -> ${v.code} ${v.reason}`);
    assert.equal(v.dispatched, false);
  }
});

test('SCHEMA: add_link / add_bookmark are SOLO -- next to any other step the task is refused pre-dispatch', () => {
  for (const solo of [al(), ab()].filter(st => has(st.intent))) {
    const v = verdict([solo, { intent: 'rewrite_paragraph', selector: { text: OTHER }, text: 'changed' }]);
    assert.equal(v.ok, false);
    assert.equal(v.dispatched, false);
  }
  assert.equal(verdict([al(), ab()]).ok, false);
});

// ---- WRITE + READ-BACK (whole unit, then matchText) ---------------------------------------------------------------------------

testLink('WRITE + READ-BACK: a whole-unit link applies once and read_formatting returns exactly the url / screen tip / displayed text; the retry is SATISFIED with no write', async () => {
  const p = product();
  const r = await p.tool.handler(req([al({ tooltip: 'tip' })]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(p.writes() >= 1, true);
  assert.deepEqual((await linksOf(p, PARENT)).map(l => [l.url, l.tip, l.text]), [[URL1, 'tip', PARENT]]);
  assert.equal((await linksOf(p, OTHER)).length, 0, 'another paragraph carries no link');
  const writes = p.writes();
  const again = await p.tool.handler(req([al({ tooltip: 'tip' })]));
  assert.equal(again.ok, true, JSON.stringify(again));
  assert.equal(p.writes(), writes, 'the retry writes nothing');
});

testLink('WRITE + READ-BACK: a matchText link covers ONLY that piece (displayed text = matchText), the rest of the paragraph and every other paragraph carry no link; a second word links independently', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([al({ matchText: 'beta', tooltip: 'one' })]))).ok, true);
  assert.deepEqual((await linksOf(p, PARENT)).map(l => [l.url, l.tip, l.text]), [[URL1, 'one', 'beta']]);
  { const rr = await p.tool.handler(req([al({ matchText: 'delta', href: 'mailto:someone@example.invalid', tooltip: 'two' })])); assert.equal(rr.ok, true, JSON.stringify(rr).slice(0, 900)); }
  assert.deepEqual((await linksOf(p, PARENT)).map(l => [l.url, l.text]), [[URL1, 'beta'], ['mailto:someone@example.invalid', 'delta']]);
  assert.equal((await linksOf(p, OTHER)).length, 0);
  const writes = p.writes();
  assert.equal((await p.tool.handler(req([al({ matchText: 'beta', tooltip: 'one' })]))).ok, true, 'the retry of the first link');
  assert.equal(p.writes(), writes);
});

testBookmark('WRITE + READ-BACK: a bookmark reads back through read_bookmarks with the text of its range -- the whole unit, or exactly matchText; the retry writes nothing', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([ab()]))).ok, true);
  assert.deepEqual((await p.readOp('read_bookmarks')).bookmarks, [{ name: 'g266_mark', text: PARENT }]);
  const q = product();
  assert.equal((await q.tool.handler(req([ab({ matchText: 'gamma', name: 'g266_range' })]))).ok, true);
  assert.deepEqual((await q.readOp('read_bookmarks')).bookmarks, [{ name: 'g266_range', text: 'gamma' }]);
  const writes = q.writes();
  assert.equal((await q.tool.handler(req([ab({ matchText: 'gamma', name: 'g266_range' })]))).ok, true);
  assert.equal(q.writes(), writes);
});

// ---- fail-closed: a named reason, nothing written ---------------------------------------------------------------------------------

test('FAIL-CLOSED: an absent / ambiguous matchText, a missing / ambiguous parent, a name that exists elsewhere, a different link on the same range -> not ok, a named reason, the document unchanged', async () => {
  const linked = (url, text) => doc().map(x => (x.text === PARENT ? { ...x, links: [{ url, tip: 'tip', text }] } : x));
  const cases = [
    [() => product(), al({ matchText: 'zzz' }), /match-not-found/],
    [() => product(), al({ matchText: 'a' }), /match-ambiguous/],
    [() => product(), ab({ matchText: 'zzz' }), /match-not-found/],
    [() => product(), ab({ matchText: 'a' }), /match-ambiguous/],
    [() => product(), al({ selector: { text: 'Nope.' } }), /target-not-found/],
    [() => product([...doc(), { text: PARENT, props: PLAIN }]), al(), /target-ambiguous/],
    [() => product(doc(), { bookmarks: [{ name: 'g266_mark', text: OTHER }] }), ab(), /bookmark-name-exists/],
    [() => product(linked('https://other.invalid/x', 'beta')), al({ matchText: 'beta', tooltip: 'tip' }), /target-already-linked/],
    [() => product(linked(URL1, PARENT)), al({ matchText: 'beta' }), /target-already-linked/],
  ];
  for (const [make, step, re] of cases.filter(([, st]) => has(st.intent))) {
    const p = make();
    const before = snapshot(p.store);
    const r = await p.tool.handler(req([step]));
    assert.equal(r.ok, false, JSON.stringify(step));
    assert.match(JSON.stringify(r), re, JSON.stringify(r).slice(0, 400));
    assert.equal(snapshot(p.store), before, `${JSON.stringify(step)} must not change the document`);
    assert.equal(p.writes(), 0);
  }
});

// ---- the gate (g266v1 ROUND4B): a bookmark must not enclose a hyperlink ------------------------------------------------------------------------------------------

testBookmark('GATE: add_bookmark on a unit that already holds a link (whole-unit or the piece it would enclose) is refused with the named reason, zero writes, the document unchanged -- and the rich READ still works afterwards', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([al({ matchText: 'beta', tooltip: 'tip' })]))).ok, true);
  const writes = p.writes();
  const before = snapshot(p.store);
  for (const step of [ab(), ab({ matchText: 'beta gamma' })]) {
    const r = await p.tool.handler(req([step]));
    assert.equal(r.ok, false, JSON.stringify(step));
    assert.match(JSON.stringify(r), /bookmark-target-holds-a-hyperlink/);
    assert.equal(p.writes(), writes);
    assert.equal(snapshot(p.store), before);
  }
  assert.equal((await linksOf(p, PARENT)).length, 1, 'read_formatting still returns the link');
  assert.deepEqual((await p.readOp('read_bookmarks')).bookmarks, [], 'read_bookmarks still answers');
});

testBookmark('GATE: a bookmark range clear of the link is written, and the reverse order (bookmark first, link over the same unit after) stays released', async () => {
  const p = product();
  assert.equal((await p.tool.handler(req([al({ matchText: 'beta', tooltip: 'tip' })]))).ok, true);
  assert.equal((await p.tool.handler(req([ab({ matchText: 'delta', name: 'g266_clear' })]))).ok, true, 'a range that does not enclose the link');
  assert.deepEqual((await p.readOp('read_bookmarks')).bookmarks, [{ name: 'g266_clear', text: 'delta' }]);
  const q = product();
  assert.equal((await q.tool.handler(req([ab()]))).ok, true);
  const r = await q.tool.handler(req([al({ tooltip: 'tip' })]));
  assert.equal(r.ok, true, JSON.stringify(r).slice(0, 400));
  assert.equal((await linksOf(q, PARENT)).length, 1);
  assert.deepEqual((await q.readOp('read_bookmarks')).bookmarks, [{ name: 'g266_mark', text: PARENT }]);
});
