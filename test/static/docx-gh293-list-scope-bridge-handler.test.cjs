'use strict';

/**
 * GH #293 L7 -- the word /5 Bridge capabilities docx.numbering.listMembers.read/v1 and docx.numbering.listDetach.write/v1
 * (documentserver-patches/9.3.5.0/euro-office-bridge-word-v1/rev5-candidate, built by build-rev5.py). Held here, offline:
 *  - /5 is /4 plus exactly those two capabilities and their shared helper: removing the two manifest entries, the two handlers and
 *    the helper from /5, and putting the /4 patchVersion back, gives the /4 block byte for byte (every /4 capability unchanged);
 *  - the contract, run in a vm against a fake editor model that holds what the LIVE probes measured on DS 9.3.5
 *    (docs/evidence/gh293-l7-detach-probe-run2-clone-20261022.json, gh293-l7-internals-probe-run1-20261022.json): the numbering
 *    collection sees a table-cell item the public read does not; GetAllNumsByAbstractNum takes the abstract OBJECT; a null override
 *    slot = no startOverride; CNum.Copy() is a new w:num on a new abstract.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-word-v1');
const READ = 'docx.numbering.listMembers.read/v1';
const WRITE = 'docx.numbering.listDetach.write/v1';
const V4 = fs.readFileSync(path.join(DIR, 'rev4-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V5 = fs.readFileSync(path.join(DIR, 'rev5-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');

/** the index of the `}` closing the `{` at `open`; strings, regex-free code assumed except the one known literal, escapes skipped */
function closingBrace(src, open) {
  let depth = 0; let quote = null;
  for (let k = open; k < src.length; k += 1) {
    const c = src[k];
    if (quote) { if (c === '\\') k += 1; else if (c === quote) quote = null; continue; }
    if (c === '"' || c === '\'' || c === '`') quote = c;
    else if (c === '{') depth += 1;
    else if (c === '}') { depth -= 1; if (depth === 0) return k; }
  }
  return -1;
}
const entryOf = (src, cap) => { const e = src.indexOf(`{ id: "${cap}"`); assert.ok(e > 0, `manifest entry ${cap}`); return src.slice(e, closingBrace(src, e) + 1) + ', '; };
const handlerOf = (src, cap) => { const h = src.indexOf(`"${cap}": function (args) {`); assert.ok(h > 0, `handler ${cap}`); return src.slice(h, closingBrace(src, src.indexOf('{', h + cap.length + 13)) + 1) + ', '; };

test('/5 = /4 + the two capabilities and their helper: without them, and with the /4 patchVersion, /5 IS the /4 block', () => {
  const h = V5.indexOf('function eoListMembers(args) {'); assert.ok(h > 0, 'helper');
  const helper = V5.slice(h, closingBrace(V5, V5.indexOf('{', h)) + 1) + ' ';
  const parts = [entryOf(V5, READ), entryOf(V5, WRITE), handlerOf(V5, READ), handlerOf(V5, WRITE), helper];
  for (const p of parts) assert.equal(V5.split(p).length, 2, 'each part exactly once');
  let back = V5; for (const p of parts) back = back.replace(p, '');
  back = back.replace('word-v1/9.3.5.0/5"', 'word-v1/9.3.5.0/4"');
  assert.equal(back, V4);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev5-candidate', 'old.txt'), 'utf8'), V4);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev5-candidate', 'new.txt'), 'utf8'), V5);
  const man = JSON.parse(fs.readFileSync(path.join(DIR, 'rev5-candidate', 'manifest.json'), 'utf8'));
  assert.equal(man.operations[0].old, V4); assert.equal(man.operations[0].new, V5);
  assert.equal(man.original_sha256, JSON.parse(fs.readFileSync(path.join(DIR, 'rev4-candidate', 'manifest.json'), 'utf8')).patched_sha256);
});

/**
 * A fake DS 9.3.5 editor: paragraphs [text, numId, lvl, pos] (pos = document order; a table-cell item is just another paragraph to the
 * numbering collection), nums { id: { abs, starts: { lvl: n } } }. Search finds a paragraph by its exact text; ApiParagraph.GetText
 * answers the text + '\r\n' (the paragraph separator).
 */
function fakeEditor(paragraphs, nums, { sortable = true, copyThrows = false, setIgnored = false } = {}) {
  let seq = 900;
  const lvlOf = (num, l) => ({ GetFormat: () => 1, GetLvlText: () => [{ Type: 2, Value: l }, { Type: 1, Value: '.' }], GetStart: () => 1, Suff: 0, Jc: 0 });
  const numObj = {};
  const mkNum = (id, abs, starts) => {
    const n = { id, abs, LvlOverride: Array.from({ length: 9 }, (_, l) => (starts[l] !== undefined ? { GetStartOverride: () => starts[l] } : null)),
      GetId() { return this.id; }, GetAbstractNumId() { return this.abs; }, GetLvl(l) { return lvlOf(this, l); },
      Copy() { if (copyThrows) throw new Error('boom'); seq += 2; return mkNum(`${seq}_2`, `${seq}_1`, starts); } };
    numObj[id] = n; return n;
  };
  for (const [id, d] of Object.entries(nums)) mkNum(id, d.abs, d.starts || {});
  const absObj = id => ({ id, GetId() { return this.id; } });
  const inner = paragraphs.map(([text, numId, lvl, pos]) => ({ text, pos, numPr: numId === null ? null : { NumId: numId, Lvl: lvl },
    GetNumPr() { return this.numPr; }, SetNumPr(id, l) { if (!setIgnored) this.numPr = { NumId: id, Lvl: l }; } }));
  const apiPara = ip => ({ Paragraph: ip, GetNumbering: () => (ip.numPr ? { GetLevelIndex: () => ip.numPr.Lvl } : null) });
  const N = { GetNum: id => numObj[id] || null, GetAbstractNum: id => absObj(id),
    GetAllNumsByAbstractNum: a => Object.values(numObj).filter(n => n.abs === a.GetId()) };
  const ld = { GetNumbering: () => N, GetAllParagraphsByNumbering: np => inner.filter(p => p.numPr && p.numPr.NumId === np.NumId) };
  const doc = { Document: ld, Search: t => inner.filter(p => p.text === t).map(p => ({ GetText: () => p.text, GetParagraph: () => apiPara(p) })) };
  const ctx = {
    Api: { GetDocument: () => doc },
    AscWord: { NumPr: function NumPr(id, l) { this.NumId = id; this.Lvl = l; }, ...(sortable ? { sortByDocumentPosition: a => a.sort((x, y) => x.pos - y.pos) } : {}) },
    ApiParagraph: function ApiParagraph(q) { this.GetText = () => `${q.text}\r\n`; },
  };
  const bridge = vm.runInNewContext(`(function(){ ${V5.replace('Api.GetEuroOfficeBridge = ', 'var f = ')}; return f(); })`, ctx)();
  const plain = r => JSON.parse(JSON.stringify(r));
  return { bridge: { manifest: () => plain(bridge.manifest()), read: (c, x) => plain(bridge.read(c, x)), write: (c, x) => plain(bridge.write(c, x)) },
    numOf: t => inner.find(p => p.text === t).numPr, absOf: t => { const np = inner.find(p => p.text === t).numPr; return np && numObj[np.NumId] ? numObj[np.NumId].abs : null; } };
}

// Word's shape: list A (no override) first, list B (a level-0 startOverride) after it, both on abstract 662
const WORD = [[['GH293 A1', '663', 0, 1], ['GH293 A2', '663', 0, 2], ['GH293 B1', '664', 0, 3], ['GH293 B2', '664', 0, 4], ['GH293 end', null, null, 5]],
  { 663: { abs: '662' }, 664: { abs: '662', starts: { 0: 1 } } }];
const members = (...t) => t.map(x => ({ text: x, level: 0 }));

test('listMembers: the whole list, the other list on the definition with its override and position, the order flags', () => {
  const { bridge } = fakeEditor(...WORD);
  const cap = bridge.manifest().capabilities.find(c => c.id === READ);
  assert.equal(cap.kind, 'READ'); assert.equal(cap.mutation, 'NO');
  const r = bridge.read(READ, { text: 'GH293 A1', level: 0 });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.members, members('GH293 A1', 'GH293 A2'));
  assert.deepEqual(r.value.starts, []);
  assert.equal(r.value.others.length, 1);
  assert.deepEqual({ ...r.value.others[0], sig: undefined }, { numId: '664', starts: [0], members: members('GH293 B1', 'GH293 B2'), sig: undefined, before: false, after: true });
  assert.deepEqual(r.value.order, { targetFirst: true, contiguous: true });
  assert.equal(r.value.sig.length, 9);
  assert.deepEqual(bridge.read(READ, { text: 'GH293 end', level: 0 }), { ok: true, value: { present: false } });
  assert.equal(bridge.read(READ, { text: 'GH293 A1', level: 1 }).reason, 'level-mismatch');
  assert.equal(bridge.read(READ, { text: 'GH293 A1', level: 0, extra: 1 }).reason, 'unsupported-property');
  assert.equal(bridge.read(READ, { text: 'No Such', level: 0 }).reason, 'target-not-found');
  assert.equal(bridge.read(WRITE, { text: 'GH293 A1', level: 0, members: [], others: 0 }).reason, 'capability-not-advertised', 'the detach is served only through write()');
});

test('listMembers sees a member the public read does not (an item in a table cell): it is IN the affected set', () => {
  const paras = [...WORD[0], ['GH293 T1', '663', 0, 6]];
  const r = fakeEditor(paras, WORD[1]).bridge.read(READ, { text: 'GH293 A1', level: 0 });
  assert.deepEqual(r.value.members, members('GH293 A1', 'GH293 A2', 'GH293 T1'));
  assert.deepEqual(r.value.order, { targetFirst: true, contiguous: false }, 'list B sits between A2 and T1');
});

test('listDetach: the whole list moves to a NEW w:num on a NEW definition; the other list is untouched', () => {
  const e = fakeEditor(...WORD);
  const r = e.bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 1 });
  assert.equal(r.ok, true); assert.deepEqual(r.value.applied, ['list-detached']); assert.equal(r.value.moved, 2); assert.deepEqual(r.value.before, { numId: '663' });
  assert.equal(e.numOf('GH293 A1').NumId, e.numOf('GH293 A2').NumId); assert.notEqual(e.numOf('GH293 A1').NumId, '663');
  assert.notEqual(e.absOf('GH293 A1'), '662');
  assert.equal(e.numOf('GH293 B1').NumId, '664'); assert.equal(e.absOf('GH293 B1'), '662');
  // retry: the list is alone on its definition now
  assert.equal(e.bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 0 }).reason, 'nothing-to-detach');
});

test('listDetach: a list AFTER another one detaches when it carries its own startOverride (list B)', () => {
  const e = fakeEditor(...WORD);
  const r = e.bridge.write(WRITE, { text: 'GH293 B1', level: 0, members: members('GH293 B1', 'GH293 B2'), others: 1 });
  assert.equal(r.ok, true); assert.equal(e.numOf('GH293 A1').NumId, '663');
});

test('listDetach refuses -- nothing moved -- when the caller saw a different set, or the count could flow between the lists', () => {
  // the caller did not see the table-cell member
  let e = fakeEditor([...WORD[0], ['GH293 T1', '663', 0, 6]], WORD[1]);
  let r = e.bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 1 });
  assert.equal(r.reason, 'members-mismatch'); assert.equal(e.numOf('GH293 A1').NumId, '663'); assert.equal(e.numOf('GH293 T1').NumId, '663');
  assert.equal(fakeEditor(...WORD).bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 2 }).reason, 'others-mismatch');
  // continuation: list B has no startOverride -> it continues A's count; detaching EITHER list changes labels
  const cont = { 663: { abs: '662' }, 664: { abs: '662' } };
  e = fakeEditor(WORD[0], cont);
  assert.equal(e.bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 1 }).reason, 'not-isolatable');
  assert.equal(e.bridge.write(WRITE, { text: 'GH293 B1', level: 0, members: members('GH293 B1', 'GH293 B2'), others: 1 }).reason, 'not-isolatable');
  assert.equal(e.numOf('GH293 A1').NumId, '663'); assert.equal(e.numOf('GH293 B1').NumId, '664');
  // non-contiguous: B between A2 and T1
  e = fakeEditor([...WORD[0], ['GH293 T1', '663', 0, 6]], WORD[1]);
  assert.equal(e.bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2', 'GH293 T1'), others: 1 }).reason, 'not-isolatable');
  // a list wrapped by another one (items before AND after)
  e = fakeEditor([['GH293 B1', '664', 0, 1], ['GH293 A1', '663', 0, 2], ['GH293 A2', '663', 0, 3], ['GH293 B2', '664', 0, 4]], { 663: { abs: '662', starts: { 0: 1 } }, 664: { abs: '662', starts: { 0: 1 } } });
  assert.equal(e.bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 1 }).reason, 'not-isolatable');
  // the editor cannot order the items -> no detach
  assert.equal(fakeEditor(...WORD, { sortable: false }).bridge.write(WRITE, { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 1 }).reason, 'order-unreadable');
  // a standalone list: nothing shares the definition
  assert.equal(fakeEditor([['GH293 X1', '7', 0, 1]], { 7: { abs: '6' } }).bridge.write(WRITE, { text: 'GH293 X1', level: 0, members: members('GH293 X1'), others: 0 }).reason, 'nothing-to-detach');
});

test('a Copy() that throws is copy-threw, a SetNumPr that does not take is write-not-observed -- never a success; bad / extra args refused', () => {
  const args = { text: 'GH293 A1', level: 0, members: members('GH293 A1', 'GH293 A2'), others: 1 };
  assert.equal(fakeEditor(...WORD, { copyThrows: true }).bridge.write(WRITE, args).reason, 'copy-threw');
  assert.equal(fakeEditor(...WORD, { setIgnored: true }).bridge.write(WRITE, args).reason, 'write-not-observed');
  const { bridge } = fakeEditor(...WORD);
  assert.equal(bridge.write(WRITE, { text: 'GH293 A1', level: 0 }).reason, 'bad-args');
  assert.equal(bridge.write(WRITE, { ...args, extra: 1 }).reason, 'unsupported-property');
  const cap = bridge.manifest().capabilities.find(c => c.id === WRITE);
  assert.equal(cap.kind, 'WRITE'); assert.equal(cap.mutation, 'TARGETED');
  assert.equal(bridge.manifest().patchVersion, 'euro-office-bridge-word-v1/9.3.5.0/5');
});
