'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * M11.5d.1 (owner: M11.5 continuous mandate, owner channel, internal tracker) -- OFFLINE acceptance of the word /2
 * CANDIDATE block (documentserver-patches/9.3.4.60/euro-office-bridge-word-v1/rev2-candidate/bridge-block.js): the
 * B18-WRITE Bridge capability docx.paragraph.numbering.remove/v1. The ACTUAL candidate text runs against a fake document
 * that records every call. PARITY: the bridge targets the paragraph exactly as the product's B18 path does
 * (src/adapters/docx/families/list.cjs commandBody: Search MatchCase, GetText({Numbering:false}) === text, one hit, the
 * live level re-checked) and removes through the SAME editor routine, CParagraph.RemoveNumPr().
 */

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-word-v1');
const read = p => fs.readFileSync(p, 'utf8');
const V1 = read(path.join(DIR, 'bridge-block.js'));
const V2 = read(path.join(DIR, 'rev2-candidate', 'bridge-block.js'));
const LIST = read(path.join(__dirname, '..', '..', 'src', 'adapters', 'docx', 'families', 'list.cjs'));
// the product's former B18 path, recorded from the last source that had it (a50bb07; M11.5d.4 moved it behind the bridge)
const PRE_CUTOVER = "const removeNumbering = () => { if (paragraph.Paragraph && typeof paragraph.Paragraph.RemoveNumPr === 'function') paragraph.Paragraph.RemoveNumPr(); };";
const bridgeOf = (src, Api) => new Function('Api', `${src.replace(/,\s*$/, '')}; return Api.GetEuroOfficeBridge();`)(Api);
const R = 'docx.paragraph.numbering.remove/v1';

/** fake document: paragraphs { text, num: { level, numId } | null }; every RemoveNumPr recorded. */
function fakeDoc(paras, { noSearch = false, searchThrows = false, noRemove = false, removeThrows = false, removeNoEffect = false, noGetNumbering = false } = {}) {
  const calls = [];
  const P = paras.map(x => ({ ...x, num: x.num ? { ...x.num } : null }));
  const api = p => ({
    ...(noGetNumbering ? {} : { GetNumbering: () => (p.num ? { GetLevelIndex: () => p.num.level } : null) }),
    Paragraph: { Pr: { NumPr: p.num ? { NumId: p.num.numId } : null },
      ...(noRemove ? {} : { RemoveNumPr: () => { calls.push(['RemoveNumPr', p.text]); if (removeThrows) throw new Error('boom'); if (!removeNoEffect) p.num = null; } }) },
  });
  const doc = {
    ...(noSearch ? {} : { Search: (t, o) => { calls.push(['Search', t, o && o.MatchCase]); if (searchThrows) throw new Error('x');
      return P.filter(p => p.text.includes(t)).map(p => ({ GetText: opts => { assert.deepEqual(opts, { Numbering: false }); return p.text; }, GetParagraph: i => (i === 0 ? api(p) : null) })); } }),
  };
  return { Api: { GetDocument: () => doc }, calls, P, removes: () => calls.filter(c => c[0] === 'RemoveNumPr') };
}
const BASE = [{ text: 'A elem', num: { level: 0, numId: 7 } }, { text: 'A elem ketto', num: { level: 1, numId: 7 } }, { text: 'sima', num: null }];
const write = (env, args, id = R) => bridgeOf(V2, env.Api).write(id, args);

test('/2 manifest: patchVersion /2, /1 capabilities byte-identical, exactly one new WRITE entry', () => {
  const m1 = bridgeOf(V1, {}).manifest(); const m2 = bridgeOf(V2, {}).manifest();
  assert.equal(m1.patchVersion, 'euro-office-bridge-word-v1/9.3.4.60/1');
  assert.equal(m2.patchVersion, 'euro-office-bridge-word-v1/9.3.4.60/2');
  assert.equal(m2.capabilities.length, m1.capabilities.length + 1);
  assert.deepEqual(m2.capabilities.slice(0, m1.capabilities.length), m1.capabilities);
  const w = m2.capabilities.filter(c => c.id === R);
  assert.equal(w.length, 1);
  assert.deepEqual([w[0].version, w[0].kind, w[0].format, w[0].mutation], [1, 'WRITE', 'DOCX', 'TARGETED']);
  for (const k of ['bridgeVersion', 'documentServerVersion', 'documentServerBuild']) assert.equal(m2[k], m1[k]);
});

test('/2 = /1 + three inserted segments + the patchVersion bump (every /1 implementation byte-identical)', () => {
  let t = V2.replace('patchVersion: "euro-office-bridge-word-v1/9.3.4.60/2"', 'patchVersion: "euro-office-bridge-word-v1/9.3.4.60/1"');
  const cut = (s, startMark, endMark) => { const i = s.indexOf(startMark); const j = s.indexOf(endMark, i); assert.ok(i > 0 && j > i, startMark); return s.slice(0, i) + s.slice(j); };
  t = cut(t, '{ id: "docx.paragraph.numbering.remove/v1"', '], };');
  t = cut(t, 'var EO_BRIDGE_WRITES = ', 'return { manifest: function');
  t = cut(t, 'write: function (capabilityId, args)', '}; },');
  assert.equal(t, V1);
});

test('class separation: read() refuses the WRITE id; write() refuses other ids; /1 has no write()', () => {
  const env = fakeDoc(BASE); const b = bridgeOf(V2, env.Api); const args = { text: 'A elem', level: 0 };
  assert.deepEqual(b.read(R, args), { ok: false, reason: 'capability-not-advertised' });
  for (const id of ['docx.paragraph.numbering.remove/v2', 'docx.numbering.levelFormat.read/v1', 'nope']) assert.deepEqual(b.write(id, args), { ok: false, reason: 'capability-not-advertised' }, id);
  assert.deepEqual(env.removes(), []);
  assert.equal(typeof bridgeOf(V1, env.Api).write, 'undefined', '/1: the B18 WRITE path is absent, detectably');
});

test('PARITY: the product B18 target rule and routine -- exact text, MatchCase, Numbering:false, one hit, live level, RemoveNumPr', () => {
  assert.match(LIST, /doc\.Search\(m\.targetText, \{ MatchCase: true \}\)/);
  assert.match(LIST, /String\(r\.GetText\(textOpts\) \|\| ''\) === m\.targetText/);
  assert.match(PRE_CUTOVER, /paragraph\.Paragraph\.RemoveNumPr\(\)/, 'the former product routine this capability moved behind the bridge');
  const code = LIST.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, ''); // the history in the comments stays
  assert.doesNotMatch(code, /RemoveNumPr/, 'M11.5d.4: the product no longer calls it directly');
  assert.match(LIST, /__eoB18Remove\(m\.targetText, m\.beforeLevel\)/, 'the product sends the SAME target text and live level to the bridge');
  const env = fakeDoc(BASE);
  assert.deepEqual(write(env, { text: 'A elem ketto', level: 1 }), { ok: true, value: { applied: ['numbering-removed'], before: { level: 1, numId: '7' } } });
  assert.deepEqual(env.removes(), [['RemoveNumPr', 'A elem ketto']]);
  assert.deepEqual(env.calls.filter(c => c[0] === 'Search'), [['Search', 'A elem ketto', true]]);
  assert.deepEqual(env.P.map(p => p.num && p.num.level), [0, null, null], 'the other list item untouched');
  // the substring hit 'A elem' also matches 'A elem ketto' in the search; the exact-text filter keeps ONE
  const env2 = fakeDoc(BASE);
  assert.equal(write(env2, { text: 'A elem', level: 0 }).ok, true);
  assert.deepEqual(env2.removes(), [['RemoveNumPr', 'A elem']]);
});

test('refusals BEFORE the routine (zero removals): arguments, target, numbering, level, missing API', () => {
  const dup = [{ text: 'X', num: { level: 0, numId: 1 } }, { text: 'X', num: { level: 0, numId: 1 } }];
  const cases = [
    [{}, BASE, null, 'bad-args'], [{ level: 0 }, BASE, null, 'bad-args'], [{ text: '', level: 0 }, BASE, null, 'bad-args'], [{ text: 1, level: 0 }, BASE, null, 'bad-args'],
    [{ text: 'A elem', level: 0, at: 1 }, BASE, null, 'unsupported-property'],
    [{ text: 'A elem' }, BASE, null, 'invalid-level'], [{ text: 'A elem', level: 9 }, BASE, null, 'invalid-level'], [{ text: 'A elem', level: -1 }, BASE, null, 'invalid-level'],
    [{ text: 'A elem', level: 0.5 }, BASE, null, 'invalid-level'], [{ text: 'A elem', level: '0' }, BASE, null, 'invalid-level'],
    [{ text: 'nincs', level: 0 }, BASE, null, 'target-not-found'],
    [{ text: 'X', level: 0 }, dup, null, 'target-ambiguous'],
    [{ text: 'sima', level: 0 }, BASE, null, 'not-numbered'],
    [{ text: 'A elem', level: 2 }, BASE, null, 'level-mismatch'],
    [{ text: 'A elem', level: 0 }, BASE, { noRemove: true }, 'RemoveNumPr-missing'],
    [{ text: 'A elem', level: 0 }, BASE, { noGetNumbering: true }, 'no-GetNumbering'],
    [{ text: 'A elem', level: 0 }, BASE, { noSearch: true }, 'document-unavailable'],
    [{ text: 'A elem', level: 0 }, BASE, { searchThrows: true }, 'search-threw'],
  ];
  for (const [args, paras, opts, reason] of cases) {
    const env = fakeDoc(paras, opts || {});
    assert.equal(write(env, args).reason, reason, JSON.stringify([args, opts]));
    assert.deepEqual(env.removes(), [], JSON.stringify(args));
  }
  assert.equal(write(fakeDoc(BASE), { text: 'A elem', level: 2 }).level, 0, 'level-mismatch reports the live level');
  assert.equal(write(fakeDoc(dup), { text: 'X', level: 0 }).count, 2);
});

test('after the routine: throw / no effect -> ok:false, exactly one call, never retried; a repeat is not-numbered (no silent success)', () => {
  for (const [opts, reason] of [[{ removeThrows: true }, 'remove-threw'], [{ removeNoEffect: true }, 'remove-not-observed']]) {
    const env = fakeDoc(BASE, opts);
    const r = write(env, { text: 'A elem', level: 0 });
    assert.deepEqual([r.ok, r.reason], [false, reason], reason);
    assert.equal(env.removes().length, 1, `${reason}: exactly one call, no retry`);
  }
  const env = fakeDoc(BASE); const b = bridgeOf(V2, env.Api);
  assert.equal(b.write(R, { text: 'A elem', level: 0 }).ok, true);
  assert.deepEqual(b.write(R, { text: 'A elem', level: 0 }), { ok: false, reason: 'not-numbered' });
  assert.equal(env.removes().length, 1);
});

test('internal-threw: an exception escaping a capability is caught at the write() boundary', () => {
  assert.deepEqual(bridgeOf(V2, { GetDocument: () => { throw new Error('x'); } }).write(R, { text: 'A', level: 0 }), { ok: false, reason: 'document-unavailable' });
  const Api = { GetDocument: () => ({ Search: () => [{ GetText: () => 'A', GetParagraph: () => { throw new Error('y'); } }] }) };
  assert.deepEqual(bridgeOf(V2, Api).write(R, { text: 'A', level: 0 }), { ok: false, reason: 'paragraph-unreadable' });
  const Bad = { GetDocument: () => ({ Search: () => [{ GetText: () => 'A', GetParagraph: () => ({ GetNumbering: () => { throw new Error('z'); } }) }] }) };
  assert.deepEqual(bridgeOf(V2, Bad).write(R, { text: 'A', level: 0 }), { ok: false, reason: 'internal-threw' });
});

test('the build is deterministic and pinned: bridge-block.js = new.txt + separator; old.txt = the canonical /1 block', () => {
  const C = path.join(DIR, 'rev2-candidate');
  assert.equal(V2, read(path.join(C, 'new.txt')) + ',');
  assert.equal(read(path.join(C, 'old.txt')), V1.replace(/,\s*$/, ''));
  const man = JSON.parse(read(path.join(C, 'manifest.json')));
  assert.equal(man.original_sha256, '8362e3a9fe07173e7ce9ece91e73d0829bdd7bd6ecf2b13a28330ae0e08b4699', 'the input is the served /1');
  assert.equal(man.operations.length, 1);
  assert.equal(man.operations[0].old, read(path.join(C, 'old.txt')));
  assert.equal(man.operations[0].new, read(path.join(C, 'new.txt')));
});
