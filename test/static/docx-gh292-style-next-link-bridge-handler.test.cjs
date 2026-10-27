'use strict';

/**
 * GH #292 S4 -- the word /4 Bridge WRITE docx.style.nextLink.write/v1 (documentserver-patches/9.3.5.0/euro-office-bridge-word-v1/
 * rev4-candidate, built by build-rev4.py). Two things are held here, offline:
 *  - /4 is /3 plus exactly that one capability: removing its manifest entry and its handler from /4, and putting the /3 patchVersion
 *    back, gives the shipped /3 block byte for byte (every /3 capability unchanged);
 *  - the handler's contract, run in a vm against a fake Api that models what the LIVE probes measured on DS 9.3.5
 *    (docs/evidence/gh292s4-next-link-internal-probe-20261022.json, gh292s4b-next-link-internal-probe-run2-20261022.json): the internal
 *    setters store ANY id they are given (an unknown id, a next on a character style) and doc.GetStyle(<unknown>) answers Normal --
 *    so the name / type / linked-elsewhere checks are the handler's own job.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-word-v1');
const CAP = 'docx.style.nextLink.write/v1';
const V3 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V4 = fs.readFileSync(path.join(DIR, 'rev4-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');

/** the index of the `}` closing the `{` at `open`; strings and escapes skipped */
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

test('/4 = /3 + the one capability: without its manifest entry and handler, and with the /3 patchVersion, /4 IS the shipped /3 block', () => {
  const e = V4.indexOf(`{ id: "${CAP}"`); assert.ok(e > 0, 'manifest entry');
  const entry = V4.slice(e, closingBrace(V4, e) + 1) + ', ';
  const h = V4.indexOf(`"${CAP}": function (args) {`); assert.ok(h > 0, 'handler');
  const handler = V4.slice(h, closingBrace(V4, V4.indexOf('{', h + CAP.length + 13)) + 1) + ', ';
  assert.equal(V4.split(entry).length, 2); assert.equal(V4.split(handler).length, 2);
  const back = V4.replace(entry, '').replace(handler, '').replace('word-v1/9.3.5.0/4"', 'word-v1/9.3.5.0/3"');
  assert.equal(back, V3);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev4-candidate', 'old.txt'), 'utf8'), V3);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev4-candidate', 'new.txt'), 'utf8'), V4);
  const man = JSON.parse(fs.readFileSync(path.join(DIR, 'rev4-candidate', 'manifest.json'), 'utf8'));
  assert.equal(man.operations[0].old, V3); assert.equal(man.operations[0].new, V4);
  assert.equal(man.original_sha256, JSON.parse(fs.readFileSync(path.join(DIR, '..', 'fresh-install.json'), 'utf8')).word.output_sha256);
});

/** a fake DS 9.3.5 document: setters store whatever they get (measured), GetStyle(<unknown>) answers Normal (measured, W-10 run4) */
function fakeDoc(defs, { throwOnSecondLink = false, setterIgnored = false } = {}) {
  let linkCalls = 0;
  const styles = defs.map(([name, type, id, next = null, link = null]) => {
    const inner = { id, next, link, GetId() { return this.id; }, Get_Next() { return this.next; }, Get_Link() { return this.link; },
      Set_Next(v) { if (!setterIgnored) this.next = v; },
      Set_Link(v) { linkCalls += 1; if (throwOnSecondLink && linkCalls === 2) throw new Error('boom'); if (!setterIgnored) this.link = v; } };
    return { GetName: () => name, GetType: () => type, Style: inner };
  });
  const normal = styles.find(s => s.GetName() === 'Normal');
  const doc = { GetStyle: n => styles.find(s => s.GetName() === n) || normal };
  const api = { GetDocument: () => doc };
  const bridge = vm.runInNewContext(`(function(){ var Api = this.Api; ${V4.replace('Api.GetEuroOfficeBridge = ', 'var f = ')}; return f(); })`, { Api: api }).call({ Api: api });
  // a vm-context result has the context's own prototypes: compare plain copies
  const plain = r => JSON.parse(JSON.stringify(r));
  return { bridge: { manifest: () => plain(bridge.manifest()), read: (c, x) => plain(bridge.read(c, x)), write: (c, x) => plain(bridge.write(c, x)) }, inner: n => styles.find(s => s.GetName() === n).Style };
}
const DEFS = [['Normal', 'paragraph', '886'], ['Body', 'paragraph', '1_1'], ['Body Char', 'run', '1_2'], ['Heading 1', 'paragraph', '827', '886', '837'], ['Heading 1 Char', 'run', '837', null, '827']];

test('next: paragraph -> paragraph is set and re-read; an equal value is no-change; a character style either side is a type-mismatch', () => {
  const { bridge, inner } = fakeDoc(DEFS);
  assert.deepEqual(bridge.manifest().capabilities.find(c => c.id === CAP).kind, 'WRITE');
  assert.deepEqual(bridge.write(CAP, { style: 'Body', field: 'next', target: 'Heading 1' }), { ok: true, value: { applied: ['next-set'], before: { value: null } } });
  assert.equal(inner('Body').next, '827');
  assert.equal(bridge.write(CAP, { style: 'Body', field: 'next', target: 'Heading 1' }).reason, 'no-change');
  assert.equal(bridge.write(CAP, { style: 'Body Char', field: 'next', target: 'Normal' }).reason, 'type-mismatch');
  assert.equal(bridge.write(CAP, { style: 'Body', field: 'next', target: 'Body Char' }).reason, 'type-mismatch');
  assert.equal(inner('Body Char').next, null, 'nothing written on a refusal');
});

test('link: BOTH sides are set; a style already linked elsewhere is refused, never re-pointed; an existing pair is no-change', () => {
  const { bridge, inner } = fakeDoc(DEFS);
  assert.deepEqual(bridge.write(CAP, { style: 'Body', field: 'link', target: 'Body Char' }), { ok: true, value: { applied: ['link-set'], before: { value: null } } });
  assert.equal(inner('Body').link, '1_2'); assert.equal(inner('Body Char').link, '1_1');
  assert.equal(bridge.write(CAP, { style: 'Body Char', field: 'link', target: 'Body' }).reason, 'no-change', 'the reverse direction of an existing pair');
  const r = bridge.write(CAP, { style: 'Body', field: 'link', target: 'Heading 1 Char' });
  assert.equal(r.reason, 'already-linked-elsewhere'); assert.equal(inner('Heading 1 Char').link, '827');
  assert.equal(bridge.write(CAP, { style: 'Normal', field: 'link', target: 'Heading 1' }).reason, 'type-mismatch', 'two paragraph styles');
});

test('names: an unknown style / target is not-found (never the Normal fallback); bad / extra / unknown-field args are refused before any read', () => {
  const { bridge, inner } = fakeDoc(DEFS);
  assert.equal(bridge.write(CAP, { style: 'No Such', field: 'next', target: 'Normal' }).reason, 'style-not-found');
  assert.equal(bridge.write(CAP, { style: 'Body', field: 'next', target: 'No Such' }).reason, 'target-not-found');
  assert.equal(inner('Body').next, null);
  assert.equal(bridge.write(CAP, { style: 'Body', field: 'next' }).reason, 'bad-args');
  assert.equal(bridge.write(CAP, { style: 'Body', field: 'next', target: 'Normal', extra: 1 }).reason, 'unsupported-property');
  assert.equal(bridge.write(CAP, { style: 'Body', field: 'basedOn', target: 'Normal' }).reason, 'unsupported-field');
  assert.equal(bridge.read(CAP, { style: 'Body', field: 'next', target: 'Normal' }).reason, 'capability-not-advertised', 'served only through write()');
});

test('a setter that does not take is write-not-observed; a throwing second side is set-threw, never a success', () => {
  assert.equal(fakeDoc(DEFS, { setterIgnored: true }).bridge.write(CAP, { style: 'Body', field: 'next', target: 'Normal' }).reason, 'write-not-observed');
  assert.equal(fakeDoc(DEFS, { throwOnSecondLink: true }).bridge.write(CAP, { style: 'Body', field: 'link', target: 'Body Char' }).reason, 'set-threw');
});
