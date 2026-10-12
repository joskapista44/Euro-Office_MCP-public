'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * M11.5e.1 (owner: M11.5 continuous mandate, owner channel, internal tracker; design an internal branch
 * #ref, elfogadva ket korrekcioval) -- OFFLINE acceptance of the word /3 CANDIDATE block
 * (documentserver-patches/9.3.4.60/euro-office-bridge-word-v1/rev3-candidate/bridge-block.js): the B21-READ
 * Bridge capability docx.section.variantState.read/v1. The ACTUAL candidate text runs against a fake document
 * that records every call. PARITY: the capability reads the SAME two internal getters, from the SAME
 * `.Section` property, as the product's B21 path (src/drivers/eurooffice/docx-driver.cjs propertyQuery's
 * `flag(sec, name)` helper) and the write-time re-check (src/adapters/docx/families/header-footer.cjs
 * commandBody).
 */

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-word-v1');
const read = p => fs.readFileSync(p, 'utf8');
const V2 = read(path.join(DIR, 'rev2-candidate', 'bridge-block.js'));
const V3 = read(path.join(DIR, 'rev3-candidate', 'bridge-block.js'));
const DRIVER = read(path.join(__dirname, '..', '..', 'src', 'drivers', 'eurooffice', 'docx-driver.cjs'));
const HEADERFOOTER = read(path.join(__dirname, '..', '..', 'src', 'adapters', 'docx', 'families', 'header-footer.cjs'));
const bridgeOf = (src, Api) => new Function('Api', `${src.replace(/,\s*$/, '')}; return Api.GetEuroOfficeBridge();`)(Api);
const R = 'docx.section.variantState.read/v1';

/** fake document: sections [{ titlePage, evenAndOdd } | opts], each turned into an ApiSection-shaped wrapper
 * whose .Section exposes Get_TitlePage()/IsEvenAndOdd() the way the product's internal getters do. */
function fakeDoc(sections, { noGetSections = false, getSectionsThrows = false, noDocument = false } = {}) {
  const calls = [];
  const wrap = s => {
    if (s === null || s === undefined) return null;
    if (s.noSection) return {};
    const inner = {};
    if (!s.noTitlePageReader) inner.Get_TitlePage = () => { calls.push(['Get_TitlePage']); if (s.titlePageThrows) throw new Error('tp-boom'); return s.titlePage; };
    if (!s.noEvenOddReader) inner.IsEvenAndOdd = () => { calls.push(['IsEvenAndOdd']); if (s.evenAndOddThrows) throw new Error('eo-boom'); return s.evenAndOdd; };
    return { Section: inner };
  };
  const secs = sections.map(wrap);
  const doc = noDocument ? null : {
    ...(noGetSections ? {} : { GetSections: () => { if (getSectionsThrows) throw new Error('sections-boom'); return secs; } }),
  };
  return { Api: { GetDocument: () => { if (noDocument) throw new Error('no-doc'); return doc; } }, calls };
}
const read_ = (env, args, id = R) => bridgeOf(V3, env.Api).read(id, args);

test('/3 manifest: patchVersion /3, /2 capabilities byte-identical, exactly one new READ entry', () => {
  const m2 = bridgeOf(V2, {}).manifest(); const m3 = bridgeOf(V3, {}).manifest();
  assert.equal(m2.patchVersion, 'euro-office-bridge-word-v1/9.3.4.60/2');
  assert.equal(m3.patchVersion, 'euro-office-bridge-word-v1/9.3.4.60/3');
  assert.equal(m3.capabilities.length, m2.capabilities.length + 1);
  assert.deepEqual(m3.capabilities.slice(0, m2.capabilities.length), m2.capabilities);
  const r = m3.capabilities.filter(c => c.id === R);
  assert.equal(r.length, 1);
  assert.deepEqual([r[0].version, r[0].kind, r[0].format, r[0].mutation], [1, 'READ', 'DOCX', 'NO']);
  for (const k of ['bridgeVersion', 'documentServerVersion', 'documentServerBuild']) assert.equal(m3[k], m2[k]);
});

test('/3 = /2 + one inserted READ capability + the patchVersion bump (every /2 implementation byte-identical)', () => {
  let t = V3.replace('patchVersion: "euro-office-bridge-word-v1/9.3.4.60/3"', 'patchVersion: "euro-office-bridge-word-v1/9.3.4.60/2"');
  const cut = (s, startMark, endMark) => { const i = s.indexOf(startMark); const j = s.indexOf(endMark, i); assert.ok(i > 0 && j > i, startMark); return s.slice(0, i) + s.slice(j); };
  t = cut(t, '{ id: "docx.section.variantState.read/v1"', '], };');
  t = cut(t, '"docx.section.variantState.read/v1": function', '}; var EO_BRIDGE_WRITES = ');
  assert.equal(t, V2);
});

test('class separation: write() refuses the READ id; read() refuses other ids; /2 lacks the capability', () => {
  const env = fakeDoc([{ titlePage: true, evenAndOdd: false }]);
  const b = bridgeOf(V3, env.Api);
  assert.deepEqual(b.write(R, { section: 0 }), { ok: false, reason: 'capability-not-advertised' });
  for (const id of ['docx.section.variantState.read/v2', 'docx.paragraph.numbering.remove/v1', 'nope']) assert.deepEqual(b.read(id, { section: 0 }), { ok: false, reason: 'capability-not-advertised' }, id);
  assert.deepEqual(bridgeOf(V2, env.Api).read(R, { section: 0 }), { ok: false, reason: 'capability-not-advertised' }, '/2: the B21 READ is absent, detectably');
  assert.deepEqual(env.calls, []);
});

test('PARITY (M11.5e.4 cutover): the product B21 read goes through this SAME Bridge capability -- no direct internal getter remains in the product', () => {
  // post-cutover, the product's READ-side (propertyQuery's variant() helper) and WRITE-side re-check (__eoB21Variant)
  // both go through docx.section.variantState.read/v1; the direct internal getter NAMES live only in this DS-side block
  const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
  assert.equal(BRIDGE_CAPABILITIES.DOCX_SECTION_VARIANT_STATE_READ.id, R);
  assert.match(DRIVER, /BRIDGE_CAPABILITIES\.DOCX_SECTION_VARIANT_STATE_READ/);
  assert.match(HEADERFOOTER, /__eoB21Variant\(/);
  assert.doesNotMatch(DRIVER.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, ''), /Get_TitlePage|IsEvenAndOdd/);
  assert.doesNotMatch(HEADERFOOTER.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, ''), /Get_TitlePage|IsEvenAndOdd/);
  // a successful read mirrors exactly what the product's propertyQuery would compute for a healthy section
  const env = fakeDoc([{ titlePage: true, evenAndOdd: false }, { titlePage: false, evenAndOdd: true }]);
  assert.deepEqual(read_(env, { section: 0 }), { ok: true, value: { titlePage: true, evenAndOdd: false } });
  assert.deepEqual(read_(env, { section: 1 }), { ok: true, value: { titlePage: false, evenAndOdd: true } });
  assert.deepEqual(env.calls, [['Get_TitlePage'], ['IsEvenAndOdd'], ['Get_TitlePage'], ['IsEvenAndOdd']]);
});

test('refusals: bad args, document/sections unavailable, out of range, missing section, reader missing/threw/non-boolean', () => {
  const S0 = [{ titlePage: true, evenAndOdd: true }];
  const cases = [
    [{}, S0, {}, 'bad-args'], [{ section: '0' }, S0, {}, 'bad-args'], [{ section: -1 }, S0, {}, 'bad-args'], [{ section: 0.5 }, S0, {}, 'bad-args'],
    [{ section: 1 }, S0, {}, 'section-out-of-range'],
    [{ section: 0 }, [null], {}, 'section-not-found'],
    [{ section: 0 }, [{ noSection: true }], {}, 'reader-missing'],
    [{ section: 0 }, [{ titlePage: true, evenAndOdd: true, noTitlePageReader: true }], {}, 'reader-missing'],
    [{ section: 0 }, [{ titlePage: true, evenAndOdd: true, noEvenOddReader: true }], {}, 'reader-missing'],
    [{ section: 0 }, [{ titlePage: true, evenAndOdd: true, titlePageThrows: true }], {}, 'reader-threw'],
    [{ section: 0 }, [{ titlePage: true, evenAndOdd: true, evenAndOddThrows: true }], {}, 'reader-threw'],
    [{ section: 0 }, [{ titlePage: 'yes', evenAndOdd: true }], {}, 'not-boolean'],
    [{ section: 0 }, [{ titlePage: true, evenAndOdd: 1 }], {}, 'not-boolean'],
    [{ section: 0 }, S0, { noDocument: true }, 'document-unavailable'],
    [{ section: 0 }, S0, { noGetSections: true }, 'document-unavailable'],
    [{ section: 0 }, S0, { getSectionsThrows: true }, 'sections-threw'],
  ];
  for (const [args, sections, opts, reason] of cases) {
    const env = fakeDoc(sections, opts);
    assert.equal(read_(env, args).reason, reason, JSON.stringify([args, opts]));
  }
});

test('the build is deterministic and pinned: bridge-block.js = new.txt + separator; old.txt = the canonical /2 block', () => {
  const C = path.join(DIR, 'rev3-candidate');
  assert.equal(V3, read(path.join(C, 'new.txt')) + ',');
  assert.equal(read(path.join(C, 'old.txt')), V2.replace(/,\s*$/, ''));
  const man = JSON.parse(read(path.join(C, 'manifest.json')));
  assert.equal(man.original_sha256, 'acd8fa14a65341ef2aa8a5c3adcf5c4e0ffb43627300732219b2e417410be924', 'the input is the served /2');
  assert.equal(man.operations.length, 1);
  assert.equal(man.operations[0].old, read(path.join(C, 'old.txt')));
  assert.equal(man.operations[0].new, read(path.join(C, 'new.txt')));
});
