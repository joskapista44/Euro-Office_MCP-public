'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const PROPERTY_QUERY = propertyQuery(); // the DEFAULT read -- lineSpacing is unconditional, not behind numberingFormat
const { BRIDGE_MAJOR, RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

/**
 * GH #231b (docs/GH231-FORMAT-TEXT-PARAGRAPH-PROPS-DECISION.md section 8): the `lineSpacing`
 * READ conversion (docx-driver.cjs's `GetSpacingLineValue` handling) was a ported, never-LIVE-
 * verified formula with a measured round-trip defect below a written multiplier of ~0.833. Per
 * the `test/static/docx-driver-numbering-read-vm.test.cjs` precedent: the fake docx driver
 * (test/fakes/fake-docx-driver.cjs) does NOT model the twips/multiplier unit conversion at all for
 * paragraph-level keys -- it passes values through verbatim. Only running the REAL `PROPERTY_QUERY`
 * string (imported from the driver, not retyped) via `vm`, against a hand-built fake `Api`/`doc`
 * with a controllable `GetSpacingLineValue()` return, can prove the conversion bug or its fix.
 *
 * The LIVE confirmation of the chosen formula (uniform n/240, decision doc 8.4/8.5) is a SEPARATE,
 * not-yet-run step (test/live/gh231b-linespacing-read-roundtrip-probe.live.cjs) -- these tests prove
 * the DRIVER CODE now computes the chosen formula correctly, not that the formula itself is the one
 * true answer. If the LIVE probe contradicts 8.4's reasoning, this test (and the fix) are revisited.
 */

function withWordBridge(api, byParagraph) {
  const cap = BRIDGE_CAPABILITIES.DOCX_LIST_LEVEL_FORMAT_READ;
  const manifest = { bridgeVersion: `${BRIDGE_MAJOR}.0.0`, documentServerVersion: RUNTIMES.docx.documentServerVersion, patchVersion: RUNTIMES.docx.patchVersion, capabilities: [{ id: cap.id, version: cap.version, kind: 'READ', mutation: 'NO' }] };
  return { ...api, GetEuroOfficeBridge: () => ({ manifest: () => manifest, read: (id, args) => (id === cap.id && byParagraph[args.paragraph] ? { ok: true, value: byParagraph[args.paragraph] } : { ok: false, reason: 'not-faked' }) }) };
}

/** One paragraph, uniquely findable by doc.Search, with a controllable ParaPr.GetSpacingLineValue. */
function makeFakeApi(rawSpacingLineValue) {
  const textPr = {}; // no Get* methods faked -- every run-level prop field is guarded by typeof===function, stays unread
  const paraPr = {
    GetSpacingLineValue: () => rawSpacingLineValue,
    // GetJc/GetSpacingBefore/GetSpacingAfter/GetIndLeft/GetIndRight/GetIndFirstLine/GetShd
    // deliberately ABSENT -- this test is about lineSpacing only, the other paragraph-level
    // keys are already covered by GH231's own vm/acceptance tests.
  };
  const linePara = {
    GetClassType: () => 'paragraph',
    GetText: () => 'LS_PARA',
    GetElementsCount: () => 0,
    GetNumbering: () => null,
    GetStyle: () => null,
    GetParaPr: () => paraPr,
  };
  const elements = [linePara];
  const doc = {
    GetElementsCount: () => elements.length,
    GetElement: i => elements[i],
    Search: text => elements.filter(e => e.GetText() === text).map(() => ({ GetText: () => text, GetTextPr: () => textPr })),
    GetAllBookmarksNames: () => [],
    GetBookmarkRange: () => null,
    GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetSections: () => [],
    GetDefaultTextPr: () => null,
  };
  return withWordBridge({ GetDocument: () => doc }, { 0: { present: false } });
}

function runPropertyQuery(fakeApi) {
  const sandbox = { Api: fakeApi, result: undefined };
  vm.createContext(sandbox);
  const raw = vm.runInContext(`(function(){ ${PROPERTY_QUERY} })()`, sandbox);
  // cross-realm plain-object normalization, same reason as the numbering vm test.
  return JSON.parse(JSON.stringify(raw));
}

function lsOf(rawSpacingLineValue) {
  const r = runPropertyQuery(makeFakeApi(rawSpacingLineValue));
  assert.equal(r.error, undefined, `PROPERTY_QUERY threw: ${r.error}`);
  const p = r.paragraphs.find(x => x.text === 'LS_PARA');
  assert.ok(p, 'the probe paragraph must be present in the read');
  return p.format;
}

// --- RED proof: the formula this unit replaces (decision doc 8.2/8.6), evaluated here as plain
// JS (NOT re-executed from any live source -- the old branch no longer exists in docx-driver.cjs)
// so the historical defect stays provable even after the fix lands.
function oldFormula(n) { return n > 10 ? (n >= 200 ? n / 240 : (n + 1) / 100) : n; }

test('RED proof: the OLD formula misreads a written lineSpacing of 0.5 (raw 120) as 1.21, not 0.5', () => {
  assert.equal(oldFormula(120), 1.21);
  assert.notEqual(oldFormula(120), 0.5);
});

test('GREEN: the fixed driver reads raw 120 (written lineSpacing 0.5) back as 0.5, not the old 1.21', () => {
  assert.equal(lsOf(120).lineSpacing, 0.5);
});

test('GREEN: the full grid (0.5, 0.75, 1, 1.15, 1.5, 2, 3) round-trips exactly through n/240, each value written as n = value * 240', () => {
  const GRID = [0.5, 0.75, 1, 1.15, 1.5, 2, 3];
  for (const v of GRID) {
    const raw = v * 240;
    const got = lsOf(raw).lineSpacing;
    assert.ok(Math.abs(got - v) < 1e-9, `written ${v} (raw ${raw}) read back as ${got}`);
  }
});

test('GREEN, boundary: a raw value just under the OLD formula\'s n>=200 branch boundary (raw 199, i.e. written ~0.829) now reads back as a clean n/240, not (n+1)/100', () => {
  const got = lsOf(199).lineSpacing;
  assert.ok(Math.abs(got - 199 / 240) < 1e-9);
  assert.notEqual(got, (199 + 1) / 100);
});

test('values that already round-tripped correctly under the OLD formula (>=200/240) are UNCHANGED by the fix', () => {
  assert.equal(lsOf(240).lineSpacing, 1.0);
  assert.equal(lsOf(360).lineSpacing, 1.5);
  assert.equal(lsOf(480).lineSpacing, 2.0);
});

test('NEGATIVE: a paragraph whose ParaPr has no GetSpacingLineValue method leaves lineSpacing unset, never guessed', () => {
  const r = runPropertyQuery((() => {
    const api = makeFakeApi(240);
    const doc = api.GetDocument();
    doc.GetElement(0).GetParaPr = () => ({});
    return api;
  })());
  const p = r.paragraphs.find(x => x.text === 'LS_PARA');
  assert.equal('lineSpacing' in p.format, false);
});

test('NEGATIVE: a non-finite raw value (NaN) leaves lineSpacing unset, never coerced', () => {
  assert.equal('lineSpacing' in lsOf(NaN), false);
});
