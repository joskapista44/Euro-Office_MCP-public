'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const PROPERTY_QUERY = propertyQuery({ numberingFormat: true }); // W-07's opt-in read (liveScope of a set_numbering task)
const { BRIDGE_MAJOR, RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

/**
 * The paragraph's numbering identity (present/level/numId/levelFormats) comes ONLY through the Word
 * DS Bridge read docx.numbering.levelFormat.read/v1 (W-B1920-C1 cutover) -- this minimal served-
 * bridge stand-in answers that one capability per paragraph index, so the test still exercises the
 * REAL query text; the W-07 `format` field itself is read through the PUBLIC ToJSON() path.
 */
function withWordBridge(api, byParagraph) {
  const cap = BRIDGE_CAPABILITIES.DOCX_LIST_LEVEL_FORMAT_READ;
  const manifest = { bridgeVersion: `${BRIDGE_MAJOR}.0.0`, documentServerVersion: RUNTIMES.docx.documentServerVersion, patchVersion: RUNTIMES.docx.patchVersion, capabilities: [{ id: cap.id, version: cap.version, kind: 'READ', mutation: 'NO' }] };
  return { ...api, GetEuroOfficeBridge: () => ({ manifest: () => manifest, read: (id, args) => (id === cap.id && byParagraph[args.paragraph] ? { ok: true, value: byParagraph[args.paragraph] } : { ok: false, reason: 'not-faked' }) }) };
}
const NUMBERED = { present: true, level: 0, numId: null, levelFormats: null };

/**
 * W-07 (M7 backlog, numbering pattern/start) -- the MANDATORY vm-test zero required (#ref,
 * "a W04-5 LECKEJEBOL #ref"): the fake docx driver (test/fakes/fake-docx-driver.cjs) does NOT
 * exercise the real PROPERTY_QUERY string at all -- it is a SEPARATE, hand-maintained read-side
 * simulation (see fake-docx-driver.cjs's own numberingOf()). A bug in PROPERTY_QUERY's numbering
 * block (e.g. the new `format` field this unit adds) could ship invisibly: every offline
 * acceptance test would stay green against the FAKE's own numbering shape, never against the
 * REAL query text.
 *
 * This test runs the ACTUAL `PROPERTY_QUERY` string (imported from the driver, not retyped) via
 * `vm`, against a hand-built fake `Api`/`doc` that includes ONE numbered paragraph and ONE plain
 * control paragraph -- the fake the fake-driver test harness never builds. It asserts the
 * resulting `numbering.format` field has the shape and values this unit's driver change
 * produces, reading the SAME nested path a maintainer's own LIVE evidence described:
 * abstractNum.<id>.lvl[ownLevel] = { numFmt: { val }, lvlText, suff, lvlJc, start }.
 */

function makeFakeApi() {
  const level = {
    GetLevelIndex: () => 0,
    GetNumbering: () => numberingObj,
  };
  const numberingObj = {
    // LIVE run1 (REAL bug, found by code-reading my own diff, NOT by this test as
    // it originally stood): ApiNumbering.ToJSON() returns a JSON STRING, not an object -- this
    // fake previously returned the OBJECT directly, which is EXACTLY why it could not catch the
    // driver treating the real (string) return as an object. Fixed to the MEASURED shape.
    ToJSON: () => JSON.stringify({
      abstractNum: {
        7: {
          lvl: [
            { numFmt: { val: 'lowerLetter' }, lvlText: '%1)', suff: 'space', lvlJc: 'left', start: 7, ilvl: 0 },
          ],
        },
      },
    }),
    // Num (the internal object levelFormats reads) is deliberately ABSENT here -- this test is
    // about the NEW public ToJSON path, not the pre-existing internal levelFormats one.
  };
  const numberedPara = {
    GetClassType: () => 'paragraph',
    GetText: () => 'NUMBERED_PARA',
    GetElementsCount: () => 0,
    GetNumbering: () => level,
    GetParaPr: () => null,
    GetStyle: () => null,
    Paragraph: null,
  };
  const controlPara = {
    GetClassType: () => 'paragraph',
    GetText: () => 'CONTROL_PARA',
    GetElementsCount: () => 0,
    GetNumbering: () => null,
    GetParaPr: () => null,
    GetStyle: () => null,
    Paragraph: null,
  };
  const elements = [numberedPara, controlPara];
  const doc = {
    GetElementsCount: () => elements.length,
    GetElement: i => elements[i],
    Search: text => elements.filter(e => e.GetText() === text).map(e => ({ GetText: () => text, GetTextPr: () => null })),
    GetAllBookmarksNames: () => [],
    GetBookmarkRange: () => null,
    GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetSections: () => [],
    GetDefaultTextPr: () => null,
  };
  return withWordBridge({ GetDocument: () => doc }, { 0: NUMBERED, 1: { present: false } });
}

function runPropertyQuery(fakeApi) {
  const sandbox = { Api: fakeApi, result: undefined };
  // PROPERTY_QUERY's own text is `try { ... return {...}; } catch (err) { return {error:...}; }`
  // -- wrapped in a function here only so `return` is legal; the query text itself is NOT
  // retyped or reconstructed, it is the REAL exported string.
  vm.createContext(sandbox);
  const raw = vm.runInContext(`(function(){ ${PROPERTY_QUERY} })()`, sandbox);
  // A vm context is a SEPARATE realm with its own Object/Array constructors -- plain objects
  // crossing the boundary fail assert.deepEqual's prototype-identity check even with identical
  // own properties. A JSON round-trip normalizes the result into THIS realm's plain objects
  // (the real bridge the driver uses, Playwright's page.evaluate, does exactly this -- it never
  // hands back a live cross-realm object either).
  return JSON.parse(JSON.stringify(raw));
}

test('vm: the REAL PROPERTY_QUERY, run against a fake Api with a numbered paragraph, carries the new numbering.format field', () => {
  const r = runPropertyQuery(makeFakeApi());
  assert.equal(r.error, undefined, `PROPERTY_QUERY threw: ${r.error}`);
  const numbered = r.paragraphs.find(p => p.text === 'NUMBERED_PARA');
  assert.ok(numbered, 'the numbered paragraph must be present in the read');
  assert.deepEqual(numbered.numbering, {
    present: true,
    level: 0,
    numId: null, // Paragraph.Pr.NumPr is not faked here -- out of this test's scope, unaffected by this unit's change
    levelFormats: null, // Num is absent on the fake numberingObj -- the PRE-EXISTING path correctly reports unread
    format: { numFmt: 'lowerLetter', lvlText: '%1)', suff: 'space', lvlJc: 'left', start: 7 },
    abstractId: '7', // GH #293 N5: the abstractNum key of the same ToJSON -- the list definition's identity
  });
});

test('vm: the control paragraph (no numbering) reads { present: false }, never a fabricated format', () => {
  const r = runPropertyQuery(makeFakeApi());
  const control = r.paragraphs.find(p => p.text === 'CONTROL_PARA');
  assert.ok(control, 'the control paragraph must be present in the read');
  assert.deepEqual(control.numbering, { present: false });
});

test('vm: a numbering whose ToJSON lacks the start key (default level) reads start: 1 -- W07-7 MEASURED resolution, #ref: absence means 1, both directions', () => {
  const api = makeFakeApi();
  const doc = api.GetDocument();
  const numbered = doc.GetElement(0);
  const level = numbered.GetNumbering();
  const realNumbering = level.GetNumbering();
  level.GetNumbering = () => ({
    ToJSON: () => JSON.stringify({ abstractNum: { 7: { lvl: [{ numFmt: { val: 'decimal' }, lvlText: '%1.', suff: 'tab', lvlJc: 'left', ilvl: 0 }] } } }),
  });
  const r = runPropertyQuery(api);
  const p = r.paragraphs.find(x => x.text === 'NUMBERED_PARA');
  assert.ok(p.numbering.format, 'format must still be read for the other fields');
  assert.equal(p.numbering.format.start, 1, 'an absent start key must read as 1 -- the W07-7 probe measured this in both directions (never-set and explicitly-set-to-1 are file-identical, w:start=1)');
  void realNumbering; // kept only to document what was overridden
});

test('vm, NEGATIVE: a numbering whose ToJSON start is present but NOT an integer stays unread -- never guessed, only the measured ABSENT case defaults to 1', () => {
  const api = makeFakeApi();
  const doc = api.GetDocument();
  const numbered = doc.GetElement(0);
  const level = numbered.GetNumbering();
  level.GetNumbering = () => ({
    ToJSON: () => JSON.stringify({ abstractNum: { 7: { lvl: [{ numFmt: { val: 'decimal' }, lvlText: '%1.', suff: 'tab', lvlJc: 'left', start: 'x', ilvl: 0 }] } } }),
  });
  const r = runPropertyQuery(api);
  const p = r.paragraphs.find(x => x.text === 'NUMBERED_PARA');
  assert.ok(p.numbering.format, 'format must still be read for the other fields');
  assert.equal('start' in p.numbering.format, false, 'a PRESENT but non-integer start must stay unread, never coerced or defaulted');
});

test('vm, NEGATIVE: a numbering with no ToJSON method at all reads format: null, never throws', () => {
  const api = makeFakeApi();
  const doc = api.GetDocument();
  const numbered = doc.GetElement(0);
  const level = numbered.GetNumbering();
  level.GetNumbering = () => ({});
  const r = runPropertyQuery(api);
  assert.equal(r.error, undefined);
  const p = r.paragraphs.find(x => x.text === 'NUMBERED_PARA');
  assert.equal(p.numbering.format, null);
});

test('opt-in only: the DEFAULT property query (every non-set_numbering read) carries no ToJSON numbering-format read -- the paragraph numbering stays the Bridge value as-is', () => {
  const sandbox = { Api: makeFakeApi(), result: undefined };
  vm.createContext(sandbox);
  const raw = JSON.parse(JSON.stringify(vm.runInContext(`(function(){ ${propertyQuery()} })()`, sandbox)));
  assert.equal(propertyQuery().includes('num.ToJSON'), false);
  assert.deepEqual(raw.paragraphs.find(p => p.text === 'NUMBERED_PARA').numbering, NUMBERED);
});
