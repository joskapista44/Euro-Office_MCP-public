'use strict';

/**
 * GH #223 (DOCX P1 READ): the DEDICATED `read_page_setup` operation on the product `docx_read`
 * surface -- a plain field-select over inventory.cjs's own per-section pageSetup fields
 * (pageWidth/pageHeight/margins/columnsCount/columnsSpaces/startPageNumber, twips), already
 * measured for the W11 page-setup WRITE side's own verifyPreservation -- no new driver reader.
 * `titlePage`/`evenAndOdd` are DELIBERATELY excluded (B21/#227, a different, runtime-gated axis).
 * Decision: docs/GH223-DOCX-PAGE-SETUP-READ-DECISION.md.
 *
 * Offline only: test/fakes/fake-docx-driver.cjs already simulates the per-field UNOBSERVABLE
 * gaps (script.pageSetupUnreadable / columnsSpacesUnreadable) the W11 WRITE-side offline tests
 * (docx-w11-page-setup.test.cjs) already rely on.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };

function harness(sections, script = {}) {
  const s = createStore([{ text: 'Body', props: { ...PLAIN } }], { doc: { sections } });
  const b = fakeBroker('docx');
  return createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(s, script) });
}

async function readPageSetup(sections, script) {
  const r = await harness(sections, script).handler({ fileId: FILE, operation: 'read_page_setup' });
  assert.equal(r.ok, true, JSON.stringify(r));
  return r.result.sections;
}

// ---- every field, fully measured ----------------------------------------------------------------------------------

test('GH #223: all six fields, fully measured -- a non-default size/margins/columns/startPageNumber round-trip verbatim', async () => {
  const sec = {
    pageWidth: 16838, pageHeight: 11906, // A4 landscape, twips
    margins: { top: 720, right: 1080, bottom: 720, left: 1080 },
    columnsCount: 2, columnsSpaces: [360], startPageNumber: 5,
  };
  const sections = await readPageSetup([sec]);
  assert.deepEqual(sections, [{ index: 0, ...sec }]);
});

// ---- a genuinely measured zero is NOT the same string as an unread gap ---------------------------------------------

test('GH #223: a measured startPageNumber of 0 passes through as the NUMBER 0, never confused with UNOBSERVABLE', async () => {
  const sections = await readPageSetup([{ pageWidth: 12240, pageHeight: 15840, margins: { top: 0, right: 0, bottom: 0, left: 0 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 0 }]);
  assert.equal(sections[0].startPageNumber, 0);
  assert.equal(typeof sections[0].startPageNumber, 'number');
  assert.deepEqual(sections[0].margins, { top: 0, right: 0, bottom: 0, left: 0 });
  for (const v of Object.values(sections[0].margins)) assert.equal(typeof v, 'number');
});

// ---- absent vs unread ------------------------------------------------------------------------------------------------

test('GH #223: the WHOLE page-setup axis unreadable -- every field carries the SAME named reason, never a default', async () => {
  const sections = await readPageSetup([{ pageWidth: 12240, pageHeight: 15840, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 1 }], { pageSetupUnreadable: true });
  assert.deepEqual(sections, [{
    index: 0, pageWidth: 'UNOBSERVABLE:no-such-getter', pageHeight: 'UNOBSERVABLE:no-such-getter',
    margins: { top: 'UNOBSERVABLE:no-such-getter', right: 'UNOBSERVABLE:no-such-getter', bottom: 'UNOBSERVABLE:no-such-getter', left: 'UNOBSERVABLE:no-such-getter' },
    columnsCount: 'UNOBSERVABLE:no-such-getter', columnsSpaces: 'UNOBSERVABLE:no-such-getter', startPageNumber: 'UNOBSERVABLE:no-such-getter',
  }]);
});

test('GH #223: ONE field (columnsSpaces) unreadable -- every other field on the SAME section stays independently measured, a single gap never widens', async () => {
  const sections = await readPageSetup([{ pageWidth: 12240, pageHeight: 15840, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, columnsCount: 3, columnsSpaces: [360, 360], startPageNumber: 1 }], { columnsSpacesUnreadable: true });
  assert.equal(sections[0].columnsSpaces, 'UNOBSERVABLE:no-such-getter');
  assert.equal(sections[0].pageWidth, 12240);
  assert.equal(sections[0].columnsCount, 3);
  assert.deepEqual(sections[0].margins, { top: 1440, right: 1440, bottom: 1440, left: 1440 });
});

// ---- multiple sections: identity/order stable, each independently measured -------------------------------------------

test('GH #223: section index/order stable across three sections, each with its own independent page setup', async () => {
  const sections = await readPageSetup([
    { pageWidth: 12240, pageHeight: 15840, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 1 },
    { pageWidth: 15840, pageHeight: 12240, margins: { top: 720, right: 720, bottom: 720, left: 720 }, columnsCount: 2, columnsSpaces: [180], startPageNumber: 1 },
    { pageWidth: 12240, pageHeight: 15840, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 7 },
  ]);
  assert.equal(sections.length, 3);
  assert.deepEqual(sections.map(s => s.index), [0, 1, 2]);
  assert.equal(sections[1].pageWidth, 15840, 'section 1 is landscape, section 0/2 are portrait');
  assert.equal(sections[0].startPageNumber, 1);
  assert.equal(sections[2].startPageNumber, 7, 'the THIRD section restarts page numbering -- independent of the other two');
});

// ---- never leaks titlePage/evenAndOdd (B21/#227's axis, excluded by design) --------------------------------------------

test('GH #223: titlePage/evenAndOdd never appear on a read_page_setup result -- a different, runtime-gated axis (B21/#227)', async () => {
  const sections = await readPageSetup([{ pageWidth: 12240, pageHeight: 15840, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 1 }]);
  assert.equal('titlePage' in sections[0], false);
  assert.equal('evenAndOdd' in sections[0], false);
  assert.equal(/titlePage|evenAndOdd/.test(JSON.stringify(sections)), false);
});

// ---- never leaks header/footer (read_headers_footers's own axis) -----------------------------------------------------

test('GH #223: header/footer never appear on a read_page_setup result -- that is read_headers_footers own axis, not duplicated here', async () => {
  const sections = await readPageSetup([{ pageWidth: 12240, pageHeight: 15840, margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 1 }]);
  assert.equal('header' in sections[0], false);
  assert.equal('footer' in sections[0], false);
});

// ---- RED CONTROL: proves the handler reads the documents OWN fixture, not a hardcoded default ------------------------

test('RED CONTROL: pageWidth/startPageNumber come from the documents OWN fixture, not a hardcoded A4/page-1 default', async () => {
  const sections = await readPageSetup([{ pageWidth: 9999, pageHeight: 8888, margins: { top: 1, right: 2, bottom: 3, left: 4 }, columnsCount: 1, columnsSpaces: [], startPageNumber: 42 }]);
  assert.equal(sections[0].pageWidth, 9999);
  assert.notEqual(sections[0].pageWidth, 12240, 'must not fall back to the US-Letter-ish fixture default');
  assert.equal(sections[0].startPageNumber, 42);
  assert.notEqual(sections[0].startPageNumber, 1, 'must not fall back to a hardcoded page-1 default');
});
