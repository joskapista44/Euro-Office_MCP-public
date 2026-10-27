'use strict';

/**
 * GH #227 (DOCX P2 READ: section flags title/even). `read_section_flags` INTERNAL -> ACCEPTED: per-section
 * titlePage/evenAndOdd through the word /3 Bridge READ capability docx.section.variantState.read/v1 (B21,
 * M11.5e.4 cutover), the SAME runtime gate the header/footer family's own variant WRITE already refuses on
 * (internalGateOf(runtime).sectionVariant, B21_BRIDGE_READ_RUNTIME_IDS). No new driver code: the driver
 * already produces this field on every section when requested (docx-driver.cjs:160-162/585-592); the only
 * new wiring is read-tools.cjs's FORMATS.docx.read opting INTO `includeSectionVariantState`, and ONLY for
 * this one operation (the SAME opt-in pattern GH #225 established for numberingFormat).
 * Decision: docs/GH227-SECTION-FLAGS-READ-DECISION.md.
 *
 * Offline only: test/fakes/fake-docx-driver.cjs's own `variantFlagsOf` (added by this unit) models the
 * THREE named reasons a flag can be unread -- not-requested (opt-out) | runtime-not-verified (gate closed)
 * | bridge-unavailable (gate open, the bridge call itself failed) -- the SAME `UNAVAILABLE:<reason>` strings
 * the real driver's sectionVariantQuery/__eoB21Unavailable produce, and the SAME variantFlag() normalizer
 * (inventory.cjs) the write-side's own preservation check (hf.flagChanges) already depends on.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { createReadTool } = require('../../src/server/read-tools.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

function harness(script = {}, doc = {}) {
  const s = createStore([{ text: 'Body', props: {} }], { doc });
  const b = fakeBroker('docx');
  return createReadTool('docx', { configSource: b.source, loadPlaywright: () => ({ ok: true }), createDriver: () => createFakeDocxDriver(s, script) });
}

async function readSectionFlags(script, doc) {
  const r = await harness(script, doc).handler({ fileId: FILE, operation: 'read_section_flags' });
  assert.equal(r.ok, true, JSON.stringify(r));
  return r.result.sections;
}

const TWO_SECTIONS = { sections: [{ titlePage: true, header: {}, footer: {} }, { titlePage: false, header: {}, footer: {} }], evenAndOdd: true };

test('GH #227: gated runtime VERIFIED + requested -- every section reads a REAL boolean, per index, never a guessed default', async () => {
  const sections = await readSectionFlags({}, TWO_SECTIONS);
  assert.deepEqual(sections, [
    { index: 0, titlePage: true, evenAndOdd: true },
    { index: 1, titlePage: false, evenAndOdd: true },
  ]);
});

test('GH #227: gated runtime NOT VERIFIED -- both flags on EVERY section read a NAMED UNAVAILABLE:runtime-not-verified, never false/0', async () => {
  const sections = await readSectionFlags({ sectionVariantRuntimeNotVerified: true }, TWO_SECTIONS);
  assert.deepEqual(sections, [
    { index: 0, titlePage: 'UNAVAILABLE:runtime-not-verified', evenAndOdd: 'UNAVAILABLE:runtime-not-verified' },
    { index: 1, titlePage: 'UNAVAILABLE:runtime-not-verified', evenAndOdd: 'UNAVAILABLE:runtime-not-verified' },
  ]);
});

test('GH #227: gate OPEN but the bridge call itself fails -- both flags read UNAVAILABLE:bridge-unavailable, the SAME reason the write-side refusal (m11-3a1-section-variant.test.cjs) already exercises under its OWN pre-existing name `flagsUnreadable`', async () => {
  const sections = await readSectionFlags({ flagsUnreadable: true }, TWO_SECTIONS);
  assert.deepEqual(sections, [
    { index: 0, titlePage: 'UNAVAILABLE:bridge-unavailable', evenAndOdd: 'UNAVAILABLE:bridge-unavailable' },
    { index: 1, titlePage: 'UNAVAILABLE:bridge-unavailable', evenAndOdd: 'UNAVAILABLE:bridge-unavailable' },
  ]);
});

test('GH #227: a single, titlePage:false/evenAndOdd:false section -- a real false is NOT collapsed into an unread/UNAVAILABLE gap (the PRESERVATION contract m11-3a1-section-variant.test.cjs already proves on the WRITE side: unknown-before==unknown-after != a real false==false)', async () => {
  const sections = await readSectionFlags({}, { sections: [{ titlePage: false, header: {}, footer: {} }], evenAndOdd: false });
  assert.deepEqual(sections, [{ index: 0, titlePage: false, evenAndOdd: false }]);
});

test('GH #227: no section at all (absent doc.sections) -- the operation reads { sections: null }, a NAMED unread, never an empty list standing in for "no sections exist"', async () => {
  const r = await harness({ sectionsUnreadable: true }).handler({ fileId: FILE, operation: 'read_section_flags' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.result.sections, null);
});
