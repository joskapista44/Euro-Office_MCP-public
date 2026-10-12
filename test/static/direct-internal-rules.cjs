'use strict';

/**
 * Direct ONLYOFFICE internal-model dependency guard -- the rules (owner decision, owner channel 2026-10-08, point 7).
 *
 * A "direct internal" use is a call or property on the editor's INTERNAL model (sheet/document model objects, the
 * `asc_*` editor methods that reach into it, the internal paragraph/print/pivot classes), as opposed to the public
 * builder API or the DS Bridge. Every such use in src/ must be covered by a grandfathered legacy entry. Anything else
 * is a NEW direct internal dependency and fails the gate.
 *
 * The legacy set is FROZEN at the groups still under migration (DS-BRIDGE-INTERNAL-MIGRATION-20261006.md):
 * B03 B10 B18 (Wave B READ removed B09, B13 and B05's READ half; the M11.5a.4 Bridge WRITE cutover removed B05's WRITE half; the B15 /5
 * RUNTIME_HYDRATION cutover removed B15 -- owner GO owner channel 2026-10-09 msg 33257). These four are the M10
 * TEMPORARY_RC_BOUNDARY set (docs/M10-ACCEPTED-DIRECT-INTERNAL-BOUNDARIES.md), POST_M10_MIGRATION_REQUIRED=YES. A migration
 * REMOVES its entry (the exact counts force that); the set may only shrink.
 * Growing it is an owner decision and has to edit KNOWN_LEGACY_SET below -- visibly, in review, never silently.
 *
 * Kept pure (no file walk) so the same rules can be shown to go red on a synthetic fixture.
 */

// B21 (owner decision M11.3a.1, after owner channel): the DOCX section variant state READ (ApiSection's internal
// Section.Get_TitlePage / IsEvenAndOdd), in src/ since 2026-09-25 (bc1c0ed) but missed by this register until then --
// an INVENTORY_DEFECT, not a new dependency. Runtime-gated like B18; migration = M11.5e.
// M11.5a.4 (owner GO owner channel, internal tracker): the B05-WRITE cutover to range.alignmentExt.write/v1 removed
// the last B05 use, so the SOURCE register drops it (the exact counts force that). Two separate notions (owner decision):
//   SOURCE_LEGACY_REGISTER (KNOWN_LEGACY_SET below)   what is physically in src/ -- 4
//   ACCEPTED_OPERATIONAL_LEGACY_SET                  the canonical operational status -- B05 stayed in it until the
//                                                    D-state (cutover product + DS /6) LIVE acceptance PASS
// M11.5a.5 (owner: M11.5a.4 CLOSED/PASS, D-state release 9e3a57f + DS /6 af42c232; final integrated acceptance PASS):
// B05-WRITE = REMOVED_FROM_LEGACY, so the operational set equals the source register again -- 4.
// M11.5b.4 (M11.5 continuous mandate, internal tracker): the B10 Bridge WRITE cutover (sheet.printSettings.write/v1)
// removed every B10 use, so the SOURCE register drops it -- 3. The accepted operational set keeps B10 until the D-state
// (cutover product + DS /7) LIVE acceptance PASS.
// M11.5b.5 (release 4a4fef4 + DS /7 3a73042b: D-state, product and DS rollback/forward LIVE PASS): B10-WRITE =
// REMOVED_FROM_LEGACY, so the operational set equals the source register again -- 3.
// M11.5c.4 (same mandate): the B03 Bridge WRITE cutover (worksheet.copy/v1) removed every B03 use, so the SOURCE register
// drops it -- 2. The accepted operational set keeps B03 until the D-state (cutover product + DS /8) LIVE acceptance PASS.
// M11.5c.5 (release d47a42c + DS /8 73c9211d: D-state, B03 retry/interrupted, product and DS rollback/forward LIVE PASS):
// B03-WRITE = REMOVED_FROM_LEGACY, so the operational set equals the source register again -- 2.
// M11.5d.4 (same mandate): the B18 Bridge WRITE cutover (docx.paragraph.numbering.remove/v1, word /2) removed every B18 use,
// so the SOURCE register drops it -- 1. The accepted operational set keeps B18 until the D-state (cutover product + DS
// word /2) LIVE acceptance PASS.
// M11.5d.5 (release ea5b2f8 + DS word /2 acd8fa14: D-state, product and DS rollback/forward LIVE PASS):
// B18-WRITE = REMOVED_FROM_LEGACY, so the operational set equals the source register again -- 1.
// M11.5e.4 (same mandate): the B21 Bridge READ cutover (docx.section.variantState.read/v1, word /3) removed every B21
// use, so the SOURCE register drops it -- 0. The accepted operational set kept B21 until the D-state (cutover product
// + DS word /3) LIVE acceptance PASS -- the two sets were ASYMMETRIC at that mid-migration step (0 vs 1), unlike every
// prior cutover's own mid-migration step.
// M11.5e.5 (release 870f3ac67cb6 + DS word /3 7aeb04e9: D-state, product and DS rollback/forward LIVE PASS):
// B21-READ = REMOVED_FROM_LEGACY, so the operational set equals the source register again -- 0 (both sets empty).
const KNOWN_LEGACY_SET = Object.freeze([]);
const ACCEPTED_OPERATIONAL_LEGACY_SET = Object.freeze([]);

// The internal-model markers. A token is the matched text, normalised (whitespace and the `(` removed).
const INTERNAL_PATTERNS = Object.freeze([
  /\.asc_[A-Za-z]+(?=\s*\()/g,         // asc_* editor/model methods
  /\bwbModel\b/g,                      // the workbook model
  /\bAscCommon[A-Za-z]*\b/g,           // internal namespaces (history, ...)
  /\.worksheet\b/g,                    // ApiWorksheet -> internal sheet model
  /\.getWorksheet(?=\s*\()/g,          // the internal workbook view
  /\bPagePrintOptions\b/g,             // the internal print model
  /\bCParagraph\b/g, /\bRemoveNumPr\b/g, // internal Word paragraph model
  /\b(?:getReadingOrder|getIndent|getShrinkToFit|setIndent|setShrinkToFit)\b/g, // internal alignment accessors
  // B21: ApiSection's internal section object, and the internal (underscore / Is*) getter names -- also as the STRING a
  // dynamic `inner[name]()` call takes, so the dynamic form cannot slip past the register
  /\.Section\b/g, /\b(?:Get|Set|Is|Has)_[A-Z]\w*\b/g, /\bIsEvenAndOdd\b/g,
]);

/**
 * The grandfathered uses: { id, file, token, count } -- `count` is EXACT. A migration that removes uses must lower
 * (or delete) the entry in the same change; a new use of an existing token in the same file breaks the count.
 */
const LEGACY = Object.freeze([
  // M11.5e.4: B21 cut over to the word /3 Bridge READ -- no direct internal use is left to grandfather.
]);

/**
 * NOT internal-model uses, listed so they cannot grow silently either: the editor SESSION API the core lifecycle
 * reads (save state), which touches no document model.
 */
const EDITOR_SESSION = Object.freeze([
  { file: 'drivers/eurooffice/xlsx-driver.cjs', token: '.asc_isDocumentCanSave', count: 1 },
  { file: 'drivers/eurooffice/docx-driver.cjs', token: '.asc_isDocumentCanSave', count: 1 },
]);

/** Strip comments line by line (block-comment lines and `//` tails), keep code and string bodies. */
function codeLines(source) {
  return source.split('\n').filter(l => !/^\s*(\/\*|\*)/.test(l)).map(l => l.replace(/(^|[^:'"`\\])\/\/.*$/, '$1'));
}

/** Count the internal tokens per file: { [file]: { [token]: n } }. `modules`: [{ file, source }]. */
function countInternal(modules) {
  const out = {};
  for (const { file, source } of modules) {
    for (const line of codeLines(source)) {
      for (const re of INTERNAL_PATTERNS) {
        for (const m of line.match(re) || []) {
          const token = m.replace(/\s+/g, '');
          out[file] = out[file] || {};
          out[file][token] = (out[file][token] || 0) + 1;
        }
      }
    }
  }
  return out;
}

/** The verdict: { newDirectInternal: [...], staleLegacy: [...], unknownIds: [...] }. */
function judge(modules, legacy = LEGACY, session = EDITOR_SESSION, known = KNOWN_LEGACY_SET) {
  const counts = countInternal(modules);
  const allowed = {};
  for (const e of [...legacy, ...session]) {
    allowed[e.file] = allowed[e.file] || {};
    allowed[e.file][e.token] = (allowed[e.file][e.token] || 0) + e.count;
  }
  const newDirectInternal = []; const staleLegacy = [];
  for (const [file, toks] of Object.entries(counts)) {
    for (const [token, n] of Object.entries(toks)) {
      const a = (allowed[file] && allowed[file][token]) || 0;
      if (n > a) newDirectInternal.push({ file, token, found: n, grandfathered: a });
      else if (n < a) staleLegacy.push({ file, token, found: n, grandfathered: a });
    }
  }
  for (const [file, toks] of Object.entries(allowed)) {
    for (const [token, a] of Object.entries(toks)) if (!(counts[file] && counts[file][token])) staleLegacy.push({ file, token, found: 0, grandfathered: a });
  }
  const unknownIds = [...new Set(legacy.map(e => e.id))].filter(id => !known.includes(id));
  return { newDirectInternal, staleLegacy, unknownIds, legacyIds: [...new Set(legacy.map(e => e.id))].sort() };
}

module.exports = { KNOWN_LEGACY_SET, ACCEPTED_OPERATIONAL_LEGACY_SET, INTERNAL_PATTERNS, LEGACY, EDITOR_SESSION, countInternal, judge, codeLines };
