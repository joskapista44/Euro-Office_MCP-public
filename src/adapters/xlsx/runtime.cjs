'use strict';

/**
 * The accepted DS runtime identity, as the adapter's OWN copy (the adapter imports nothing outside its tree): the
 * driver's session identity (core/runtime-identity.cjs, owner decision (B), internal tracker) is compared against it as
 * the adapter-side line of the runtime-gated internal paths. A static test pins it to the core value.
 *
 * GH #326 / #327 (owner decision, path A): every list below widens from its single
 * 9.3.5.0/16 entry to /16 + /17 (compatibility-first -- a DS /16 -> /17 apply needs no product outage). /17 adds
 * ONLY the (not yet used) sheet.protection.read/v1 Bridge READ capability; every WRITE handler stays byte-identical
 * on it, mirroring core's own widening.
 *
 * GH #300 L1 / #298 D: the GENERAL accepted list (XLSX_RUNTIME_IDS only) widens again, to /16 + /17 + /18 -- /18
 * adds ONLY the two new READ capabilities (sheet.tables.read/v1, sheet.hyperlinks.read/v1). The six Bridge-WRITE
 * lists below carry /18 too, mirroring core: every pre-existing handler is byte-identical on /17 and /18
 * (test/static/ds935-bridge18-write-handler-identity.test.cjs).
 */
const XLSX_RUNTIME_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const XLSX_RUNTIME_17_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/17';
const XLSX_RUNTIME_18_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/18';
const XLSX_RUNTIME_19_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/19';
// GH #354/#355: /21 = /19 + the /20 READs + sheet.printSettings.write/v2; in the general list and the Bridge-WRITE lists (byte-identity warrant)
const XLSX_RUNTIME_21_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/21';
// GH #338-#342: /22 = /21 + sheet.protection.write/v1 + sheet.cellProtection.write/v1; in the general list and the Bridge-WRITE lists (byte-identity warrant)
const XLSX_RUNTIME_22_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/22';
// GH #356/#357: /23 = /22 + sheet.printPageCount.read/v1; in the general list and the Bridge-WRITE lists (byte-identity warrant)
const XLSX_RUNTIME_23_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/23';
// GH #338-#342 (docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md): /24 = /23 + workbook.protection.read/v1 +
// workbook.protection.write/v1 + sheet.table.write/v1's totalsRow accepting value:false; in the general list and
// the nine Bridge-WRITE lists on the byte-identity warrant (test/static/ds935-bridge24-write-handler-identity.test.cjs)
const XLSX_RUNTIME_24_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/24';
// W2 regression fix (internal tracker): /25 = /24 + ONE handler-body change only, to workbook.protection.read/v1
// (the engine's null shape -- an empty <workbookProtection/> element -- now answers {lockStructure:false}); no
// capability added or removed, every other handler byte-identical (test/static/ds935-bridge25-write-handler-identity.test.cjs)
const XLSX_RUNTIME_25_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/25';
// Review fix (internal tracker, f0451f47): /26 = /25 + ONE guard change in TWO existing handlers (worksheet.copy/v1,
// sheet.table.write/v1 -- the SAME workbook.protection.read/v1 getter, not asc_isProtectedWorkbook()); no capability
// added or removed, every other handler byte-identical (test/acceptance/ds935-bridge26-worksheet-copy-table-write-protection.test.cjs)
const XLSX_RUNTIME_26_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/26';
const XLSX_RUNTIME_27_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/27';
// GH #361/#363 pre-install readiness gate: /28 = /27 + sheet.table.autoFilter.criteria.read/v1 (GH #361) +
// sheet.cells.compiledStyle.read/v1 (GH #363); every /27 handler byte-identical on it (test/static/ds935-
// bridge28-table-autofilter-criteria-write-handler-identity.test.cjs). The candidate's patchVersion
// originally baked its own directory suffix in (unlike every earlier candidate's clean "/NN") -- fixed at
// the source -- see core/runtime-identity.cjs's own note.
const XLSX_RUNTIME_28_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
// GH #365 option C U1: /29 = /28 + sheet.veryHidden.read/v1; every /28
// handler byte-identical on it (test/static/ds935-bridge29-veryhidden-write-handler-identity.test.cjs).
// See core/runtime-identity.cjs's own note for the full widening rationale (the ONE deliberate
// exclusion, XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS, is the same there and here).
const XLSX_RUNTIME_29_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/29';
/** The explicit accepted list -- /16 + /17 + /18, pinned to core ACCEPTED_RUNTIMES by a static test. */
const XLSX_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The B05-WRITE Bridge capability's runtime -- /16 + /17 + /18, pinned to the core value by a static test. */
const XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The B10-WRITE Bridge capability's runtime -- /16 + /17 + /18, pinned to the core value by a static test. */
const XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The B03-WRITE Bridge capability's runtime -- /16 + /17 + /18, pinned to the core value by a static test. */
const XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The chart-style gapOverlap/seriesOrder WRITE capabilities' own runtime -- /16 + /17 + /18, pinned to the core value by a static test. */
const XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The chart-delete WRITE capability's own runtime -- /16 + /17 + /18, pinned to the core value by a static test. */
const XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The chart.series.marker.write/v1 / chart.series.lineStyle.write/v1 WRITE capabilities' own runtime -- /16 + /17 + /18,
 * pinned to the core value by a static test. */
const XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The sheet.table.write/v1 capability's own runtime -- /19 ONLY (the capability does not exist on /16../18), pinned to the core value by a static test. */
const XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The sheet.printSettings.write/v2 capability's own runtime -- /21 ONLY (the capability does not exist on /16../19), pinned to the core value by a static test. */
const XLSX_B26_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The sheet.protection.write/v1 + sheet.cellProtection.write/v1 capabilities' own runtime -- /22 + /23, pinned to the core value by a static test. */
const XLSX_B27_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

/** The workbook.protection.read/v1 + workbook.protection.write/v1 capabilities' own runtime -- /24 ONLY (the
 * capabilities do not exist on /16../23), pinned to the core value by a static test. The next free Bridge-write
 * label after /22+/23's shipped XLSX_B27_BRIDGE_WRITE_RUNTIME_IDS. */
const XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);

// GH #346 unit 2: the chart.series.dataPoint.style.write/v1 capability's own runtime -- /27 ONLY (the capability
// does not exist on /16..26), pinned to the core value by a static test. NOT yet wired to any adapter family.
// DELIBERATELY NOT WIDENED to DS revision /28 (the pre-install readiness gate): this label's "B29" is an internal Bridge-write
// sequence number, unrelated to the DS revision number "/28" -- see core/runtime-identity.cjs's own note.
const XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_27_ID]);

module.exports = { XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_18_ID, XLSX_RUNTIME_19_ID, XLSX_RUNTIME_21_ID, XLSX_RUNTIME_22_ID, XLSX_RUNTIME_23_ID, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID, XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B26_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B27_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS, XLSX_RUNTIME_IDS, XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS };
