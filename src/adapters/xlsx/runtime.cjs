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
 */
const XLSX_RUNTIME_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const XLSX_RUNTIME_17_ID = '9.3.5|euro-office-bridge-v1/9.3.5.0/17';
/** The explicit accepted list -- /16 + /17, pinned to core ACCEPTED_RUNTIMES by a static test. */
const XLSX_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

/** The B05-WRITE Bridge capability's runtime -- /16 + /17, pinned to the core value by a static test. */
const XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

/** The B10-WRITE Bridge capability's runtime -- /16 + /17, pinned to the core value by a static test. */
const XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

/** The B03-WRITE Bridge capability's runtime -- /16 + /17, pinned to the core value by a static test. */
const XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

/** The chart-style gapOverlap/seriesOrder WRITE capabilities' own runtime -- /16 + /17, pinned to the core value by a static test. */
const XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

/** The chart-delete WRITE capability's own runtime -- /16 + /17, pinned to the core value by a static test. */
const XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

/** The chart.series.marker.write/v1 / chart.series.lineStyle.write/v1 WRITE capabilities' own runtime -- /16 + /17,
 * pinned to the core value by a static test. */
const XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID]);

module.exports = { XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID, XLSX_RUNTIME_IDS, XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B24_BRIDGE_WRITE_RUNTIME_IDS };
