'use strict';

/**
 * The accepted DS runtime identity, as the adapter's OWN copy (the adapter imports nothing outside its tree): the
 * driver's session identity (core/runtime-identity.cjs, owner decision (B), internal tracker) is compared against it as
 * the adapter-side line of the runtime-gated internal paths. A static test pins it to the core value.
 */
const XLSX_RUNTIME_ID = '9.3.4|euro-office-bridge-v1/9.3.4.60/5';
/** M11.5a.2 / M11.5b.1 / M11.5c.1: the explicit accepted list (/5 + /6 + /7 + /8), pinned to core ACCEPTED_RUNTIMES by a static test. */
const XLSX_RUNTIME_IDS = Object.freeze([XLSX_RUNTIME_ID, '9.3.4|euro-office-bridge-v1/9.3.4.60/6', '9.3.4|euro-office-bridge-v1/9.3.4.60/7', '9.3.4|euro-office-bridge-v1/9.3.4.60/8']);

/** M11.5a.4 / M11.5b.1 / M11.5c.1: the B05-WRITE Bridge capability's runtimes (/6 + /7 + /8), pinned to the core value by a static test. */
const XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze(['9.3.4|euro-office-bridge-v1/9.3.4.60/6', '9.3.4|euro-office-bridge-v1/9.3.4.60/7', '9.3.4|euro-office-bridge-v1/9.3.4.60/8']);

/** M11.5b.4 / M11.5c.1: the B10-WRITE Bridge capability's runtimes (/7 + /8), pinned to the core value by a static test. */
const XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze(['9.3.4|euro-office-bridge-v1/9.3.4.60/7', '9.3.4|euro-office-bridge-v1/9.3.4.60/8']);

/** M11.5c.4: the B03-WRITE Bridge capability's own runtime (/8 only), pinned to the core value by a static test. */
const XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS = Object.freeze(['9.3.4|euro-office-bridge-v1/9.3.4.60/8']);

module.exports = { XLSX_RUNTIME_ID, XLSX_RUNTIME_IDS, XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS, XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS };
