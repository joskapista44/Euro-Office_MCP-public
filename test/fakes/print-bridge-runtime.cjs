'use strict';

const fs = require('node:fs');
const path = require('node:path');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const { b10BridgeWritePrelude } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): was the 9.3.4.60 rev7-candidate bridge-block
// (the first 9.3.4.60 revision to carry sheet.printSettings.write/v1) -- retired with the rest of the 9.3.4.60
// family. The 9.3.5.0 block carries the SAME capability (MEASURED: grep '"sheet.printSettings.write/v1"' ->
// present), and self-reports the identity XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES now actually accepts.
const V935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');

/**
 * M11.5b.4 (DS9346: now the 9.3.5.0 block): runs the PRODUCT's print command -- the driver's injected B10 client +
 * the adapter's lowered body -- against the REAL bridge block, over a fake sheet `ws` ({ GetName, worksheet:
 * <recording model> }) and AscCommon. The model calls the old direct-internal body made are now made by
 * sheet.printSettings.write/v1; the tests pin them end to end.
 */
function runViaBridge(op, ws, ascCommon) {
  const api = { GetSheet: n => (n === ws.GetName() ? ws : null) };
  new Function('Api', 'AscCommon', `${V935};`)(api, ascCommon);
  return new Function('Api', b10BridgeWritePrelude() + print.lower(op).body)(api);
}

module.exports = { runViaBridge, V935 };
