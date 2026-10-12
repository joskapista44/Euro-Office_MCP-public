'use strict';

const fs = require('node:fs');
const path = require('node:path');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const { b10BridgeWritePrelude } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

const V7 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'rev7-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');

/**
 * M11.5b.4: runs the PRODUCT's print command -- the driver's injected B10 client + the adapter's lowered body -- against
 * the REAL /7 bridge block, over a fake sheet `ws` ({ GetName, worksheet: <recording model> }) and AscCommon. The model
 * calls the old direct-internal body made are now made by sheet.printSettings.write/v1; the tests pin them end to end.
 */
function runViaBridge(op, ws, ascCommon) {
  const api = { GetSheet: n => (n === ws.GetName() ? ws : null) };
  new Function('Api', 'AscCommon', `${V7};`)(api, ascCommon);
  return new Function('Api', b10BridgeWritePrelude() + print.lower(op).body)(api);
}

module.exports = { runViaBridge, V7 };
