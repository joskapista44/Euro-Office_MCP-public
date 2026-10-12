'use strict';

/**
 * The M11 READ-tool fakes (moved verbatim from test/acceptance/m11-3a-read-tools.test.cjs for reuse by the M11.3b.2
 * release acceptance): a recording fake broker, getters-only fake XLSX / DOCX editors behind a fake Playwright, and the
 * non-mutation assertion. Only the editor, the broker and the timings are fakes; drivers and adapters are the real ones.
 */

const assert = require('node:assert/strict');
const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');
const { createBrokerConfigSource } = require('../../src/server/broker-config-source.cjs');
const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { installBridge } = require('./ds-bridge-runtime.cjs');

const BROKER = 'http://127.0.0.1:9/api/euro-broker';
const TOKEN = 'tok-zero-FIXTURE-0000';
const FILE = '1302148';

function editConfig(fileId, fileType) {
  return {
    document: { fileType, key: `KEY-${fileId}`, permissions: { edit: true }, title: `fixture.${fileType}`, url: 'https://ds.example/download',
      referenceData: { fileKey: fileId, instanceId: 'instfixture' } },
    documentType: fileType === 'xlsx' ? 'cell' : 'word',
    editorConfig: { user: { id: 'instfixture_zero', name: 'zero' }, callbackUrl: 'https://nc.example/callback' },
    documentServerUrl: 'https://ds.example/', token: 'ds-jwt-FIXTURE',
  };
}

/** A fake broker (own token -> own edit config). Records every request. */
function fakeBroker(fileType) {
  const calls = [];
  const fetchImpl = async (url, init) => {
    const body = JSON.parse(init.body);
    calls.push({ bearer: String(init.headers.Authorization).replace(/^Bearer /, ''), body });
    return { status: 200, json: async () => ({ ok: true, status: 200, config: editConfig(body.fileId, fileType) }) };
  };
  return { calls, source: createBrokerConfigSource({ brokerUrl: BROKER, readToken: () => TOKEN, fetchImpl }) };
}

const fakeWs = url => ({ url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { if (this.handlers[ev]) this.handlers[ev]({ payload }); } });
const AUTH_FRAME = `42${JSON.stringify(['message', { type: 'auth', result: 1 }])}`;

// ---- the XLSX fake editor: an in-memory, getters-only Api ----------------------------------------------------------

const ro = (o, log) => new Proxy(o, { get(t, k) { if (typeof k === 'string' && /^[A-Z]/.test(k) && !(k in t)) log.push(`missing:${k}`); return t[k]; } });

function xlsxApi(log) {
  const data = { A1: 'Item', B1: 'Amount', A2: 'Rent', B2: 1200, A3: 'Food', B3: 340 };
  const colNum = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const cells = a1 => { const [p, q] = a1.replace(/\$/g, '').split(':'); const m1 = /^([A-Z]+)(\d+)$/.exec(p); const m2 = /^([A-Z]+)(\d+)$/.exec(q || p);
    const out = []; for (let r = Number(m1[2]); r <= Number(m2[2]); r += 1) for (let c = colNum(m1[1]); c <= colNum(m2[1]); c += 1) out.push(`${colName(c)}${r}`); return out; };
  const cell = a => ro({ GetAddress: () => a, GetValue: () => (a in data ? String(data[a]) : ''), GetFormula: () => (a in data ? String(data[a]) : ''),
    MergeArea: ro({ GetAddress: () => a }, log), GetNumberFormat: () => 'General', GetFillColor: () => null }, log);
  const validation = ro({ GetType: () => 'xlValidateWholeNumber', GetAlertStyle: () => 'xlValidAlertStop', GetOperator: () => 'xlBetween', GetFormula1: () => '0', GetFormula2: () => '10000' }, log);
  const range = a1 => ro({ GetAddress: () => a1, GetValue: () => cell(cells(a1)[0]).GetValue(), ForEach: fn => cells(a1).forEach(a => fn(cell(a))),
    GetValidation: () => (a1 === 'B2:B3' ? validation : null) }, log);
  const ws = ro({ GetName: () => 'Budget', GetIndex: () => 0, GetUsedRange: () => range('A1:B3'), GetRange: a1 => range(a1), GetAllDrawings: () => [] }, log);
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'Budget' ? ws : null), GetActiveSheet: () => ws,
    WorksheetFunction: ro({ COUNT: r => (typeof data[r.GetAddress()] === 'number' ? 1 : 0) }, log) };
  installBridge(api);
  return ro(api, log);
}

function xlsxEditor() {
  const log = [];
  const Api = xlsxApi(log);
  const frame = {
    url: () => 'https://ds.example/9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b/web-apps/apps/spreadsheeteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && typeof arg.body === 'string') {
        log.push(arg.body === HANDSHAKE_BODY ? 'handshake' : 'read-body');
        try { return { ok: true, outcome: 'command-completed', result: new Function('Api', arg.body)(Api) ?? null }; } catch (err) { log.push(`threw:${err}`); return { ok: false, outcome: 'call-threw', error: String(err) }; }
      }
      const src = String(fn);
      if (src.includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: false }; }
      if (src.includes('forceSave')) { log.push('forceSave'); return { ok: true, sent: true }; }
      return { ok: true };
    },
  };
  const ctl = { wsHandler: null, ws: null };
  const page = { on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; }, waitForFunction: async () => {}, waitForTimeout: async () => {}, frames: () => [frame],
    // the editor <-> DS socket of the session opens with the page (authenticated, no save frames)
    goto: async () => { ctl.ws = fakeWs(`wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY-${FILE}&EIO=4&transport=websocket`); ctl.wsHandler(ctl.ws); ctl.ws.emit('framereceived', AUTH_FRAME); },
    evaluate: async fn => { if (String(fn).includes('requestClose')) log.push('requestClose'); return { ok: true, detail: 'requestClose-called' }; } };
  let launches = 0;
  const browser = { newPage: async () => page, close: async () => { log.push('browser-close'); if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); } };
  return { log, api: { dispatch: 0, forceSave: 0 }, launches: () => launches, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => { launches += 1; return browser; } } } }) };
}

// the three generated property-query bodies (B21: default | not requested | runtime not verified)
const PROPERTY_QUERIES = Object.freeze({ [docx.propertyQuery()]: 'default',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'not-requested' })]: 'opt-out',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'runtime-not-verified' })]: 'gated' });

// ---- the DOCX fake editor: answers the driver's real property query ------------------------------------------------

const HTML = '<h1>Budget</h1><p>Hello world</p><ul><li>one</li><li>two</li></ul><table><tr><td>a</td><td>b</td></tr></table>';
const PROPERTY_RESULT = {
  paragraphs: [{ text: 'Hello world', props: { bold: false }, format: { alignment: 'left' }, nonRunChildren: 0, runsCoverText: true,
    runs: [{ text: 'Hello world', start: 0, end: 11, format: { bold: false } }], hyperlinks: [], noteRefs: { footnotes: 0, endnotes: 0 },
    numbering: { level: 0, numId: 7 } }],
  document: { bookmarks: [{ name: 'bm1', text: 'Hello' }], footnotes: ['fn'], endnotes: [], inMainBody: true,
    tables: [{ rows: [[{ text: 'a', paragraphs: 1, simple: true, format: {} }, { text: 'b', paragraphs: 1, simple: true, format: {} }]], gridCalc: [100, 200] }],
    defaultText: { fontFamily: 'Arial', fontSize: 11 },
    sections: [{ header: { default: { paragraphs: 1, simple: true, paras: [{ text: 'Head', format: {} }] } }, footer: {}, titlePage: true, evenAndOdd: false }] },
};

function docxEditor() {
  const log = [];
  const ctl = { wsHandler: null, ws: null };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-x/web-apps/apps/documenteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && arg.body === HANDSHAKE_BODY) { log.push('handshake'); return { ok: true, outcome: 'command-completed', result: null }; }
      if (arg && PROPERTY_QUERIES[arg.body]) { log.push('read-query'); log.push(`read-query:${PROPERTY_QUERIES[arg.body]}`); return { ok: true, outcome: 'command-completed', result: PROPERTY_RESULT }; }
      if (arg && arg.body) { log.push('command'); return { ok: true, outcome: 'command-completed', result: null }; }
      const src = String(fn);
      if (src.includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: false }; }
      if (src.includes('e.forceSave()')) { log.push('forceSave'); return { ok: true, sent: true }; }
      if (src.includes('pluginMethod_GetFileHTML() ||')) { log.push('readLive'); return HTML; }
      return true;
    },
  };
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; },
    // the editor <-> DS socket of the session opens with the page (authenticated, no save frames)
    goto: async () => { ctl.ws = fakeWs(`wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY-${FILE}&EIO=4&transport=websocket`); ctl.wsHandler(ctl.ws); ctl.ws.emit('framereceived', AUTH_FRAME); },
    waitForFunction: async () => {}, waitForTimeout: async () => {}, frames: () => [frameObj],
    evaluate: async fn => { if (String(fn).includes('requestClose')) log.push('requestClose'); return { ok: true, detail: 'requestClose-called' }; },
  };
  let launches = 0;
  const browser = { newPage: async () => page, close: async () => { log.push('browser-close'); if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); } };
  return { log, api: { dispatch: 0, forceSave: 0 }, launches: () => launches, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => { launches += 1; return browser; } } } }) };
}

const NON_MUTATION = (ed, out) => {
  const { log } = ed;
  assert.deepEqual(ed.api, { dispatch: 0, forceSave: 0 }, 'driver dispatch/forceSave never called on a READ');
  // WRITE=0 / DOCUMENT_MUTATION=0: no write command; HISTORY_MUTATION=0 / SAVE_REQUIRED=NO: no forceSave, close settled without save traffic
  assert.equal(log.includes('command'), false, 'no write command dispatched');
  assert.equal(log.includes('forceSave'), false, 'no forceSave');
  assert.equal(log.some(l => l.startsWith('threw:')), false, 'no body touched a missing (write) API');
  assert.match(String(out.close), /no-save-traffic/, `close: ${out.close}`);
};

module.exports = { BROKER, TOKEN, FILE, editConfig, fakeBroker, xlsxEditor, docxEditor, NON_MUTATION, PROPERTY_QUERIES, PROPERTY_RESULT, HTML };
