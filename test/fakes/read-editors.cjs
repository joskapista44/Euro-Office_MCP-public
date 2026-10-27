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

/** The layout the fake XLSX editor answers (L-09): B is 20.43 wide, C hidden, row 2 is 30 high, row 3 hidden. */
const LAYOUT = Object.freeze({ widths: Object.freeze({ B: 20.43 }), hiddenCols: Object.freeze(['C']), heights: Object.freeze({ 2: 30 }), hiddenRows: Object.freeze([3]) });

/** The page layout the fake XLSX editor answers (L-10b): landscape, headings on, the public default margins in mm. */
const PAGE = Object.freeze({ orientation: 'xlLandscape', printGridlines: false, printHeadings: true, topMargin: 19.1, bottomMargin: 19.1, leftMargin: 17.8, rightMargin: 17.8 });

function xlsxApi(log) {
  const data = { A1: 'Item', B1: 'Amount', A2: 'Rent', B2: 1200, A3: 'Food', B3: 340 };
  const colNum = s => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const cells = a1 => { const [p, q] = a1.replace(/\$/g, '').split(':'); const m1 = /^([A-Z]+)(\d+)$/.exec(p); const m2 = /^([A-Z]+)(\d+)$/.exec(q || p);
    const out = []; for (let r = Number(m1[2]); r <= Number(m2[2]); r += 1) for (let c = colNum(m1[1]); c <= colNum(m2[1]); c += 1) out.push(`${colName(c)}${r}`); return out; };
  const cell = a => ro({ GetAddress: () => a, GetValue: () => (a in data ? String(data[a]) : ''), GetFormula: () => (a in data ? String(data[a]) : ''),
    MergeArea: ro({ GetAddress: () => a }, log), GetNumberFormat: () => 'General', GetFillColor: () => null }, log);
  const validation = ro({ GetType: () => 'xlValidateWholeNumber', GetAlertStyle: () => 'xlValidAlertStop', GetOperator: () => 'xlBetween', GetFormula1: () => '0', GetFormula2: () => '10000' }, log);
  // L-22e (card an internal branch): a type 'custom' rule with a
  // RELATIVE formula on C2:C4 -- the public Formula1 getter reports it relative to the SHEET's own
  // current active cell (A1, this fixture's fixed answer), not C2 (the rule's own range). This is
  // the EXACT shape the LIVE candidate (a maintainer, card comment #ref) found read_validations getting
  // wrong: the driver captured activeCell but the read-tool projection did not re-anchor formula1
  // before returning it. read-tools.cjs must shift it back to '=C2>0' (matching what was written).
  const customValidation = ro({ GetType: () => 'xlValidateCustom', GetAlertStyle: () => 'xlValidAlertStop', GetOperator: () => 'xlBetween', GetFormula1: () => '=A1>0', GetFormula2: () => '' }, log);
  // L-09: the column / row layout getters (GetEntireColumn / GetEntireRow), shaped as the W4b LIVE probe measured them: a hidden column reads width 0, a hidden row
  // height null, an unset row reads the engine default 15. LAYOUT is the fixture the L-09 tests assert against.
  const entireColumn = a1 => { const l = /^\$?([A-Z]+)/.exec(a1)[1]; const h = LAYOUT.hiddenCols.includes(l);
    return ro({ GetColumnWidth: () => (h ? 0 : (LAYOUT.widths[l] ?? 8.43)), GetHidden: () => h }, log); };
  const entireRow = a1 => { const r = /(\d+)$/.exec(a1)[1]; const h = LAYOUT.hiddenRows.includes(Number(r));
    return ro({ GetRowHeight: () => (h ? null : (LAYOUT.heights[r] ?? 15)), GetHidden: () => h }, log); };
  const range = a1 => ro({ GetAddress: () => a1, GetValue: () => cell(cells(a1)[0]).GetValue(), ForEach: fn => cells(a1).forEach(a => fn(cell(a))),
    GetValidation: () => (a1 === 'B2:B3' ? validation : (a1 === 'C2:C4' ? customValidation : null)), GetEntireColumn: () => entireColumn(a1), GetEntireRow: () => entireRow(a1) }, log);
  // X-O05/4: one fake chart over the Budget series -- series index 0, point 0 (Rent, 1200) falls through to SERIES
  // level (findDLblByIdx returns null there); point 1 (Food, 340) carries a POINT-level override. `.Chart.chart.
  // plotArea.charts[0].series[i].dLbls` is the REAL bridge-block's own internal shape (installBridge below runs the
  // actual chart.dataLabels.effective.read/v1 implementation, not an imitation) -- not a public Builder API surface,
  // so this one object is plain, not `ro()`-wrapped (nothing here should ever log a "missing:" public getter).
  // X-O05/4 round 2 (a maintainer's LIVE r1, internal tracker): GetValues() is a FORMULA STRING in reality ('$B$2:$B$3'),
  // never a plain array -- the fixture matches that now (a single-column, 2-row range: B2=Rent, B3=Food).
  const chart1 = {
    GetName: () => 'Spend', GetChartType: () => 'bar', GetTitle: () => null, GetLegendPos: () => 'none',
    GetAllSeries: () => [{ GetName: () => 'Amount', GetValues: () => '=Budget!$B$2:$B$3', GetCatFormula: () => '=Budget!$A$2:$A$3' }],
    GetHorAxisTitle: () => null, GetVerAxisTitle: () => null,
    GetDataLabels: () => ({ showSerName: false, showCatName: false, showVal: false, showPercent: false }),
    Chart: { chart: { plotArea: { charts: [{ dLbls: null, series: [{ dLbls: {
      showSerName: false, showCatName: false, showVal: false, showPercent: false,
      findDLblByIdx: idx => (idx === 1 ? { showSerName: true, showCatName: false, showVal: false, showPercent: false } : null),
    } }] }] } } },
  };
  // X-O08 (/14): one fake shape -- the SAME public-getter set M7-S2b measured LIVE (kind/name/parent/pos/
  // width/height/rotation/flip, docs/evidence/M7-S2B-DRAWING-PROBE.md), plus a stroke (X-O08-7: width PUBLIC-readable).
  // GetLine()/GetFill() are plain, not `ro()`-wrapped: the Bridge's OWN handler calls them on this same object when a
  // bridge revision serves FILL_LINE_COLOR_READ; the served fixture bridge is pinned at /5 (ds-bridge-runtime.cjs's
  // own manifest.json), so that capability degrades to CAPABILITY_MISSING here, same as every pre-/14 runtime.
  const shape1 = {
    GetClassType: () => 'shape', GetName: () => 'Logo', GetParentSheet: () => ({ GetName: () => 'Budget' }),
    GetPosition: () => ({ fromCol: 2, colOffset: 0, fromRow: 1, rowOffset: 0 }),
    GetWidth: () => 914400, GetHeight: () => 457200, GetRotation: () => 0, GetFlipH: () => false, GetFlipV: () => false,
    GetLine: () => ({ GetWidth: () => 19050, GetFill: () => ({ GetType: () => 'solid' }) }),
    GetFill: () => ({ GetType: () => 'solid' }),
  };
  // L-10b: the public page getters (set_page_layout reads, W8b) and a sheet with NO AutoFilter (GetAutoFilter -> null = the driver's { range: null, filters: [] }). The print / freeze / filter-FIELD reads go
  // through DS Bridge capabilities that need the editor's internal model, which this fake does not carry: they surface as NAMED 'UNOBSERVABLE:BRIDGE:...' gaps (asserted as such).
  // L-22e: the sheet's own active cell, fixed at A1 -- NEVER C2 (the custom rule's own range above),
  // so the fixture actually exercises the re-anchor rather than coincidentally matching it.
  const ws = ro({ GetName: () => 'Budget', GetIndex: () => 0,
    GetPageOrientation: () => PAGE.orientation, GetPrintGridlines: () => PAGE.printGridlines, GetPrintHeadings: () => PAGE.printHeadings,
    GetTopMargin: () => PAGE.topMargin, GetBottomMargin: () => PAGE.bottomMargin, GetLeftMargin: () => PAGE.leftMargin, GetRightMargin: () => PAGE.rightMargin, GetAutoFilter: () => null, GetUsedRange: () => range('A1:B3'), GetRange: a1 => range(a1), GetAllDrawings: () => [shape1],
    GetAllCharts: () => [chart1], GetAllPivotTables: () => [], GetActiveCell: () => cell('A1') }, log);
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

// the four generated property-query bodies (B21: default | not requested | runtime not verified;
// GH #225: not requested + numberingFormat -- read_numbering's own extra ToJSON() call, requested
// ONLY for that operation, read-tools.cjs FORMATS.docx.read)
const PROPERTY_QUERIES = Object.freeze({ [docx.propertyQuery()]: 'default',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'not-requested' })]: 'opt-out',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'runtime-not-verified' })]: 'gated',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'not-requested', numberingFormat: true })]: 'opt-out-numbering',
  // GH #269b: read_comments' own extra read (includeComments), requested ONLY for read_comments
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'not-requested', includeComments: true })]: 'opt-out-comments' });

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

// GH #269b: the rows the driver's page-side GetAllComments() read returns (the driver's own field names: solved, not resolved)
const COMMENT_ROWS = [{ id: 'c1', text: 'a note', author: 'Ann', userId: 'u1', timeUtc: 1792420516000, quote: 'Hello', solved: false, replies: [{ text: 'a reply', author: 'Bob', userId: 'u2' }] }];

function docxEditor() {
  const log = [];
  const ctl = { wsHandler: null, ws: null };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-x/web-apps/apps/documenteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && arg.body === HANDSHAKE_BODY) { log.push('handshake'); return { ok: true, outcome: 'command-completed', result: null }; }
      if (arg && PROPERTY_QUERIES[arg.body]) { log.push('read-query'); log.push(`read-query:${PROPERTY_QUERIES[arg.body]}`); return { ok: true, outcome: 'command-completed', result: PROPERTY_QUERIES[arg.body] === 'opt-out-comments' ? { ...PROPERTY_RESULT, document: { ...PROPERTY_RESULT.document, comments: COMMENT_ROWS } } : PROPERTY_RESULT }; }
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

module.exports = { LAYOUT, PAGE, BROKER, TOKEN, FILE, editConfig, fakeBroker, xlsxEditor, docxEditor, NON_MUTATION, PROPERTY_QUERIES, PROPERTY_RESULT, HTML };
