'use strict';

/**
 * M11.3a.1 / M11.5e.4 -- B21 the DOCX section VARIANT STATE read, CUT OVER (owner: M11.5 continuous mandate, card
 * a4e91e93 #ref) from the internal getters (ApiSection.Section.Get_TitlePage / IsEvenAndOdd) to the word /3 Bridge
 * READ capability docx.section.variantState.read/v1:
 *   A  the product READ opts out (readLive { includeSectionVariantState: false }) -> the Bridge is never called,
 *      and the accepted READ projections are byte-identical;
 *   C  without the opt-out the shared readLive is unchanged: on the verified (/3) runtime the WRITE family still gets
 *      the state;
 *   G  runtime gate: on a runtime without the capability (/1, /2, or fully unknown) the Bridge is never called
 *      (UNAVAILABLE:runtime-not-verified);
 *   H  three states (true | false | UNAVAILABLE:<why>); the header/footer WRITE family refuses an unknown state before
 *      dispatch and never counts unknown-before == unknown-after as preserved.
 * The generated query bodies run against an in-memory Api whose bridge counts every read() call.
 */

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');

const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const hf = require('../../src/adapters/docx/families/header-footer.cjs');
const { createReadTool } = require('../../src/server/read-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

const FILE = '1302148';
const VERIFIED_PATCH = 'euro-office-bridge-word-v1/9.3.4.60/1';
const VERIFIED_PATCH_V2 = 'euro-office-bridge-word-v1/9.3.4.60/2';
const VERIFIED_PATCH_V3 = 'euro-office-bridge-word-v1/9.3.4.60/3';
const CAP = BRIDGE_CAPABILITIES.DOCX_SECTION_VARIANT_STATE_READ;
const HTML = '<h1>Budget</h1><p>Hello world</p><table><tr><td>a</td></tr></table>';

/**
 * An in-memory Word Api: one paragraph, one section with a header part; the bridge's read() is counted. The capability
 * is advertised ONLY when `patchVersion` is /3 and `mode !== 'capability-missing'` -- matching the real /2 vs /3 blocks.
 */
function wordApi(spy, { patchVersion = VERIFIED_PATCH, mode = 'ok', sections = [{ titlePage: true, evenAndOdd: false }] } = {}) {
  // the paragraph/run format readers the query touches answer null (unread) -- the format read is not under test here
  const getters = o => new Proxy(o, { get: (t, k) => (k in t ? t[k] : (typeof k === 'string' && /^(Get|Is)/.test(k) ? () => null : undefined)) });
  const pr = () => getters({});
  const rng = () => getters({ GetTextPr: pr });
  const run = getters({ GetClassType: () => 'run', GetText: () => 'Head', GetRange: rng });
  const headerPara = getters({ GetClassType: () => 'paragraph', GetText: () => 'Head', GetElementsCount: () => 1, GetElement: () => run, GetRange: rng, GetParaPr: pr, GetTextPr: pr });
  const header = { GetElementsCount: () => 1, GetElement: () => headerPara };
  const section = { GetHeader: v => (v === 'default' ? header : null), GetFooter: () => null };
  const paragraph = { GetClassType: () => 'paragraph', GetText: () => 'Hello world', GetElementsCount: () => 0, GetElement: () => null };
  const doc = { GetElementsCount: () => 1, GetElement: i => (i === 0 ? paragraph : null), GetAllTables: () => [], Search: () => [], GetSections: () => [section],
    GetAllBookmarksNames: () => [], GetFootnotesFirstParagraphs: () => [], GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null };
  const hasCap = patchVersion === VERIFIED_PATCH_V3 && mode !== 'capability-missing';
  const manifest = { bridgeVersion: '1.0.0', documentServerVersion: '9.3.4', patchVersion, capabilities: hasCap ? [{ id: CAP.id, version: 1, kind: 'READ', format: 'DOCX', mutation: 'NO' }] : [] };
  const read = (capabilityId, args) => {
    spy.reads += 1;
    if (mode === 'throw') throw new Error('bridge-boom');
    if (mode === 'refuse') return { ok: false, reason: 'section-out-of-range' };
    if (mode === 'malformed') return { ok: true, value: { titlePage: 'yes', evenAndOdd: false } };
    const s = sections[args.section];
    return s ? { ok: true, value: { titlePage: s.titlePage, evenAndOdd: s.evenAndOdd } } : { ok: false, reason: 'section-out-of-range' };
  };
  return { GetDocument: () => doc, GetEuroOfficeBridge: () => ({ manifest: () => manifest, read }) };
}

const run = (Api, body) => new Function('Api', body)(Api); // eslint-disable-line no-new-func

// ---- the generated body itself -------------------------------------------------------------------------------------

test('the generated body: on word /3 it calls the Bridge READ ONCE per section; opt-out / gated bodies call it 0 times and do not even name the capability id', () => {
  const spy = { reads: 0 };
  const r = run(wordApi(spy, { patchVersion: VERIFIED_PATCH_V3 }), docx.propertyQuery());
  assert.equal(spy.reads, 1, 'one Bridge read per section carries BOTH flags');
  assert.deepEqual([r.document.sections[0].titlePage, r.document.sections[0].evenAndOdd], [true, false]);
  for (const reason of ['not-requested', 'runtime-not-verified']) {
    const s0 = { reads: 0 };
    const body = docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: reason });
    assert.equal(body.includes(CAP.id), false, `${reason}: the body does not even name the capability`);
    const x = run(wordApi(s0, { patchVersion: VERIFIED_PATCH_V3 }), body);
    assert.equal(s0.reads, 0, reason);
    assert.deepEqual([x.document.sections[0].titlePage, x.document.sections[0].evenAndOdd], [`UNAVAILABLE:${reason}`, `UNAVAILABLE:${reason}`]);
    assert.equal(x.document.sections[0].header.default.paras[0].text, 'Head', 'the PUBLIC header read is untouched');
  }
  assert.equal(docx.PROPERTY_QUERY, docx.propertyQuery(), 'the exported constant stays the default body');
});

test('any Bridge failure (capability missing, explicit refusal, a thrown read, or a malformed shape) is UNAVAILABLE:bridge-unavailable on BOTH flags, never null or a guessed boolean', () => {
  for (const mode of ['capability-missing', 'refuse', 'throw', 'malformed']) {
    const spy = { reads: 0 };
    const sec = run(wordApi(spy, { patchVersion: VERIFIED_PATCH_V3, mode }), docx.propertyQuery()).document.sections[0];
    assert.deepEqual([sec.titlePage, sec.evenAndOdd], ['UNAVAILABLE:bridge-unavailable', 'UNAVAILABLE:bridge-unavailable'], mode);
  }
  const inv = normalizeInventory({ html: HTML, paragraphs: [], document: { sections: [{ header: {}, footer: {} }] } });
  assert.deepEqual([inv.document.sections[0].titlePage, inv.document.sections[0].evenAndOdd], ['UNAVAILABLE:unread', 'UNAVAILABLE:unread'], 'absent -> UNAVAILABLE, not null');
});

// ---- the real driver ---------------------------------------------------------------------------------------------------

function editor(apiOpts) {
  const spy = { reads: 0 };
  const Api = wordApi(spy, apiOpts);
  const ctl = { wsHandler: null, ws: null };
  const frame = {
    url: () => 'https://ds.example/9.3.4-x/web-apps/apps/documenteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && typeof arg.body === 'string') {
        try { return { ok: true, outcome: 'command-completed', result: run(Api, arg.body) ?? null }; } catch (err) { return { ok: false, outcome: 'call-threw', error: String(err) }; }
      }
      const src = String(fn);
      if (src.includes('asc_isDocumentCanSave')) return { ok: true, canSave: false };
      if (src.includes('pluginMethod_GetFileHTML() ||')) return HTML;
      return true;
    },
  };
  const ws = url => ({ url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { if (this.handlers[ev]) this.handlers[ev]({ payload }); } });
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; }, waitForFunction: async () => {}, waitForTimeout: async () => {}, frames: () => [frame],
    goto: async () => { ctl.ws = ws(`wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY-${FILE}&EIO=4&transport=websocket`); ctl.wsHandler(ctl.ws); ctl.ws.emit('framereceived', `42${JSON.stringify(['message', { type: 'auth', result: 1 }])}`); },
    evaluate: async () => ({ ok: true, detail: 'requestClose-called' }),
  };
  const browser = { newPage: async () => page, close: async () => { if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); } };
  return { spy, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => browser } } }) };
}

const configSource = { kind: 'BROKER', getEditorConfig: async ({ fileId }) => ({ document: { fileType: 'docx', key: `KEY-${fileId}`, referenceData: { fileKey: fileId } }, documentServerUrl: 'https://ds.example/' }) };
const drivers = [];
after(async () => { for (const d of drivers) await d.close().catch(() => null); });

async function liveRead(apiOpts, request) {
  const ed = editor(apiOpts);
  const d = docx.createEuroOfficeDocxDriver({ configSource, loadPlaywright: ed.loadPlaywright, pollMs: 5, closeAckTimeoutMs: 200, timeoutMs: 2000 });
  drivers.push(d);
  const o = await d.open({ fileId: FILE, operation: 'READ' });
  assert.equal(o.ok, true, JSON.stringify(o));
  const r = request === undefined ? await d.readLive() : await d.readLive(null, request);
  await d.close();
  drivers.pop();
  return { r, spy: ed.spy, inv: r.ok ? normalizeInventory(r.payload) : null };
}

test('the handshake precondition: the fake serves manifests the handshake can parse; B21 is proven on word /3 ONLY', () => {
  assert.equal(run(wordApi({ reads: 0 }), HANDSHAKE_BODY) !== undefined, true);
  assert.deepEqual([...docx.B21_BRIDGE_READ_RUNTIMES], [`9.3.4|${VERIFIED_PATCH_V3}`]);
  assert.ok(Object.isFrozen(docx.B21_BRIDGE_READ_RUNTIMES));
});

test('SHARED DEFAULT UNCHANGED: readLive() on word /3 reads the Bridge once and the WRITE family gets a known state', async () => {
  const { r, spy, inv } = await liveRead({ patchVersion: VERIFIED_PATCH_V3 });
  assert.equal(r.ok, true, r.detail);
  assert.equal(r.payload.runtime.internalGate.sectionVariant, true);
  assert.equal(spy.reads, 1);
  assert.deepEqual([inv.document.sections[0].titlePage, inv.document.sections[0].evenAndOdd], [true, false]);
});

test('RUNTIME GATE: without the B21 Bridge READ capability (word /1, word /2, or a fully unknown runtime) readLive() never calls the Bridge for it -> UNAVAILABLE:runtime-not-verified', async () => {
  for (const [label, patchVersion] of [['word /1 (no B21 capability)', VERIFIED_PATCH], ['word /2 (no B21 capability)', VERIFIED_PATCH_V2], ['fully unknown', 'euro-office-bridge-word-v1/9.3.4.60/4']]) {
    const { r, spy, inv } = await liveRead({ patchVersion });
    assert.equal(r.ok, true, r.detail);
    assert.equal(r.payload.runtime.internalGate.sectionVariant, false, label);
    assert.equal(spy.reads, 0, label);
    assert.deepEqual([inv.document.sections[0].titlePage, inv.document.sections[0].evenAndOdd], ['UNAVAILABLE:runtime-not-verified', 'UNAVAILABLE:runtime-not-verified'], label);
  }
});

test('OPT-OUT: readLive({ includeSectionVariantState: false }) on word /3 calls the Bridge 0 times', async () => {
  const { r, spy, inv } = await liveRead({ patchVersion: VERIFIED_PATCH_V3 }, { includeSectionVariantState: false });
  assert.equal(r.ok, true, r.detail);
  assert.equal(spy.reads, 0);
  assert.equal(inv.document.sections[0].titlePage, 'UNAVAILABLE:not-requested');
  assert.equal(inv.document.sections[0].header.default.paras[0].text, 'Head');
});

test('PRODUCT READ: the five accepted DOCX operations call the Bridge READ 0 times, with output IDENTICAL to a default-read projection (on word /3, where the control DOES read it)', async () => {
  const OPS = ['read_text', 'read_formatting', 'read_tables', 'read_headers_footers', 'read_bookmarks'];
  const tool = (forceDefault, spies) => createReadTool('docx', { configSource, loadPlaywright: () => ({ ok: true }), createDriver: () => {
    const ed = editor({ patchVersion: VERIFIED_PATCH_V3 });
    spies.push(ed.spy);
    const d = docx.createEuroOfficeDocxDriver({ configSource, loadPlaywright: ed.loadPlaywright, pollMs: 5, closeAckTimeoutMs: 200, timeoutMs: 2000 });
    if (forceDefault) { const real = d.readLive.bind(d); d.readLive = () => real(); } // the control: the shared default read
    drivers.push(d);
    return d;
  } });
  const productSpies = [];
  const controlSpies = [];
  for (const op of OPS) {
    const p = await tool(false, productSpies).handler({ fileId: FILE, operation: op });
    const c = await tool(true, controlSpies).handler({ fileId: FILE, operation: op });
    assert.equal(p.ok, true, JSON.stringify(p));
    assert.deepEqual(p.result, c.result, `${op}: READ_OUTPUT_PARITY`);
  }
  assert.deepEqual(productSpies.map(s => s.reads), [0, 0, 0, 0, 0], 'ACCEPTED_DOCX_BRIDGE_READS=0');
  assert.deepEqual(controlSpies.map(s => s.reads), [1, 1, 1, 1, 1], 'positive control: the default read does call the Bridge, once');
});

// ---- the WRITE family: unknown -> fail closed ------------------------------------------------------------------------

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const part = (...texts) => ({ paras: texts.map(t => ({ text: t, props: { ...PLAIN }, alignment: 'left' })) });
const sec = h => ({ header: { default: null, even: null, title: null, ...h }, footer: { default: null, even: null, title: null }, titlePage: false });

async function hfRun(script) {
  const s = createStore([{ text: 'Body first.', props: { ...PLAIN } }], { doc: { sections: [sec({ default: part('Old') })], evenAndOdd: false } });
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const result = await withSession({ driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    docxAdapter, createExecutor({ id: 'b21', steps: [{ intent: 'edit_header', section: { index: 0, of: 1 }, text: 'New', variant: 'default' }] }));
  return { result, x: result.execution, m: result.metrics };
}

test('UNKNOWN STATE -> REFUSE before dispatch: a default-variant header edit with an unavailable section state writes nothing', async () => {
  const known = await hfRun({});
  assert.equal(known.x.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'positive control: a known state still edits');
  const { x, m } = await hfRun({ flagsUnreadable: true });
  assert.equal(x.appliedOperations[0].status, OPERATION_STATUS.FAILED);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  assert.match(JSON.stringify(x.appliedOperations[0].evidence), /section-variant-state-unavailable-section-0/);
});

test('PRESERVATION: unknown-before == unknown-after is NOT preserved; a real false == false is', () => {
  const op = { variant: 'default', sectionIndex: 0 };
  const u = 'UNAVAILABLE:runtime-not-verified';
  assert.deepEqual(hf.flagChanges([{ titlePage: u, evenAndOdd: u }], [{ titlePage: u, evenAndOdd: u }], op), ['0:titlePage-unknown', '0:evenAndOdd-unknown']);
  assert.deepEqual(hf.flagChanges([{ titlePage: false, evenAndOdd: false }], [{ titlePage: false, evenAndOdd: false }], op), []);
});

// ---- WRITE-path dispatch gate -----------------------------------------------------------------------------------------

test('WRITE RUNTIME GATE: a title/even command carries bridgeSectionVariantRead and the driver refuses it off word /3, before anything is sent; default carries neither flag', () => {
  const target = v => ({ kind: 'section_part', sectionIndex: 0, sectionCount: 1, part: 'header', variant: v, create: false, activate: false });
  for (const v of ['title', 'even']) {
    const cmd = hf.lower({ intent: 'edit_header', text: 'x' }, target(v));
    assert.equal(cmd.meta.bridgeSectionVariantRead, true);
    for (const runtime of [`9.3.4|${VERIFIED_PATCH}`, `9.3.4|${VERIFIED_PATCH_V2}`, '9.3.4|euro-office-bridge-word-v1/9.3.4.60/4', null]) {
      const refused = docx.internalWriteRefusal(cmd, runtime);
      assert.deepEqual([refused.ok, refused.dispatched, refused.outcome, refused.proven], [false, false, 'b21-bridge-read-runtime-not-proven', [...docx.B21_BRIDGE_READ_RUNTIMES]], String(runtime));
    }
    assert.equal(docx.internalWriteRefusal(cmd, `9.3.4|${VERIFIED_PATCH_V3}`), null, 'word /3: let through');
  }
  const d = hf.lower({ intent: 'edit_header', text: 'x' }, target('default'));
  assert.equal(d.meta.bridgeSectionVariantRead, undefined);
  assert.equal(d.meta.internal, undefined);
});
