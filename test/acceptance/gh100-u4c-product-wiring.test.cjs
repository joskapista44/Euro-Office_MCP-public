'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { book, FILE } = require('../fakes/write-wiring.cjs');

/**
 * GH #100 U4c -- a maintainer's review #ref: the read/write WIRING through the REAL product assembly
 * (buildServerTools), not through createReadTool/createWriteTool constructed by hand. The U4c code
 * (session-reuse-rule.cjs, lifecycle.cjs, read-tools.cjs) was correct, but server-tools.cjs never
 * passed `reuse` to xlsx_read -- so in the shipped product xlsx_read NEVER took or held a session;
 * the entire feature only ever lived on the test seam (half-wired-feature-must-not-land).
 *
 * RED-first baseline: on the pre-fix server-tools.cjs (`xlsx_read: createReadTool('xlsx',
 * { configSource, browsersPath, ...readTool, gate })`, no `reuse`), every assertion below that
 * checks a read taking/holding a session fails -- xlsx_read always opens fresh and always closes.
 */

const ON = { EURO_MCP_SESSION_REUSE: '1', EURO_MCP_SESSION_IDLE_MS: '60000', EURO_MCP_SESSION_MAX_AGE_MS: '600000', EURO_MCP_SESSION_MAX_CALLS: '50' };

function wrapDriver(d, ctl) {
  return Object.assign(d, {
    resetForNextCall() { return { ok: true }; },
    async health() { return { alive: true, canSave: false, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 }; },
    documentKey() { return ctl.key; },
    async release() { return { ok: true, detail: 'kept', confirmation: d.calls.forceSave > 0 ? { state: 'confirmed', reason: 'forcesave-confirmed' } : null }; },
  });
}

function setup() {
  const store = createStore(book());
  const ctl = { key: 'KEY-1' };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: ctl.key } }; } };
  const drivers = [];
  const createDriver = () => { const d = wrapDriver(createFakeXlsxDriver(store, {}), ctl); drivers.push(d); return d; };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  let t = 0;
  const session = { ready: { wait: async () => {} }, persistence: { wait: async () => {} }, now: () => (t += 1) };
  const built = buildServerTools({
    config: { broker: { url: 'http://127.0.0.1:1/b' } },
    env: ON,
    readTool: { configSource, runtime, createDriver },
    writeTool: { configSource, runtime, createDriver, session },
  });
  const sumOpen = () => drivers.reduce((a, d) => a + d.calls.open, 0);
  const sumClose = () => drivers.reduce((a, d) => a + d.calls.close, 0);
  const step = v => ({ intent: 'write_range', id: 'w', sheet: 'Budget', range: 'F2', values: [[v]] });
  const write = v => built.tools.xlsx_execute_task.handler({ fileId: FILE, task: { id: `t${v}`, steps: [step(v)] } });
  const read = () => built.tools.xlsx_read.handler({ fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'F2:F2' });
  return { built, drivers, sumOpen, sumClose, write, read };
}

test('product wiring: xlsx_read, assembled by buildServerTools with reuse ON, holds a session instead of closing it', async () => {
  const s = setup();
  const r1 = await s.read();
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.equal(s.sumOpen(), 1);
  assert.equal(s.sumClose(), 0, 'xlsx_read, as wired into the real server tool set, must HOLD a clean read');

  const r2 = await s.read();
  assert.equal(r2.ok, true, JSON.stringify(r2));
  assert.equal(s.sumOpen(), 1, 'the second read must TAKE the first xlsx_read call`s session, not open a new one');
});

test('product wiring: write -> read -> write through the REAL assembled tools, ONE shared session (the exact sequence zero\'s review named)', async () => {
  const s = setup();
  const w1 = await s.write(1);
  assert.equal(w1.ok, true, JSON.stringify(w1));
  assert.equal(s.sumOpen(), 1);
  assert.equal(s.sumClose(), 0);

  const r1 = await s.read();
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.equal(s.sumOpen(), 1, 'xlsx_read must take the session xlsx_execute_task left warm -- the product wiring, not a hand-built tool, must do this');
  assert.equal(s.sumClose(), 0);

  const w2 = await s.write(2);
  assert.equal(w2.ok, true, JSON.stringify(w2));
  assert.equal(s.sumOpen(), 1, 'the second write must take back the session the read used -- the SAME registry slot the write side shares with xlsx_read');
  assert.equal(s.sumClose(), 0);
});

// GH #219 3: docx's OWN warm-session wiring, through the REAL buildServerTools -- the driver-surface
// tripwire this file used to carry (GH #220 1b closed it: docx-driver.cjs now has release/
// resetForNextCall/health/documentKey) is no longer this file's job to guard; docs/GH220-DOCX-
// WRITE-LIFECYCLE-DESIGN.md section 5 (as amended by GH #220 1b) is the record. What THIS pair of
// tests proves: (a) EURO_MCP_SESSION_REUSE alone (the xlsx switch, ON) must NOT turn docx reuse on
// by itself; (b) with the SECOND, docx-specific switch also on, docx_read behaves exactly like
// xlsx_read already does above (holds a clean read instead of closing it).
const BOTH_ON = { ...ON, EURO_MCP_DOCX_SESSION_REUSE: '1' };

function docxSetup(env) {
  const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');
  const store = createStore([{ text: 'Hello', props: {} }]);
  const ctl = { key: 'KEY-1' };
  const drivers = [];
  const createDriver = () => { const d = Object.assign(createFakeDocxDriver(store, { saveObserver: true, documentKey: () => ctl.key }), {}); drivers.push(d); return d; };
  const configSource = { kind: 'BROKER', async getEditorConfig() { return { document: { key: ctl.key } }; } };
  const runtime = { check: () => ({ ok: true }), load: () => ({ ok: false, detail: 'unused' }) };
  const built = buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1/b' } }, env, readTool: { configSource, runtime, createDriver } });
  const sumOpen = () => drivers.reduce((a, d) => a + d.calls.open, 0);
  const sumClose = () => drivers.reduce((a, d) => a + d.calls.close, 0);
  const read = () => built.tools.docx_read.handler({ fileId: '1303626', operation: 'read_text' });
  return { built, drivers, sumOpen, sumClose, read };
}

test('product wiring: EURO_MCP_SESSION_REUSE ON alone (the xlsx switch) does NOT turn docx_read reuse on -- two reads, two opens, two closes', async () => {
  const s = docxSetup(ON);
  const r1 = await s.read();
  assert.equal(r1.ok, true, JSON.stringify(r1));
  const r2 = await s.read();
  assert.equal(r2.ok, true, JSON.stringify(r2));
  assert.equal(s.sumOpen(), 2, 'EURO_MCP_DOCX_SESSION_REUSE was never set -- docx_read opens fresh every time, unchanged by this unit');
  assert.equal(s.sumClose(), 2);
});

test('product wiring: BOTH switches on -- docx_read, assembled by the REAL buildServerTools, holds a session instead of closing it', async () => {
  const s = docxSetup(BOTH_ON);
  const r1 = await s.read();
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.equal(s.sumOpen(), 1);
  assert.equal(s.sumClose(), 0, 'docx_read, as wired into the real server tool set with its own switch ON, must HOLD a clean read');
  const r2 = await s.read();
  assert.equal(r2.ok, true, JSON.stringify(r2));
  assert.equal(s.sumOpen(), 1, 'the second read must TAKE the first docx_read call\'s session, not open a new one');
});

// GH #219 3: docx_read <-> docx_execute_task cross-tool sharing cannot be exercised BEHAVIOURALLY at
// this (product-registry) level today -- DOCX_WRITE_ALLOWLIST is Object.freeze([]), so every
// docx_execute_task step is refused REFUSED_NOT_ALLOWLISTED at PRE_DISPATCH, before any driver ever
// opens (docx-write-tools.cjs:49-50); no task can reach registry.take()/hold() through this tool while
// the allowlist stays empty (releasing it is an owner decision, out of this unit's scope). What this
// unit CAN and does prove offline is the WIRING: docx_read and docx_execute_task are constructed from
// the SAME `docxReuse` object (identical key fields), so sharing is correct BY CONSTRUCTION, ready the
// moment an intent is released -- the behavioural proof (write holds, read takes, write takes back) is
// exactly what `gh220-docx-write-lifecycle.test.cjs` already demonstrates at the `createDocxWriteExecution`
// level (below the allowlist gate), and what `docx-write-lifecycle.live.cjs` will measure against the
// real DS editor once run.
test('product wiring: docx_read and docx_execute_task are built from the SAME reuse object (same key fields) -- cross-tool sharing is wired correctly by construction, even though no task can reach it today (DOCX_WRITE_ALLOWLIST is empty)', () => {
  const fs = require('node:fs');
  const path = require('node:path');
  const src = fs.readFileSync(path.join(__dirname, '..', '..', 'src', 'server', 'server-tools.cjs'), 'utf8');
  const readLine = src.split('\n').find(l => l.includes("docx_read: createReadTool"));
  const writeLine = src.split('\n').find(l => l.includes('docx_execute_task: createDocxWriteTool'));
  assert.ok(readLine && writeLine, 'both construction lines must exist');
  assert.match(readLine, /reuse:\s*docxReuse\b/);
  assert.match(writeLine, /reuse:\s*docxReuse\b/, 'the SAME identifier, not a second, lookalike object');
});
