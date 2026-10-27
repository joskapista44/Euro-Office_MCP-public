'use strict';

/**
 * M11.4a.2 offline wiring harness: the product WRITE tool (`createWriteTool('xlsx')`, the canonical executor path) with
 * the fake XLSX transport (test/fakes/fake-xlsx-driver.cjs) behind the createDriver seam. The fake records the open
 * arguments, so the WRITE open mode is observable; `wrap(driver, counters)` injects transport faults.
 */

const { createStore, createFakeXlsxDriver } = require('./fake-xlsx-driver.cjs');
const { createWriteTool } = require('../../src/server/write-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');

const FILE = '1303626';

/** A workbook for the six intents: values, a merge and a chart (charts are seeded through the store). */
const book = () => [
  { name: 'Budget', cells: { A1: { v: 'Name' }, A2: { v: 'apple' }, A3: { v: 'pear' }, A4: { v: 'plum' }, A5: { v: 'kiwi' },
    B1: { v: 'Qty' }, B2: { v: 3 }, B3: { v: 7 }, B4: { v: 1 }, B5: { v: 9 }, C1: { v: 'merged' } }, merges: ['C1:D1'],
  objects: { charts: [], pivots: [], names: [] } },
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, A2: { v: 'a' }, B2: { v: 1 }, A3: { v: 'b' }, B3: { v: 2 } },
    objects: { charts: [], pivots: [], names: [] } },
];

/** Test SETUP only (not the product path): the chart R1 on K is created by the canonical set_chart, so its series
 * carry the shape the fake's chart reader serves (a hand-written chart object would read UNOBSERVABLE). */
async function seedChart(store) {
  const { withSession } = require('../../src/core/session/lifecycle.cjs');
  const { createExecutor } = require('../../src/core/execution/executor.cjs');
  const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
  const t = { id: 'seed', steps: [{ intent: 'set_chart', sheet: 'K', name: 'R1', range: 'A1:B3', type: 'bar', title: 'Sales' }] };
  let n = 0;
  const r = await withSession({ driver: createFakeXlsxDriver(store, {}), fileId: 'seed', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  if (!r.ok) throw new Error(`chart seed failed: ${r.outcome}`);
  return store;
}

/**
 * harness({ store, script, wrap, gate, persistence }) -> { tool, gate, n, drivers, opens }
 * n: { runtime, drivers, open, dispatch, forceSave, close } summed over every driver this harness created.
 */
function harness({ store = createStore(book()), script = {}, wrap = null, gate = createSessionGate(), persistence = null, configSource = { kind: 'BROKER' }, onTiming = null, keyRotation = null, reuse = null } = {}) {
  const drivers = [];
  const opens = [];
  const n = { runtime: 0, drivers: 0 };
  const createDriver = (format, opts) => {
    n.drivers += 1;
    if (format !== 'xlsx') throw new Error(`unexpected format ${format}`);
    if (!opts || opts.configSource !== configSource) throw new Error('the WRITE driver must get the broker configSource');
    const d = createFakeXlsxDriver(store, script);
    const open = d.open.bind(d);
    d.open = async args => { opens.push(args); return open(args); };
    drivers.push(d);
    return wrap ? wrap(d, n) : d;
  };
  const runtime = { check: () => { n.runtime += 1; return { ok: true }; }, load: () => ({ ok: false, detail: 'unused by the fake' }) };
  let t = 0;
  // GH #76 MODEL 1: keyRotation overrides verifyStoredContent's bounded poll (timeoutMs/pollMs/now/wait) so a test
  // never actually waits real wall-clock time.
  const session = { ready: { wait: async () => {} }, persistence: persistence || { wait: async () => {} }, now: () => (t += 1), ...(keyRotation ? { keyRotation } : {}) };
  const tool = createWriteTool('xlsx', { configSource, gate, runtime, createDriver, session, ...(onTiming ? { onTiming } : {}), ...(reuse ? { reuse } : {}) });
  const sum = k => drivers.reduce((a, d) => a + d.calls[k], 0);
  return { tool, gate, n, drivers, opens, store, calls: () => ({ open: sum('open'), dispatch: sum('dispatch'), forceSave: sum('forceSave'), close: sum('close') }) };
}

const req = (steps, id = 't') => ({ fileId: FILE, task: { id, steps } });

module.exports = { FILE, book, seedChart, harness, req, createStore };
