'use strict';

/**
 * GH #140 U1 (internal tracker): a STATIC, code-grounded check of docs/GH140-PIVOT-CONTRACT.md's
 * own claims -- repo-internal only (no network, no live editor, no git history), the same pattern
 * as test/static/gh139-chart-symmetry-matrix.test.cjs / gh141-cell-format-symmetry-matrix.test.cjs.
 *
 * The CENTRAL claim this file pins: `read_charts`'s `pivots[]` field (what a product-surface
 * `xlsx_read` caller actually receives) is a BARE NAME LIST, not the richer {name, source, table,
 * layout} shape `withProvenPivots()` produces -- contrary to what
 * docs/CAPABILITY-COMPLETION-STATUS-20261012.md's X-O10 row currently states. This is checked by
 * calling the driver's own `readQuery()` generator (the same technique
 * gh141-cell-format-symmetry-matrix.test.cjs uses for READ-side claims) and by running the FAKE
 * driver end-to-end through the real read_charts handler -- two independent code paths, same answer.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { XLSX_WRITE_ALLOWLIST } = require('../../src/server/write-tools.cjs');
const { OPERATIONS, STATUS } = require('../../src/server/read-tools.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const pivots = require('../../src/adapters/xlsx/families/pivots.cjs');
const pivotRefresh = require('../../src/adapters/xlsx/families/pivot-refresh.cjs');
const pivotSource = require('../../src/adapters/xlsx/pivot-source.cjs');

// ---- 1. T7 / section 2: create_pivot and refresh_pivot are adapter-complete. GH #140 U3 CLOSED this for
// create_pivot (the allowlisting itself is now present, test/acceptance/gh140-u3-create-pivot-release.test.cjs
// has the full RED-first behavioural coverage); refresh_pivot stays the U4 unit. ----

test('create_pivot and refresh_pivot are FULL adapter families (plan/classify/lower/verify); both are released on the product WRITE surface (GH #140 U3 create_pivot, U4 refresh_pivot)', () => {
  assert.deepEqual(pivots.INTENTS, Object.freeze(['create_pivot']));
  assert.deepEqual(pivotRefresh.INTENTS, Object.freeze(['refresh_pivot']));
  for (const fn of ['plan', 'classify', 'lower', 'verifyPrimitive', 'verifyPreservation', 'wholeTaskChecks', 'windowsOf']) {
    assert.equal(typeof pivots[fn], 'function', `pivots.cjs should export ${fn}`);
    assert.equal(typeof pivotRefresh[fn], 'function', `pivot-refresh.cjs should export ${fn}`);
  }
  assert.ok(XLSX_WRITE_ALLOWLIST.includes('create_pivot'), 'create_pivot is allowlisted since GH #140 U3');
  assert.ok(XLSX_WRITE_ALLOWLIST.includes('refresh_pivot'), 'refresh_pivot is allowlisted since GH #140 U4');
});

// ---- 2. GH #140 U2 CLOSED this: read_pivots was a status-table name only (REFUSED_PARTIAL, never
// dispatchable) at U1 time -- it is now ACCEPTED and dispatchable (test/acceptance/
// gh140-u2-read-pivots-release.test.cjs has the full RED-first behavioural coverage). ----

test('read_pivots is now ACCEPTED and dispatchable (GH #140 U2) -- U1 found it PARTIAL; this is the closure, not a drift', () => {
  assert.equal(OPERATIONS.xlsx.read_pivots.status, STATUS.ACCEPTED);
  assert.ok(require('../../src/server/read-tools.cjs').inputSchema('xlsx').properties.operation.enum.includes('read_pivots'));
});

// ---- 3. THE CENTRAL FINDING: read_charts's pivots[] field is name-only. Two independent checks: ----

// 3a. The driver's OWN generated read-script source: the exact line that builds chartReads[n].pivots.
test('the driver\'s chartReads builder maps pivots to BARE NAMES ((ws.GetAllPivotTables()||[]).map(p=>p.GetName())), never to an object carrying source/table/layout', () => {
  const script = readQuery([], { chartSheets: ['Sheet1'] });
  const m = script.match(/const pivots = ([^\n]+);\s*\n\s*return \{ charts, pivots \};/);
  assert.ok(m, 'the chartReads pivots-builder line should be findable in the generated script');
  assert.match(m[1], /GetAllPivotTables\(\)\s*\|\|\s*\[\]\)\.map\(p\s*=>\s*p\.GetName\(\)\)/, `chartReads pivots builder: ${m[1]}`);
  assert.ok(!m[1].includes('source') && !m[1].includes('table') && !m[1].includes('layout'), 'the chartReads pivots builder must not carry source/table/layout');
});

// 3b. GH #140 U2 CLOSED this too: at U1 time, inventory.cjs's `objects` axis (withProvenPivots()'s
// rich {name,source,table,layout} shape) had NO product-surface READ consumer. read_pivots is now
// exactly that consumer -- read_charts's OWN handler is unchanged (still projects inv.chartReads,
// the thin name-only axis), so the fix is additive, not a widening of read_charts.
test('read-tools.cjs now reads inv.objects via read_pivots (GH #140 U2) -- read_charts is untouched, still projecting inv.chartReads only', () => {
  const fs = require('node:fs');
  const src = fs.readFileSync(require.resolve('../../src/server/read-tools.cjs'), 'utf8');
  assert.ok(src.includes('inv.objects'), 'read_pivots should read inv.objects');
  assert.ok(src.includes('inv.chartReads'), 'read_charts should still read inv.chartReads (the thin, name-only axis)');
  // isolate read_charts's OWN handler body and confirm it still never mentions inv.objects
  const m = src.match(/read_charts: \(inv, a\) => \{[\s\S]*?\n  \},/);
  assert.ok(m, 'read_charts handler should be findable');
  assert.ok(!m[0].includes('inv.objects'), 'read_charts\'s own handler must not have been widened to read inv.objects');
});

test('pivot-source.cjs\'s provenSource()/withProvenPivots() (the RICH, content-proven {name,source,table,layout} shape) IS real, tested, exported code -- it exists, it is simply never projected to a READ operation', () => {
  assert.equal(typeof pivotSource.provenSource, 'function');
  assert.equal(typeof pivotSource.withProvenPivots, 'function');
  // the function's OWN shape: given a proof it cannot verify, it returns UNOBSERVABLE, never a guess
  assert.match(pivotSource.provenSource('not-an-object', 'bogus-runtime', ['S']), /^UNOBSERVABLE:/);
});

// ---- 4. RED CONTROL: a fabricated shape must correctly be reported as absent / not matching --
// proves the checks above discriminate, not pass by construction. ----

test('RED CONTROL: a fabricated chartReads pivots shape (an object literal, not GetName()) is correctly absent from the real generated script', () => {
  const script = readQuery([], { chartSheets: ['Sheet1'] });
  assert.ok(!script.includes('pivots: (ws.GetAllPivotTables() || []).map(p => ({ name: p.GetName(), source:'));
});

test('RED CONTROL: a made-up intent is correctly absent from XLSX_WRITE_ALLOWLIST', () => {
  assert.ok(!XLSX_WRITE_ALLOWLIST.includes('create_pivot_definitely_not_real'));
});

test('RED CONTROL: a made-up read operation is correctly absent from OPERATIONS.xlsx', () => {
  assert.ok(!Object.prototype.hasOwnProperty.call(OPERATIONS.xlsx, 'read_pivots_definitely_not_real'));
});

// ---- 5. T3/T4 schema shape: create_pivot's fields are exactly rows/columns/sum (one each,
// three distinct names) -- no multi-field layer, no aggregation-function choice. ----

test('create_pivot\'s own FIELDS list (T3/T4 rows) is exactly the 8 named fields, and the plan() refuses a 4th row/column/sum-shaped field', () => {
  const bad = pivots.plan({ intent: 'create_pivot', name: 'P', sheet: 'S', range: 'A1:C3', rows: 'r', columns: 'c', sum: 'r', pivotSheet: 'H', at: 'A1' }, 0);
  assert.equal(bad.ok, false, 'rows/columns/sum must be three DIFFERENT names');
  assert.match(bad.detail, /three different header names/);
});

test('refresh_pivot never writes a layout or aggregation function -- its own FIELDS list is exactly pivot/sourceSheet/sourceRange', () => {
  const bad = pivotRefresh.plan({ intent: 'refresh_pivot', pivot: 'P', sourceSheet: 'S', sourceRange: 'A1:C3', rows: 'r' }, 0);
  assert.equal(bad.ok, false, 'refresh_pivot should refuse an unknown field like rows');
  assert.match(bad.detail, /does not support field\(s\): rows/);
});
