'use strict';

/**
 * L-11/3 (kiosztó: zero, kártya an internal branch; the owner 36660
 * kérdésének a válasza): a leltár L-11 megjegyzése ">5000 cella tiltva"-t mond -- MA MÁR HAMIS. A read
 * bound owner-dontés C (PR #160) 5000-ről 45000-re emelte (xlsx-driver.cjs MAX_READ_CELLS), ÉS
 * filter_range NEM tiered (nincs a TIERED_INTENTS-ben, a saját `layout` extrája nincs a
 * TIERED_EXTRAS-ban -- adapter.cjs 191-192. sor, l11-filter-range-release.test.cjs már pinneli ezt a
 * HATÁRT, ezt a fájlt NEM lazítja). A nem-tiered refúzió (refusalOf, xlsx-driver.cjs ~1104.sor) a read
 * bound FÖLÖTT jön -- tehát a határ filter_range-re MA 45000/45001, nem 5000/5001.
 *
 * REVIEW/LIVE FIX (a maintainer, kártya-komment #ref): a filter_range saját X-L05 kapuja (a kézzel elrejtett
 * sorok felismerése) a driver `layout` olvasatát kéri a CÉL LAPRA -- és annak a lapnak a SAJÁT
 * maxCol+maxRow összege a driver EGY MÁSIK, FÜGGETLEN határa alá esik: xlsx-driver.cjs
 * MAX_LAYOUT_ITEMS = 2000. A megosztott fake driver (test/fakes/fake-xlsx-driver.cjs) ezt a 2000-es
 * sapkát NEM modellezte (a layout-olvasata korlátlan) -- ez a fájl ezért MOSTANTÓL EGY SAJÁT, a fájlban
 * élő réteggel (wrapWithLayoutBound) teszi rá a valódi MAX_LAYOUT_ITEMS döntést, UGYANAZZAL a
 * technikával, mint a wrapWithReadBound a read bound-ra. A workbook-wide READ bound (45000/45001) és a
 * CÉL LAP SAJÁT layout összege (2000/2001) KÉT FÜGGETLEN kapu: az alábbi fixture-ök ezért a nagy
 * tömeget egy BYSTANDER lapra teszik (amit filter_range SOHA nem kér layout-ra), hogy a READ bound
 * teszteket a layout kapu ne zavarja; a layout kapu SAJÁT tesztje (2001 sor+oszlop) a CÉL lapot magát
 * alakítja a határ fölé.
 *
 * Ez a fájl a DÖNTÉSI ÚTAT pinneli a PRODUKTUM write-tool szintjén (createWriteTool, a valódi executor
 * út): egy fake XLSX könyvön keresztül, AMELYNEK readLive-ja a VALÓDI tierOf/withLimitNote/MAX_READ_CELLS
 * ÉS MAX_LAYOUT_ITEMS döntést alkalmazza (saját küszöb, nem egy kitalált kisebb szám) -- a megosztott
 * test/fakes/fake-xlsx-driver.cjs-t NEM módosítja, hanem a harness() `wrap` pontján EGY, ebben a fájlban
 * élő rétegként teszi rá. A valódi LIVE proof (egy ~40 000 cellás lapon, write_range darabokban
 * validateWriteRequest-en át, read_filter visszaolvasással és időméréssel, ÉS egy szándékosan a
 * layout-sapka fölé alakított lapon): test/live/release-l11-3-filter-large-book-product.live.cjs (zero
 * futtatja).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { tierOf, withLimitNote, MAX_READ_CELLS, MAX_LAYOUT_ITEMS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { colName } = require('../../src/adapters/xlsx/a1.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);

/**
 * The filter table itself stays tiny (A1:A2, 2 cells -- CONSTRAINTS.filter_range's own rangeProblem()
 * caps a filter_range step's RANGE at MAX_CELLS = 1000, same as write_range; a book-sized range would be
 * REFUSED_SHAPE before anything reaches the driver).
 *
 * A BYSTANDER sheet ('B') carries the REST of the workbook-wide bound, far from the filter TARGET ('F').
 * filter_range's own liveExtras() only ever asks for the TARGET sheet's layout (xlsx-filter-rows-
 * unobservable's gate reads p.layout[op.selector.sheet] exclusively) -- the driver computes a layout
 * entry for EVERY sheet in the workbook once ANY layout is asked, but a bystander sheet going over
 * MAX_LAYOUT_ITEMS only makes ITS OWN entry a gap; filter_range never looks at it. So F can stay tiny
 * and layout-safe while B alone determines the workbook-wide read bound (boundOf() sums bounding-box
 * AREA across every sheet; a single far cell at (col, row) on B gives B's own bbox area exactly
 * col*row). This isolates the READ bound tests from the layout-item cap, same as the review already
 * isolated boundOf()'s AREA metric from a naive cell-count assumption.
 */
function filterBook(totalBound) {
  const F = { name: 'F', cells: { A1: { v: 'id' }, A2: { v: 'x' } } }; // bbox 1x2 = 2
  const O = { name: 'O', cells: { A1: { v: 'keep' } } }; // bbox 1x1 = 1
  const bArea = totalBound - 2 - 1;
  // an A1 anchor is required -- a single far cell alone would give B a 1x1 bbox, not a 1 x bArea one
  const B = { name: 'B', cells: bArea > 0 ? { A1: { v: 'anchor' }, [`A${bArea}`]: { v: 'bulk' } } : {} };
  return [F, O, B];
}
const flt = () => ({ intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:A2', field: 1, value: 'x' });

/**
 * A sheet shaped directly, by (cols, rows): a single far cell at (colName(cols), rows) gives the sheet
 * an exact used-range bbox of cols x rows -- both the AREA (boundOf()) and the LAYOUT SUM (cols+rows,
 * the MAX_LAYOUT_ITEMS gate) follow directly from the shape, no bystander involved. Used for the cases
 * where the TARGET sheet's own shape, not the workbook-wide total, is what is under test.
 */
function shapedFilterBook(cols, rows) {
  const cells = { A1: { v: 'id' }, A2: { v: 'x' } };
  if (cols > 1 || rows > 2) cells[`${colName(cols)}${rows}`] = { v: 'far' };
  return [{ name: 'F', cells }, { name: 'O', cells: { A1: { v: 'keep' } } }];
}

/**
 * A readLive wrapper that applies the driver's OWN tier decision (tierOf, on the sheets the fake
 * already reports) to the fake's otherwise-untruncated read: when the bound makes the request
 * 'refused' (filter_range is never tiered, so this is exactly "bound > MAX_READ_CELLS"), the real
 * driver's readRange() hits the global cell-count ceiling before it can finish enumerating that sheet
 * (filter_range has no windows to dedupe against, so bound IS the enumerated count) -- `truncated`
 * turns true and withLimitNote's IMPLEMENTATION_LIMIT entry is appended, exactly as
 * xlsx-driver.cjs:1551 does. Below the bound nothing changes: the fake's full read stands.
 */
function wrapWithReadBound(driver) {
  return {
    ...driver,
    async readLive(handle, request = {}) {
      const r = await driver.readLive(handle, request);
      if (!r.ok) return r;
      const tier = tierOf(r.payload.sheets, request || {});
      if (tier.tier !== 'refused') return r;
      return { ...r, payload: { ...r.payload, truncated: true, unsupported: withLimitNote(r.payload.unsupported || [], tier) } };
    },
  };
}

/**
 * A readLive wrapper applying the driver's OWN MAX_LAYOUT_ITEMS cap to the fake's otherwise-uncapped
 * layout read (test/fakes/fake-xlsx-driver.cjs's layoutOf() mirrors the real used-range-derived
 * maxCol/maxRow logic but never enforces the 2000-item ceiling -- the #ref LIVE finding this
 * harness was blind to before this fix). Per sheet: maxCol/maxRow read back from the SIZE of the
 * fake's own (uncapped) cols/rows maps -- the same quantity the real driver's own used-range scan
 * would have produced -- and if their sum is over MAX_LAYOUT_ITEMS, that ONE sheet's entry becomes
 * 'UNOBSERVABLE:layout-truncated', exactly as xlsx-driver.cjs:676 does. Other sheets' entries (and a
 * gap the fake itself already produced) pass through unchanged.
 */
function wrapWithLayoutBound(driver) {
  return {
    ...driver,
    async readLive(handle, request = {}) {
      const r = await driver.readLive(handle, request);
      if (!r.ok || !r.payload || !r.payload.layout) return r;
      const layout = {};
      for (const [sheet, axis] of Object.entries(r.payload.layout)) {
        if (typeof axis === 'string') { layout[sheet] = axis; continue; }
        const maxCol = Object.keys(axis.cols || {}).length;
        const maxRow = Object.keys(axis.rows || {}).length;
        layout[sheet] = maxCol + maxRow > MAX_LAYOUT_ITEMS ? 'UNOBSERVABLE:layout-truncated' : axis;
      }
      return { ...r, payload: { ...r.payload, layout } };
    },
  };
}
const wrapFilterLiveBounds = driver => wrapWithLayoutBound(wrapWithReadBound(driver));

const run = book => {
  const h = harness({ store: createStore(book), wrap: wrapFilterLiveBounds });
  return { h, call: () => h.tool.handler(req([flt()])) };
};
const runBound = n => run(filterBook(n));
const runShaped = (cols, rows) => run(shapedFilterBook(cols, rows));

test('MAX_READ_CELLS is 45000 (the leltár\'s ">5000" note is stale -- pinned here so this file stays honest if it moves again)', () => {
  assert.equal(MAX_READ_CELLS, 45000);
});

test('MAX_LAYOUT_ITEMS is 2000 (the driver\'s INDEPENDENT layout-read cap, the #ref LIVE finding)', () => {
  assert.equal(MAX_LAYOUT_ITEMS, 2000);
});

test('filter_range stays outside the tiered observation (the owner rule this file does not loosen)', () => {
  assert.notEqual(liveScope({ id: 't', steps: [flt()] }).tiered, true);
});

test('CONSTRAINTS caps a filter_range step\'s OWN range at 1000 cells -- the large-book case must grow the SHEET, not the filter range', () => {
  const { validateWriteRequest } = require('../../src/server/write-tools.cjs');
  const over = validateWriteRequest(req([{ intent: 'filter_range', id: 's', sheet: 'F', range: 'A1:A1001', field: 1, value: 'x' }]));
  assert.equal(over.ok, false);
  assert.match(JSON.stringify(over), /over the 1000 ceiling/);
});

for (const n of [5001, 44000]) {
  test(`${n} cells workbook-wide (above the old 5000 limit, within the new 45000 one), the TARGET sheet layout-safe (bystander carries the bulk): APPLIED, the other sheets untouched, retry SATISFIED 0/0`, async () => {
    const { h, call } = runBound(n);
    const r = await call();
    assert.equal(r.ok, true, J(r).slice(0, 400));
    assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
    const F = h.store.durable.find(x => x.name === 'F');
    assert.equal(F.autoFilter.range, 'A1:A2');
    assert.ok(F.autoFilter.filters.some(f => f.field === 1 && f.on !== false), J(F.autoFilter));
    assert.deepEqual(h.store.durable.find(x => x.name === 'O').cells, { A1: { v: 'keep' } });
    const d0 = h.calls().dispatch;
    const again = await call();
    assert.equal(again.ok, true, J(again).slice(0, 300));
    assert.deepEqual((again.operations || []).map(o => o.status), ['SATISFIED']);
    assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
  });
}

test('45001 cells workbook-wide (above the new 45000 limit): refused, 0 dispatch, the named IMPLEMENTATION_LIMIT reason', async () => {
  const { h, call } = runBound(45001);
  const r = await call();
  assert.equal(r.ok, false, J(r).slice(0, 400));
  assert.equal(h.calls().dispatch, 0);
  assert.match(J(r), /xlsx-observation-truncated/);
  assert.match(J(r), /IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation \(bound 45001\)/);
  assert.deepEqual(h.store.durable.find(x => x.name === 'O').cells, { A1: { v: 'keep' } }, 'a refused observation never reaches dispatch, let alone another sheet');
});

test('the READ-BOUND boundary itself: 45000 still APPLIED, 45001 the first to refuse (no off-by-one)', async () => {
  const ok = await runBound(45000).call();
  assert.equal(ok.ok, true, J(ok).slice(0, 300));
  assert.deepEqual((ok.operations || []).map(o => o.status), ['APPLIED']);
  const refused = await runBound(45001).call();
  assert.equal(refused.ok, false);
});

test('POS_WIDE (the LIVE harness\'s own shape, zero\'s #ref fix): the TARGET sheet itself is 1900 rows x 20 cols -- area 38000 (clears the old 5000 ceiling), layout sum 1920 (under MAX_LAYOUT_ITEMS): APPLIED, retry SATISFIED', async () => {
  const { h, call } = runShaped(20, 1900);
  const r = await call();
  assert.equal(r.ok, true, J(r).slice(0, 400));
  assert.deepEqual((r.operations || []).map(o => o.status), ['APPLIED']);
  const d0 = h.calls().dispatch;
  const again = await call();
  assert.equal(again.ok, true, J(again).slice(0, 300));
  assert.deepEqual((again.operations || []).map(o => o.status), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('POS_TALL / the LAYOUT boundary itself: the TARGET sheet is shaped to 2000 sum (2 cols + 1998 rows) -> still APPLIED; 2001 sum (2 cols + 1999 rows) -> refused named xlsx-filter-rows-unobservable/layout-truncated, 0 dispatch (no off-by-one, independent of the read bound)', async () => {
  const atCap = await runShaped(2, 1998).call();
  assert.equal(atCap.ok, true, J(atCap).slice(0, 300));
  assert.deepEqual((atCap.operations || []).map(o => o.status), ['APPLIED']);

  const { h, call } = runShaped(2, 1999);
  const r = await call();
  assert.equal(r.ok, false, J(r).slice(0, 400));
  assert.equal(h.calls().dispatch, 0);
  assert.match(J(r), /xlsx-filter-rows-unobservable/);
  assert.match(J(r), /layout-truncated/);
});
