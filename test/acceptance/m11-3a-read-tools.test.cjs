'use strict';

/**
 * M11.3a -- READ tools xlsx_read / docx_read: handlers + schemas, WITHHELD. OFFLINE acceptance (owner GO, owner channel
 * 2026-10-09, after the M11.2 closure).
 *
 * The REAL product handler (src/server/read-tools.cjs) runs end to end against the REAL drivers and adapters:
 *   MCP args -> schema -> the M11.2 broker configSource (fake broker, own token) -> the canonical driver open(READ) +
 *   readLive + close -> normalizeInventory -> projection.
 * The editor is fake: XLSX command bodies run against an in-memory Api that has GETTERS ONLY (a write call would throw
 * and fail the read; every missing API name the body touches is logged), DOCX answers the real property query.
 * Proven: the accepted reads return structured results with DOCUMENT_MUTATION=0 / HISTORY_MUTATION=0 / WRITE=0 /
 * SAVE_REQUIRED=NO; every refused request (unknown / PARTIAL / NOT_MEASURED / INTERNAL / WRITE intent, malformed
 * arguments, credential injection, requester spoof, wrong fileId type) stops PRE-DISPATCH -- no broker call, no driver;
 * the tools stayed WITHHELD in M11.3a (neither listed nor callable); M11.3b.2 released exactly these two (see below).
 */

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');
const { PassThrough } = require('node:stream');

const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const xlsx = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { OPERATIONS, STATUS, createReadTool, validateReadRequest, writeIntents, inputSchema } = require('../../src/server/read-tools.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createMcpServer } = require('../../src/server/mcp.cjs');
const { exposedTools, TOOL_ROWS } = require('../../src/server/tools.cjs');

const { PAGE, BROKER, TOKEN, FILE, fakeBroker, xlsxEditor, docxEditor, NON_MUTATION } = require('../fakes/read-editors.cjs');

const drivers = [];
after(async () => { for (const d of drivers) await d.close().catch(() => null); });

/** The real read tool on the real driver; only the editor, the broker and the timings are fakes. */
function tool(format, ed, broker) {
  const createDriver = (fmt, deps) => {
    assert.equal(deps.configSource.kind, 'BROKER', 'the handler hands the driver the broker configSource');
    const ctor = fmt === 'xlsx' ? xlsx.createEuroOfficeXlsxDriver : docx.createEuroOfficeDocxDriver;
    const d = ctor({ ...deps, pollMs: 5, closeAckTimeoutMs: 200, timeoutMs: 2000 });
    drivers.push(d);
    // the driver-API level proof: a READ never calls the write/save surface of the canonical driver
    for (const m of ['dispatch', 'forceSave']) { const real = d[m].bind(d); d[m] = (...a) => { ed.api[m] += 1; return real(...a); }; }
    return d;
  };
  return createReadTool(format, { configSource: broker.source, loadPlaywright: ed.loadPlaywright, createDriver });
}

async function noDirectFetch(fn) {
  const real = globalThis.fetch;
  let direct = 0;
  globalThis.fetch = async () => { direct += 1; throw new Error('the product READ path must not fetch Nextcloud directly'); };
  try { const out = await fn(); assert.equal(direct, 0, 'no direct Nextcloud / Basic-auth fetch'); return out; } finally { globalThis.fetch = real; }
}

// ---- inventory ----------------------------------------------------------------------------------------------------

const accepted = f => Object.entries(OPERATIONS[f]).filter(([, o]) => o.status === STATUS.ACCEPTED).map(([k]) => k);
const byStatus = s => ['xlsx', 'docx'].flatMap(f => Object.entries(OPERATIONS[f]).filter(([, o]) => o.status === s).map(([k]) => `${f}:${k}`));

test('INVENTORY: the accepted READ allowlist, every catalogue row carries evidence, the 77 routed intents are WRITE', () => {
  // X-O05/4: read_charts PARTIAL -> ACCEPTED (effectiveLabels, per-point, via chart.dataLabels.effective.read/v1) --
  // conditional on this round's LIVE acceptance; revert this line too if that run does not PASS.
  // X-O08 (/14): read_drawings PARTIAL -> ACCEPTED too (parent/kind/name/pos/width/height/rotation/flip: M7 S2b;
  // shape fill/line via drawing.fillLineColor.read/v1, bridge /14; the X-O09 evidence this status carried alone is
  // a narrow, unrelated image-height preservation note, same precedent as read_pivots's own M7-H note).
  // L-08 unit 2: read_conditional_formats PARTIAL -> ACCEPTED (the write-schema mirror, LIVE l08-r1).
  // L-09 unit 1: read_layout (NEW, was not in the catalogue) -> ACCEPTED. L-10: read_freeze_panes (NEW) -> ACCEPTED; L-10b: read_page_setup / read_filter (NEW) -> ACCEPTED (LIVE on a pre-seeded fixture).
  // GH #137 U1: read_runs (NEW) -> ACCEPTED -- the product READ mirror of format_run (X-F04), released before that
  // family's own WRITE (test/acceptance/gh137-u1-read-runs-release.test.cjs).
  // GH #140 U2: read_pivots PARTIAL -> ACCEPTED -- a dedicated projection of inventory.cjs's already-
  // computed, content-proven pivot read (pivot-source.cjs withProvenPivots()), never a read_charts widening.
  // GH #133: read_sheet_state (sheet visibility + the active sheet) -> ACCEPTED, a field-select over the existing sheetVisibility / panes axes.
  // GH #224: read_styles (NEW, dedicated) -> ACCEPTED -- the style DEFINITIONS of the canonical inventory's document.styles; scope = the GetAllStyles enumeration.
  // GH #225: read_numbering PARTIAL -> ACCEPTED -- present/type/numId/level/levelFormats (Word Bridge,
  // B19 parity) + format fields (ToJSON, the SAME path numbering.cjs's own WRITE-side verification uses).
  // GH #268b: read_images PARTIAL -> ACCEPTED (the release flip, after the g268v1 SEED_OK): the page-side GetAllImages + ToJSON projection, no image bytes.
  // GH #269b: read_comments NOT_MEASURED -> ACCEPTED (g269bv2 LIVE PASS): the main document's comments, the range anchor a named gap.
  // GH #227: read_section_flags INTERNAL -> ACCEPTED -- titlePage/evenAndOdd through the SAME B21 Bridge
  // READ (docx.section.variantState.read/v1) the write-side header/footer family already runtime-gates.
  assert.deepEqual(accepted('xlsx'), ['workbook_overview', 'read_range', 'read_validations', 'read_charts', 'read_pivots', 'read_drawings', 'read_conditional_formats', 'read_layout', 'read_freeze_panes', 'read_sheet_state', 'read_page_setup', 'read_filter', 'read_comments', 'read_print_area', 'read_defined_names', 'read_runs']);
  assert.deepEqual(accepted('docx'), ['read_text', 'read_formatting', 'read_tables', 'read_headers_footers', 'read_page_setup', 'read_bookmarks', 'read_numbering', 'read_styles', 'read_notes', 'read_images', 'read_comments', 'read_section_flags']);
  assert.equal(byStatus(STATUS.PARTIAL).length, 1); // GH #268b release flip: PARTIAL 2 -> 1; L-08 + X-O08: 6 -> 4; GH #140 U2: 4 -> 3; GH #225: 3 -> 2; GH #226 3: read_notes 2 -> 1; GH #268b: read_images NOT_MEASURED -> PARTIAL 1 -> 2 (the release flip is a separate commit)
  assert.equal(byStatus(STATUS.NOT_MEASURED).length, 1); // GH #269b: read_comments 2 -> 1; GH #268b: 3 -> 2
  assert.equal(byStatus(STATUS.INTERNAL).length, 0); // GH #227: 1 -> 0, the last INTERNAL docx op released
  for (const f of ['xlsx', 'docx']) for (const [k, o] of Object.entries(OPERATIONS[f])) assert.ok(o.evidence && o.evidence.length > 3, `${f}:${k} evidence`);
  assert.equal(writeIntents().size, 77);
  for (const f of ['xlsx', 'docx']) for (const k of Object.keys(OPERATIONS[f])) assert.equal(writeIntents().has(k), false, `${k} is not a write intent`);
});

test('SCHEMA: fileId + operation required, only accepted operations offered, no credential or requester field, additionalProperties false', () => {
  for (const f of ['xlsx', 'docx']) {
    const s = inputSchema(f);
    assert.deepEqual(s.required, ['fileId', 'operation']);
    assert.equal(s.additionalProperties, false);
    assert.equal(s.properties.fileId.type, 'string');
    assert.deepEqual(s.properties.operation.enum, accepted(f));
    for (const k of Object.keys(s.properties)) assert.equal(/user|pass|secret|token|cred|vault|key|requester|uid|owner|actor/i.test(k), false, `${f}.${k}`);
  }
});

// ---- valid reads through the real stack -----------------------------------------------------------------------------

test('VALID_XLSX_READ: read_range + workbook_overview + read_validations + read_charts + read_drawings -- broker READ config, real driver/adapter, non-mutating', async () => {
  await noDirectFetch(async () => {
    const b = fakeBroker('xlsx');
    const ed = xlsxEditor();
    const t = tool('xlsx', ed, b);
    const r = await t.handler({ fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'A1:B3' });
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.complete, true);
    assert.deepEqual(r.result.cells.map(c => [c.address, c.v]), [['A1', 'Item'], ['B1', 'Amount'], ['A2', 'Rent'], ['B2', '1200'], ['A3', 'Food'], ['B3', '340']]);
    // GitHub #72: each cell carries its value type; without format/style no other field appears
    assert.deepEqual(r.result.cells.map(c => [c.address, c.type]), [['A1', 'text'], ['B1', 'text'], ['A2', 'text'], ['B2', 'number'], ['A3', 'text'], ['B3', 'number']]);
    for (const c of r.result.cells) assert.deepEqual(Object.keys(c).sort(), ['address', 'f', 'merge', 'type', 'v']);
    assert.equal(b.calls.length, 1);
    assert.equal(b.calls[0].bearer, TOKEN, 'the own token only');
    assert.deepEqual(Object.keys(b.calls[0].body).sort(), ['fileId', 'instance', 'op', 'requestedMode']);
    assert.equal(b.calls[0].body.requestedMode, 'READ');
    assert.ok(ed.log.includes('read-body'), 'the canonical readLive query ran (core dispatch of the READ)');
    NON_MUTATION(ed, r);

    const o = await t.handler({ fileId: FILE, operation: 'workbook_overview' });
    assert.equal(o.ok, true, JSON.stringify(o));
    assert.deepEqual(o.result.sheets, [{ name: 'Budget', index: 0, used: 'A1:B3' }]);
    NON_MUTATION(ed, o);

    const v = await t.handler({ fileId: FILE, operation: 'read_validations', sheet: 'Budget', range: 'B2:B3' });
    assert.equal(v.ok, true, JSON.stringify(v));
    assert.deepEqual(v.result.validation, { type: 'xlValidateWholeNumber', alertStyle: 'xlValidAlertStop', operator: 'xlBetween', formula1: '0', formula2: '10000' });
    NON_MUTATION(ed, v);

    // L-22e GREEN: a type 'custom' rule with a RELATIVE formula on C2:C4 -- the fixture's GetFormula1()
    // answers '=A1>0' (the sheet's own active cell, A1, never C2) through the SAME real readQuery the
    // driver generates. read_validations must re-anchor it to the range's own first cell, matching what
    // was actually written, never returning the raw, active-cell-shifted text.
    const vc = await t.handler({ fileId: FILE, operation: 'read_validations', sheet: 'Budget', range: 'C2:C4' });
    assert.equal(vc.ok, true, JSON.stringify(vc));
    assert.deepEqual(vc.result.validation, { type: 'xlValidateCustom', alertStyle: 'xlValidAlertStop', operator: 'xlBetween', formula1: '=C2>0', formula2: '' });
    assert.ok(!('activeCell' in vc.result.validation), 'activeCell is an internal field, never part of this tool\'s own public contract');
    NON_MUTATION(ed, vc);

    // L-09 unit 1: read_layout through the REAL driver's layout read (GetEntireColumn / GetEntireRow on the fake editor): B 20.43 wide, C hidden, row 2 30 high, row 3 hidden.
    const lay = await t.handler({ fileId: FILE, operation: 'read_layout', sheet: 'Budget', range: 'B2:C3' });
    assert.equal(lay.ok, true, JSON.stringify(lay));
    assert.deepEqual(lay.result.columns.B, { width: 20.43, hidden: false });
    assert.equal(lay.result.columns.C.hidden, true);
    assert.equal(lay.result.columns.C.width, null);
    assert.match(lay.result.columns.C.reason, /^HIDDEN:/);
    assert.deepEqual(lay.result.rows[2], { height: 30, hidden: false });
    assert.equal(lay.result.rows[3].hidden, true);
    assert.equal(lay.result.rows[3].height, null);
    const layMissing = await t.handler({ fileId: FILE, operation: 'read_layout', sheet: 'NoSuchSheet', range: 'A1:B2' });
    assert.equal(layMissing.result.error, 'SHEET_NOT_FOUND', 'a non-existent sheet is named, never a guessed default layout');
    NON_MUTATION(ed, lay);

    // L-10: read_freeze_panes through the REAL driver. The freeze read goes through a DS Bridge capability that needs the editor's internal model, which the fake does not carry, so it surfaces
    // as a NAMED gap -- never "not frozen" (the LIVE acceptance is test/live/release-l10-sheet-settings-product.live.cjs, LIVE r1 green).
    const fp = await t.handler({ fileId: FILE, operation: 'read_freeze_panes', sheet: 'Budget' });
    assert.equal(fp.ok, true, JSON.stringify(fp));
    assert.equal(fp.result.frozen, null);
    assert.match(fp.result.reason, /^UNOBSERVABLE:BRIDGE:/);
    NON_MUTATION(ed, fp);
    const fpMissing = await t.handler({ fileId: FILE, operation: 'read_freeze_panes', sheet: 'NoSuchSheet' });
    assert.equal(fpMissing.result.error, 'SHEET_NOT_FOUND', 'a non-existent sheet is named');
    NON_MUTATION(ed, fpMissing);
    // L-10b: read_page_setup / read_filter through the REAL driver. The page getters answer on the fake (landscape, headings on, the public margins) and a sheet without an AutoFilter reads { range: null, filters: [] };
    // the print half goes through a DS Bridge capability the fake does not carry, so it surfaces as a NAMED gap and the two halves fail independently.
    const ps = await t.handler({ fileId: FILE, operation: 'read_page_setup', sheet: 'Budget' });
    assert.equal(ps.ok, true, JSON.stringify(ps));
    assert.deepEqual(ps.result.page, { orientation: 'landscape', printGridlines: PAGE.printGridlines, printHeadings: PAGE.printHeadings, topMargin: PAGE.topMargin, bottomMargin: PAGE.bottomMargin, leftMargin: PAGE.leftMargin, rightMargin: PAGE.rightMargin });
    assert.match(ps.result.print.reason, /^UNOBSERVABLE:BRIDGE:/, 'the two halves fail independently');
    NON_MUTATION(ed, ps);
    const fl = await t.handler({ fileId: FILE, operation: 'read_filter', sheet: 'Budget' });
    assert.equal(fl.ok, true, JSON.stringify(fl));
    assert.deepEqual(fl.result, { sheet: 'Budget', range: null, filters: [], engineShapes: [] });
    NON_MUTATION(ed, fl);
    // GH #134: read_print_area through the REAL driver -- the defined-names read is the DS Bridge workbook.definedNames.read/v1
    const pa = await t.handler({ fileId: FILE, operation: 'read_print_area', sheet: 'Budget' });
    assert.equal(pa.ok, true, JSON.stringify(pa));
    // the fake editor carries no internal model, so the names read is a NAMED gap -- never { set: false } (absence) on an unread reader
    assert.deepEqual(pa.result, { sheet: 'Budget', printArea: 'UNOBSERVABLE:BRIDGE:model-missing' });
    NON_MUTATION(ed, pa);
    // GH #134 P2: read_defined_names through the REAL driver -- the same defined-names read; on the fake editor a NAMED gap, never an empty list
    const dn = await t.handler({ fileId: FILE, operation: 'read_defined_names' });
    assert.equal(dn.ok, true, JSON.stringify(dn));
    assert.deepEqual(dn.result, { names: 'UNOBSERVABLE:BRIDGE:model-missing' });
    NON_MUTATION(ed, dn);
    for (const op of ['read_page_setup', 'read_filter', 'read_print_area']) {
      const miss = await t.handler({ fileId: FILE, operation: op, sheet: 'NoSuchSheet' });
      assert.equal(miss.result.error, 'SHEET_NOT_FOUND', `${op}: a non-existent sheet is named`);
    }

    // X-O05/4: effectiveLabels through the REAL chart.dataLabels.effective.read/v1 bridge-block (installBridge
    // runs the actual patch code, not an imitation) -- point 1 (Food, 340) carries the fixture's POINT override,
    // point 0 (Rent, 1200) falls through to SERIES, exactly the point->series->chart precedence the design covers.
    const ch = await t.handler({ fileId: FILE, operation: 'read_charts' });
    assert.equal(ch.ok, true, JSON.stringify(ch));
    assert.equal(ch.result.sheets.length, 1);
    const spend = ch.result.sheets[0].charts[0];
    assert.equal(spend.name, 'Spend');
    assert.deepEqual(spend.effectiveLabels['0:0'], { level: 'SERIES', showSerName: false, showCatName: false, showVal: false, showPercent: false });
    assert.deepEqual(spend.effectiveLabels['0:1'], { level: 'POINT', showSerName: true, showCatName: false, showVal: false, showPercent: false });
    NON_MUTATION(ed, ch);

    const chOne = await t.handler({ fileId: FILE, operation: 'read_charts', sheet: 'Budget' });
    assert.deepEqual(chOne.result.sheets.map(s => s.sheet), ['Budget'], 'an explicit sheet scopes the result to just that sheet');
    const chMissing = await t.handler({ fileId: FILE, operation: 'read_charts', sheet: 'NoSuchSheet' });
    assert.equal(chMissing.result.sheets[0].reason, 'UNOBSERVABLE:sheet-missing', 'a non-existent sheet names the gap, never a guessed empty list');
    NON_MUTATION(ed, chMissing);

    // X-O08 (/14): the SAME public-getter fields M7-S2b measured LIVE (parent/kind/name/pos/width/height/
    // rotation/flip, docs/evidence/M7-S2B-DRAWING-PROBE.md) + strokeWidth (X-O08-7, PUBLIC-readable) + the
    // bridge-backed fillLineColor.
    // DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the fixture bridge (ds-bridge-runtime.cjs)
    // now loads the 9.3.5.0 block, which carries the FULL capability set (including fillLineColor) -- so the
    // capability IS advertised here. The gap is now at the GETTER level, not the capability level: this fixture's
    // shape lacks the internal chain the handler needs, so each sub-field degrades to 'UNOBSERVABLE:internal-chain-broken'.
    const dr = await t.handler({ fileId: FILE, operation: 'read_drawings' });
    assert.equal(dr.ok, true, JSON.stringify(dr));
    assert.equal(dr.result.sheets.length, 1);
    const logo = dr.result.sheets[0].drawings[0];
    assert.equal(logo.kind, 'shape');
    assert.equal(logo.name, 'Logo');
    assert.equal(logo.parent, 'Budget');
    assert.deepEqual(logo.pos, { fromCol: 2, colOffset: 0, fromRow: 1, rowOffset: 0 });
    assert.deepEqual([logo.width, logo.height, logo.rotation, logo.flipH, logo.flipV], [914400, 457200, 0, false, false]);
    assert.equal(logo.strokeWidth, 19050, 'PUBLIC GetLine().GetWidth() -- X-O08-7');
    assert.deepEqual(logo.fillLineColor, { fill: 'UNOBSERVABLE:internal-chain-broken', line: 'UNOBSERVABLE:internal-chain-broken' }, 'DS9346: the capability is now present, the fixture\'s own getter chain is not, never a guessed colour');
    NON_MUTATION(ed, dr);

    const drOne = await t.handler({ fileId: FILE, operation: 'read_drawings', sheet: 'Budget' });
    assert.deepEqual(drOne.result.sheets.map(s => s.sheet), ['Budget'], 'an explicit sheet scopes the result to just that sheet');
    const drMissing = await t.handler({ fileId: FILE, operation: 'read_drawings', sheet: 'NoSuchSheet' });
    assert.equal(drMissing.result.sheets[0].reason, 'UNOBSERVABLE:sheet-missing', 'a non-existent sheet names the gap, never a guessed empty list');
    NON_MUTATION(ed, drMissing);

    assert.equal(ed.launches(), 21); // GH #134: +2 (read_print_area on Budget + on NoSuchSheet); P2: +1 (read_defined_names); L-22e: +1 (read_validations on C2:C4, the custom/relative re-anchor case)
  });
});

test('VALID_DOCX_READ: the twelve accepted operations -- broker READ config, real driver/adapter, non-mutating, no PARTIAL/INTERNAL field', async () => {
  await noDirectFetch(async () => {
    const b = fakeBroker('docx');
    const ed = docxEditor();
    const t = tool('docx', ed, b);
    const results = {};
    for (const op of accepted('docx')) {
      const r = await t.handler({ fileId: FILE, operation: op });
      assert.equal(r.ok, true, `${op}: ${JSON.stringify(r)}`);
      assert.equal(r.complete, true);
      NON_MUTATION(ed, r);
      results[op] = r.result;
      // gridCalc (PARTIAL), the internal gate: never projected.
      // GH #225: 'numbering' is read_numbering's OWN accepted field now -- excluded from ITS OWN check,
      // still checked against every OTHER operation's result (a leak there would still be a defect).
      // GH #227: 'titlePage'/'evenAndOdd' is read_section_flags's OWN accepted field now -- same rule.
      const leak = new RegExp([op === 'read_numbering' ? null : 'numbering', 'removeNumPr', 'gridCalc', op === 'read_section_flags' ? null : 'titlePage', op === 'read_section_flags' ? null : 'evenAndOdd', 'internalGate'].filter(Boolean).join('|'));
      assert.equal(leak.test(JSON.stringify(r.result)), false, `${op} leaks a refused field`);
    }
    // GH #269b: read_comments -- the driver's rows projected with `solved` named `resolved`, the range anchor ALWAYS the named gap, complete:true
    assert.deepEqual(results.read_comments, { comments: [{ id: 'c1', text: 'a note', author: 'Ann', userId: 'u1', timeUtc: 1792420516000, quote: 'Hello', resolved: false, replies: [{ text: 'a reply', author: 'Bob', userId: 'u2' }], anchor: 'UNOBSERVABLE:no-public-anchor-getter' }], complete: true });
    assert.deepEqual(results.read_text.blocks.map(x => x.kind), ['heading', 'paragraph', 'list', 'table']);
    assert.equal(results.read_text.blocks[1].text, 'Hello world');
    assert.deepEqual(results.read_text.blocks[2].items, ['one', 'two']);
    assert.equal(results.read_formatting.units[1].text, 'Hello world');
    assert.deepEqual(results.read_formatting.units[1].format, { alignment: 'left' });
    assert.deepEqual(results.read_formatting.defaultText, { fontFamily: 'Arial', fontSize: 11 });
    assert.deepEqual(results.read_tables.tables[0].rows[0].map(c => c.text), ['a', 'b']);
    assert.equal(results.read_headers_footers.sections[0].header.default.paras[0].text, 'Head');
    // GH #223: the shared fixture seeds no page-setup fields at all -- every field reads the
    // driver's own UNOBSERVABLE:unread, never a guessed 0/default (the SAME discipline the W11
    // WRITE side's own inventory.cjs measured()/measuredArr() already enforces).
    assert.deepEqual(results.read_page_setup.sections, [{
      index: 0, pageWidth: 'UNOBSERVABLE:unread', pageHeight: 'UNOBSERVABLE:unread',
      margins: { top: 'UNOBSERVABLE:unread', right: 'UNOBSERVABLE:unread', bottom: 'UNOBSERVABLE:unread', left: 'UNOBSERVABLE:unread' },
      columnsCount: 'UNOBSERVABLE:unread', columnsSpaces: 'UNOBSERVABLE:unread', startPageNumber: 'UNOBSERVABLE:unread',
    }]);
    assert.deepEqual(results.read_bookmarks.bookmarks, [{ name: 'bm1', text: 'Hello' }]);
    // GH #225: against THIS shared fixture (one paragraphs[] entry, text 'Hello world', attached to a
    // plain <p> -- not a <li>), the numbering it carries (level/numId, no `present` key) is INCONSISTENT
    // by the SAME rule families/list.cjs's own listState() already enforces (a non-list-item paragraph's
    // numbering must read present:false) -- fail-closed, not silently accepted. 'Budget'/'one'/'two' have
    // no matching paragraphs[] entry at all (hits.length!==1 in inventory.cjs's numberingOf) -> unread.
    // The POSITIVE (present:true, format fields, shared-numId) cases are test/acceptance/
    // gh225-docx-numbering-read.test.cjs's own purpose-built fixtures, not this shared one.
    assert.deepEqual(results.read_numbering.items, [
      { text: 'Budget', kind: 'heading', numbering: null, numberingGap: 'unread' },
      { text: 'Hello world', kind: 'paragraph', numbering: null, numberingGap: 'inconsistent' },
      { text: 'one', kind: 'list_item', numbering: null, numberingGap: 'unread' },
      { text: 'two', kind: 'list_item', numbering: null, numberingGap: 'unread' },
    ]);
    // GH #224: read_styles against THIS shared fixture (no document.styles in the canned property result) -> the map
    // was never read: a NAMED gap (styles:null), NEVER an empty list that would read as "the document has no styles".
    assert.equal(results.read_styles.styles, null);
    assert.match(results.read_styles.scope, /GetAllStyles/);
    // GH #227: against THIS shared fixture, the canned property result's own section already carries
    // titlePage:true/evenAndOdd:false (read-editors.cjs PROPERTY_RESULT.document.sections[0]) -- the
    // SAME values read_page_setup's and read_headers_footers' own section-0 objects are built from.
    assert.deepEqual(results.read_section_flags.sections, [{ index: 0, titlePage: true, evenAndOdd: false }]);
    // GH #226 3: against THIS shared fixture the canned property result carries no note chains -> both kinds NAMED unread (null) and the read NOT complete, never an empty list.
    assert.deepEqual(results.read_notes, { footnotes: null, endnotes: null, complete: false });
    assert.equal(b.calls.length, 12); // GH #269b: +1 (read_comments); GH #268b: +1 (read_images)
    assert.ok(b.calls.every(c => c.bearer === TOKEN && c.body.requestedMode === 'READ'));
    assert.equal(ed.log.filter(l => l === 'read-query').length, 12, 'the canonical property query ran once per read');
    // M11.3a.1: every product READ asks without the B21 section variant state -- GH #225's own
    // read_numbering is the ONE call that also opts INTO numberingFormat. GH #227's own
    // read_section_flags is the ONE call that opts INTO the section variant state itself --
    // and this fixture's handshake runtime is NOT B21-verified (internalGateOf(runtimeId).
    // sectionVariant is false here), so the request still resolves to the GATED query shape
    // (sectionVariantUnavailable: 'runtime-not-verified'), the SAME 'gated' key read-editors.cjs's
    // PROPERTY_QUERIES map already names for exactly this case -- proving the wiring asks (opts
    // in), not that this particular fixture's runtime happens to be proven for B21.
    assert.equal(ed.log.filter(l => l === 'read-query:opt-out').length, 9);
    assert.equal(ed.log.filter(l => l === 'read-query:opt-out-numbering').length, 1, 'read_numbering alone requests the extra ToJSON() pass');
    assert.equal(ed.log.filter(l => l === 'read-query:opt-out-comments').length, 1, 'read_comments alone requests the comments read (includeComments)');
    assert.equal(ed.log.filter(l => l === 'read-query:gated').length, 1, 'read_section_flags alone opts into the section variant state (gated on this fixture\'s unverified runtime)');
  });
});

// ---- pre-dispatch refusals --------------------------------------------------------------------------------------------

const REFUSALS = [
  ['UNKNOWN_INTENT', 'xlsx', { fileId: FILE, operation: 'read_everything' }, 'REFUSED_UNKNOWN'],
  // GH #140 U2: read_pivots moved PARTIAL -> ACCEPTED (see INVENTORY) -- xlsx has no PARTIAL
  // operation left, so the PARTIAL-refusal path is exercised on docx only below.
  // GH #225: read_numbering ALSO moved PARTIAL -> ACCEPTED -- read_table_grid is the one that stays.
  ['PARTIAL intent (docx)', 'docx', { fileId: FILE, operation: 'read_table_grid' }, 'REFUSED_PARTIAL'],
  // GH #269b: read_comments moved NOT_MEASURED -> ACCEPTED -- read_print_rendering (xlsx) is the NOT_MEASURED row left.
  ['NOT_MEASURED intent', 'xlsx', { fileId: FILE, operation: 'read_print_rendering' }, 'REFUSED_NOT_MEASURED'],
  // GH #227: read_section_flags moved INTERNAL -> ACCEPTED -- it was the LAST docx:INTERNAL operation
  // (byStatus(STATUS.INTERNAL).length is now 0, see INVENTORY above), so no live REFUSED_INTERNAL
  // example remains; the generic `REFUSED_${op.status}` mechanism (checkOperation) is unchanged and
  // would fire again the moment any operation is next marked INTERNAL.
  ['WRITE_INTENT (xlsx)', 'xlsx', { fileId: FILE, operation: 'write_range', sheet: 'Budget', range: 'A1' }, 'REFUSED_WRITE'],
  ['WRITE_INTENT (docx intent on the xlsx tool)', 'xlsx', { fileId: FILE, operation: [...writeIntents()].find(i => !OPERATIONS.xlsx[i] && !/_range|_sheet|_rows|_columns/.test(i)) }, 'REFUSED_WRITE'],
  ['malformed args: not an object', 'docx', ['1302148'], 'INVALID_ARGS'],
  ['malformed args: missing range', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget' }, 'INVALID_ARGS'],
  ['malformed args: bad range', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'Budget!A1' }, 'INVALID_ARGS'],
  ['malformed args: flag type', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'A1', format: 'yes' }, 'INVALID_ARGS'],
  ['malformed args: unknown field', 'docx', { fileId: FILE, operation: 'read_text', depth: 3 }, 'UNKNOWN_FIELD'],
  ['malformed args: field of another operation', 'xlsx', { fileId: FILE, operation: 'workbook_overview', sheet: 'Budget' }, 'UNKNOWN_FIELD'],
  ['range over the read ceiling', 'xlsx', { fileId: FILE, operation: 'read_range', sheet: 'Budget', range: 'A1:Z1800' }, 'LIMIT'],
  ['CREDENTIAL_INJECTION: password', 'docx', { fileId: FILE, operation: 'read_text', password: 'x' }, 'CREDENTIAL_FIELD'],
  ['CREDENTIAL_INJECTION: user', 'xlsx', { fileId: FILE, operation: 'workbook_overview', user: 'maintainer' }, 'CREDENTIAL_FIELD'],
  ['CREDENTIAL_INJECTION: token', 'docx', { fileId: FILE, operation: 'read_text', token: 'abc' }, 'CREDENTIAL_FIELD'],
  ['CREDENTIAL_INJECTION: key_name', 'docx', { fileId: FILE, operation: 'read_text', key_name: 'nc/maintainer' }, 'CREDENTIAL_FIELD'],
  ['REQUESTER_SPOOF: requester', 'docx', { fileId: FILE, operation: 'read_text', requester: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['REQUESTER_SPOOF: uid', 'xlsx', { fileId: FILE, operation: 'workbook_overview', uid: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['REQUESTER_SPOOF: onBehalfOf', 'docx', { fileId: FILE, operation: 'read_text', onBehalfOf: 'maintainer' }, 'REQUESTER_OVERRIDE'],
  ['wrong fileId type: number', 'docx', { fileId: 1302148, operation: 'read_text' }, 'INVALID_ARGS'],
  ['wrong fileId type: path', 'xlsx', { fileId: '/Budget.xlsx', operation: 'workbook_overview' }, 'INVALID_ARGS'],
  ['wrong fileId type: missing', 'docx', { operation: 'read_text' }, 'INVALID_ARGS'],
];

for (const [name, format, args, code] of REFUSALS) {
  test(`PRE-DISPATCH refusal -- ${name} -> ${code}: no broker call, no driver, no browser`, async () => {
    const b = fakeBroker(format);
    let created = 0;
    const t = createReadTool(format, { configSource: b.source, loadPlaywright: () => { throw new Error('no browser'); }, createDriver: () => { created += 1; throw new Error('no driver'); } });
    const r = await t.handler(args);
    assert.equal(r.ok, false);
    assert.equal(r.stage, 'PRE_DISPATCH');
    assert.equal(r.code, code, JSON.stringify(r));
    assert.equal(r.dispatched, false);
    assert.equal(b.calls.length, 0);
    assert.equal(created, 0);
    assert.equal(JSON.stringify(r).includes(TOKEN), false, 'SECRET_LEAK=NO');
  });
}

test('the refusal classification is the validator\'s, independent of the handler (same codes)', () => {
  for (const [, format, args, code] of REFUSALS) assert.equal(validateReadRequest(format, args).code, code);
});

// ---- exposure (M11.3a: withheld; M11.3b.2, owner decision 2026-10-10: the two READ tools released) ------------------

test('M11.4b.1 exposure (GH #228 proposed state): tools/list = product_status + xlsx_read + docx_read + xlsx_execute_task + docx_execute_task', async () => {
  const config = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: BROKER, instance: 'production', timeoutMs: 1000 } };
  const { registry, tools } = buildServerTools({ config, configIdentity: 'fixture' });
  // M11.4a.1 built xlsx_execute_task; M11.4b.1 released it. GH #219 3 built docx_execute_task
  // too (same shape, same gate) but left it withheld; GH #228 (PROPOSED, owner decision pending) serves it:
  // `tools` (the exposed subset below) now carries all five.
  assert.deepEqual(Object.keys(registry).sort(), ['docx_execute_task', 'docx_read', 'product_status', 'xlsx_execute_task', 'xlsx_read']);
  assert.deepEqual(Object.keys(tools), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.deepEqual(TOOL_ROWS.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'YES').map(r => r.TOOL).sort(), []);

  const input = new PassThrough();
  const output = new PassThrough();
  const lines = [];
  output.on('data', d => lines.push(...String(d).split('\n').filter(Boolean).map(l => JSON.parse(l))));
  const log = { info() {}, error() {}, warn() {}, debug() {} };
  const server = createMcpServer({ tools, info: { version: 'test' }, log, input, output });
  const done = server.start();
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/list' })}\n`);
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'xlsx_execute_task', arguments: {} } })}\n`);
  input.write(`${JSON.stringify({ jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'docx_execute_task', arguments: {} } })}\n`);
  input.end();
  await done;
  const byId = Object.fromEntries(lines.map(m => [m.id, m]));
  assert.deepEqual(byId[1].result.tools.map(x => x.name), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.equal(byId[2].error, undefined, 'xlsx_execute_task is callable (an empty task is refused by its own schema)');
  assert.equal(byId[2].result.structuredContent.ok, false);
  assert.equal(byId[3].error, undefined, 'docx_execute_task is callable (an empty call is refused by its own schema, pre-dispatch)');
  assert.equal(byId[3].result.structuredContent.ok, false);
});

test('the product default with no browsers directory configured: an accepted READ stops at BROWSER_RUNTIME_UNAVAILABLE, no broker call', async () => {
  const b = fakeBroker('docx');
  const { registry } = buildServerTools({ config: { broker: { url: BROKER } } });
  const t = createReadTool('docx', { configSource: b.source });
  const r = await t.handler({ fileId: FILE, operation: 'read_text' });
  assert.equal(r.ok, false);
  assert.equal(r.stage, 'RUNTIME');
  assert.equal(r.code, 'BROWSER_RUNTIME_UNAVAILABLE');
  assert.equal(b.calls.length, 0, 'the browser runtime check refuses before the config is fetched');
  assert.equal(typeof registry.docx_read.handler, 'function');
});
