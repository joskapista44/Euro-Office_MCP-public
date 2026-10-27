'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');

/**
 * DOCX `table-cell` family (M4 Wave 4a): replace_table_cell (P1, PRIMITIVE_INTEGRAL).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   target      docx-agent-preflight.cjs:66-73 / docx-agent-w012-live-lowering.cjs:28-33
 *               (exactly one cell in all tables whose normalized text is the selector text)
 *   lower       docx-session-tables.cjs:9-13, :32-33 (Search + Select, then frame.evaluate
 *               selectCell + asc_GetSelectedText + asc_AddText) -- the backdoor the owner brief forbids
 *   retry       docx-agent-idempotency.cjs:165-173 (old 0 + new 1 -> SATISFIED; old>1 or new>1 ->
 *               ambiguous; 0/0 -> terminal-unprovable)
 *   L1/L3       docx-tables.cjs:107-200 (set_cell_text; table shape :163; every unaffected cell :185)
 *   L2          docx-agent-whole-verify.cjs:75 (old 0, new 1)
 *
 * Measured live before migrating (test/live/docx-table-probe*.live.cjs, 2026-09-24):
 *   - the public read works inside callCommand: GetAllTables / GetRowsCount / GetRow /
 *     GetCellsCount / GetCell / GetContent().GetText (the text ends in '\t' or '\r\n'; a merged
 *     cell joins its paragraphs with '\r\n'); the HTML projection renders a merged cell colspan=2;
 *   - the rewrite family's Search-range replacement is a SILENT NO-OP inside a cell (probe 1);
 *   - the cell's OWN content API replaces the whole cell text (probe 2, all four candidates
 *     `stage: done`, the fresh re-read equal to the in-session read): keep the first paragraph,
 *     remove every other element, replace the first paragraph's runs with ONE run carrying the old
 *     first paragraph's effective text properties. Bold/italic survived on a formatted cell; a
 *     two-paragraph merged cell became one paragraph and kept colspan=2.
 *   Font family/size survival was not readable from probe 2's HTML; here L3 asserts it through the
 *   driver's table read, so the LIVE acceptance measures it.
 *
 * Where the target differs from the source, stated:
 *   identity    the driver's PUBLIC table read is the cell authority (not the HTML projection,
 *               which joins a merged cell's paragraphs without a separator); its shape must agree
 *               with the HTML tables, else FAIL_CLOSED. The cell resolves from TABLE identity +
 *               CELL identity (owner, W4a closure point 2): an optional tableAnchor (the full text
 *               of a cell) names ONE table, the cell text must be unique in that scope; text
 *               outside the tables, or in another table named away by the anchor, is never an
 *               ambiguity. The command resolves the SAME scope and identity -- Search is not used
 *               (it matches substrings, probe 1).
 *   write       only a cell whose elements are all paragraphs made only of runs (a nested table,
 *               a hyperlink, a note reference or a bookmark marker -> FAIL_CLOSED before dispatch,
 *               and the command re-checks it): replacing the whole cell text must not silently
 *               destroy structure. The whole cell text is replaced; a multi-paragraph cell becomes
 *               one paragraph (the first paragraph's paragraph properties and text properties stay).
 *   retry       counted in the selector's table scope, cells only: SATISFIED needs old 0 and new
 *               exactly once (the source counted paragraphs too); with task.recovery naming this operation -> FAIL_CLOSED
 *               preservation-baseline-unavailable (the P1 baseline is gone once the text changed).
 *               old 1 + new 1 -> FAIL_CLOSED before dispatch (the source applied, then L2 failed).
 *   L3          every other cell (every table): text, paragraph count, content shape, first-
 *               paragraph properties unchanged; the target keeps every first-paragraph property the
 *               read names (bold, italic, underline, strike, fontFamily, fontSize, style, alignment --
 *               W4a run 2 read all but fontFamily; since W4c the family is read per non-empty run,
 *               column probe 1, and a MIXED family is not claimed); table shapes, HTML colspan and
 *               rowspan unchanged; the text units outside the tables unchanged IN ORDER.
 *   selector    { text }, { tableCell: { text } }, or both with the same text (the source's form);
 *               no ref, no row/column hints.
 *   text        non-empty, no tab or line break, different from the old text.
 *   chained     two steps on the same cell in one task are not supported: the first step's L2
 *               (old 0, new 1) fails closed after the second.
 */

const CELL_KEYS = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'style', 'alignment'];

/**
 * The cell's semantic identity: TABLE identity + CELL identity (owner, W4a closure point 2).
 *   { text } | { tableCell: { text } } | the source's canonical { tableCell: { text }, text }
 *   (test_docx_agent_table_cell_terminal_live.cjs; the combined form only when both name the SAME text)
 *   tableCell.tableAnchor -- the full text of a cell that names exactly ONE table; without it, the
 *   cell text must be unique across all tables (two tables holding it -> FAIL_CLOSED, never a guess).
 * Returns { text, tableAnchor } or null.
 */
function cellSelector(sel) {
  if (!sel || typeof sel !== 'object') return null;
  const keys = Object.keys(sel).sort().join(',');
  const tc = sel.tableCell && typeof sel.tableCell === 'object' ? sel.tableCell : null;
  const tcKeys = tc ? Object.keys(tc).sort().join(',') : '';
  const cell = tc && typeof tc.text === 'string' && (tcKeys === 'text' || (tcKeys === 'tableAnchor,text' && typeof tc.tableAnchor === 'string'))
    ? { text: tc.text, tableAnchor: tcKeys === 'text' ? null : tc.tableAnchor } : null;
  if (keys === 'text' && typeof sel.text === 'string') return { text: sel.text, tableAnchor: null };
  if (keys === 'tableCell' && cell) return cell;
  if (keys === 'tableCell,text' && cell && typeof sel.text === 'string' && norm(sel.text) === norm(cell.text)) return cell;
  return null;
}

function plan(step, index) {
  if (step.intent !== 'replace_table_cell') return { ok: false, detail: `step ${index}: ${step.intent} is not a table-cell intent` };
  const sel = cellSelector(step.selector);
  const old = sel && sel.text;
  if (!sel || !norm(old)) return { ok: false, detail: `step ${index}: a replace_table_cell selector is { text }, { tableCell: { text [, tableAnchor] } } or the combined form with the same non-empty text` };
  if (sel.tableAnchor !== null && (!norm(sel.tableAnchor) || norm(sel.tableAnchor) === norm(old) || norm(sel.tableAnchor) === norm(step.text))) {
    return { ok: false, detail: `step ${index}: a tableAnchor must be a non-empty cell text other than the old and the new cell text (it must survive the replacement)` };
  }
  if (typeof step.text !== 'string' || !norm(step.text)) return { ok: false, detail: `step ${index}: replace_table_cell requires a non-empty text` };
  if (/[\t\r\n]/.test(step.text)) return { ok: false, detail: `step ${index}: replace_table_cell text must not contain a tab or line break` };
  if (norm(step.text) === norm(old)) return { ok: false, detail: `step ${index}: replace_table_cell text equals the cell text` };
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, oldText: norm(old), tableAnchor: sel.tableAnchor === null ? null : norm(sel.tableAnchor), text: step.text, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) } };
}

// ---- the observation ----------------------------------------------------------------------------

/**
 * The driver's table read, accepted only when its shape (tables, rows, cells per row) equals the
 * HTML projection's -- two independent reads of the same tables; a disagreement means one of them
 * does not see what the other does (e.g. a nested table), and nothing is claimed.
 */
function tableRead(primary) {
  const d = primary && primary.document;
  const tables = d ? d.tables : null;
  if (!Array.isArray(tables)) return { ok: false, reason: 'tables-unreadable' };
  const html = ((primary && primary.blocks) || []).filter(b => b.kind === 'table');
  // a VERTICAL merge (row probe 4, measured): the projection renders rowspan on the top cell and
  // OMITS the continuation cell, while the driver read keeps it as an empty cell -- the two reads
  // no longer describe the same grid for THAT table, and a row delete on the top row lost the
  // merged content (row probe 3). This function refuses the moment ANY table in the document has
  // one -- DOCUMENT-WIDE, not scoped to the table actually being acted on. W06-3 (point 2,
  // LIVE-measured READ_DIVERGENCE: a merge-free table's two reads agree regardless of what
  // another table carries) found this too wide for `replace_table_cell` and the column family,
  // which now use `tableReadWithVerticalMerges` (below) instead -- PER-TABLE reconciled, with
  // their own callers (`cellState`/`classify`, table-column.cjs's `locate`) refusing a merge only
  // in the TARGET table. `tableRead` itself is unchanged and still exported; nothing in this file
  // calls it anymore as of W06-3, but a future caller wanting the document-wide behaviour can.
  if (html.some(t => (t.rows || []).some(r => r.some(c => c.rowspan > 1)))) return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:vertical-merge' };
  const shape = ts => JSON.stringify(ts.map(t => (t.rows || []).map(r => r.length)));
  if (shape(tables) !== shape(html)) return { ok: false, reason: 'table-read-shape-mismatch' };
  const cells = [];
  tables.forEach((t, ti) => t.rows.forEach((row, ri) => row.forEach((cell, ci) => cells.push({ at: `${ti}:${ri}:${ci}`, cell }))));
  return { ok: true, tables, html, cells };
}

/**
 * M5 (owner decision after Talk 28963, internal tracker): the table read for the ROW family on tables
 * WITH a vertical merge. The projection renders rowspan on the owner and OMITS the continuation cell;
 * the driver read keeps it as an (empty) cell. The two reads are RECONCILED row by row instead of being
 * refused: driver cells(r) == projected cells(r) + the owners from rows above whose rowspan covers r.
 * Any row that does not reconcile -> 'vertical-merge-grid-unreconciled' (fail closed: an uncertain grid
 * is never guessed). Per table, `vspans` lists every vertical merge as [top, bottom] row indices.
 */
function tableReadWithVerticalMerges(primary) {
  const d = primary && primary.document;
  const tables = d ? d.tables : null;
  if (!Array.isArray(tables)) return { ok: false, reason: 'tables-unreadable' };
  const html = ((primary && primary.blocks) || []).filter(b => b.kind === 'table');
  if (tables.length !== html.length) return { ok: false, reason: 'table-read-shape-mismatch' };
  const vspans = [];
  for (let ti = 0; ti < tables.length; ti += 1) {
    const hrows = html[ti].rows || [];
    if (hrows.length !== tables[ti].rows.length) return { ok: false, reason: 'table-read-shape-mismatch' };
    const spans = [];
    hrows.forEach((row, r) => row.forEach(c => { if (c.rowspan > 1) spans.push([r, r + c.rowspan - 1]); }));
    for (let r = 0; r < hrows.length; r += 1) {
      const covered = spans.filter(([top, bottom]) => top < r && r <= bottom).length;
      if (tables[ti].rows[r].length !== hrows[r].length + covered) return { ok: false, reason: spans.length ? 'vertical-merge-grid-unreconciled' : 'table-read-shape-mismatch', table: ti, row: r };
    }
    if (spans.some(([, bottom]) => bottom >= hrows.length)) return { ok: false, reason: 'vertical-merge-grid-unreconciled', table: ti };
    vspans.push(spans);
  }
  const cells = [];
  tables.forEach((t, ti) => t.rows.forEach((row, ri) => row.forEach((cell, ci) => cells.push({ at: `${ti}:${ri}:${ci}`, cell }))));
  return { ok: true, tables, html, cells, vspans };
}

/**
 * The table scope of the selector: with a tableAnchor, the ONE table that has a cell with exactly
 * that text (none / several -> not ok); without, every table. Text outside the tables, and cells of
 * other tables when a tableAnchor names the table, are never an ambiguity (owner, W4a closure 2).
 */
function scopeOf(read, tableAnchor) {
  if (tableAnchor === null || tableAnchor === undefined) return { ok: true, prefix: null, cells: read.cells };
  const tis = [...new Set(read.cells.filter(c => c.cell.text === tableAnchor).map(c => tableOf(c.at)))];
  if (tis.length !== 1) return { ok: false, reason: tis.length ? 'table-anchor-ambiguous' : 'table-anchor-not-found', tables: tis.length };
  return { ok: true, prefix: tis[0], cells: read.cells.filter(c => tableOf(c.at) === tis[0]) };
}

/**
 * GH #291 N3 (LIVE, docs/evidence/gh291-n3-nested-cell-probe-20261022.json): the cells of a nested table -- one level down, bound
 * to its holding cell (cell.nestedTables) -- are cells too: `at` = `<holder at>><nested table>:<row>:<cell>`, their table is the
 * nested table itself (an anchor of the HOLDING table does not reach them: measured, the command finds no such cell there).
 * The write's own command reaches them through GetAllTables(); the holding cell's read text runs on into the nested text.
 * GH #291 N3x (LIVE, docs/evidence/gh291-nested-remaining-probe-20261023.json): TWO levels down the same -- only that cell changes,
 * and every holding cell above it (`holders`, nearest first) reads the new text in place of the old. Deeper is not measured: such a
 * cell is not in this list (its write ends `table-cell-terminal-unprovable`).
 */
const NESTED_DEPTH_MEASURED = 2;
const tableOf = at => { const i = at.lastIndexOf('>'); const head = at.slice(i + 1).split(':')[0]; return i < 0 ? head : `${at.slice(0, i)}>${head}`; };
function withNestedCells(read) {
  const nested = [];
  const walk = (at, cell, holders) => {
    if (holders.length >= NESTED_DEPTH_MEASURED || !Array.isArray(cell.nestedTables)) return;
    cell.nestedTables.forEach((nt, n) => nt.rows.forEach((row, r) => row.forEach((c, k) => {
      const here = { at: `${at}>${n}:${r}:${k}`, cell: c, holder: at, holders: [at, ...holders] };
      nested.push(here);
      walk(here.at, c, here.holders);
    })));
  };
  for (const { at, cell } of read.cells) walk(at, cell, []);
  return { ...read, cells: [...read.cells, ...nested] };
}
const occurrences = (hay, needle) => { let n = 0; for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + 1)) n += 1; return n; };

function cellState(primary, op) {
  // W06-3 (point 2): per-table reconciled, not the document-wide blanket -- `classify`
  // below refuses a vertical merge only in the SPECIFIC table the resolved target lives in.
  const read0 = tableReadWithVerticalMerges(primary);
  if (!read0.ok) return read0;
  const read = withNestedCells(read0);
  const scope = scopeOf(read, op.tableAnchor);
  if (!scope.ok) return scope;
  const old = scope.cells.filter(c => c.cell.text === op.oldText);
  const neu = scope.cells.filter(c => c.cell.text === norm(op.text));
  return { ok: true, read, scope, old: { cells: old }, neu: { cells: neu }, oldTotal: old.length, newTotal: neu.length };
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = cellState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason } };
  const counts = { oldCount: s.oldTotal, newCount: s.newTotal };
  if (s.oldTotal > 1 || s.newTotal > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-cell-ambiguous', ...counts } };
  if (s.oldTotal === 1) {
    if (s.newTotal === 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-cell-conflicting-old-and-new-state', ...counts } };
    const target = s.old.cells[0];
    // GH #291 N3 / N3x: a nested cell -- every holding cell's read text carries the old text; L3 expects exactly that one occurrence replaced
    if (target.holder !== undefined) {
      const proven = target.holders.every(h => { const holder = s.read.cells.find(c => c.at === h); return !!holder && occurrences(holder.cell.text, op.oldText) === 1; });
      if (!proven) return { state: 'FAIL_CLOSED', evidence: { reason: 'nested-table-holder-text-unprovable', at: target.at } };
    }
    // GH #291 T-02 / T-03: a vertical merge in the target table no longer refuses. LIVE-measured
    // (docs/evidence/gh291-vmerge-table-probe-20261021.json): this command on a cell in the owner's row, the
    // continuation's row, below the merge, and on the vMerge OWNER itself changes only that cell; the saved
    // vMerge restart/continue and gridSpan stay. A continuation cell (empty text) is never a target (plan
    // requires a non-empty old text); L3 still compares every cell and the HTML spans (rowspan included).
    if (!target.cell.simple) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-cell-not-plain-paragraphs-of-runs', at: target.at } };
    // GH #289: a field (PAGE ...) or an image reads as runs with no text -- the whole-cell replacement would drop it silently
    if (!Array.isArray(target.cell.inline)) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-cell-inline-content-unreadable', at: target.at } };
    if (target.cell.inline.length) return { state: 'FAIL_CLOSED', evidence: { reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:table-cell-replacement-would-drop-inline:${target.cell.inline.join(',')}`, at: target.at } };
    return { state: 'APPLY', evidence: { from: op.oldText, at: target.at, baseline: target.cell.format } };
  }
  if (s.newTotal === 0) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-cell-terminal-unprovable', ...counts } };
  if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable', at: s.neu.cells[0].at } };
  return { state: 'SATISFIED', evidence: { text: op.text, at: s.neu.cells[0].at, proof: 'fresh-live-old-text-absent-new-text-once-in-a-cell' } };
}

function resolve(primary, selector) {
  // W06-3 (point 2): the per-table reconciled read, so a target in a merge-free table can
  // still be located even when another table in the document carries a vertical merge.
  const read0 = tableReadWithVerticalMerges(primary);
  const up = r => r.toUpperCase().replace(/-/g, '_');
  if (!read0.ok) return { ok: false, target: null, selector, ambiguity: up(read0.reason) };
  const read = withNestedCells(read0);
  const sel = cellSelector(selector);
  if (!sel) return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const tableAnchor = sel.tableAnchor === null ? null : norm(sel.tableAnchor);
  const scope = scopeOf(read, tableAnchor);
  if (!scope.ok) return { ok: false, target: null, selector, ambiguity: up(scope.reason), candidates: scope.tables };
  const text = norm(sel.text);
  const hits = scope.cells.filter(c => c.cell.text === text);
  if (hits.length !== 1) return { ok: false, target: null, selector, ambiguity: hits.length ? 'AMBIGUOUS' : 'NOT_FOUND', candidates: hits.length };
  return { ok: true, target: { kind: 'table_cell', text, tableAnchor, at: hits[0].at }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'table_cell' && target.text === op.oldText && target.tableAnchor === op.tableAnchor ? { ok: true } : { ok: false, dispatched: false, detail: 'table-cell-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'table', action: 'cell', targetText: target.text, tableAnchor: target.tableAnchor, text: op.text });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (table probe 2, the content-API candidate): the cell is found by its FULL
 * normalized text across every table; it must be the only one, and hold only paragraphs made only
 * of runs; then the first paragraph stays, every other element goes, and its runs become one run
 * with the old first paragraph's effective text properties, and the font family of its non-empty
 * runs when they agree (the paragraph range reads it undefined: a leading empty run in the default
 * font, W4c column probe 1). Any guard miss returns unchanged.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
  const doc = Api.GetDocument();
  const cellsOf = t => { const out = []; for (let r = 0; r < t.GetRowsCount(); r += 1) { const row = t.GetRow(r); for (let c = 0; c < row.GetCellsCount(); c += 1) out.push(row.GetCell(c)); } return out; };
  const textOf = cell => norm(cell.GetContent().GetText(textOpts));
  let tables = doc.GetAllTables() || [];
  if (m.tableAnchor !== null) {
    tables = tables.filter(t => cellsOf(t).some(cell => textOf(cell) === m.tableAnchor));
    if (tables.length !== 1) return;
  }
  const hits = [];
  for (const t of tables) for (const cell of cellsOf(t)) if (textOf(cell) === m.targetText) hits.push(cell);
  if (hits.length !== 1) return;
  const content = hits[0].GetContent();
  const n = content.GetElementsCount();
  if (n < 1) return;
  for (let i = 0; i < n; i += 1) {
    const e = content.GetElement(i);
    if (!e || e.GetClassType() !== 'paragraph') return;
    for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') return; }
  }
  const first = content.GetElement(0);
  const pr = first.GetRange().GetTextPr();
  const fams = new Set();
  for (let j = 0; j < first.GetElementsCount(); j += 1) {
    const ch = first.GetElement(j);
    if (String(ch.GetText(textOpts) || '') === '') continue;
    const rp = ch.GetRange().GetTextPr();
    const f = rp && typeof rp.GetFontFamily === 'function' ? rp.GetFontFamily() : undefined;
    fams.add(f ? String(f) : '');
  }
  const font = fams.size === 1 ? [...fams][0] : '';
  for (let i = n - 1; i > 0; i -= 1) content.RemoveElement(i);
  first.RemoveAllElements();
  const run = Api.CreateRun();
  run.AddText(m.text);
  if (pr) run.SetTextPr(pr);
  if (font) run.SetFontFamily(font);
  first.AddElement(run);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

// the shape of every table, the nested ones too (GH #291 N3): the cell positions
const shapeOf = read => JSON.stringify(read.cells.map(c => c.at));
const pick = f => CELL_KEYS.filter(k => f[k] !== undefined).map(k => [k, f[k]]);

/** L1: the target cell (same place, same table shape) holds exactly the new text in one paragraph; old 0, new 1. */
function verifyPrimitive(op, before, after) {
  const b = cellState(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  if (b.old.cells.length !== 1) return { ok: false, evidence: { reason: 'before-target-cell-not-unique', candidates: b.old.cells.length } };
  const at = b.old.cells[0].at;
  const a = cellState(after.primary, op);
  if (!a.ok) return { ok: false, evidence: { at, reason: `after-${a.reason}` } };
  const cell = (a.read.cells.find(c => c.at === at) || {}).cell || null;
  const sameShape = shapeOf(a.read) === shapeOf(b.read);
  const ok = sameShape && !!cell && cell.text === norm(op.text) && cell.paragraphs === 1 && a.oldTotal === 0 && a.newTotal === 1;
  return { ok, evidence: { at, sameShape, cellText: cell ? cell.text : null, paragraphs: cell ? cell.paragraphs : null, oldCount: a.oldTotal, newCount: a.newTotal } };
}

/**
 * L3, PRIMITIVE_INTEGRAL (source docx-tables.cjs:163/185): table shapes and HTML spans unchanged;
 * every OTHER cell unchanged (text, paragraph count, content shape, first-paragraph properties);
 * the target keeps every first-paragraph property the fresh pre-mutation read could name (a
 * property the read returns undefined for -- mixed runs -- is not claimed); every text unit
 * outside the tables unchanged.
 */
function verifyPreservation(op, before, after) {
  const b = cellState(before && before.primary, op);
  if (!b.ok || b.old.cells.length !== 1) return { ok: false, evidence: { reason: b.ok ? 'before-target-cell-not-unique' : `before-${b.reason}` } };
  // W06-3 (point 2): the SAME per-table reconciled read as cellState/classify -- an
  // unrelated table's vertical merge must not make verification itself unreadable for an
  // operation classify() just allowed to dispatch (the `html`/`cells` shape this function reads
  // afterward is byte-identical between `tableRead` and `tableReadWithVerticalMerges`, see
  // tableReadWithVerticalMerges's own doc comment -- only the OK/refusal check differs).
  const a0 = tableReadWithVerticalMerges(after.primary);
  if (!a0.ok) return { ok: false, evidence: { reason: `after-${a0.reason}` } };
  const a = withNestedCells(a0);
  const at = b.old.cells[0].at;
  // GH #291 N3 / N3x: every holding cell of a nested target reads the new text in place of the old (classify proved it occurs once)
  const holders = b.old.cells[0].holders || [];
  const expectedOf = c => (holders.includes(c.at) ? { ...c.cell, text: c.cell.text.replace(op.oldText, norm(op.text)) } : c.cell);
  const sameShape = shapeOf(a) === shapeOf(b.read);
  const spans = r => JSON.stringify(r.html.map(t => t.rows.map(row => row.map(c => [c.span, c.rowspan]))));
  const sameSpans = spans(a) === spans(b.read);
  const afterAt = new Map(a.cells.map(c => [c.at, c.cell]));
  const key = c => JSON.stringify([c.text, c.paragraphs, c.simple, pick(c.format)]);
  const changed = sameShape ? b.read.cells.filter(c => c.at !== at && (!afterAt.get(c.at) || key(afterAt.get(c.at)) !== key(expectedOf(c)))).map(c => c.at) : null;
  const target = afterAt.get(at);
  // a MIXED font family (runs that disagree, W4c column probe 1) is not one family: never claimed
  const wanted = pick(b.old.cells[0].cell.format).filter(([k, v]) => !(k === 'fontFamily' && v === 'MIXED'));
  const targetProps = { expected: Object.fromEntries(wanted), ok: !!target && wanted.every(([k, v]) => JSON.stringify(target.format[k]) === JSON.stringify(v)) };
  const units = p => textUnits(p).map(u => norm(u.text));
  const sameUnits = JSON.stringify(units(after.primary)) === JSON.stringify(units(before.primary));
  const ok = sameShape && sameSpans && changed !== null && changed.length === 0 && targetProps.ok && sameUnits;
  return { ok, evidence: { at, sameShape, sameSpans, changedCells: changed, targetProperties: { ...targetProps, actual: target ? Object.fromEntries(pick(target.format)) : null }, textUnitsOutsideTablesUnchanged: sameUnits, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/** L2 (source docx-agent-whole-verify.cjs:75, stricter): old text absent, new text exactly once, in a cell. */
function wholeTaskChecks(op, finalObservation) {
  const s = cellState(finalObservation.primary, op);
  return [{ name: `${op.id}: table cell "${op.oldText}" -> "${norm(op.text)}" (in its table: old 0, new exactly once)`, ok: s.ok && s.oldTotal === 0 && s.newTotal === 1 }];
}

// refs: none (intent inventory) -- the cell text is not a text unit a later step can name
function producedText() {
  return null;
}

module.exports = { plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, tableRead, cellState, cellSelector, scopeOf, CELL_KEYS , tableReadWithVerticalMerges };
