'use strict';

const cells = require('./cells.cjs');

/**
 * XLSX `page` family: set_page_layout (M6 wave 8). Audit PP1-PP4 MIGRATE on the PUBLIC ApiWorksheet getters /
 * setters (Get/SetPageOrientation, Get/SetPrintGridlines, Get/SetPrintHeadings).
 *
 * The reader is PUBLIC; the driver serves it only on a build listed in EO_PAGE_LAYOUT_PER_SHEET_RUNTIMES -- the W8b
 * probe (2026-09-29, case L4) proved there that the getters read the TARGET sheet, not the active one (the W4b2 trap
 * on GetFreezePanes), in session and after reopen. Any other build: refused before the write.
 *
 * Contract: set_page_layout { sheet, orientation?: 'portrait' | 'landscape', printGridlines?: boolean,
 * printHeadings?: boolean, topMargin? / bottomMargin? / leftMargin? / rightMargin?: millimetres } -- at least one.
 * The margin unit is MEASURED: millimetres (W8: SetLeftMargin(72) read back 72 after reopen; W8b: the public
 * defaults 19.1 / 17.8 equal the internal millimetre margins). A margin holds within 0.01 mm.
 * SATISFIED: every requested field already holds. L1 the requested fields; L3 the target's UNrequested fields, every
 * other sheet's layout, every cell. Retry SATISFIED 0/0.
 */

const INTENTS = Object.freeze(['set_page_layout']);
const ORIENT = Object.freeze({ portrait: 'xlPortrait', landscape: 'xlLandscape' });
const MARGINS = Object.freeze(['topMargin', 'bottomMargin', 'leftMargin', 'rightMargin']);
const MAX_MARGIN_MM = 200;
const FIELDS = Object.freeze(['orientation', 'printGridlines', 'printHeadings', ...MARGINS]);

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const unknown = Object.keys(step).filter(k => !['intent', 'id', 'sheet', ...FIELDS].includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const want = {};
  if ('orientation' in step) {
    if (!Object.hasOwn(ORIENT, step.orientation)) return { ok: false, detail: `${at}: orientation must be 'portrait' or 'landscape'` };
    want.orientation = ORIENT[step.orientation];
  }
  for (const k of ['printGridlines', 'printHeadings']) {
    if (!(k in step)) continue;
    if (typeof step[k] !== 'boolean') return { ok: false, detail: `${at}: ${k} must be true or false` };
    want[k] = step[k];
  }
  for (const k of MARGINS) {
    if (!(k in step)) continue;
    if (!(Number.isFinite(step[k]) && step[k] >= 0 && step[k] <= MAX_MARGIN_MM)) return { ok: false, detail: `${at}: ${k} must be millimetres in [0, ${MAX_MARGIN_MM}]` };
    want[k] = step[k];
  }
  if (!Object.keys(want).length) return { ok: false, detail: `${at} requires orientation, printGridlines, printHeadings and / or a margin` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'page',
      // a sheet-level setting: no cell footprint; it claims the sheet's layout (two layout steps on ONE sheet clash)
      selector: Object.freeze({ sheet: step.sheet, range: 'A1', cellFree: true, claims: Object.freeze([`page:${step.sheet}`]) }),
      want: Object.freeze(want),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const readable = x => !!x && typeof x === 'object' && typeof x.orientation === 'string' && typeof x.printGridlines === 'boolean' && typeof x.printHeadings === 'boolean' && MARGINS.every(k => Number.isFinite(x[k]));
const holds = (st, want) => Object.entries(want).every(([k, v]) => (MARGINS.includes(k) ? Math.abs(st[k] - v) < 0.01 : st[k] === v));

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const reads = p.pageReads;
  const bad = !reads ? ['not-read'] : p.sheets.map(x => x.name).filter(n => !readable(reads[n]));
  if (bad.length) return { ok: false, reason: 'xlsx-page-layout-unobservable', sheets: bad.slice(0, 20), why: reads ? bad.map(n => reads[n]).slice(0, 5) : 'not-read' };
  return { ok: true, st: reads[op.selector.sheet] };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, want: op.want };
  return holds(g.st, op.want) ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-page-layout' } } : { state: 'APPLY', evidence: { ...where, now: g.st } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet }) };
}

function lower(op) {
  const { sheet } = op.selector;
  const w = op.want;
  const calls = [
    'orientation' in w ? `ws.SetPageOrientation(${JSON.stringify(w.orientation)});` : '',
    'printGridlines' in w ? `ws.SetPrintGridlines(${w.printGridlines});` : '',
    'printHeadings' in w ? `ws.SetPrintHeadings(${w.printHeadings});` : '',
    ...MARGINS.filter(k => k in w).map(k => `ws.Set${k[0].toUpperCase()}${k.slice(1)}(${w[k]});`),
  ].filter(Boolean).join('\n  ');
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${calls}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'page', action: op.intent, sheet, want: op.want }),
  };
}

function verifyPrimitive(op, _before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  return { ok: holds(g.st, op.want), evidence: { sheet: op.selector.sheet, want: op.want, now: g.st, proof: 'fresh-live-page-layout' } };
}

/** L3: every cell; the target's UNrequested fields; every other sheet's layout. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary.pageReads;
  const a = after.primary.pageReads;
  const rest3 = reads => (reads ? Object.fromEntries(Object.entries(reads).map(([s, st]) => [s, s === op.selector.sheet && readable(st) ? Object.fromEntries(FIELDS.map(k => [k, k in op.want ? null : st[k]])) : st])) : null);
  const kept = !!b && !!a && JSON.stringify(rest3(b)) === JSON.stringify(rest3(a));
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherLayoutKept: kept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  return [{ name: `${op.selector.sheet}: page layout ${JSON.stringify(op.want)}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ pageSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
