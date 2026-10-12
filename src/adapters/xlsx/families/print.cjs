'use strict';

const { XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS } = require('../runtime.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `print` family (M6 wave 8): the page-setup / print settings the W8b probe (2026-09-29,
 * docs/evidence/m6-w8b-probe-20260929.json) proved to WRITE and to READ BACK after save + reopen on 9.3.4-9475e9f1.
 * Since M11.5b.4 the write goes through the DS Bridge WRITE sheet.printSettings.write/v1 (/7), which runs the SAME
 * engine setters the old direct-internal path called, in the same order (PARITY, M11.5b.1).
 * The reader is the driver's printReads, through the DS Bridge sheet.printSettings.read/v1 (B09, Wave B READ /4); a
 * missing bridge or another patch is unobservable and refuses.
 * Print area and print titles are NOT here: they persisted only through the editor view OUTSIDE callCommand (W8b),
 * which is not our dispatch path.
 *
 * Only the measured shapes (each fixes an audit defect where the source had one):
 *   set_print_fit { sheet, fitToWidth, fitToHeight } -- SATISFIED only when both numbers hold AND the sheet's
 *     fitToPage FLAG reads on (audit PP5: the source set the flag opportunistically and never read it); M7
 *     backlog X-P (m7-s28-probe1 MEASURED): fitToWidth=0 AND fitToHeight=0 (fit-to-page
 *     OFF) is now also accepted -- SATISFIED when both numbers hold AND the flag reads OFF (the flag's
 *     expected state flips with this one shape; every other fit value still requires it ON)
 *   set_header_footer { sheet, slot: 'oddHeader' | 'oddFooter', value } -- a non-empty code, or null = clear the slot
 *     (M7 F2 Q9: footer set, footer clear, header clear MEASURED); first / even slots (PP17 / PP18) are not measured
 *   set_page_break { sheet, axis: 'row' | 'column', at, remove? } -- a manual break stored as id at-1 (the source's
 *     convention, W8b measured id 10 for at 11; F2 Q11: a column break at E is id 4); `remove: true` removes exactly
 *     that manual break (F2 Q12, row and column), the others stay; where it lands in print is NOT measured (PP19 L2);
 *     an automatic break or a reset is refused
 *   set_first_page_number { sheet, value } -- a positive integer, or null = clear the number (F2 Q14: the number goes,
 *     the useFirstPage flag stays as it was)
 *   set_print_centering { sheet, horizontal: true | false } or { sheet, vertical: true | false } -- one axis per
 *     step (F2 Q15: vertical true and horizontal false MEASURED; vertical false widened by --
 *     the driver's vCenter read is `=== true`, so a never-set sheet ALSO reads false; a false request is
 *     SATISFIED at once on such a sheet, same as any other axis's untouched-default shortcut)
 * L1 the requested setting; L3 every other print setting of the sheet, every other sheet's print settings, every
 * cell. Retry SATISFIED 0/0.
 */

/**
 * B10-WRITE through the DS Bridge (M11.5b.4, owner: M11.5 continuous mandate, internal tracker): the measured shapes
 * below (the contract the owner approved for the former EO_INTERNAL_DEPENDENCY WRITE, 2026-09-30 / M7 GO 2026-10-03 /
 * #ref #ref) are written ONLY through sheet.printSettings.write/v1, ONE setting per command, after the prewrite READ.
 * The capability exists only on /7: on any other runtime the step is refused before the write (the driver's
 * B10_BRIDGE_WRITE_RUNTIMES is the second line, the in-editor gate the third). No direct-internal fallback exists. The
 * receipt is never the verification: L1/L3 run on the fresh sheet.printSettings.read/v1. Proving the settings does not
 * prove a printed page number or the printed appearance.
 */
const B10_BRIDGE_WRITE_RUNTIMES = XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS;

const INTENTS = Object.freeze(['set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering', 'set_print_scale', 'set_paper_size', 'set_header_footer_flag']);
const FIELDS = Object.freeze({
  set_print_fit: ['fitToWidth', 'fitToHeight'],
  set_header_footer: ['slot', 'value'],
  set_page_break: ['axis', 'at', 'remove'],
  set_first_page_number: ['value'],
  set_print_centering: ['horizontal', 'vertical'],
  set_print_scale: ['scale'],
  set_paper_size: ['paper'],
  set_header_footer_flag: ['flag', 'value'],
});
/** The sheet-level setting a step claims (and, for L3, the fields of its own sheet it may change). */
const axisOf = (intent, w) => ({ set_print_fit: 'fit', set_print_scale: 'fit', set_paper_size: 'paper', set_header_footer_flag: w.flag, set_header_footer: w.slot, set_page_break: w.breaks, set_first_page_number: 'firstPage', set_print_centering: w.center }[intent]);
/**
 * Owner OD-C / OD-A (owner channel): the paper sizes are EXACTLY the editor's own menu (web-apps view/Toolbar.js
 * pageSize items, the mm the UI passes to asc_changeDocSize); the file stores a size id found by an exact mm match
 * (sdkjs Serialize.js getSizeByWH, 2024-2032, every size below is in that table) -- any other width / height is saved as A4
 * silently, so nothing else is admitted.
 */
const PAPER = Object.freeze({
  'US Letter': [215.9, 279.4], 'US Legal': [215.9, 355.6], A4: [210, 297], A5: [148, 210], B5: [176, 250],
  'Envelope #10': [104.8, 241.3], 'Envelope DL': [110, 220], Tabloid: [279.4, 431.8], A3: [297, 420],
  'Tabloid Oversize': [296.9, 457.2], 'ROC 16K': [196.8, 273], 'Envelope Choukei 3': [120, 235], 'Super B/A3': [305, 487],
});
const BREAK_AXIS = Object.freeze({ row: 'rowBreaks', column: 'colBreaks' });
const colNum = s => (/^[A-Z]{1,3}$/.test(s) ? [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) : 0);

const posInt = (x, max) => Number.isInteger(x) && x >= 1 && x <= max;

/** X-P07 (OD-A, owner channel; p42 / S31b measured: texts + flags persist): the odd, first and even slots; a
 * first / even text is stored independently of its enabling flag (set_header_footer_flag), as the editor stores it. */
const HF_SLOTS = Object.freeze(['oddHeader', 'oddFooter', 'firstHeader', 'firstFooter', 'evenHeader', 'evenFooter']);
const HF_FLAGS = Object.freeze({ differentFirst: 'diffFirst', differentOddEven: 'diffOddEven' });

function headerWant(step, at) {
  if (!HF_SLOTS.includes(step.slot)) return { error: `${at}: slot must be one of ${HF_SLOTS.join(', ')}` };
  if (step.value !== null && (typeof step.value !== 'string' || step.value === '' || step.value.length > 255)) return { error: `${at}: value must be a non-empty code, or null to clear the slot` };
  return { want: { slot: step.slot, value: step.value } };
}

function breakWant(step, at) {
  if (!(step.axis in BREAK_AXIS)) return { error: `${at}: axis must be 'row' or 'column'` };
  if ('remove' in step && step.remove !== true) return { error: `${at}: remove must be true when given` };
  const n = step.axis === 'row' ? (posInt(step.at, 1048576) ? step.at : 0) : colNum(step.at);
  if (n < 2 || (step.axis === 'column' && n > 16384)) return { error: `${at}: at must be a row number >= 2 (row) or a column letter after A (column)` };
  return { want: { breaks: BREAK_AXIS[step.axis], breakId: n - 1, remove: step.remove === true } };
}

function centerWant(step, at) {
  const keys = ['horizontal', 'vertical'].filter(k => k in step);
  if (keys.length !== 1) return { error: `${at}: give exactly one of horizontal / vertical` };
  if (keys[0] === 'horizontal' && typeof step.horizontal !== 'boolean') return { error: `${at}: horizontal must be true or false` };
  if (keys[0] === 'vertical' && typeof step.vertical !== 'boolean') return { error: `${at}: vertical must be true or false` };
  return keys[0] === 'horizontal' ? { want: { center: 'hCenter', value: step.horizontal } } : { want: { center: 'vCenter', value: step.vertical } };
}

/** set_print_scale (the UI's asc_SetPrintScale(0, 0, n): fit OFF + the scale) and set_paper_size. */
function pageWant(step, at) {
  if (step.intent === 'set_header_footer_flag') {
    return Object.hasOwn(HF_FLAGS, step.flag) && typeof step.value === 'boolean' ? { want: { flag: HF_FLAGS[step.flag], value: step.value } } : { error: `${at}: flag must be 'differentFirst' or 'differentOddEven' with a boolean value` };
  }
  if (step.intent === 'set_print_scale') {
    return Number.isInteger(step.scale) && step.scale >= 10 && step.scale <= 400 ? { want: { scale: step.scale } } : { error: `${at}: scale must be an integer percentage in [10, 400]` };
  }
  return Object.hasOwn(PAPER, step.paper) ? { want: { paper: step.paper, width: PAPER[step.paper][0], height: PAPER[step.paper][1] } } : { error: `${at}: paper must be one of the editor's sizes: ${Object.keys(PAPER).join(', ')}` };
}

/** OD-A / OD-C page intents, table-driven: what holds and the fields the step may change on its own sheet. */
const PAGE = Object.freeze({
  set_print_scale: { holds: (st, w) => st.fitW === 0 && st.fitH === 0 && st.fitFlag === false && st.scale === w.scale, own: () => ['fitW', 'fitH', 'fitFlag', 'scale'] },
  set_paper_size: { holds: (st, w) => st.width === w.width && st.height === w.height, own: () => ['width', 'height'] },
  set_header_footer_flag: { holds: (st, w) => st[w.flag] === w.value, own: w => [w.flag] },
});
const wantOf = (step, at) => (Object.hasOwn(PAGE, step.intent) ? pageWant(step, at) : wantBase(step, at));
const holds = (st, op) => (Object.hasOwn(PAGE, op.intent) ? PAGE[op.intent].holds(st, op.want) : holdsBase(st, op));
const ownFields = op => (Object.hasOwn(PAGE, op.intent) ? PAGE[op.intent].own(op.want) : ownBase(op));

function wantBase(step, at) {
  switch (step.intent) {
    case 'set_print_fit':
      if (!(Number.isInteger(step.fitToWidth) && step.fitToWidth >= 0 && step.fitToWidth <= 100) || !(Number.isInteger(step.fitToHeight) && step.fitToHeight >= 0 && step.fitToHeight <= 100)) return { error: `${at}: fitToWidth / fitToHeight must be integers in [0, 100]` };
      return { want: { fitW: step.fitToWidth, fitH: step.fitToHeight } };
    case 'set_header_footer': return headerWant(step, at);
    case 'set_page_break': return breakWant(step, at);
    case 'set_first_page_number':
      if (step.value !== null && !posInt(step.value, 32767)) return { error: `${at}: value must be a positive integer, or null to clear the number` };
      return { want: { firstPage: step.value } };
    default: return centerWant(step, at);
  }
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const unknown = Object.keys(step).filter(k => !['intent', 'id', 'sheet', ...FIELDS[step.intent]].includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const w = wantOf(step, at);
  if (w.error) return { ok: false, detail: w.error };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'print',
      // a sheet-level setting: no cell footprint; it claims its axis of the sheet
      selector: Object.freeze({ sheet: step.sheet, range: 'A1', cellFree: true, claims: Object.freeze([`print:${step.sheet}:${axisOf(step.intent, w.want)}`]) }),
      want: Object.freeze(w.want),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const readable = x => !!x && typeof x === 'object' && Number.isInteger(x.fitW) && Number.isInteger(x.fitH) && typeof x.fitFlag === 'boolean'
  && (x.oddHeader === null || typeof x.oddHeader === 'string') && Array.isArray(x.rowBreaks) && typeof x.useFirstPage === 'boolean' && typeof x.hCenter === 'boolean' && typeof x.scale === 'number' && typeof x.width === 'number' && typeof x.height === 'number';
/** The F2-widened fields a step reads on its own sheet (absent on an unwidened read -> unobservable, refused). */
const textOrNull = x => x === null || typeof x === 'string';
const WIDE = Object.freeze({ oddFooter: textOrNull, vCenter: x => typeof x === 'boolean', colBreaks: Array.isArray,
  firstHeader: textOrNull, firstFooter: textOrNull, evenHeader: textOrNull, evenFooter: textOrNull, diffFirst: x => typeof x === 'boolean', diffOddEven: x => typeof x === 'boolean' });
const needsWide = op => op.intent === 'set_header_footer_flag' || (op.intent === 'set_header_footer' && op.want.slot !== 'oddHeader') || (op.intent === 'set_page_break' && op.want.breaks === 'colBreaks') || (op.intent === 'set_print_centering' && op.want.center === 'vCenter');
const wideReadable = st => Object.entries(WIDE).every(([k, ok]) => k in st && ok(st[k]));

const fitOff = w => w.fitW === 0 && w.fitH === 0;

function holdsBase(st, op) {
  const w = op.want;
  switch (op.intent) {
    // fitW=fitH=0 is fit-to-page OFF (m7-s28-fit-off-probe MEASURED): the flag flips to OFF
    // with it, instead of the ON every other fit value requires.
    // Owner 1/a (owner channel): OFF also means scale 100, as the editor's UI fit-off does (p53: an editor fit
    // left scale 16, the old fit-off kept it).
    case 'set_print_fit': return st.fitW === w.fitW && st.fitH === w.fitH && st.fitFlag === !fitOff(w) && (!fitOff(w) || st.scale === 100);
    case 'set_header_footer': return st[w.slot] === w.value;
    case 'set_page_break': { const has = st[w.breaks].some(b => b.id === w.breakId && b.man === true); return w.remove ? !st[w.breaks].some(b => b.id === w.breakId) : has; }
    case 'set_first_page_number': return w.firstPage === null ? st.firstPage === null : st.useFirstPage === true && st.firstPage === w.firstPage;
    default: return st[w.center] === w.value;
  }
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const reads = p.printReads;
  const bad = !reads ? ['not-read'] : p.sheets.map(x => x.name).filter(n => !readable(reads[n]));
  if (bad.length) return { ok: false, reason: 'xlsx-print-unobservable', sheets: bad.slice(0, 20), why: reads ? bad.map(n => reads[n]).slice(0, 5) : 'not-read' };
  const st = reads[op.selector.sheet];
  const gap = stepGap(op, st);
  return gap ? { ok: false, ...gap } : { ok: true, st };
}

/** M7 F2: the step's own sheet must serve the widened fields it reads; a removal names a MANUAL break (an automatic
 * one at the same id is not measured). */
function stepGap(op, st) {
  if (needsWide(op) && !wideReadable(st)) return { reason: 'xlsx-print-unobservable', sheets: [op.selector.sheet], why: 'widened-fields-not-read' };
  const auto = op.intent === 'set_page_break' && op.want.remove && st[op.want.breaks].some(b => b.id === op.want.breakId && b.man !== true);
  return auto ? { reason: 'xlsx-print-break-not-manual', breakId: op.want.breakId } : null;
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, want: op.want };
  if (holds(g.st, op)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-print-settings' } };
  const runtime = p.runtime ? p.runtime.identity : null;
  if (!B10_BRIDGE_WRITE_RUNTIMES.includes(runtime)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-print-bridge-write-runtime-not-proven', runtime, proven: B10_BRIDGE_WRITE_RUNTIMES } };
  return { state: 'APPLY', evidence: { ...where, now: g.st } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet }) };
}

/**
 * The ONE Bridge WRITE a step makes: the capability's arguments and the receipt it must return (`applied`, the
 * sheet.printSettings.read/v1 field names in the capability's own order -- the driver's client checks it exactly).
 */
function bridgeArgs(op) {
  const w = op.want;
  switch (op.intent) {
    case 'set_print_fit': return { args: { setting: 'fit', fitToWidth: w.fitW, fitToHeight: w.fitH }, requested: fitOff(w) ? ['fitW', 'fitH', 'fitFlag', 'scale'] : ['fitW', 'fitH', 'fitFlag'] };
    case 'set_print_scale': return { args: { setting: 'scale', scale: w.scale }, requested: ['fitW', 'fitH', 'fitFlag', 'scale'] };
    case 'set_paper_size': return { args: { setting: 'paper', width: w.width, height: w.height }, requested: ['width', 'height'] };
    case 'set_first_page_number': return { args: { setting: 'firstPage', firstPage: w.firstPage }, requested: ['firstPage'] };
    case 'set_header_footer': return { args: { setting: 'headerFooter', slot: w.slot, value: w.value }, requested: [w.slot] };
    case 'set_header_footer_flag': return { args: { setting: 'headerFooterFlag', flag: w.flag === 'diffFirst' ? 'differentFirst' : 'differentOddEven', value: w.value }, requested: [w.flag] };
    case 'set_page_break': return { args: { setting: 'pageBreak', axis: w.breaks === 'colBreaks' ? 'column' : 'row', id: w.breakId, remove: w.remove }, requested: [w.breaks] };
    case 'set_print_centering': return { args: { setting: 'center', axis: w.center === 'hCenter' ? 'horizontal' : 'vertical', value: w.value }, requested: [w.center] };
    default: return null; // an intent without a Bridge mapping never becomes some other setting's write
  }
}

/**
 * One command, one Bridge WRITE. __eoB10Gate / __eoB10Write / __eoB10Fail come from the driver's injected client (no
 * client -> ReferenceError -> caught below -> not ok). The gate runs ONCE, before any mutation; a failed write stops the
 * command: no second write, no fallback, no rollback.
 */
function lower(op) {
  const { sheet } = op.selector;
  const mapped = bridgeArgs(op);
  if (!mapped) return { kind: 'callCommand', body: `return { ok: false, stage: 'no-bridge-mapping', intent: ${JSON.stringify(op.intent)} };`, meta: Object.freeze({ family: 'print', action: op.intent, sheet, want: op.want, bridgePrintWrite: true }) };
  const { args, requested } = mapped;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const g10 = __eoB10Gate(name); if (!g10.ok) return g10;
  const w = __eoB10Write(${JSON.stringify(args)}, ${JSON.stringify(requested)}); if (!w.ok) return __eoB10Fail(w);
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'print', action: op.intent, sheet, want: op.want, bridgePrintWrite: true }),
  };
}

function verifyPrimitive(op, _before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  return { ok: holds(g.st, op), evidence: { sheet: op.selector.sheet, want: op.want, now: g.st, proof: 'fresh-live-print-settings' } };
}

/** The fields one step may change on its own sheet (L3 masks exactly these). A first-page CLEAR may not touch the
 * useFirstPage flag (F2: it stays as it was). */
function ownBase(op) {
  const w = op.want;
  switch (op.intent) {
    case 'set_print_fit': return fitOff(w) ? ['fitW', 'fitH', 'fitFlag', 'scale'] : ['fitW', 'fitH', 'fitFlag'];
    case 'set_header_footer': return [w.slot];
    case 'set_page_break': return [w.breaks];
    case 'set_first_page_number': return w.firstPage === null ? ['firstPage'] : ['useFirstPage', 'firstPage'];
    default: return [w.center];
  }
}

/** A break step changes exactly its own break: an add keeps every other, a removal keeps every other too. */
function breaksKept(op, b, a) {
  if (op.intent !== 'set_page_break') return true;
  const k = op.want.breaks;
  const [x, y] = [b && b[op.selector.sheet], a && a[op.selector.sheet]];
  if (!x || !y || !Array.isArray(x[k]) || !Array.isArray(y[k])) return false;
  return x[k].filter(z => z.id !== op.want.breakId).every(z => y[k].some(v => v.id === z.id && v.man === z.man))
    && y[k].filter(z => z.id !== op.want.breakId).length === x[k].filter(z => z.id !== op.want.breakId).length;
}

function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary.printReads;
  const a = after.primary.printReads;
  const own = ownFields(op);
  const mask = reads => (reads ? Object.fromEntries(Object.entries(reads).map(([s, st]) => [s, s === op.selector.sheet && readable(st) ? { ...st, ...Object.fromEntries(own.map(k => [k, null])) } : st])) : null);
  const kept = !!b && !!a && JSON.stringify(mask(b)) === JSON.stringify(mask(a)) && breaksKept(op, b, a);
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherPrintSettingsKept: kept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  return [{ name: `${op.selector.sheet}: ${op.intent} ${JSON.stringify(op.want)}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ printSheets: ['*'] });

module.exports = { INTENTS, B10_BRIDGE_WRITE_RUNTIMES, bridgeArgs, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
