'use strict';

const { parseRange, rangeText } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `format_run` family: per-run (character-range) font writes inside ONE cell (M7 backlog
 * X-F04 -- design doc `docs/XF04-RUN-FONT-DESIGN-20261007.md`,
 * OWNER-APPROVED).
 *
 * A SEPARATE family from `format.cjs`'s `format_range` (not a `font.runs` extension of it): the
 * run-addressing (character start/length) and the single-cell scope are a different
 * granularity than `format_range`'s flat, per-cell key model -- the design doc's own point 1
 * weighed both and the owner approved the separate-family path, the same precedent `cf.cjs`
 * already set for a differently-shaped "kind of formatting".
 *
 * Contract (F04 LIVE-measured): `GetCharacters(start, len).GetFont().SetBold/SetItalic/
 * SetSize/SetName/SetColor` -- all 5 axes persistent on ONE run, the neighbour run and the
 * control cell untouched. Only these 5 axes are in scope (the X-F05 whole-cell effects --
 * underline/strikeout/subscript/superscript -- are `format_range`'s own, unrelated capability).
 *
 * Run-map (design-review correction): the driver enumerates every run PUBLICLY, character by
 * character (`GetCharacters(i, 1).GetFont()`, the SAME `fontAt` the F04 probe and `format.cjs`'s
 * own whole-cell font reader already use), merging consecutive identical font-tuples. This is
 * N getter calls for an N-character cell -- bounded by `xlsx-format-run-text-too-long` (zero
 * 256 accepted as a TEMPORARY ceiling, pending a LIVE timing measurement in the
 * acceptance harness).
 *
 * L1 = every requested run's requested axes hold on a fresh run-map read. L3 = the SAME
 * run-map, before vs after: every character OUTSIDE every requested run's window must carry
 * the identical font-tuple, and the cell's text itself must be unchanged (a length change makes
 * the run-map's position-based comparison meaningless -- its own named failure,
 * `run-map-length-changed`, never a silent mismatch). Retry = the terminal state on a fresh
 * run-map read -> SATISFIED, 0 write / 0 command.
 *
 * BOUNDARY (design review point 4, stated explicitly): the run-map -- and therefore L1/L3both --
 * compares ONLY the 5 MEASURED axes (bold, italic, name, size, color). `format_range`'s own
 * whole-cell effects (underline, strikeout, subscript, superscript) are NOT part of the
 * run-map at all; a character outside every requested run keeping those axes unchanged is NOT
 * measured here one way or the other -- this family simply never looks at them.
 *
 * The SAVED FILE's own bytes (`file-verifier.cjs` v21, `readRichTextRuns`) are the LIVE
 * ACCEPTANCE's external evidence, never this family's own L3 (design review: the product does
 * not read the saved file at runtime; file-verifier is a test tool).
 */

const INTENTS = Object.freeze(['format_run']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'cell', 'runs']);
const RUN_MAP_CHAR_CEILING = 256; // TEMPORARY, pending a LIVE timing measurement (see the harness's timing row)
const HEX = /^#[0-9A-Fa-f]{6}$/;
const FONT_KEYS = Object.freeze(['bold', 'italic', 'name', 'size', 'color']);

const rgbOf = hex => [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16));
const intOf = hex => { const [r, g, b] = rgbOf(hex); return r * 65536 + g * 256 + b; };

/** A run's `font` sub-object -> its validated want entries ({bold?,italic?,name?,size?,color?}
 * in EDITOR terms, color as the RGB int), or an error string. Only the 5 MEASURED axes --
 * unlike `format.cjs`'s own `font` key, no underline/strikeout/subscript/superscript here. */
function runFontWant(f) {
  if (!f || typeof f !== 'object' || Array.isArray(f) || !Object.keys(f).length) return { error: 'run.font must be a non-empty object' };
  const bad = Object.keys(f).filter(k => !FONT_KEYS.includes(k));
  if (bad.length) return { error: `run.font key(s) not supported: ${bad.join(', ')}` };
  if ('bold' in f && typeof f.bold !== 'boolean') return { error: 'run.font.bold must be true or false' };
  if ('italic' in f && typeof f.italic !== 'boolean') return { error: 'run.font.italic must be true or false' };
  if ('name' in f && !(typeof f.name === 'string' && f.name && f.name.length <= 100)) return { error: 'run.font.name must be a non-empty string up to 100 chars' };
  if ('size' in f && !(Number.isFinite(f.size) && f.size >= 1 && f.size <= 409)) return { error: 'run.font.size must be a number in [1, 409]' };
  if ('color' in f && !(typeof f.color === 'string' && HEX.test(f.color))) return { error: "run.font.color must be '#RRGGBB'" };
  const want = {};
  if ('bold' in f) want.bold = f.bold;
  if ('italic' in f) want.italic = f.italic;
  if ('name' in f) want.name = f.name;
  if ('size' in f) want.size = f.size;
  if ('color' in f) want.color = intOf(f.color);
  return { want };
}

/** The `runs` array -> validated `{start, length, want}` entries (document order), or an error
 * string. Plan-time only: shape + mutual overlap -- NOT the cell's text length (that needs the
 * LIVE read, checked in `gate()`). */
function runsWant(runs) {
  if (!Array.isArray(runs) || runs.length === 0) return { error: 'runs must be a non-empty array' };
  const parsed = [];
  for (const [i, r] of runs.entries()) {
    if (!r || typeof r !== 'object' || Array.isArray(r)) return { error: `runs[${i}] must be an object` };
    const bad = Object.keys(r).filter(k => !['start', 'length', 'font'].includes(k));
    if (bad.length) return { error: `runs[${i}] key(s) not supported: ${bad.join(', ')}` };
    if (!Number.isInteger(r.start) || r.start < 1) return { error: `runs[${i}].start must be an integer >= 1` };
    if (!Number.isInteger(r.length) || r.length < 1) return { error: `runs[${i}].length must be an integer >= 1` };
    const fw = runFontWant(r.font);
    if (fw.error) return { error: `runs[${i}]: ${fw.error}` };
    parsed.push({ start: r.start, length: r.length, want: fw.want });
  }
  const sorted = [...parsed].sort((a, b) => a.start - b.start);
  for (let i = 1; i < sorted.length; i += 1) {
    if (sorted[i].start < sorted[i - 1].start + sorted[i - 1].length) return { error: 'runs must not overlap' };
  }
  return { runs: parsed };
}

function plan(step, index) {
  const at = `step ${index}: format_run`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.cell);
  if (!g || g.count !== 1) return { ok: false, detail: `${at} requires a single A1 cell (not a range)` };
  const rw = runsWant(step.runs);
  if (rw.error) return { ok: false, detail: `${at}: ${rw.error}` };
  const cellAddr = rangeText(g);
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'format_run',
      family: 'format_run',
      selector: Object.freeze({ sheet: step.sheet, range: cellAddr }),
      cells: Object.freeze([Object.freeze({ a: cellAddr })]),
      runs: Object.freeze(rw.runs.map(r => Object.freeze({ ...r, want: Object.freeze(r.want) }))),
      footprint: Object.freeze({ style: cellAddr }),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** The run-map key `readLive`'s `runReads` is keyed by -- `sheet!cell`, the SAME `sheet!range`
 * string shape `cf.cjs`'s own `cfReads` uses. */
const runKey = op => `${op.selector.sheet}!${op.selector.range}`;

/** `primary.runReads[key]` -> `{ ok: true, text, runs }` or `{ ok: false, reason }` -- the
 * driver's own UNOBSERVABLE markers translated to this family's named reasons. */
function runMapOf(primary, op) {
  const v = primary && primary.runReads ? primary.runReads[runKey(op)] : undefined;
  if (v === undefined) return { ok: false, reason: 'xlsx-format-run-unobservable', why: 'not-read' };
  if (typeof v === 'string') {
    if (v === 'UNOBSERVABLE:text-too-long') return { ok: false, reason: 'xlsx-format-run-text-too-long' };
    if (v === 'UNOBSERVABLE:empty-cell') return { ok: false, reason: 'xlsx-format-run-on-empty-cell' };
    return { ok: false, reason: 'xlsx-format-run-unobservable', why: v };
  }
  return { ok: true, text: v.text, runs: v.runs };
}

/** The shared font-tuple of `runMap` over `[start, start+length)`, or `null` if that window is
 * NOT uniform (spans more than one of the map's own merged segments) -- "not satisfied yet",
 * never an error. */
function windowFont(runMap, start, length) {
  const end = start + length; // exclusive
  const covering = runMap.filter(r => r.start < end && r.start + r.length > start);
  if (covering.length !== 1) return null;
  const seg = covering[0];
  if (seg.start > start || seg.start + seg.length < end) return null;
  return seg.font;
}

const holdsWant = (font, want) => font !== null && Object.entries(want).every(([k, v]) => font[k] === v);

function cellRefusal(primary, op, cell, a) {
  if (!cell || !('f' in cell)) return { reason: 'xlsx-target-format-unread', cell: a };
  if (cell.f !== '') return { reason: 'xlsx-format-run-on-formula-unmeasured', cell: a };
  if (cell.merge && cell.merge !== a) return { reason: 'xlsx-format-run-in-merged-area-refused', cell: a, mergeArea: cell.merge };
  return null;
}

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported };
  const s = primary.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: primary.sheets.map(x => x.name) };
  const a = op.selector.range;
  const cell = cellAt(primary, op.selector.sheet, a);
  const refused = cellRefusal(primary, op, cell, a);
  if (refused) return { ok: false, ...refused };
  const rm = runMapOf(primary, op);
  if (!rm.ok) return { ok: false, ...rm };
  for (const r of op.runs) {
    if (r.start + r.length - 1 > rm.text.length) return { ok: false, reason: 'xlsx-format-run-out-of-bounds', run: r, textLength: rm.text.length };
  }
  return { ok: true, cell, runMap: rm };
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, cell: op.selector.range, runs: op.runs.map(r => ({ start: r.start, length: r.length, keys: Object.keys(r.want) })) };
  const satisfied = op.runs.every(r => holdsWant(windowFont(g.runMap.runs, r.start, r.length), r.want));
  return satisfied ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-run-map-every-run' } } : { state: 'APPLY', evidence: where };
}

function resolve(primary, selector, ctx = {}) {
  const g = gate(primary, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

/** The design review's own correction: the FONT object's (GetCharacters(...).GetFont()) MEASURED
 * setter names are `SetBold`/`SetItalic`/`SetName`/`SetSize`/`SetColor` -- NOT the RANGE-level
 * `SetFontName`/`SetFontSize`/`SetFontColor` `format.cjs`'s own `r.SetFontXxx` calls use on a
 * DIFFERENT object. Using the wrong name on `SetColor` threw AFTER `SetBold` had already run,
 * a genuine PARTIAL WRITE -- fixed structurally below (point 2), not just by the name fix. */
const FONT_SETTER = Object.freeze({ bold: 'SetBold', italic: 'SetItalic', name: 'SetName', size: 'SetSize', color: 'SetColor' });

function lower(op) {
  const { sheet, range } = op.selector;
  const J = JSON.stringify;
  const colour = hex => `Api.CreateColorFromRGB(${hex})`; // hex is already the editor rgbOf(...) triple, joined by the caller
  // Correction (LIVE run2, own diagnostic xf04-diag-two-handles): resolving EVERY
  // run's font-object HANDLE before writing to any of them (my own v2 fix for the run1 bug)
  // has a MEASURED side effect -- a LATER GetCharacters(...).GetFont() call invalidates an
  // EARLIER handle's writes (RUN1 read back unchanged, in-session and after a reopen). The
  // MEASURED rule (F04 probe, and the timing cell's own single-run pattern): get ONE handle,
  // write to it IMMEDIATELY, before getting the next. The mechanism (does a later GetCharacters
  // invalidate the earlier handle, or does the second handle's write silently win) is NOT
  // measured; the RULE is.
  //
  // The setter-existence pre-check (design review point 2, still required -- partial writes must
  // still be refused before ANY write) therefore runs on ONE handle only: method existence on
  // the font object is a PROTOTYPE-level fact, not run-dependent, so checking the UNION of every
  // requested setter name on the FIRST run's handle is equivalent to checking each run's own
  // handle -- without needing a second, invalidating handle just to check it.
  const allKeys = [...new Set(op.runs.flatMap(r => Object.keys(r.want)))];
  const guard = `const cf0 = c.GetCharacters(${op.runs[0].start}, ${op.runs[0].length}).GetFont();
  if (${allKeys.map(k => `typeof cf0.${FONT_SETTER[k]} !== 'function'`).join(' || ')}) return { ok: false, stage: 'setter-missing', which: ${J(allKeys)} };`;
  // Runs AFTER the first get their OWN handle, but ONLY right before their own writes (never
  // ahead of time) -- run 0 reuses cf0 (the SAME handle the guard already resolved, written to
  // immediately right after, never left idle while another handle is fetched).
  const runLines = op.runs.map((r, i) => {
    const w = r.want;
    const cf = i === 0 ? 'cf0' : `cf${i}`;
    const getter = i === 0 ? '' : `const ${cf} = c.GetCharacters(${r.start}, ${r.length}).GetFont();\n    `;
    const setters = [];
    if ('bold' in w) setters.push(`${cf}.SetBold(${J(w.bold)});`);
    if ('italic' in w) setters.push(`${cf}.SetItalic(${J(w.italic)});`);
    if ('name' in w) setters.push(`${cf}.SetName(${J(w.name)});`);
    if ('size' in w) setters.push(`${cf}.SetSize(${J(w.size)});`);
    if ('color' in w) {
      const [rr, gg, bb] = [16, 8, 0].map(sh => (w.color >> sh) & 255);
      setters.push(`${cf}.SetColor(${colour(`${rr}, ${gg}, ${bb}`)});`);
    }
    return `{ ${getter}${setters.join(' ')} }`;
  });
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${J(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const c = ws.GetRange(${J(range)});
  ${guard}
  ${runLines.join('\n  ')}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'format_run', sheet, cell: range, runs: op.runs.map(r => ({ start: r.start, length: r.length, want: r.want })) }),
  };
}

/** L1: every requested run's requested axes hold on a fresh run-map read. design review point 3:
 * a 'wrong' run's evidence carries the ACTUALLY observed `windowFont` next to `want`, not just
 * start/length -- the run1 bug's own diagnosis needed reading the source for this, since the
 * evidence alone didn't show WHICH axis mismatched or what it read instead. */
function verifyPrimitive(op, before, after) {
  const a = after && after.primary;
  const rm = runMapOf(a, op);
  if (!rm.ok) return { ok: false, evidence: rm };
  const wrong = op.runs.filter(r => !holdsWant(windowFont(rm.runs, r.start, r.length), r.want)).map(r => ({ start: r.start, length: r.length, want: r.want, observed: windowFont(rm.runs, r.start, r.length) }));
  return { ok: wrong.length === 0, evidence: { sheet: op.selector.sheet, cell: op.selector.range, wrong, proof: 'fresh-live-run-map-every-run' } };
}

/** L3: the SAME run-map, before vs after -- every character OUTSIDE every requested run's
 * window must be unchanged, and the text itself unchanged (a length change is its own named
 * failure, never a silent position-mismatch). Plus the usual cross-cell content preservation
 * every family builds on. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = runMapOf(before && before.primary, op);
  const a = runMapOf(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-format-run-unobservable', before: b.ok || b.reason, after: a.ok || a.reason } };
  if (b.text.length !== a.text.length) return { ok: false, evidence: { ...rest.evidence, reason: 'run-map-length-changed', before: b.text.length, after: a.text.length } };
  const inWindow = i => op.runs.some(r => i >= r.start && i < r.start + r.length);
  const changed = [];
  for (let i = 1; i <= b.text.length; i += 1) {
    if (inWindow(i)) continue;
    const bf = windowFont(b.runs, i, 1);
    const af = windowFont(a.runs, i, 1);
    if (JSON.stringify(bf) !== JSON.stringify(af)) changed.push(i);
  }
  // X-L13/6: once tiered, report preservationScope FULL -- the run-map comparison above never leaves the
  // target's OWN cell (runReads is keyed by the exact requested (sheet,cell) pair, xlsx-driver.cjs's own
  // `runAsked` loop -- never a used-range scan), so nothing is reduced by tiering. Below 45000 unchanged.
  const tiered = before && before.primary && before.primary.tier === 'tiered';
  return { ok: rest.ok && changed.length === 0, evidence: { ...rest.evidence, charsCompared: b.text.length, charsChanged: changed.slice(0, 20), charsChangedCount: changed.length, ...(tiered ? { preservationScope: 'FULL' } : {}) } };
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const rm = runMapOf(p, op);
  const wrong = rm.ok ? op.runs.filter(r => !holdsWant(windowFont(rm.runs, r.start, r.length), r.want)).map(r => `${r.start}+${r.length}`) : ['unobservable'];
  return [{ name: `${op.selector.sheet}!${op.selector.range} holds every requested run's font`, ok: rm.ok && wrong.length === 0, detail: wrong.length ? { wrong } : null }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range }];
const liveExtras = op => ({ runCells: [{ sheet: op.selector.sheet, cell: op.selector.range }] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, RUN_MAP_CHAR_CEILING };
