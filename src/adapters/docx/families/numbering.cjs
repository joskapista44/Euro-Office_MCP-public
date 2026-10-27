'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { listState, levelReferencedBy, sharedAbstractWith, listMembersOf } = require('./list.cjs');

/**
 * DOCX `numbering` family (M7 backlog W-07, numbering pattern/start -- a NEW capability,
 * registered via src/adapters/docx/capabilities.cjs's NEW_INTENTS, disjoint from the 28/28
 * migration inventory, the owner decision's option b).
 * See docs/W07-NUMBERING-DESIGN-20261007.md for the full design.
 *
 * `set_numbering` FINE-TUNES a paragraph's EXISTING list membership (induló érték / pontos
 * sablon vagy egyedi formátum / utótag) -- it never creates one. `families/list.cjs` owns
 * STRUCTURE (apply_list/set_list_level/set_list_type/indent/outdent/remove_list); this family
 * reuses its `listState` (the ONLY helper list.cjs actually exports that this family needs) for
 * paragraph targeting and current-state reading -- list.cjs itself is NOT modified.
 *
 * SELF-CAUGHT CORRECTION (design doc v2 cited `keyText`/`isListed` as exported from list.cjs,
 * list.cjs:486-491 -- re-checking the real `module.exports` there, only `listState` is: `keyText`
 * and `isListed` are list.cjs-internal, unexported one-liners. Re-implemented LOCALLY below
 * (`keyTextOf`/`isListedState`) rather than exporting them from list.cjs, honoring the reviewer's
 * decision that list.cjs stays untouched.
 *
 * Measured (W07-1a..1d, W07-3):
 *   - `doc.CreateNumbering()` (arity 1, no-arg tried first and it is the one this probe chain
 *     used throughout) returns an ApiNumbering; `.GetLevel(0)` an ApiNumberingLevel.
 *   - the paragraph setter is `SetNumPr` (CONFIRMED working; `SetNumbering` also exists but is
 *     NOT this family's write path -- `list.cjs`'s own `apply_list` already owns assignment, this
 *     family only re-templates an EXISTING assignment).
 *   - `ApiNumberingLevel.SetCustomType(sType, sTextFormat, sAlign)` is the CANONICAL write for
 *     format/text/alignment -- `SetTemplateType(sType, sym)` is NEVER called by this family's
 *     `lower()` because it sets `lvlJc` as an UNDOCUMENTED, template-dependent side effect
 *     ('1.'/'1)'/'I.' -> 'right', 'a.'/others -> 'left', measured); `SetCustomType`
 *     takes an EXPLICIT align and is deterministic (confirmed: it also RECOVERS a list's lvlJc
 *     back to 'left' after a prior `SetTemplateType` call, W07-3 probe (c) -- i.e. this
 *     family's canonical path is ALSO the fix for a list `set_list_type` (list.cjs:363-369) may
 *     have right-aligned; that EXISTING product behaviour is a separate, named, NOT-fixed-here
 *     finding -- see the design doc point 2).
 *   - `SetStart(n)` persists (own abstractNum lvl).
 *   - `SetSuff('tab'|'space'|'none')` -- all three MEASURED: file words
 *     tab/space/nothing respectively; the ToJSON `suff` key's spelling for 'none' is UNMEASURED
 *     (review: "a ToJSON-kulcsot az acceptance-ben merd") -- L1 below accepts EITHER
 *     'none' or 'nothing' for that one case, named in code, not guessed silently.
 *   - the `ToJSON()`-derived `format.start` key is PRESENT only when EXPLICITLY set via
 *     SetStart(n); it is ABSENT on a default-level numbering (whose file nonetheless implicitly
 *     writes w:start=1) -- absence is read as UNOBSERVABLE, never assumed to mean 1.
 */

const CODES = Object.freeze(['1.', '1)', 'a.', 'I.', 'bullet', 'none']);
const SUFFIXES = Object.freeze(['tab', 'space', 'none']);
// Every `code` translates to an explicit (numFmt, lvlText, align) triple --
// the ONLY align value measured so far is 'left' (W07-1d/W07-3); a caller-requested OTHER align
// is therefore refused at plan time (v1 scope), never silently guessed.
const CODE_TABLE = Object.freeze({
  '1.': { numFmt: 'decimal', lvlText: '%1.' },
  '1)': { numFmt: 'decimal', lvlText: '%1)' },
  'a.': { numFmt: 'lowerLetter', lvlText: '%1.' },
  'I.': { numFmt: 'upperRoman', lvlText: '%1.' },
  bullet: { numFmt: 'bullet', lvlText: null }, // lvlText from `symbol` (default '-') -- filled in plan()
  none: { numFmt: 'none', lvlText: '' },
});
const DEFAULT_BULLET_SYMBOL = '-';
const MEASURED_ALIGN = Object.freeze(['left']);

// This is the SAME setter-not-in-STEP_KEYS situation page-setup.cjs's own
// comment documents (page-setup.cjs:40-48) -- field-contract.cjs's STEP_KEYS stays closed to the
// 28/28, and its own acceptance test asserts that 1:1, so a 29th/30th key there would break it.
// The fail-closed-unknown-field protection is re-implemented HERE, locally.
const FIELDS = Object.freeze(['code', 'custom', 'symbol', 'start', 'suffix']);
// GH #248: `sharedDefinition` is NOT a numbering field but the caller's POLICY for a definition other items sit on: 'propagate' (the default, the W-07 behaviour -- the co-members show the new value,
// L3 expects it) or 'refuse' (the product: a write that would change another item's numbering fails closed BEFORE any dispatch).
const SHARED_POLICIES = Object.freeze(['propagate', 'refuse']);
// GH #293 L7 (owner GO B, 2026-10-22): `scope` -- 'item' (the default, the released semantics: the target's level, and any other item it
// would change is a refusal) or 'list' (EXPLICIT: every item of the target's numbering at the target's level changes, named in the result;
// a definition shared with another list is first detached through the word /5 Bridge, only when the isolation is provable).
const SCOPES = Object.freeze(['item', 'list']);
const STEP_KEYS = Object.freeze(['intent', 'id', 'dependsOn', 'selector', ...FIELDS, 'sharedDefinition', 'scope']);

const isPosInt = v => Number.isInteger(v) && v > 0;

/** keyText/isListed re-implemented locally -- list.cjs exports `listState` only (see header). */
function keyTextOf(step) {
  if (step.selector && step.selector.ref !== undefined) return typeof step.refText === 'string' ? step.refText : null;
  return step.selector ? step.selector.text : null;
}
const isListedState = s => s.type === 'bullet' || s.type === 'numbered';

function validateCustom(custom, index) {
  if (!custom || typeof custom !== 'object' || Object.keys(custom).sort().join(',') !== 'align,format,text') {
    return { ok: false, detail: `step ${index}: set_numbering custom requires exactly { format, text, align }` };
  }
  if (typeof custom.format !== 'string' || !custom.format) return { ok: false, detail: `step ${index}: set_numbering custom.format must be a non-empty string` };
  if (typeof custom.text !== 'string') return { ok: false, detail: `step ${index}: set_numbering custom.text must be a string` };
  if (!MEASURED_ALIGN.includes(custom.align)) return { ok: false, detail: `step ${index}: set_numbering custom.align '${custom.align}' is not measured -- only ${MEASURED_ALIGN.join(', ')} in v1` };
  return { ok: true };
}

function plan(step, index) {
  if (step.intent !== 'set_numbering') return { ok: false, detail: `step ${index}: ${step.intent} is not a numbering intent` };
  for (const key of Object.keys(step)) {
    if (!STEP_KEYS.includes(key)) return { ok: false, detail: `field-unsupported: step ${index} (set_numbering).${key}` };
  }
  const sel = step.selector;
  if (!sel || typeof sel !== 'object') return { ok: false, detail: `step ${index}: set_numbering requires a selector` };
  const keys = Object.keys(sel);
  const textSel = keys.length === 1 && typeof sel.text === 'string' && sel.text.trim() !== '';
  const refSel = keys.length === 1 && typeof sel.ref === 'string' && sel.ref !== '';
  if (!textSel && !refSel) return { ok: false, detail: `step ${index}: a numbering selector is { text } or { ref } (got: ${keys.join(', ') || 'nothing'})` };

  if (step.sharedDefinition !== undefined && !SHARED_POLICIES.includes(step.sharedDefinition)) return { ok: false, detail: `step ${index}: set_numbering sharedDefinition must be one of ${SHARED_POLICIES.join(', ')}` };
  if (step.scope !== undefined && !SCOPES.includes(step.scope)) return { ok: false, detail: `step ${index}: set_numbering scope must be one of ${SCOPES.join(', ')}` };
  if (step.code !== undefined && step.custom !== undefined) return { ok: false, detail: `step ${index}: set_numbering code and custom are mutually exclusive` };
  if (step.symbol !== undefined && step.code !== 'bullet') return { ok: false, detail: `step ${index}: set_numbering symbol is only valid with code: 'bullet'` };

  const fields = {};
  if (step.code !== undefined) {
    if (!CODES.includes(step.code)) return { ok: false, detail: `step ${index}: set_numbering code must be one of ${CODES.join(', ')}` };
    const row = CODE_TABLE[step.code];
    const lvlText = step.code === 'bullet' ? (step.symbol !== undefined ? step.symbol : DEFAULT_BULLET_SYMBOL) : row.lvlText;
    if (step.code === 'bullet' && typeof lvlText !== 'string') return { ok: false, detail: `step ${index}: set_numbering symbol must be a string` };
    fields.format = Object.freeze({ numFmt: row.numFmt, lvlText, align: 'left' });
  } else if (step.custom !== undefined) {
    const v = validateCustom(step.custom, index); if (!v.ok) return v;
    fields.format = Object.freeze({ numFmt: step.custom.format, lvlText: step.custom.text, align: step.custom.align });
  }
  if (step.start !== undefined) {
    if (!isPosInt(step.start)) return { ok: false, detail: `step ${index}: set_numbering start must be a positive integer` };
    fields.start = step.start;
  }
  if (step.suffix !== undefined) {
    if (!SUFFIXES.includes(step.suffix)) return { ok: false, detail: `step ${index}: set_numbering suffix must be one of ${SUFFIXES.join(', ')}` };
    fields.suffix = step.suffix;
  }
  if (!fields.format && fields.start === undefined && fields.suffix === undefined) {
    return { ok: false, detail: `step ${index}: set_numbering requires at least one of code, custom, start, suffix` };
  }

  const op = {
    id: step.id || `s${index}`, intent: step.intent, selector: sel, fields: Object.freeze(fields),
    ...(step.sharedDefinition === 'refuse' ? { sharedDefinition: 'refuse' } : {}),
    ...(step.scope === 'list' ? { scope: 'list' } : {}),
    preservationMode: 'PRIMITIVE_INTEGRAL',
    ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
  };
  return { ok: true, operation: op };
}

// ---- the observation ------------------------------------------------------------------------

/** Which requested fields already hold the wanted value, reading each via the SAME ToJSON-sourced
 * `format` the driver/inventory populate -- never the write args. */
function fieldMatches(fmt, fields) {
  const out = {};
  if (fields.format) {
    out.format = !!fmt && fmt.numFmt === fields.format.numFmt && fmt.lvlText === fields.format.lvlText && fmt.lvlJc === fields.format.align;
  }
  if (fields.start !== undefined) {
    const hasStart = !!fmt && Object.prototype.hasOwnProperty.call(fmt, 'start');
    out.start = hasStart ? fmt.start === fields.start : 'UNOBSERVABLE';
  }
  if (fields.suffix !== undefined) {
    // the ToJSON spelling for 'none' is unmeasured -- 'nothing' accepted too, named here.
    const want = fields.suffix === 'none' ? ['none', 'nothing'] : [fields.suffix];
    out.suffix = !!fmt && want.includes(fmt.suff);
  }
  return out;
}

function numberingState(primary, text) {
  const s = listState(primary, text);
  if (!s.ok) return { ok: false, reason: s.reason };
  if (!isListedState(s)) return { ok: false, reason: 'set-numbering-target-not-listed' };
  return { ok: true, state: s };
}

// ---- classify / resolve ---------------------------------------------------------------------

/**
 * GH #248: the OTHER items that sit on the target's numbering AT THE TARGET'S LEVEL -- the level format is per numbering and level, so these are the items a write would change too.
 * Items of the same numbering on a DIFFERENT level are not changed by this level's format (their `%1.%2` style labels that quote this level are an unmeasured exception, named in the decision doc).
 */
function sharedWith(primary, state) {
  const want = norm(state.text);
  return numberingUnits(primary).filter(u => u.numId !== null && u.numId === state.numId && u.level === state.level && u.text !== want).map(u => u.text);
}

function classify(primary, op) {
  const text = keyTextOf(op);
  const ns = numberingState(primary, text);
  if (!ns.ok) return { state: 'FAIL_CLOSED', evidence: { reason: ns.reason, text } };
  const s = ns.state;
  const fmt = s.node && s.node.numbering ? s.node.numbering.format : null;
  const matches = fieldMatches(fmt, op.fields);
  const unobservable = Object.entries(matches).filter(([, v]) => v === 'UNOBSERVABLE').map(([k]) => k);
  // An UNOBSERVABLE field (e.g. `start` absent from ToJSON on a level that was
  // never explicitly set) is NEVER evidence of a match -- "hianyzo start != bizonyitott 1". It
  // must ALSO never BLOCK the attempt: re-writing an already-correct value is harmless and
  // idempotent, and the write itself makes the field OBSERVABLE afterward (verifyPrimitive below
  // fails closed if it still is not, which would indicate the measured expectation was wrong).
  const done = unobservable.length === 0 && Object.values(matches).every(v => v === true);
  if (done) {
    // GH #293 L7: a list-scope no-op names the items it covers (every member shares the target's level format)
    const mem = op.scope === 'list' ? listMembersOf(primary, s, text) : null;
    return { state: 'SATISFIED', evidence: { text, fields: op.fields, proof: 'fresh-live-numbering-format-matches-every-requested-field',
      ...(mem && mem.ok ? { affectedItems: affectedOf(mem, s) } : {}) } };
  }
  if (op.scope === 'list') return classifyList(primary, op, s, text, fmt, unobservable);
  if (op.sharedDefinition === 'refuse') {
    // no numId read = sharing cannot be observed: refuse rather than assume the definition is the target's alone
    if (s.numId === null || s.numId === undefined) return { state: 'FAIL_CLOSED', evidence: { reason: 'set-numbering-numid-unreadable', text } };
    // GH #293 L7: the co-members come from the editor's own numbering collection too (the inventory misses a table-cell item); not provable -> refuse
    const mem = listMembersOf(primary, s, text);
    if (!mem.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-members-unprovable:${mem.reason}`, text } };
    const others = [...new Set([...sharedWith(primary, s), ...mem.coMembers])];
    if (others.length) return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-shared-definition:${s.numId}:level${s.level}:${others.length}`, text, numId: s.numId, level: s.level, coMembers: others } };
    // GH #293 N5: another list (its own numId) on the same abstractNum would change too
    const inventoryLists = sharedAbstractWith(primary, s);
    if (inventoryLists === null) return { state: 'FAIL_CLOSED', evidence: { reason: 'set-numbering-abstract-definition-unreadable', text } };
    const otherLists = [...new Set([...inventoryLists, ...mem.otherLists])];
    if (otherLists.length) return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-shared-abstract-definition:${s.abstractId}:${otherLists.length}`, text, numId: s.numId, abstractId: s.abstractId, otherListItems: otherLists } };
    // GH #293 N6: a format (code / custom / symbol) or counter (start) change of this level also changes the labels of other-level items that quote it ("%1.%2"); a suffix does not
    if (op.fields.format !== undefined || op.fields.start !== undefined) {
      const quoting = [...new Set([...levelReferencedBy(primary, s), ...mem.quoting])];
      if (quoting.length) return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-level-referenced-by-other-levels:${s.numId}:level${s.level}:${quoting.length}`, text, numId: s.numId, level: s.level, referencedBy: quoting } };
    }
  }
  return { state: 'APPLY', evidence: { text, fields: op.fields, observed: fmt, unobservable } };
}

/** GH #293 L7: the items a list-scope write changes -- every member of the target's numbering at the target's level, document order */
const affectedOf = (mem, s) => mem.members.filter(m => m.level === s.level).map(m => m.text);
const coversLevels = (starts, members) => members.every(m => Array.isArray(starts) && starts.includes(m.level));

/**
 * GH #293 L7, scope 'list' (EXPLICIT). The affected set is the editor's own: every member of the target's numbering (word /5 Bridge read,
 * table cells included) -- not provable -> FAIL_CLOSED. A label that quotes the level from a deeper level still refuses (N6). No other
 * w:num on the definition -> APPLY the level write (it changes exactly the members at that level). Another list on the definition ->
 * APPLY with a detach FIRST, only when the count is provably isolated (measured: DS continues a count across two w:num on one definition
 * unless the later one has a startOverride; docs/evidence/gh293-l7-detach-probe-run2-clone-20261022.json), else FAIL_CLOSED, named.
 */
function classifyList(primary, op, s, text, fmt, unobservable) {
  if (s.numId === null || s.numId === undefined) return { state: 'FAIL_CLOSED', evidence: { reason: 'set-numbering-numid-unreadable', text } };
  const mem = listMembersOf(primary, s, text);
  if (!mem.ok) return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-list-members-unprovable:${mem.reason}`, text } };
  if ((op.fields.format !== undefined || op.fields.start !== undefined) && mem.quoting.length) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-level-referenced-by-other-levels:${s.numId}:level${s.level}:${mem.quoting.length}`, text, numId: s.numId, level: s.level, referencedBy: mem.quoting } };
  }
  const affectedItems = affectedOf(mem, s);
  const v = mem.value;
  if (v.others.length === 0) return { state: 'APPLY', evidence: { text, fields: op.fields, observed: fmt, unobservable, scope: 'list', affectedItems, detach: null } };
  const otherListItems = mem.otherLists;
  if (!v.order || typeof v.order.contiguous !== 'boolean' || typeof v.order.targetFirst !== 'boolean') {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'set-numbering-list-order-unreadable', text, otherListItems } };
  }
  const why = !v.order.contiguous ? 'another-list-inside-it'
    : v.others.some(o => o.before === true && o.after === true) ? 'wrapped-by-another-list'
      : !v.order.targetFirst && !coversLevels(v.starts, v.members) ? 'continues-a-preceding-list'
        : v.others.some(o => o.after === true && !coversLevels(o.starts, o.members)) ? 'a-following-list-continues-it' : null;
  if (why) return { state: 'FAIL_CLOSED', evidence: { reason: `set-numbering-list-not-isolatable:${why}`, text, otherListItems } };
  return { state: 'APPLY', evidence: { text, fields: op.fields, observed: fmt, unobservable, scope: 'list', affectedItems, otherListItems,
    detach: { text, level: s.level, members: v.members, others: v.others.length } } };
}

function resolve(primary, selector, ctx = {}) {
  let text = selector && selector.text;
  if (selector && selector.ref !== undefined) {
    const ref = ctx.refs && ctx.refs[selector.ref];
    text = ref && ref.evidence && ref.evidence.text;
    if (typeof text !== 'string') return { ok: false, target: null, selector, ambiguity: 'REF_WITHOUT_TEXT' };
  }
  const ns = numberingState(primary, text);
  if (!ns.ok) return { ok: false, target: null, selector, ambiguity: ns.reason.replace(/^set-numbering-|^list-/, '').toUpperCase() };
  // GH #293 L7: a list-scope write carries its plan (the detach the fresh targeting observation proved isolated, or none) to lower()
  const op = ctx.operation;
  let listPlan;
  if (op && op.scope === 'list') {
    const c = classifyList(primary, op, ns.state, text, null, []);
    if (c.state !== 'APPLY') return { ok: false, target: null, selector, ambiguity: 'LIST_SCOPE_NOT_PROVABLE', detail: c.evidence.reason };
    listPlan = { detach: c.evidence.detach, affectedItems: c.evidence.affectedItems };
  }
  return { ok: true, target: { kind: 'numbering_item', text: ns.state.text, state: ns.state, ...(listPlan ? { listPlan } : {}) }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  if (!target || !target.state) return { ok: false, dispatched: false, detail: 'numbering-state-unreadable' };
  if (!isListedState(target.state)) return { ok: false, dispatched: false, detail: 'set-numbering-target-not-listed' };
  return { ok: true };
}

// ---- lower ------------------------------------------------------------------------------------

function lower(op, target) {
  const s = target.state;
  const detach = target.listPlan && target.listPlan.detach ? target.listPlan.detach : null;
  const meta = Object.freeze({
    family: 'numbering', action: 'setNumbering', targetText: s.text, ownLevel: s.level, numId: s.numId,
    fields: op.fields,
    // GH #293 L7: the detach through the word /5 Bridge WRITE (the driver prepends __eoL7Gate / __eoL7Detach and reads the body's receipt)
    ...(detach ? { bridgeListDetach: true, detach } : {}),
  });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand. The target by its exact text, the live LEVEL re-checked against the fresh
 * observation before any mutation (list.cjs's own re-check shape, list.cjs:342-344). ALL-OR-
 * NOTHING (the W04-5 lesson): every requested setter's EXISTENCE is checked first;
 * if any is missing, NO write happens at all, not a partial one.
 */
/**
 * GH #293 L7: the list-scope detach, spliced into commandBody ONLY when the plan carries one (every other body stays byte-identical).
 * The setter checks first, then the Bridge gate (a refusal writes nothing), the detach, and the level write on the COPY -- the level is
 * read again from the paragraph, never reused from before the detach. The body's { ok: true } is the receipt the driver requires.
 */
const LIST_DETACH_BLOCK = `  if (m.fields.start !== undefined && typeof level.SetStart !== 'function') return;
  if (m.fields.suffix !== undefined && typeof level.SetSuff !== 'function') return;
  const g = __eoL7Gate();
  if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' };
  const d = __eoL7Detach(g.bridge, m.detach);
  if (!d.ok) return { ok: false, stage: 'list-detach', reason: d.reason, mutation: d.mutation };
  const moved = paragraph.GetNumbering();
  if (!moved || moved.GetLevelIndex() !== m.ownLevel) return { ok: false, stage: 'after-detach', reason: 'level-not-reread', mutation: 'DETACHED' };
  if (needsCustom) moved.SetCustomType(m.fields.format.numFmt, m.fields.format.lvlText, m.fields.format.align);
  if (m.fields.start !== undefined) moved.SetStart(m.fields.start);
  if (m.fields.suffix !== undefined) moved.SetSuff(m.fields.suffix);
  return { ok: true, detached: d.moved };
`;

function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  if (!doc) return;
  const hits = (doc.Search(m.targetText, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === m.targetText);
  if (hits.length !== 1) return;
  const paragraph = hits[0].GetParagraph(0);
  if (!paragraph) return;
  const level = typeof paragraph.GetNumbering === 'function' ? paragraph.GetNumbering() : null;
  if (!level) return;
  const liveLevel = typeof level.GetLevelIndex === 'function' ? level.GetLevelIndex() : null;
  if (liveLevel !== m.ownLevel) return;
  const needsCustom = !!m.fields.format;
  if (needsCustom && typeof level.SetCustomType !== 'function') return;
${meta.bridgeListDetach === true ? LIST_DETACH_BLOCK : ''}  if (m.fields.start !== undefined && typeof level.SetStart !== 'function') return;
  if (m.fields.suffix !== undefined && typeof level.SetSuff !== 'function') return;
  if (needsCustom) level.SetCustomType(m.fields.format.numFmt, m.fields.format.lvlText, m.fields.format.align);
  if (m.fields.start !== undefined) level.SetStart(m.fields.start);
  if (m.fields.suffix !== undefined) level.SetSuff(m.fields.suffix);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** L1: every requested field now matches, via the SAME session-sourced ToJSON format (LIVE_VERIFY
 * -- never dependent on the saved file). */
function verifyPrimitive(op, before, after) {
  const text = keyTextOf(op);
  const b = numberingState(before && before.primary, text);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = numberingState(after && after.primary, text);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const fmt = a.state.node && a.state.node.numbering ? a.state.node.numbering.format : null;
  const matches = fieldMatches(fmt, op.fields);
  const unobservable = Object.entries(matches).filter(([, v]) => v === 'UNOBSERVABLE').map(([k]) => k);
  const ok = unobservable.length === 0 && Object.values(matches).every(Boolean);
  if (op.scope !== 'list') return { ok, evidence: { text, fields: op.fields, observed: fmt, unobservable } };
  // GH #293 L7: the list kept exactly its members (the editor's own set, before and after), and after a detach it shares no definition
  const mb = listMembersOf(before.primary, b.state, text);
  const ma = listMembersOf(after.primary, a.state, text);
  const sameMembers = mb.ok && ma.ok && JSON.stringify(mb.members) === JSON.stringify(ma.members);
  const isolated = ma.ok && ma.value.others.length === 0;
  const detached = mb.ok && mb.value.others.length > 0;
  return { ok: ok && sameMembers && (!detached || isolated),
    evidence: { text, fields: op.fields, observed: fmt, unobservable, scope: 'list', affectedItems: ma.ok ? affectedOf(ma, a.state) : null, sameMembers, detached, isolated } };
}

/** Every text unit of a state, numbering-item numId attached when present (mirrors list.cjs's
 * own `units()`, which this family does NOT import -- the shape it needs is one line shorter:
 * no type/level comparison, just the numId identity). */
function numberingUnits(primary) {
  const out = [];
  for (const b of (primary && primary.blocks) || []) {
    if (['paragraph', 'heading'].includes(b.kind)) out.push({ text: norm(b.text), numId: null });
    if (b.kind === 'list') for (const item of b.items || []) out.push({ text: norm(item.text), numId: item.numbering && item.numbering.present ? item.numbering.numId : null, level: item.numbering ? item.numbering.level : null, format: item.numbering ? item.numbering.format : null });
  }
  return out;
}

/**
 * L3, PRIMITIVE_INTEGRAL. Design doc point 3 (accepted): a numbering's definition is
 * SHARED -- every OTHER item on the SAME numId must show the SAME new value (this is the intended
 * effect, not a side effect); every item on a DIFFERENT numId must be unchanged; the document body
 * text otherwise unchanged (same `textUnits`/`norm` convention as list.cjs/page-setup.cjs).
 */
function verifyPreservation(op, before, after) {
  const text = keyTextOf(op);
  const b = numberingState(before && before.primary, text);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  if (op.scope === 'list') return verifyListPreservation(op, before, after, text, b.state);
  const numId = b.state.numId;
  const ua = numberingUnits(after.primary);
  const want = norm(b.state.text);
  const targetCount = ua.filter(u => u.text === want).length;
  // GH #248: under the 'refuse' policy the write is level-specific and no same-level co-member exists, so only a same-level item could show the new value; the same numbering's OTHER levels must keep theirs
  const refuse = op.sharedDefinition === 'refuse';
  const coMembersAfter = ua.filter(u => u.text !== want && u.numId === numId && (!refuse || u.level === b.state.level));
  const sharedMatch = coMembersAfter.map(u => fieldMatches(u.format, op.fields));
  const sharedOk = sharedMatch.every(m => Object.entries(m).every(([, v]) => v === true));
  // self-caught (own adversarial acceptance test, "L3 FAIL: ... co-member on the OLD value"):
  // comparing text#numId alone is BLIND to a collateral FORMAT change on another numId -- the
  // key must include the format itself, or a real divergence passes silently.
  const otherKey = u => `${u.text}#${u.numId}#${JSON.stringify(u.format)}`;
  const otherNumIdsBefore = numberingUnits(before.primary).filter(u => u.numId !== null && u.numId !== numId).map(otherKey).sort();
  const otherNumIdsAfter = ua.filter(u => u.numId !== null && u.numId !== numId).map(otherKey).sort();
  const otherNumIdsUnchanged = JSON.stringify(otherNumIdsBefore) === JSON.stringify(otherNumIdsAfter);
  const bodyBefore = textUnits(before.primary).map(u => norm(u.text)).sort();
  const bodyAfter = textUnits(after.primary).map(u => norm(u.text)).sort();
  const bodyUnchanged = JSON.stringify(bodyBefore) === JSON.stringify(bodyAfter);
  const levelKey = u => `${u.text}#${u.level}#${JSON.stringify(u.format)}`;
  const otherLevelsBefore = refuse ? numberingUnits(before.primary).filter(u => u.numId === numId && u.text !== want && u.level !== b.state.level).map(levelKey).sort() : [];
  const otherLevelsAfter = refuse ? ua.filter(u => u.numId === numId && u.text !== want && u.level !== b.state.level).map(levelKey).sort() : [];
  const otherLevelsUnchanged = JSON.stringify(otherLevelsBefore) === JSON.stringify(otherLevelsAfter);
  const ok = targetCount === 1 && sharedOk && otherNumIdsUnchanged && bodyUnchanged && otherLevelsUnchanged;
  return { ok, evidence: { targetText: b.state.text, targetCount, numId, coMemberCount: coMembersAfter.length, sharedOk, otherNumIdsUnchanged, bodyUnchanged, ...(refuse ? { otherLevelsUnchanged } : {}) } };
}

/**
 * GH #293 L7, L3 for scope 'list': the list is the editor's own member set (before), by text. Every item OUTSIDE it is unchanged (text,
 * numId and level format -- the other list on a shared definition included); inside it, the target level's items show the requested
 * fields and the other levels' items keep theirs; the body text is unchanged. (A detach gives the list a NEW numId, so a numId-keyed
 * comparison would read the list itself as "another numbering changed" -- the member set is the key here.)
 */
function verifyListPreservation(op, before, after, text, bState) {
  const mb = listMembersOf(before.primary, bState, text);
  if (!mb.ok) return { ok: false, evidence: { reason: `before-members-${mb.reason}` } };
  const inList = new Set(mb.members.map(m => m.text));
  const ub = numberingUnits(before.primary);
  const ua = numberingUnits(after.primary);
  const key = u => `${u.text}#${u.numId}#${u.level}#${JSON.stringify(u.format)}`;
  const outside = units => units.filter(u => !inList.has(u.text)).map(key).sort();
  const outsideUnchanged = JSON.stringify(outside(ub)) === JSON.stringify(outside(ua));
  const levelItems = ua.filter(u => inList.has(u.text) && u.level === bState.level);
  const levelOk = levelItems.every(u => Object.values(fieldMatches(u.format, op.fields)).every(v => v === true));
  const otherLevel = units => units.filter(u => inList.has(u.text) && u.level !== bState.level).map(u => `${u.text}#${u.level}#${JSON.stringify(u.format)}`).sort();
  const otherLevelsUnchanged = JSON.stringify(otherLevel(ub)) === JSON.stringify(otherLevel(ua));
  const listNumIds = new Set(ua.filter(u => inList.has(u.text) && u.numId !== null).map(u => u.numId));
  const bodyBefore = textUnits(before.primary).map(u => norm(u.text)).sort();
  const bodyAfter = textUnits(after.primary).map(u => norm(u.text)).sort();
  const bodyUnchanged = JSON.stringify(bodyBefore) === JSON.stringify(bodyAfter);
  const ok = outsideUnchanged && levelOk && otherLevelsUnchanged && listNumIds.size === 1 && bodyUnchanged;
  return { ok, evidence: { targetText: bState.text, scope: 'list', members: mb.members.length, levelItems: levelItems.length, outsideUnchanged, levelOk, otherLevelsUnchanged, oneNumbering: listNumIds.size === 1, bodyUnchanged } };
}

/** L2: the final state matches the task's request. */
function wholeTaskChecks(op, finalObservation) {
  const text = keyTextOf(op);
  const ns = numberingState(finalObservation.primary, text);
  const fmt = ns.ok && ns.state.node && ns.state.node.numbering ? ns.state.node.numbering.format : null;
  const matches = ns.ok ? fieldMatches(fmt, op.fields) : {};
  return [{ name: `${op.id}: numbering target holds the requested fields ${JSON.stringify(op.fields)}`, ok: ns.ok && Object.values(matches).every(v => v === true) }];
}

// refs: not a producer -- set_numbering never changes the target's own text.
function producedText() {
  return null;
}

module.exports = {
  CODES, SUFFIXES, CODE_TABLE, MEASURED_ALIGN, STEP_KEYS, SCOPES,
  plan, classify, resolve, preDispatchCheck, lower, commandBody,
  verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  numberingState, fieldMatches, keyTextOf,
};
