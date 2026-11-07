'use strict';

const { EXACT_PARAGRAPH_LOOKUP } = require('./paragraph-lookup.cjs');
const { norm, countExact, resolveTextUnit } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');

/**
 * DOCX `text` family: rewrite_paragraph (the M4 vertical slice).
 *
 * EXTRACT_SEMANTICS from the source oracle (joskapista44/euro-mcp f8c6d6cf, read-only):
 *   classify   docx-agent-idempotency.cjs:38-72   W9.2-R restart-safe identity
 *   lower      coedit-docx-target-text.cjs:32      exact-range rewrite, measured runtime facts
 *   verify     docx-session-basic-edit.cjs:62-72   old gone, new once
 *   preserve   docx-session-basic-edit.cjs:20      seven run properties, JSON-equal
 *
 * Contract placement (the difference from the source, stated):
 *   - classification is the M3 per-operation `classifyOperation`, from the FRESH observation only;
 *     receipt, history, process memory and stale fingerprints are not inputs (they are not passed).
 *   - property preservation is PRIMITIVE_INTEGRAL (owner review R1 after the M4 first checkpoint):
 *     part of the rewrite's semantic success, exactly as in the source, where a property mismatch
 *     made the primitive fail (docx-session-basic-edit.cjs:71-72). The executor runs it before the
 *     register decision; L1 PASS + properties not preserved -> DISPATCHED_OUTCOME_UNKNOWN, never
 *     APPLIED. The check itself is the separate verifyPreservation hook (L3), not folded into L1.
 *   - whole-task verification is STRONGER than the source (which checked "new text present once"
 *     only): the old text must be absent too, and a task that carried structural anchors must see
 *     the terminal text in that same context.
 */

const PRESERVED_PROPERTIES = Object.freeze(['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language']);
const ANCHOR_FIELDS = Object.freeze(['kind', 'before', 'after', 'headingPath']);
const REWRITABLE_KINDS = Object.freeze(['paragraph', 'heading']);
// Set by the adapter planner from an EXPLICIT task.recovery (adapter.cjs); never from stored state.
const UNRESOLVED = 'UNRESOLVED_UNCERTAINTY';

/**
 * GH #359 (1): `commandBody`'s named returns, for every guard that fires BEFORE the command's first
 * mutating engine call. No edit was ever ATTEMPTED on this path, so a driver reading one of these
 * back may safely treat it as "no write" (docx-driver.cjs's dispatch()).
 *
 * `rewrite-live-addtext-failed` (review feedback, point 2) is deliberately NOT in this list, even
 * though it fires before any OTHER mutating call succeeds: `AddText` IS the mutating call, and a
 * `!== true` return does not prove the engine's own change log stayed empty (a partial insert is
 * not ruled out). The same reasoning already excludes the post-AddText rollback branches
 * (text.cjs 290-301) -- `AddText` failing keeps the unconditional "write attempted" default, it
 * just never gets a name added to the no-write set.
 */
const REWRITE_NO_WRITE_REASONS = Object.freeze([
  'rewrite-live-no-document-or-empty-target',
  'rewrite-live-paragraph-not-found',
  'rewrite-live-range-text-mismatch',
  'rewrite-live-hyperlink-refused',
  'rewrite-live-not-paragraph',
  'rewrite-live-paragraph-text-not-unique',
]);

/**
 * `step.expectedPreservation` (OPTIONAL, owner decision R2 point 6): an explicit expected property
 * contract from the caller -- a subset of the seven properties. Never required; when present it is
 * the only thing that can prove a terminal state's properties once the original baseline is gone.
 */
function validExpectation(value) {
  if (value === undefined) return { ok: true };
  if (!value || typeof value !== 'object' || Array.isArray(value)) return { ok: false, detail: 'expectedPreservation must be an object' };
  const keys = Object.keys(value);
  if (keys.length === 0) return { ok: false, detail: 'expectedPreservation is empty' };
  const unknown = keys.filter(k => !PRESERVED_PROPERTIES.includes(k));
  if (unknown.length) return { ok: false, detail: `expectedPreservation has unknown properties: ${unknown.join(', ')}` };
  return { ok: true };
}

function plan(step, index) {
  if (typeof step.text !== 'string') return { ok: false, detail: `step ${index}: rewrite_paragraph requires text` };
  if (!step.selector || typeof step.selector !== 'object') return { ok: false, detail: `step ${index}: rewrite_paragraph requires a selector` };
  const expectation = validExpectation(step.expectedPreservation);
  if (!expectation.ok) return { ok: false, detail: `step ${index}: ${expectation.detail}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'rewrite_paragraph',
      selector: step.selector,
      text: step.text,
      preservationMode: 'PRIMITIVE_INTEGRAL',
      ...(step.expectedPreservation ? { expectedPreservation: Object.freeze({ ...step.expectedPreservation }) } : {}),
    },
  };
}

function anchorsOf(selector) {
  return ANCHOR_FIELDS.filter(k => {
    const v = selector && selector[k];
    return v !== undefined && v !== null && !(Array.isArray(v) && v.length === 0) && v !== '';
  });
}

/**
 * The derived preservation expectation: the seven properties of the FRESH pre-mutation target.
 *
 * GH #359 (3/3), measured LIVE (three run-count shapes -- MULTIRUN, MULTIRUN_PB, MULTIRUN3):
 * a property the source paragraph's runs disagree on is ABSENT
 * from target.props (the engine's own multi-run TextPr read returns undefined for a mixed value,
 * the same measured behaviour comparable()'s caller already documents for the whole-paragraph
 * formatting family), and on all three measured shapes the rewrite's new, single run
 * DETERMINISTICALLY takes on the FIRST source run's own value for that property. So a property
 * missing here, when the source paragraph's runs are readable and there is more than one of them,
 * falls back to the first run's own format -- every other property (present in target.props, or
 * missing for any OTHER reason: runs unreadable, a single run, or the first run's own value for
 * that property also unset) is left UNCHANGED, so verifyPreservation's fail-closed comparison
 * below still fires exactly as before this change.
 *
 * GH #359 kovetes, MERT LIVE (prod DS cell /26, egy vegyes betumeretu bekezdesen): a fenti fallback
 * 11-et (run-format, PONT) adott a tarolt/tarsitott target.props 22 (FELPONT) helyett -- a ket mezo
 * KULONBOZO egysegben jon (docx-driver.cjs: a paragraph-szintu props.fontSize a NYERS GetFontSize()
 * felpontban, a futasnyi (run) format.fontSize ugyanazt az erteket mar PONTRA osztva adja, / 2-vel).
 * A fontSize fallback ezert a run-format ertekét *2-vel konvertalja a felpontos target.props-konvenciora,
 * MIELOTT beirna az expected-be -- csak ezt az EGY propertyt erinti, a tobbi hat valtozatlan marad.
 */
function derivedExpectation(target) {
  if (!target || !target.props) return null;
  // Built in PRESERVED_PROPERTIES order directly (not comparable() + a later insertion) so a
  // collapsed-from-the-first-run key sits where it would if target.props had carried it all
  // along -- verifyPreservation's JSON.stringify comparison is key-ORDER sensitive, and
  // `actual` (comparable(now.props) below) is always built in this same canonical order.
  const first = Array.isArray(target.runs) && target.runs.length > 1 && target.runs[0] ? target.runs[0].format : null;
  const expected = {};
  for (const k of PRESERVED_PROPERTIES) {
    if (target.props[k] !== undefined) expected[k] = target.props[k];
    else if (first && first[k] !== undefined) expected[k] = k === 'fontSize' ? first[k] * 2 : first[k];
  }
  return expected;
}

function propertiesMatch(actual, expected) {
  return Object.keys(expected).every(k => JSON.stringify(actual[k]) === JSON.stringify(expected[k]));
}

/**
 * A terminal text is present. May that be called SATISFIED? (owner decision R2)
 *
 *   explicit expectedPreservation   -> SATISFIED only if the terminal run properties match it;
 *                                      unreadable or different -> FAIL_CLOSED
 *   caller says this operation ended in unresolved uncertainty (task.recovery), no explicit
 *   expectation                     -> FAIL_CLOSED preservation-baseline-unavailable: the original
 *                                      baseline is gone, so the P1 postcondition cannot be proven
 *   neither                         -> SATISFIED by the terminal contract (the normal idempotent
 *                                      retry). A fresh invocation that carries no recovery context
 *                                      is indistinguishable from a new task; nothing in it says a
 *                                      previous run was uncertain (information-theoretic limit).
 */
function terminalVerdict(target, op, proof, extra = {}) {
  if (op.expectedPreservation) {
    const actual = derivedExpectation(target);
    if (!actual) return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-terminal-properties-unreadable', ...extra } };
    if (!propertiesMatch(actual, op.expectedPreservation)) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-terminal-preservation-mismatch', expected: op.expectedPreservation, actual, ...extra } };
    }
    return { state: 'SATISFIED', evidence: { text: op.text, ...extra, proof, preservation: 'terminal-properties-match-explicit-expectation' } };
  }
  if (op.recovery === UNRESOLVED) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable', ...extra } };
  }
  return { state: 'SATISFIED', evidence: { text: op.text, ...extra, proof } };
}

/**
 * W9.2-R, re-expressed, plus the P1 retry rule (owner decision R2). A rewrite is TRANSFORMATIVE:
 * once it succeeds, the text identity the task selected by is gone -- and with it the property
 * baseline. So:
 *   original target resolves  -> already terminal? terminalVerdict; else the new text must not
 *                                already exist elsewhere (two contradicting states -> FAIL_CLOSED);
 *                                APPLY with the DERIVED expectation from this fresh state (an
 *                                explicit expectation that contradicts it -> FAIL_CLOSED)
 *   original target ambiguous -> FAIL_CLOSED (the starting point itself is uncertain)
 *   original target absent    -> a terminal state is locatable ONLY with the task's own structural
 *                                anchors (kind / before / after / headingPath), then terminalVerdict;
 *                                no anchors -> FAIL_CLOSED ("the expected text occurs once" is not enough)
 */
function classify(primary, op) {
  const resolved = resolveStableTarget(primary, op.selector);
  if (resolved.ok) {
    if (norm(resolved.target.text) === norm(op.text)) {
      return terminalVerdict(resolved.target, op, 'fresh-live-selector-already-terminal');
    }
    // the rewrite pattern deletes a hyperlink in its target (measured LIVE, paragraph-lookup-hyperlink-shift-probe W2): refused
    // here, before any dispatch -- the command's own holdsHyperlink guard would only end in DISPATCHED_OUTCOME_UNKNOWN
    if (Array.isArray(resolved.target.hyperlinks) && resolved.target.hyperlinks.length > 0) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-target-holds-hyperlink', hyperlinks: resolved.target.hyperlinks.length } };
    }
    if (resolveStableTarget(primary, { text: op.text }).ambiguity !== 'NOT_FOUND') {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-conflicting-old-and-new-state' } };
    }
    const expectedPreservation = derivedExpectation(resolved.target);
    if (op.expectedPreservation && expectedPreservation && !propertiesMatch(expectedPreservation, op.expectedPreservation)) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-explicit-expectation-contradicts-baseline', expected: op.expectedPreservation, baseline: expectedPreservation } };
    }
    return { state: 'APPLY', evidence: { from: norm(resolved.target.text), expectedPreservation } };
  }
  if (resolved.ambiguity === 'AMBIGUOUS') {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-ambiguous-old-target', candidates: resolved.candidates } };
  }
  // M5 sectionHeading + paragraphOrdinal: the selector names a position in a section, not a text, so
  // an unresolved one has no anchor-based terminal fallback -- its own reason, before any dispatch
  if (/^(SECTION|PARAGRAPH)_/.test(String(resolved.ambiguity))) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `rewrite-${String(resolved.ambiguity).toLowerCase().replace(/_/g, '-')}`, available: resolved.available, candidates: resolved.candidates } };
  }
  // GH #307: `emptyDocument` no longer resolves once the write has landed -- the retry proof is NOT
  // anchor-based (there is no text to search by): the document is terminal when it is, again, exactly
  // one paragraph whose text is the task's OWN text. Anything else (someone else's text, more than one
  // paragraph) is a NAMED failure, never the generic "insufficient selector" (the selector was fine).
  if (resolved.ambiguity === 'EMPTY_DOCUMENT_NOT_EMPTY') {
    const blocks = (primary && primary.blocks) || [];
    if (blocks.length === 1 && blocks[0].kind === 'paragraph' && norm(blocks[0].text) === norm(op.text)) {
      return terminalVerdict(blocks[0], op, 'empty-document-retry-already-terminal');
    }
    return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-empty-document-retry-target-not-terminal' } };
  }
  const anchors = anchorsOf(op.selector);
  if (anchors.length === 0) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'rewrite-terminal-identity-insufficient-selector' } };
  }
  const terminal = resolveStableTarget(primary, terminalSelector(op, anchors));
  if (!terminal.ok && hasSuccessor(op)) {
    // W2.1: a declared later list step may have turned the rewritten paragraph into a list item
    const t = terminalUnit(primary, op, anchors);
    if (t.ok) return terminalVerdict(t.node, op, 'fresh-live-terminal-context-proven-as-text-unit', { anchors, successors: op.successorTransforms });
  }
  if (!terminal.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `rewrite-terminal-target-${String(terminal.ambiguity).toLowerCase()}`, anchors } };
  }
  return terminalVerdict(terminal.target, op, 'fresh-live-terminal-context-proven', { anchors });
}

function terminalSelector(op, anchors = anchorsOf(op.selector)) {
  const s = { text: op.text };
  for (const k of anchors) s[k] = op.selector[k];
  // GH #359 (v2): beforeAlt/afterAlt (adapter.cjs's attachAnchorAlternates) are planner metadata on
  // op.selector, not one of ANCHOR_FIELDS -- carry them into the terminal selector too, or
  // resolveStableTarget's matchesAnchor never sees them and a neighbour rewritten by another
  // operation of this task fails this check even though classify()'s own top-level resolution
  // (which reads op.selector directly) already accepts it.
  if (op.selector && op.selector.beforeAlt !== undefined) s.beforeAlt = op.selector.beforeAlt;
  if (op.selector && op.selector.afterAlt !== undefined) s.afterAlt = op.selector.afterAlt;
  return s;
}

function resolve(primary, selector) {
  const r = resolveStableTarget(primary, selector);
  if (r.ok && !REWRITABLE_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

/**
 * The editor command. The body re-expresses the source's measured runtime behaviour (the facts are
 * recorded in the source header, coedit-docx-target-text.cjs:1-29): exact Search range, unique;
 * GetParagraph(0); the exact target must occur once in its paragraph; ApiRange has no SetText, so
 * the replacement is inserted BEFORE the intact target, re-formatted with the old run properties,
 * and the OLD text is deleted through a FRESH exact range (the original range is stale after
 * AddText). Any failed step rolls back what it inserted. An empty replacement deletes the target.
 * The old text is verified at its KNOWN post-insert offset (targetOffset + replacement.length), never searched for: a replacement that CONTAINS the old text (GH #228) is a normal case.
 *
 * `meta` is a structured description of the same intent, for test transports that cannot run the
 * editor API. The real driver executes `body`; nothing else.
 */
function lower(op, target) {
  // GH #307: `emptyDocument` has no OLD TEXT to search by -- EXACT_PARAGRAPH_LOOKUP's whole premise
  // (find the paragraph that contains this exact string) does not apply to an empty string, so this
  // is a DIFFERENT command, not a degenerate case of the text-anchored one (measured live, GH #307
  // probe: commandBody('', newText) runs the text-anchored command and does nothing at all, via its
  // own `if (!doc || !targetText) return;`).
  if (op.selector && op.selector.emptyDocument === true) {
    const newText = String(op.text == null ? '' : op.text);
    return {
      kind: 'callCommand',
      body: emptyDocumentCommandBody(newText),
      meta: Object.freeze({ family: 'text', action: 'rewriteEmptyDocument', newText }),
    };
  }
  const oldText = String(target.text || '');
  const newText = String(op.text == null ? '' : op.text);
  return {
    kind: 'callCommand',
    body: commandBody(oldText, newText),
    meta: Object.freeze({ family: 'text', action: 'rewrite', oldText, newText }),
  };
}

function commandBody(targetText, replacement) {
  return `
try {
  const targetText = ${JSON.stringify(targetText)};
  const replacement = ${JSON.stringify(replacement)};
  const textOpts = { Numbering: false, NewLineSeparator: '\\n' };
  const doc = Api.GetDocument();
  if (!doc || !targetText) return 'rewrite-live-no-document-or-empty-target';
${EXACT_PARAGRAPH_LOOKUP}
  const paragraph = exactParagraphOf(targetText);
  if (!paragraph || typeof paragraph.GetRange !== 'function') return 'rewrite-live-paragraph-not-found';
  const range = unitRangeOf(targetText);
  if (!range || String(range.GetText(textOpts) || '') !== targetText) return 'rewrite-live-range-text-mismatch';
  // a rewrite over a hyperlink's text deletes the link (measured LIVE, docs/evidence/paragraph-lookup-hyperlink-shift-probe-20261030.json W2):
  // not decided, so a paragraph holding one is not rewritten -- refused before any insert, as before the lookup mapped link positions
  if (holdsHyperlink(paragraph)) return 'rewrite-live-hyperlink-refused';
  if (typeof paragraph.GetClassType === 'function' && paragraph.GetClassType() !== 'paragraph') return 'rewrite-live-not-paragraph';
  const paragraphText = String(paragraph.GetText(textOpts) || '');
  const targetOffset = paragraphText.indexOf(targetText);
  if (targetOffset < 0 || paragraphText.indexOf(targetText, targetOffset + targetText.length) !== -1) return 'rewrite-live-paragraph-text-not-unique';
  // GH #359 (2/2): plainRunPadBudget comes from EXACT_PARAGRAPH_LOOKUP above -- the same run-boundary
  // widening unitRangeOf uses at the top of this body. A single-run paragraph keeps the original
  // [-1, 0, +1] window unchanged; a multi-run one widens past +1 as the measured run count demands.
  const exactRange = (p, start, text) => {
    const baseEnd = start + text.length;
    const maxPad = plainRunPadBudget(p);
    for (let pad = -1; pad <= maxPad + 1; pad += 1) {
      const end = baseEnd + pad;
      if (end < start) continue;
      try { const c = p.GetRange(start, end); if (c && String(c.GetText(textOpts) || '') === text) return c; } catch (_) {}
    }
    return null;
  };
  if (!replacement) { range.Delete(); return 'ok'; }
  const oldTextPr = range.GetTextPr();
  if (range.AddText(replacement, 'before') !== true) return 'rewrite-live-addtext-failed';
  const rollback = () => { try { const r = exactRange(paragraph, targetOffset, replacement); if (r) r.Delete(); } catch (_) {} };
  const inserted = exactRange(paragraph, targetOffset, replacement);
  if (!inserted) { rollback(); return 'rewrite-live-rollback-insert-not-found'; }
  if (oldTextPr) {
    if (typeof inserted.SetTextPr !== 'function' || inserted.SetTextPr(oldTextPr) == null) { rollback(); return 'rewrite-live-rollback-settextpr-failed'; }
  }
  if (String(inserted.GetText(textOpts) || '') !== replacement) { rollback(); return 'rewrite-live-rollback-inserted-text-mismatch'; }
  const afterInsert = String(paragraph.GetText(textOpts) || '');
  // GH #228: the old text stands at a KNOWN offset after the insert -- check it THERE (a search would also find the old text inside a replacement that contains it)
  const oldOffset = targetOffset + replacement.length;
  if (afterInsert.slice(targetOffset, oldOffset) !== replacement || afterInsert.slice(oldOffset, oldOffset + targetText.length) !== targetText) { rollback(); return 'rewrite-live-rollback-after-insert-mismatch'; }
  const freshOld = exactRange(paragraph, oldOffset, targetText);
  if (!freshOld || String(freshOld.GetText(textOpts) || '') !== targetText) { rollback(); return 'rewrite-live-rollback-old-text-not-found'; }
  if (freshOld.Delete() !== true) { rollback(); return 'rewrite-live-rollback-delete-failed'; }
  return 'ok';
} catch (_) { return 'rewrite-live-exception'; }
`;
}

/**
 * GH #307: the `emptyDocument` command. The precondition (exactly one element, a paragraph, trimmed
 * text empty) is checked INSIDE the command too -- defense in depth, the same discipline
 * REWRITE_SELECTOR's rewriteSelectorBranch uses at the schema layer: the resolver proved it against
 * a FRESH pre-dispatch observation, but the live document may have moved by the time this body
 * actually runs, so the command refuses to write rather than trust a stale proof. A blank
 * replacement is refused too -- there is no "old text" to delete here, so an empty replacement is
 * not a meaningful operation (unlike `commandBody`, where it deletes the target).
 */
function emptyDocumentCommandBody(replacement) {
  return `
try {
  const replacement = ${JSON.stringify(replacement)};
  if (!replacement) return;
  const doc = Api.GetDocument();
  if (!doc || doc.GetElementsCount() !== 1) return;
  const paragraph = doc.GetElement(0);
  if (!paragraph || typeof paragraph.GetClassType !== 'function' || paragraph.GetClassType() !== 'paragraph') return;
  const textOpts = { Numbering: false };
  const text = String(typeof paragraph.GetText === 'function' ? paragraph.GetText(textOpts) : '');
  if (text.trim() !== '') return;
  if (typeof paragraph.AddText !== 'function') return;
  paragraph.AddText(replacement);
} catch (_) {}
`;
}

/**
 * Pre-dispatch refusal: without a readable property baseline the preservation requirement cannot
 * be proven afterwards, so the write is refused BEFORE dispatch (source BE:47, same rule).
 *
 * GH #307: `emptyDocument` is EXEMPT -- its target carries no run by construction (an empty
 * paragraph has nothing to read a baseline FROM), so an unreadable baseline here is the EXPECTED
 * shape, not a suspicious one. Requiring `target.props` for this branch would refuse every write
 * this selector could ever make.
 */
function preDispatchCheck(op, target) {
  if (op.selector && op.selector.emptyDocument === true) return { ok: true };
  if (!target.props) return { ok: false, dispatched: false, detail: 'rewrite-property-baseline-unreadable' };
  return { ok: true };
}

/** L1: the fresh LIVE terminal state. Dispatch success is never enough. */
function verifyPrimitive(op, before, after) {
  const target = before ? resolveStableTarget(before.primary, op.selector).target : null;
  const oldText = target ? target.text : null;
  const oldCount = oldText === null ? null : countExact(after.primary, oldText);
  const newCount = op.text ? countExact(after.primary, op.text) : 0;
  const ok = oldText !== null && oldCount === 0 && (op.text ? newCount === 1 : true);
  return { ok, evidence: { text: op.text, oldText, oldCount, newCount, proof: 'fresh-live-terminal-state' } };
}

function comparable(props) {
  const out = {};
  for (const k of PRESERVED_PROPERTIES) if (props && props[k] !== undefined) out[k] = props[k];
  return out;
}

/**
 * L3, PRIMITIVE_INTEGRAL: the rewritten run carries the DERIVED expectation -- the seven properties
 * of the target in the fresh pre-mutation observation (`before`), compared as the source does.
 * Operation-scoped, never persisted: not a mutation target, not a receipt, not a persistence authority.
 */
function verifyPreservation(op, before, after) {
  // GH #307: an empty target has no run property to preserve -- what the write must preserve instead is
  // the document's SHAPE: before it, exactly one empty paragraph (the resolver's own precondition);
  // after it, still exactly one block, a paragraph, holding the task's own text. Anything else (an
  // extra paragraph, a different kind, other text) fails -- not a vacuous pass.
  if (op.selector && op.selector.emptyDocument === true) {
    const blocks = (after && after.primary && after.primary.blocks) || [];
    const ok = blocks.length === 1 && blocks[0].kind === 'paragraph' && norm(blocks[0].text) === norm(op.text);
    return { ok, evidence: { reason: ok ? 'empty-document: one paragraph, the task text' : 'empty-document: the document is not exactly one paragraph holding the task text', blocks: blocks.length } };
  }
  const target = before ? resolveStableTarget(before.primary, op.selector).target : null;
  const expected = derivedExpectation(target);
  if (!expected) return { ok: false, evidence: { reason: 'no derived preservation expectation (before-properties unreadable)' } };
  if (!op.text) return { ok: true, evidence: { reason: 'target deleted; no run left to compare' } };
  const now = resolveStableTarget(after.primary, { text: op.text }).target;
  if (!now || !now.props) return { ok: false, evidence: { reason: 'after-properties unreadable' } };
  const actual = comparable(now.props);
  // GH #359 (3/3): name which properties derivedExpectation filled in from the first source run
  // (target.props itself omitted them) -- so a caller can see that the SECOND run's own
  // formatting did not survive the rewrite, without re-deriving it from target.props/target.runs.
  const formattingCollapsedToFirstRun = target && target.props
    ? PRESERVED_PROPERTIES.filter(k => target.props[k] === undefined && expected[k] !== undefined)
    : [];
  return {
    ok: JSON.stringify(expected) === JSON.stringify(actual),
    evidence: {
      expected, after: actual, expectationSource: 'derived-from-fresh-pre-mutation-observation', properties: PRESERVED_PROPERTIES,
      ...(formattingCollapsedToFirstRun.length ? { formattingCollapsedToFirstRun } : {}),
    },
  };
}

/** L2 contribution for one rewrite operation, read-only against the final observation. */
/**
 * W2.1 successor-aware L2: a LATER operation of this task declaredly turned the rewritten paragraph
 * into a list item (planner metadata `successorTransforms`). The rewrite's final postconditions are
 * then checked on the representation-free TEXT UNIT: the new text exactly once as a unit, its
 * structural anchors kept -- `kind: paragraph` accepts a list item (the declared transform), a
 * heading requirement is NOT relaxed; before/after are the neighbouring units; the heading path is
 * the containing block's.
 */
const hasSuccessor = op => Array.isArray(op.successorTransforms) && op.successorTransforms.length > 0;

/** The rewrite's terminal text unit under a declared successor, with its anchors checked. */
function terminalUnit(primary, op, anchors) {
  const u = resolveTextUnit(primary, op.text);
  if (!u.ok) return { ok: false, present: false, reason: u.reason };
  const sel = op.selector;
  const kinds = sel.kind === 'paragraph' ? ['paragraph', 'list_item'] : [sel.kind];
  const inContext = (sel.kind === undefined || kinds.includes(u.unit.kind))
    && (!sel.before || u.unit.prev.includes(norm(sel.before)))
    && (!sel.after || u.unit.next.includes(norm(sel.after)))
    && (!sel.headingPath || samePath(u.unit.block.locator && u.unit.block.locator.headingPath, sel.headingPath));
  return { ok: inContext || anchors.length === 0, present: true, unit: u.unit, node: u.unit.node };
}

function successorAwareChecks(op, primary, anchors) {
  const t = terminalUnit(primary, op, anchors);
  const checks = [{ name: `${op.id}: final text present once (as a text unit; declared successor ${op.successorTransforms.join(', ')})`, ok: t.present }];
  if (anchors.length) checks.push({ name: `${op.id}: final text in its structural context (${anchors.join(', ')}; representation-free)`, ok: t.ok });
  return checks;
}

const samePath = (a, b) => Array.isArray(a) && Array.isArray(b) && a.length === b.length && a.every((x, i) => x.level === b[i].level && norm(x.text) === norm(b[i].text));

function wholeTaskChecks(op, finalObservation) {
  const primary = finalObservation.primary;
  const checks = [];
  if (op.text && hasSuccessor(op)) {
    checks.push(...successorAwareChecks(op, primary, anchorsOf(op.selector)));
    const original = resolveStableTarget(primary, op.selector);
    checks.push({ name: `${op.id}: original target text absent`, ok: !original.ok || norm(original.target.text) === norm(op.text) });
    return checks;
  }
  if (op.text) {
    checks.push({ name: `${op.id}: final text present once`, ok: countExact(primary, op.text) === 1 });
  }
  const anchors = anchorsOf(op.selector);
  if (anchors.length && op.text) {
    const terminal = resolveStableTarget(primary, terminalSelector(op, anchors));
    checks.push({ name: `${op.id}: final text in its structural context (${anchors.join(', ')})`, ok: terminal.ok });
  }
  // The original target must not survive. Resolved by the task's own selector against the FINAL
  // state: if it still resolves to a block whose text is not the terminal text, the rewrite is not done.
  const original = resolveStableTarget(primary, op.selector);
  checks.push({
    name: `${op.id}: original target text absent`,
    ok: !original.ok || norm(original.target.text) === norm(op.text),
  });
  return checks;
}

/** The semantic text a rewrite leaves behind (for a later operation's ref). */
function producedText(op) {
  return typeof op.text === 'string' ? op.text : null;
}

module.exports = {
  plan, classify, resolve, lower, commandBody, emptyDocumentCommandBody, preDispatchCheck, verifyPrimitive, verifyPreservation,
  wholeTaskChecks, anchorsOf, derivedExpectation, PRESERVED_PROPERTIES, UNRESOLVED, producedText, REWRITE_NO_WRITE_REASONS,
};
