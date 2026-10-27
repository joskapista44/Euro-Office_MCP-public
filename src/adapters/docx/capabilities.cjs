'use strict';

/**
 * NEW_INTENTS -- post-M7 capabilities that are NOT part of the 28/28 source-migration inventory
 * (src/adapters/docx/intents.cjs). Owner decision after the W11 design review (option b): the 28/28
 * accounting gate (test/static/docx-intent-accounting.test.cjs:41 `assert.deepEqual(a.extra, [])`)
 * stays untouched -- a 29th row in the SAME `INTENTS` array would be read as an INVENTED intent
 * and go RED. A genuinely new capability (no source intent, no source file:line oracle to cite)
 * gets its OWN, DISJOINT list here instead, with its own closed gate
 * (test/static/docx-new-capabilities.test.cjs).
 *
 * Row shape mirrors intents.cjs's own `I(...)` rows (intent, family, migrationClass, batchSafe,
 * preservationMode, preservationEvidence, targetModule, status, source.evidence) -- but NOT the
 * same helper function: intents.cjs's `audited()` looks a row up in `PRESERVATION_AUDIT`, which
 * only holds the pinned 28 source-migrated intents; calling it for an unlisted intent throws
 * (intents.cjs:92-96). This file states `preservationMode`/`preservationEvidence` directly, per
 * row, citing the LIVE probe evidence + the design doc instead of a source file:line.
 *
 * migrationClass: intents.cjs's CLASSES enum (ADAPT / EXTRACT_SEMANTICS / REWRITE_FROM_CONTRACT /
 * EO_LIMIT) describes how a SOURCE code path was migrated -- none of the four fit a capability
 * that never had a source path at all, so this file does NOT reuse that enum. Its own single
 * class, NEW_CAPABILITY, is used instead; this is a naming choice, stated here rather than forcing
 * an existing migration-class label onto something that was never migrated.
 */

const CLASSES = Object.freeze(['NEW_CAPABILITY']);
const STATUS = Object.freeze({ NEW: 'M7_NEW_CAPABILITY' });

function N(intent, family, fields) {
  return Object.freeze({ intent, family, migrationClass: 'NEW_CAPABILITY', batchSafe: 'NONE', ...fields });
}

const NEW_INTENTS = Object.freeze([
  N('define_style', 'style-definition', {
    targetModule: 'src/adapters/docx/families/style-definition.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: "no source intent -- target-authored placement: docs/W10-STYLE-DEFINITION-DESIGN-20261007.md point 4 (L1 via the driver's own name-keyed, getter-sourced style read; every OTHER style and the document body unchanged, as part of success -- the same owner placement rule intents.cjs applies to section-global intents)",
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-w10-1a-style-definition-typeof-probe.live.cjs + docx-w10-1b-style-write-probe.live.cjs + docx-w10-1c-style-write-probe.live.cjs (CreateStyle/GetStyle method surface + write + two-reopen persistence, accepted) + docx-w10-2-style-extras-probe.live.cjs (accepted: Api.RGB over Api.CreateRGBColor for color, SetSpacingAfter(n,false), Heading 1 cascade NOT independently observable on the paragraph); design: docs/W10-STYLE-DEFINITION-DESIGN-20261007.md (v2, pending approval)',
    },
    note: "v1 ships text{bold,size,color} and paragraph{align,indentLeft,spacingAfter} on a NEW paragraph style; `text.color:'auto'` is SetColor(Api.AutoColor()) (GH #292 S8) and the flags take false (an explicit off, not a removal); `basedOn` (GH #292 S3) names an existing paragraph style, set with ApiStyle.SetBasedOn (define_style only). A name that already exists is ALWAYS refused (never a guessed SATISFIED on a content match) -- design doc point 5.",
  }),
  N('modify_style', 'style-definition', {
    targetModule: 'src/adapters/docx/families/style-definition.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- same placement as define_style above; L3 additionally requires every OTHER style unchanged (the edited style itself is expected to change, by definition)',
    source: {
      evidence: 'same LIVE probe evidence as define_style above',
    },
    note: 'v1 refuses a `name` that does not already exist (FAIL_CLOSED, named). The cascading effect on paragraphs ALREADY using the modified style is NOT independently verified (W10-2: no effective-size getter found live on the paragraph/run) -- verifyPreservation states only what IS measured (every other style + the body unchanged), never a guessed claim about rendered appearance.',
  }),
  N('set_page_setup', 'page-setup', {
    targetModule: 'src/adapters/docx/families/page-setup.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/W11-PAGE-SETUP-DESIGN-20261007.md point 3 (per-field SATISFIED; the body, every other field of the same section, and every other section unchanged, as part of success -- the same owner placement rule intents.cjs applies to the section-global edit_header/edit_footer intents)',
    source: {
      evidence: 'LIVE evidence, not a source oracle: section reflection + write probes (method surface; write + two-reopen persistence, 16/16 cross-axis attribution) and the acceptance harness test/live/docx-w11-page-setup-live-acceptance.live.cjs (docs/evidence/docx-w11-page-setup-live-acceptance-run2-PASS-20261007.json); design: docs/W11-PAGE-SETUP-DESIGN-20261007.md (approved, option b)',
    },
    note: 'v1 ships columns/startPageNumber on the write side; size (SetPageSize’s 3rd argument, likely orientation, never measured) and margins (the 4-argument order never discriminated -- the W11-1b evidence used an identical value at all four positions) are validated at plan time but their WRITE is refused with a named FAIL_CLOSED reason, pending two dedicated probes -- design doc points 2/3, flagged before this file was written.',
  }),
  N('format_table', 'table-format', {
    targetModule: 'src/adapters/docx/families/table-format.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/W04-TABLE-FORMAT-DESIGN-20261007.md (v3) point 3 (per-field SATISFIED; the body, every other table, and every other field of the SAME table unchanged, as part of success -- the same owner placement rule intents.cjs applies to the section-global edit_header/edit_footer intents, and set_page_setup above applies to sections)',
    source: {
      evidence: 'LIVE evidence, not a source oracle: the W04-1a reflection, W04-1b/1c write, W04-2 ToJSON and W04-4 table-format probes (the 5 remaining border edges + \'none\' style, the 3 remaining cellMargins edges, and the style NAME-to-id resolution), all accepted, and the acceptance harness test/live/docx-w04-table-format-live-acceptance.live.cjs; design: docs/W04-TABLE-FORMAT-DESIGN-20261007.md v3',
    },
    note: 'v1 ships all six fields (borders -- all 6 edges + \'none\' style, shading, cellMargins -- all 4 edges, indent, width, style) end to end -- every field in this intent is MEASURED, none validated-but-refused (unlike set_page_setup above). Cell-level axes (SetCellBorderTop, cell-level SetShd/SetCellMarginTop/SetWidth) are explicitly OUT of v1: no measured readback (W04-2).',
  }),
  N('set_numbering', 'numbering', {
    targetModule: 'src/adapters/docx/families/numbering.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/W07-NUMBERING-DESIGN-20261007.md point 4 (per-field SATISFIED, LIVE_VERIFY via ToJSON, never the saved file) + point 3 (every OTHER item sharing the same numId reflects the new value, as part of success -- the SHARED-definition semantics are the intended effect, not a side effect) + point 7 (every OTHER numId/list and the document body unchanged)',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-w07-1a-numbering-typeof-probe.live.cjs (method surface) + docx-w07-1b-numbering-write-probe.live.cjs + its W07-1c/1d fixes + docx-w07-3-numbering-fine-tune-probe.live.cjs; design: docs/W07-NUMBERING-DESIGN-20261007.md (v1 approved, v2/v3 incorporating the measured values)',
    },
    note: 'v1 ships code/custom (format+text+align, align measured to \'left\' only -- any other align refused at plan time), start, and suffix (tab/space/none, all three measured). SetRestart, any ilvl != 0, and a custom align other than \'left\' are NOT in v1 -- unmeasured, named in the design doc\'s own "Mi NEM tartozik bele". The `start` field reads UNOBSERVABLE (named, never a guessed match) when the live ToJSON omits the start key, which it does on a default-level numbering even though the saved file implicitly writes w:start=1 there -- this NEVER yields a false SATISFIED, and it does NOT refuse the step either (own offline acceptance test, docx-w07-numbering.test.cjs): the write still proceeds (idempotent if already correct), and verifyPrimitive fails closed only if start is STILL unobservable after the write.',
  }),
  N('set_image_size', 'image', {
    targetModule: 'src/adapters/docx/families/image.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/GH268-DOCX-IMAGES-READ-INVENTORY-WRITE-RESEARCH.md section 4 (per-image SATISFIED via the fresh read_images row; every OTHER image and the rest of the body unchanged, as part of success -- same owner placement rule as set_page_setup/format_table above)',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-images-probe.live.cjs R3_SIZE (SetSize on an inline image, saved, reopened, widthEmu/heightEmu both the new value, getters agree with the ToJSON extent); design: docs/GH268-DOCX-IMAGES-READ-INVENTORY-WRITE-RESEARCH.md section 7 point 2',
    },
    note: 'v1 ships widthEmu/heightEmu via SetSize, addressed by { paragraph, ordinal? } or { index } (NOT docPr.id/name -- editor-session-local, GH #268b). No aspect-ratio lock, no crop. Only the inline case is LIVE-measured (R3); the static source shows one placement-independent codepath, so a floating target is not refused, but is not independently LIVE-proven either -- named in the decision doc, not claimed as measured.',
  }),
  N('set_image_wrap', 'image', {
    targetModule: 'src/adapters/docx/families/image.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- same placement as set_image_size above',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-images-probe.live.cjs R4_WRAP (SetWrappingStyle(\'square\') on an inline image -> drawingType anchor, saved, reopened, placement floating / wrap square); design: docs/GH268-DOCX-IMAGES-READ-INVENTORY-WRITE-RESEARCH.md section 7 point 3',
    },
    note: 'v1 ships ONLY wrap:\'square\', ONLY from an inline source (the one transition R4 measured) -- refused at plan time AND re-checked inside the callCommand for every other value/source combination (tight/through/topAndBottom/behind/inFront, an already-anchored source, or reverting to inline), all NEM-MÉRT. No position-setting.',
  }),
  N('resolve_comment', 'comments', {
    targetModule: 'src/adapters/docx/families/comments.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/GH269c-DOCX-RESOLVE-COMMENT-DECISION.md (the body, the comment count, every OTHER comment and every field of the target comment except its solved state unchanged, as part of success -- the forgery surface of the author/time/text setters is checked, not assumed)',
    source: {
      evidence: 'LIVE evidence, not a source oracle: docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md 3./7. + the g269v2 LIVE ROUND R6 (ApiComment.SetSolved(true) then (false), PASS); the acceptance harness is test/live/docx-resolve-comment-product.live.cjs (NOT yet run)',
    },
    note: 'v1 identifies the comment by text + author + quote TOGETHER (exact; GetId stability is NOT MEASURED, never the key); zero or several matches are refused before dispatch. NOT in scope: a new comment or reply (the author identity is an owner decision), Delete, any author/time/text setter. UNRELEASED: not on any WRITE allowlist until a LIVE PASS.',
  }),
  N('reply_comment', 'comment-write', {
    targetModule: 'src/adapters/docx/families/comment-write.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/GH269d-DOCX-REPLY-ADD-COMMENT-DECISION.md (the body, every other comment and its replies, and every field of the target comment unchanged; the new text is verified to carry the SESSION author)',
    source: {
      evidence: 'LIVE evidence, not a source oracle: docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md 3./4. + the g269 probe R3/R4/R5 (explicit-author and no-author comment, AddReply); the acceptance harness is test/live/docx-reply-add-comment-product.live.cjs (NOT yet run)',
    },
    note: 'v1 replies to ONE existing comment identified by target {text, author, quote} (exact, exactly one match); the AUTHOR is the editor session identity (owner decision): AddReply is called with the text ONLY, the task has no author field. Retry: a reply with this text by the session author exists -> SATISFIED. UNRELEASED: not on any WRITE allowlist until a LIVE PASS.',
  }),
  N('add_comment_on_match', 'comment-write', {
    targetModule: 'src/adapters/docx/families/comment-write.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: docs/GH269d-DOCX-REPLY-ADD-COMMENT-DECISION.md (the body, every other comment and its replies, and every field of the target comment unchanged; the new text is verified to carry the SESSION author)',
    source: {
      evidence: 'LIVE evidence, not a source oracle: docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md 3./4. + the g269 probe R3/R4/R5 (explicit-author and no-author comment, AddReply); the acceptance harness is test/live/docx-reply-add-comment-product.live.cjs (NOT yet run)',
    },
    note: 'v1 comments ONE body text range found by matchText (+ optional 1-based occurrence; an ambiguous match without occurrence is refused before dispatch); the AUTHOR is the editor session identity (owner decision): AddComment is called with the text ONLY, the task has no author field. The retry key is text + quote + session author, NOT the occurrence (stated limit). UNRELEASED: not on any WRITE allowlist until a LIVE PASS.',
  }),
]);

module.exports = { NEW_INTENTS, CLASSES, STATUS };
