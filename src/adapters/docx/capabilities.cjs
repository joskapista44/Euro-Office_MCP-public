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
  N('insert_table', 'table-insert', {
    targetModule: 'src/adapters/docx/families/table-insert.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: every OTHER body block stays unchanged in sequence and content; the ONLY tolerated extra, and only when the anchor was the LAST body element before the write, is exactly one empty paragraph immediately after the new table (the measured engine behaviour: the body cannot end on a table) -- any other shape is L3 FAIL',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh308-create-table-probe.live.cjs, run gh308-tableprobe-1793158850683-882e64: doc.AddElement(anchorIndex + 1, table) lands exactly after the anchor; doc.Push(table) (the known-working call elsewhere) lands at the document end and, when the anchor was the last body element, leaves a trailing empty paragraph after it; Api.CreateTable(cols, rows) returns null (no throw) on a 0/negative argument; the fresh reopen and the product\'s own read_tables both see the written table in place',
    },
    note: 'v1 ships selector { afterParagraph } (a unique body-level paragraph/heading; a table/list, or a text that only occurs inside a table cell, is refused the same way every other family\'s resolveStableTarget call already refuses it), rows/cols (1..50, a schema sanity cap -- not a measured limit) and an optional cells row-by-cols text matrix (no tab or line break; every cell blank when omitted, matching Api.CreateTable\'s own default). No table-style/border write (table-format.cjs already covers that AFTER a table exists), no image-in-cell (#306). Joined DOCX_WRITE_ALLOWLIST provisionally -- the RELEASE-ACCEPTANCE harness (test/live/docx-gh308-insert-table-product.live.cjs) has not yet been run by its author.',
  }),
  N('insert_section_break', 'section-break-insert', {
    targetModule: 'src/adapters/docx/families/section-break-insert.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: a break adds no body element, so every body block is unchanged; the section the anchor closes keeps every OTHER field it already had; every section before it is untouched at its own index; every section after it shifts by exactly one (the split inserts one new entry) and is otherwise unchanged -- any other shape is L3 FAIL',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh313-section-break-probe.live.cjs, run gh313-sectionprobe-1793165686897-6de018: the anchor paragraph has no InsertSection method (SURFACE: GetSection/SetSection only); doc.CreateSection(anchor) creates the break and returns the section the anchor now CLOSES (not the one starting after it); OOXML\'s <w:type> describes how the section carrying it BEGINS, so the requested type belongs on the NEXT section, found by re-locating the paragraph immediately after the anchor and calling its own GetSection() directly (never doc.GetSections() indexing -- a section object is a new wrapper per call, per this driver\'s own existing comment); the fresh reopen and the product\'s own read_page_setup both see the new section count and its distinct columnsCount values',
    },
    note: 'v1 ships selector { afterParagraph } (same anchor contract as insert_table) and type: nextPage | continuous (required). Columns are NOT part of this intent -- set_page_setup already addresses a section by {index, of}; once the break exists, it sets the new section\'s columns (the card\'s own point 4). Section-to-body-position correlation (classify/verify) is read off a NEW, self-validating inventory field (docx-driver.cjs: closingBlockIndex, UNOBSERVABLE whenever its own section count cannot be cross-checked against doc.GetSections().length) -- v1 refuses rather than guesses whenever that read is unobservable. Joined DOCX_WRITE_ALLOWLIST provisionally, same as insert_table -- the RELEASE-ACCEPTANCE harness (test/live/docx-gh313-insert-section-break-product.live.cjs) has not yet been run by its author.',
  }),
  N('insert_image', 'image-insert', {
    targetModule: 'src/adapters/docx/families/image-insert.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: every OTHER image is preserved as a multiset (placement/wrap/size/paragraph key) and exactly one new one appears; every OTHER body block is unchanged in sequence and content -- position \'after\' additionally requires the ONE new block to be exactly the empty paragraph this write creates, immediately after the anchor (the SAME discipline insert_table already uses for its own new block, canonicalBlock REUSED from table-insert.cjs unchanged)',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh306-insert-image-probe.live.cjs, commit 749a4f607 (two fixes made on review: route B anchor lookup by TEXT, not identity -- a new-per-call wrapper never matches by ===, the third measured instance of this after #308/#313; media sha256 resolved through each stored row\'s OWN embed id): both routes LIVE PASS -- Api.CreateImage + paragraph.AddDrawing (position \'end\') and Api.CreateImage + a new paragraph + doc.AddElement(anchorIndex+1, ...) (position \'after\', the SAME mechanism #308 measured for a table) both land a PNG and a JPEG data: URI, the saved word/media/* bytes sha256-match the original input, a requested EMU size round-trips within ±1 (the SAME open drift the XLSX insert_image precedent already accepts), and the product\'s own ALREADY-ACCEPTED read_images sees the new image at the right size',
    },
    note: 'v1 ships `data:image/png;base64,...` OR `data:image/jpeg;base64,...` source (1 MiB cap, the SAME margin the XLSX insert_image schema already uses -- 1,400,000 base64 chars), widthEmu/heightEmu (positive integer, the set_image_size convention). selector is EITHER { afterParagraph } with position: \'after\' (default, a new paragraph) | \'end\' (appended to the anchor\'s own paragraph) OR, since GH #306b 2. resz (LIVE-measured, docx-gh306b-header-image-probe/-docpr-id-probe.live.cjs), { section: {index,of}, part: \'header\'|\'footer\', variant: \'default\' } -- appends to the END of an EXISTING header/footer part\'s own last paragraph (the part must already exist; no position field, no creation, variant restricted to \'default\' -- title/even unmeasured for image insertion). RETRY IDENTITY for position \'after\' / the header-footer shape is position-based, not text-based (the new paragraph / the part\'s existing last paragraph carries no text OF THE IMAGE) -- a NAMED, ACCEPTED limit (the family\'s own header comment), the same class image.cjs already accepts for SOURCE identity. No alt-text/title (no public setter, GH #268 section 3), no table-cell/floating target, no #317 two-column layout (a later round), no NC-file source (data: URI only, a separate jogosultsagi-ut decision).',
  }),
  N('insert_paragraphs', 'paragraphs-insert', {
    targetModule: 'src/adapters/docx/families/paragraphs-insert.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: the N new blocks (text + requested style/format, in order, immediately after the anchor) are the ONLY new content; the prefix through the anchor and everything after the N new blocks is unchanged, in sequence (the SAME block-sequence discipline insert_table already uses for its own new block, canonicalBlock REUSED from table-insert.cjs unchanged); the body block count grows by exactly N',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh322-bulk-insert-probe.live.cjs, commit 48afe5074: 50 paragraphs inserted after one anchor in ONE callCommand (doc.AddElement at strictly increasing indices, the SAME mechanism #308/#306 already measured for a table/image) land in the requested order, 24 ms JS-side, correct order/text after a fresh reopen, the control paragraph unaffected; the SAME 50 through the existing, released per-paragraph product path (50 insert_paragraph_after steps, one task call) took 38 773 ms -- three orders of magnitude slower, the reason this composite intent exists; 200 paragraphs in one callCommand: 36 ms, the session survives, the save lands',
    },
    note: 'selector { afterParagraph }; paragraphs: 1..200 items, each { text (no tab/line break), style? (an EXISTING paragraph style name, resolved by GetName()===name, the W-10-9 guard), format? (a SUBSET of format_text\'s own keys -- bold, italic, underline, strike, fontFamily, fontSize, color, alignment, spaceBefore, spaceAfter -- validated by the SAME normalizeWanted formatting.cjs/header-footer.cjs already reuse) }. ALL-OR-NOTHING: every requested style resolves BEFORE any paragraph is inserted (an unknown style refuses the WHOLE batch, nothing written); the insertion loop additionally rolls back every paragraph THIS call inserted on an unexpected runtime failure, highest index first. Retry identity compares each candidate block\'s OWN already-attached `.format` (inventory.cjs attachProperties) via formatting.cjs\'s exported compare() -- never a guessed match. No table/image in the batch (#308/#306 cover those as their own intents), no MAX_STEPS change.',
  }),
  N('insert_shape', 'shape-insert', {
    targetModule: 'src/adapters/docx/families/shape-insert.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: every OTHER shape is preserved as a multiset (text/size/own-paragraph key) and exactly one new one appears; every image is unchanged; every OTHER body block is unchanged in sequence and content and the ONE new block is exactly the empty paragraph this write creates, immediately after the anchor (canonicalBlock REUSED from table-insert.cjs unchanged; only the anchor block gets the mso-border-*-alt:none equivalence)',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh309-shape-probe.live.cjs (round 1, write side: Api.CreateShape rect + flowChartProcess with fill/stroke/own text + a new paragraph + doc.AddElement(anchorIndex+1, ...) land INLINE, the stored word/document.xml carries wps:wsp / prstGeom / wp:extent cx,cy byte-exact / srgbClr / a:ln w=19050 / the text in w:txbxContent) and test/live/docx-gh309-shape-read-probe.live.cjs (round 2, read side: doc.GetAllShapes() sees the shape and its own text and size on a fresh reopen, GetAllImages() does not list it) -- both LIVE PASS (prod NC/DS, disposable copy). The product\'s read_text / read_images do NOT see a shape; the L1 here is built on the driver\'s own document.shapes rows (docx-shapes-read.cjs), never on the empty paragraph alone',
    },
    note: 'v1 ships selector { afterParagraph } (the insert_table / insert_image anchor contract) and shape { type: rect | flowChartProcess | roundRect | ellipse | triangle | trapezoid | flowChartDecision | flowChartTerminator | rightArrow | chevron | homePlate (the MEASURED presets; an unknown name is stored as a rect by the engine, silently -- hence the allowlist), widthMm / heightMm (integers 1..400), fill / stroke (6-digit hex), text (one plain line, <= 500 chars) }. The stroke width is the one measured (19050 EMU). RETRY IDENTITY is position + own text + size based (the shape paragraph carries no text of its own) -- a NAMED, ACCEPTED limit (the family\'s own header comment); an UNREADABLE shape list fails closed instead of guessing. An optional placement { wrap: square | topAndBottom | inFront | behind, horizontal: { margin + align } | { page + offsetMm }, vertical: { paragraph | page + offsetMm } } makes the shape floating (GH #309 part 3, test/live/docx-gh309-floating-probe.live.cjs LIVE PASS); it reads back through the ToJSON projection and is part of the retry identity; tight / through are not offered (tight is stored behind the text). NOT claimed: the stroke COLOUR / preset round-trip through the product read (the LIVE harness reads them from the stored XML), other references (column / character / margin offsets) or a vertical align, header / footer / table-cell targets, multi-paragraph shape text, a native diagram (#323) or a vector-image fallback (#306).',
  }),
  N('insert_chart', 'chart-insert', {
    targetModule: 'src/adapters/docx/families/chart-insert.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: every OTHER chart is preserved as a multiset (type/size/own-paragraph/series/categories key) and exactly one new one appears; every image and every shape is unchanged; every OTHER body block is unchanged in sequence and content and the ONE new block is exactly the empty paragraph this write creates, immediately after the anchor (canonicalBlock REUSED from table-insert.cjs unchanged; only the anchor block gets the mso-border-*-alt:none equivalence -- the insert_shape discipline)',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh323-native-chart-probe.live.cjs (Api.CreateChart bar + SetTitle + SetLegendPos(bottom) + AddDrawing survives save + fresh reopen; stored word/charts/chart1.xml carries c:barChart, both series names, all 12 categories and values, the title, c:legendPos b; an embedded workbook is written; GetAllImages() does not list the chart), test/live/docx-gh323-chart-readback-and-types-probe.live.cjs (14/15 candidate types land, stored element per type; line returns null; the Word API series objects have no GetName/GetValues) and test/live/docx-gh323-chart-tojson-series-probe.live.cjs (the chart own ToJSON plotArea series caches carry name/categories/values in session and after reopen) -- all LIVE PASS',
    },
    note: 'v1 ships selector { afterParagraph } (the insert_shape anchor contract) and chart { type: bar | barStacked | barStackedPercent | horizontalBar | horizontalBarStacked | area | areaStacked | pie | doughnut (the measured set; pie/doughnut take exactly one series), categories (1..100 single-line strings), series (1..10 of { name, values: one finite number per category }, unique names), widthMm / heightMm (integers 10..400), title? (one line, <= 200 chars), legend?: bottom (the measured position) }. Style index 24 (the measured one). L1 verifies type, size, categories and every series name and value through the chart own ToJSON read-back (docx-charts-read.cjs); the title and legend are NOT read back by L1 (the LIVE product harness checks them in the stored XML). RETRY IDENTITY: position + the chart own data + size; an UNREADABLE chart list fails closed. NOT offered: line (CreateChart returns null), scatter, stock, 3D, editing an existing chart, table-cell / header / footer targets, floating placement, number formats, colours.',
  }),
  N('delete_image', 'image', {
    targetModule: 'src/adapters/docx/families/image.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: every OTHER image (part, paragraph index, placement, wrap, size, paragraph text) is preserved as a multiset, the image count drops by exactly one, the body unchanged',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh296-image-replace-delete-probe.live.cjs (run p296-r1, R1 header default + R4 header even, inline: ApiDrawing.Delete() returned true, the stored part keeps its text and loses its drawing, every other image and the body unchanged) + docx-images-probe.live.cjs R6_DELETE (a body floating image)',
    },
    note: 'v1 ships delete_image on { paragraph, ordinal? } or { index }, plus an optional selector.part assertion (\'body\' | \'header|footer:<section>:<default|title|even>\', the read_images row\'s own part label; a mismatch fails closed). The image is re-identified inside the sandbox by its paragraph text, extent and drawingType before the delete; the command returns a receipt. A retry after a landed delete fails closed (the image is gone) -- never SATISFIED.',
  }),
  N('replace_image', 'image', {
    targetModule: 'src/adapters/docx/families/image.cjs',
    status: STATUS.NEW,
    preservationMode: 'PRIMITIVE_INTEGRAL',
    preservationEvidence: 'no source intent -- target-authored placement: every OTHER image is preserved as a multiset, the image count is unchanged, the replacement sits at the target\'s own part + paragraph index with the same placement / wrap / paragraph text at the expected size (±1 EMU), the body unchanged',
    source: {
      evidence: 'LIVE probe evidence, not a source oracle: test/live/docx-gh296-image-replace-delete-probe.live.cjs (run p296-r1, R2 footer default PNG -> JPEG at the old size, R3 header title at a new size, R5 body floating, R6 body inline: doc.ReplaceDrawing(old, Api.CreateImage(...), ...) returned true, the stored media bytes sha256-match the new input, the floating anchor/square placement kept, every other image and the body unchanged) + docx-images-probe.live.cjs R5_REPLACE',
    },
    note: 'v1 ships replace_image { selector (as delete_image), source: data:image/png|jpeg base64 URI (the insert_image contract, 1 MiB cap), widthEmu + heightEmu together or neither (neither = keep the current size) }. ReplaceDrawing keeps the old drawing\'s placement. VERIFY LIMIT: the in-session read carries no image bytes, so L1 proves position, placement, size and the in-sandbox receipt; the bytes are proven by the LIVE harness from the stored file. A retry replaces again with the same bytes -- never SATISFIED.',
  }),
]);

module.exports = { NEW_INTENTS, CLASSES, STATUS };
