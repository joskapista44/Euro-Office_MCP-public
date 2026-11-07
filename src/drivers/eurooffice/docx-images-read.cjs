'use strict';

/**
 * GH #268b (docs/GH268B-DOCX-READ-IMAGES-PRODUCT-DECISION.md): the PAGE-SIDE image read of the DOCX driver's readLive. `readImagesInto` is injected into the callCommand body as its own SOURCE
 * (`IMAGES_READ_CALL`), so it MUST stay self-contained: no outer variable, no require, no regex literal with a backslash, no backtick / dollar-brace (the body is a template literal).
 *
 * PUBLIC API only: doc.GetAllImages() (ApiDocument inherits it from ApiDocumentContent) and, per image, GetParentParagraph().GetText() + the PROJECTION of ApiDrawing.ToJSON(): docPr.descr / title,
 * extent.cx / cy (EMU), drawingType, wrapType. ToJSON ALSO serialises the image BYTES (graphic ... blipFill.rasterImageId, base64) -- they are parsed here and DROPPED: only the primitive fields
 * above are copied into the row, and a field of an unexpected type is null, so an object (or a string of bytes) can never ride along in a field. The result is rows of primitives; a failed read is
 * 'UNOBSERVABLE:<reason>' (a failed ToJSON marks the JSON-derived fields of THAT row via `json`), never a guessed value.
 *
 * GH #306 part 4: the placement projection (behindDoc, hRel / hAlign / hOffset, vRel / vAlign / vOffset), the same fields docx-shapes-read.cjs reads for a shape.
 *
 * GH #306b: `part` (body/header/footer). This read's OWN scope text used to claim "header / footer
 * images ... are not listed" -- LIVE-measured FALSE (docx-gh306b-header-image-probe.live.cjs, LIVE
 * PASS, 2026-10-28): doc.GetAllImages() lists them, unlabeled. There is NO public image-to-part path
 * (LIVE-measured, docx-gh306b-image-part-discriminator-probe.live.cjs: GetParent/GetParentContent/
 * GetParentDocumentContent/GetDocumentContent do not exist on an image or its GetParentParagraph();
 * GetSection()/GetPosInParent() on the parent paragraph do not discriminate). The RAW enumeration
 * direction that works (LIVE PASS, docx-gh306b-image-part-enumeration-probe.live.cjs): walk the
 * containers (body + table cells via a recursive walk; every section's header/footer parts, every
 * variant) using `paragraph.GetAllDrawingObjects()` filtered to `GetClassType() === 'image'` -- the
 * SAME pattern GH #309's own docx-shapes-read.cjs already uses for shapes.
 *
 * THE JOIN BACK TO `doc.GetAllImages()`: identity (`===`) is LIVE-MEASURED FALSE against this exact
 * product code path (2026-10-28: 0/3 images matched, even inside one synchronous page-side
 * execution -- the fourth instance of the wrapper-identity-never-matches class). The replacement is
 * `ToJSON().docPr.id`: a FOLLOW-UP probe (docx-gh306b-image-docpr-id-probe.live.cjs, LIVE PASS,
 * 2026-10-28) measured it readable+unique on BOTH enumerations, with the two id-sets identical,
 * WITHIN one callCommand execution. The SAME probe also measured the id's own limit: it is a
 * runtime-generated number, NOT the stored XML docPr id (that one was 1/2 in a separate probe) --
 * valid ONLY as a same-call join key, never persisted or compared across sessions. Because that
 * uniqueness was measured on ONE fixture (3 images, no duplicates), the code below does NOT assume
 * it holds generally: it checks, on every read, (a) every id is readable on both sides, (b) no id
 * repeats on either side, and (c) the two id-sets match element-for-element -- if ANY of the three
 * fails, EVERY row's part is the matching NAMED gap, never a guessed/missing value (a bare null
 * looked like a valid absence once already and was not -- the dangerous direction for a silent
 * defect, measured the first time this join was attempted by identity).
 *
 * The EXISTING `doc.GetAllImages()` array order and `index` numbering are kept BYTE-IDENTICAL to
 * before this change regardless of the join's own fate -- `set_image_size`/`set_image_wrap`
 * (image.cjs's own commandBody) address an image by `doc.GetAllImages()[selector.index]` directly,
 * so renumbering would silently misaddress an UNRELATED image; `part` is an ADDITIONAL field, never
 * a reordering.
 */
function readImagesInto(doc, docInfo) {
  const unobservable = e => 'UNOBSERVABLE:' + String((e && e.message) || e).slice(0, 80);
  const str = v => (typeof v === 'string' ? v : null);
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  const docPrIdOf = j => {
    const dp = j && typeof j === 'object' && j.docPr && typeof j.docPr === 'object' ? j.docPr : null;
    const id = dp ? dp.id : undefined;
    // readable means a primitive (string/number) id, never an object/array/boolean standing in
    // for one -- an id of an unexpected type is treated the same as a missing one, UNREADABLE.
    if (typeof id === 'number' && isFinite(id)) return id;
    if (typeof id === 'string' && id !== '') return id;
    return null;
  };
  try {
    if (!doc || typeof doc.GetAllImages !== 'function') { docInfo.images = 'UNOBSERVABLE:no-GetAllImages'; return; }
    const allImages = doc.GetAllImages() || [];
    const allImagesJson = allImages.map(im => {
      let j = null;
      let err = null;
      try { j = JSON.parse(im.ToJSON(false, false)); } catch (e) { err = e; }
      return { j, err };
    });
    const allIds = allImagesJson.map(entry => docPrIdOf(entry.j));

    // GH #306b: the part-map, built by walking every container (body, table cells, every section's
    // header/footer part + variant) and recording each walked image's own docPr.id alongside its
    // part label -- the join target for the GetAllImages() side, below.
    const walkedIds = [];
    const walkedParts = [];
    // GH #317 follow-up: the image's paragraph index among its part's OWN top-level elements (the SAME
    // GetElement(i) index the header/footer write addresses); null for an image inside a table cell.
    // The paragraph TEXT cannot serve as that identity: GetText() drops a real tab (GH #317 LIVE).
    const walkedParaIdx = [];
    const walkContainer = (container, label, topLevel) => {
      if (!container || typeof container.GetElementsCount !== 'function') return;
      let n = 0;
      try { n = container.GetElementsCount(); } catch { return; }
      for (let i = 0; i < n; i += 1) {
        let el = null;
        try { el = container.GetElement(i); } catch { continue; }
        if (!el || typeof el.GetClassType !== 'function') continue;
        let cls = null;
        try { cls = el.GetClassType(); } catch { cls = null; }
        if (cls === 'paragraph' && typeof el.GetAllDrawingObjects === 'function') {
          let objs = [];
          try { objs = el.GetAllDrawingObjects() || []; } catch { objs = []; }
          for (const d of objs) {
            let dcls = null;
            try { dcls = d.GetClassType(); } catch { dcls = null; }
            if (dcls !== 'image') continue;
            let wj = null;
            try { wj = JSON.parse(d.ToJSON(false, false)); } catch { wj = null; }
            walkedIds.push(docPrIdOf(wj));
            walkedParts.push(label);
            walkedParaIdx.push(topLevel ? i : null);
          }
        } else if (cls === 'table' && typeof el.GetRowsCount === 'function') {
          try {
            const rc = el.GetRowsCount();
            for (let r = 0; r < rc; r += 1) {
              const row = el.GetRow(r);
              const cc = row.GetCellsCount();
              for (let c = 0; c < cc; c += 1) walkContainer(row.GetCell(c).GetContent(), label, false);
            }
          } catch { /* unread table content, skip -- a resulting join mismatch catches it */ }
        }
      }
    };
    walkContainer(doc, 'body', true);
    let secs = [];
    try { secs = typeof doc.GetSections === 'function' ? (doc.GetSections() || []) : []; } catch { secs = []; }
    for (let si = 0; si < secs.length; si += 1) {
      const sec = secs[si];
      const variants = ['default', 'even', 'title'];
      for (let vi = 0; vi < variants.length; vi += 1) {
        const v = variants[vi];
        let hdr = null;
        let ftr = null;
        try { hdr = sec.GetHeader(v, false); } catch { hdr = null; }
        try { ftr = sec.GetFooter(v, false); } catch { ftr = null; }
        if (hdr) walkContainer(hdr, 'header:' + si + ':' + v, true);
        if (ftr) walkContainer(ftr, 'footer:' + si + ':' + v, true);
      }
    }

    // GH #306b (LIVE PASS, docPr.id probe, 2026-10-28): the join key is docPr.id, measured
    // readable+unique on both enumerations with identical id-sets -- WITHIN one callCommand
    // execution only (the id is runtime-generated, never persisted). That measurement covered ONE
    // fixture; it does not license assuming uniqueness generally, so every read re-checks all three
    // conditions before trusting the join, and fails closed on the first one that does not hold.
    const allReadable = arr => arr.every(id => id !== null);
    const hasDuplicates = arr => new Set(arr).size !== arr.length;
    let partFailReason = null;
    if (!allReadable(allIds) || !allReadable(walkedIds)) {
      partFailReason = 'UNOBSERVABLE:part-docpr-id-unreadable';
    } else if (hasDuplicates(allIds) || hasDuplicates(walkedIds)) {
      partFailReason = 'UNOBSERVABLE:part-docpr-id-duplicate';
    } else {
      const aSorted = allIds.slice().sort();
      const wSorted = walkedIds.slice().sort();
      const setsMatch = aSorted.length === wSorted.length && aSorted.every((v, i) => v === wSorted[i]);
      if (!setsMatch) partFailReason = 'UNOBSERVABLE:part-docpr-id-set-mismatch';
    }
    const idToPart = new Map();
    const idToParaIdx = new Map();
    if (!partFailReason) {
      for (let i = 0; i < walkedIds.length; i += 1) { idToPart.set(walkedIds[i], walkedParts[i]); idToParaIdx.set(walkedIds[i], walkedParaIdx[i]); }
    }

    docInfo.images = allImages.map((im, index) => {
      const entry = allImagesJson[index];
      const row = { index, part: partFailReason || idToPart.get(allIds[index]), paragraphIndex: partFailReason || idToParaIdx.get(allIds[index]), descr: null, title: null, cx: null, cy: null, drawingType: null, wrapType: null, behindDoc: null, hRel: null, hAlign: null, hOffset: null, vRel: null, vAlign: null, vOffset: null, paragraph: null, json: null };
      try { const p = im.GetParentParagraph(); if (p) { const t = String(p.GetText({ Numbering: false })); let n = t.length; while (n > 0 && (t.charCodeAt(n - 1) === 10 || t.charCodeAt(n - 1) === 13)) n -= 1; row.paragraph = t.slice(0, n); } } catch (e) { row.paragraph = unobservable(e); }   // GetText ends with CR LF (g268v1 LIVE: "GH268 para A\r\n"); the driver's other paragraph reads strip it -- no backslash / regex here (template-literal injection)
      if (entry.err) { row.json = unobservable(entry.err); return row; }
      const j = entry.j;
      if (j && typeof j === 'object') {
        const dp = j.docPr && typeof j.docPr === 'object' ? j.docPr : null;
        if (dp) { row.descr = str(dp.descr); row.title = str(dp.title); }
        const ex = j.extent && typeof j.extent === 'object' ? j.extent : null;
        if (ex) { row.cx = num(ex.cx); row.cy = num(ex.cy); }
        row.drawingType = str(j.drawingType);
        row.wrapType = str(j.wrapType);
        // GH #306 part 4: the SAME placement projection docx-shapes-read.cjs reads for a shape (behindDoc, positionH / positionV relativeFrom + align or posOffset in EMU);
        // LIVE-measured for an image (docx-gh306-cell-floating-image-probe.live.cjs): it matched the stored wp:anchor behindDoc / positionH / positionV for every floating variant.
        row.behindDoc = [true, false].includes(j.behindDoc) ? j.behindDoc : null;
        const ph = Object(j.positionH);
        const pv = Object(j.positionV);
        row.hRel = str(ph.relativeFrom); row.hAlign = str(ph.align); row.hOffset = num(ph.posOffset);
        row.vRel = str(pv.relativeFrom); row.vAlign = str(pv.align); row.vOffset = num(pv.posOffset);
      }
      return row;
    });
  } catch (e) { docInfo.images = unobservable(e); }
}

const IMAGES_READ_CALL = `(${readImagesInto.toString()})(doc, docInfo);`;

module.exports = { readImagesInto, IMAGES_READ_CALL };
