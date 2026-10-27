'use strict';

/**
 * GH #268b (docs/GH268B-DOCX-READ-IMAGES-PRODUCT-DECISION.md): the PAGE-SIDE image read of the DOCX driver's readLive. `readImagesInto` is injected into the callCommand body as its own SOURCE
 * (`IMAGES_READ_CALL`), so it MUST stay self-contained: no outer variable, no require, no regex literal with a backslash, no backtick / dollar-brace (the body is a template literal).
 *
 * PUBLIC API only: doc.GetAllImages() (ApiDocument inherits it from ApiDocumentContent) and, per image, GetParentParagraph().GetText() + the PROJECTION of ApiDrawing.ToJSON(): docPr.descr / title (docPr.id / name are deliberately NOT read: editor-session-local, g268v1 LIVE),
 * extent.cx / cy (EMU), drawingType, wrapType. ToJSON ALSO serialises the image BYTES (graphic ... blipFill.rasterImageId, base64) -- they are parsed here and DROPPED: only the primitive fields
 * above are copied into the row, and a field of an unexpected type is null, so an object (or a string of bytes) can never ride along in a field. The result is rows of primitives; a failed read is
 * 'UNOBSERVABLE:<reason>' (a failed ToJSON marks the JSON-derived fields of THAT row via `json`), never a guessed value.
 */
function readImagesInto(doc, docInfo) {
  const unobservable = e => 'UNOBSERVABLE:' + String((e && e.message) || e).slice(0, 80);
  const str = v => (typeof v === 'string' ? v : null);
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  try {
    if (!doc || typeof doc.GetAllImages !== 'function') { docInfo.images = 'UNOBSERVABLE:no-GetAllImages'; return; }
    docInfo.images = (doc.GetAllImages() || []).map((im, index) => {
      const row = { index, descr: null, title: null, cx: null, cy: null, drawingType: null, wrapType: null, paragraph: null, json: null };
      try { const p = im.GetParentParagraph(); if (p) { const t = String(p.GetText({ Numbering: false })); let n = t.length; while (n > 0 && (t.charCodeAt(n - 1) === 10 || t.charCodeAt(n - 1) === 13)) n -= 1; row.paragraph = t.slice(0, n); } } catch (e) { row.paragraph = unobservable(e); }   // GetText ends with CR LF (g268v1 LIVE: "GH268 para A\r\n"); the driver's other paragraph reads strip it -- no backslash / regex here (template-literal injection)
      let j = null;
      try { j = JSON.parse(im.ToJSON(false, false)); } catch (e) { row.json = unobservable(e); }
      if (j && typeof j === 'object') {
        const dp = j.docPr && typeof j.docPr === 'object' ? j.docPr : null;
        if (dp) { row.descr = str(dp.descr); row.title = str(dp.title); }
        const ex = j.extent && typeof j.extent === 'object' ? j.extent : null;
        if (ex) { row.cx = num(ex.cx); row.cy = num(ex.cy); }
        row.drawingType = str(j.drawingType);
        row.wrapType = str(j.wrapType);
      }
      return row;
    });
  } catch (e) { docInfo.images = unobservable(e); }
}

const IMAGES_READ_CALL = `(${readImagesInto.toString()})(doc, docInfo);`;

module.exports = { readImagesInto, IMAGES_READ_CALL };
