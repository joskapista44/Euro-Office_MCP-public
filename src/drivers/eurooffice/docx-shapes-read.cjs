'use strict';

/**
 * GH #309 (insert_shape): the PAGE-SIDE shape read of the DOCX driver's readLive. `readShapesInto` is injected into the callCommand body as its own SOURCE (`SHAPES_READ_CALL`),
 * so it MUST stay self-contained: no outer variable, no require, no regex literal with a backslash, no backtick / dollar-brace (the body is a template literal) -- the SAME contract
 * docx-images-read.cjs states for IMAGES_READ_CALL.
 *
 * LIVE-measured (test/live/docx-gh309-shape-read-probe.live.cjs, fresh reopen of the stored document): doc.GetAllShapes() lists a DrawingML shape (class 'shape') and NOT an image;
 * GetAllImages() does NOT list a shape; a shape's own text is GetDocContent().GetElement(0).GetText(); GetWidth() / GetHeight() read back the requested EMU exactly; the shape's parent
 * paragraph carries NO text of its own, so a paragraph-level GetAllDrawingObjects() on the body element is the link from a shape to its place in the body.
 *
 * WHY A DOCUMENT-LEVEL ROW LIST (like `images`) AND NOT A FIELD ON THE PARAGRAPH BLOCK: the per-paragraph read in docx-driver.cjs skips every paragraph without text
 * (`if (!text.trim()) continue`), and a shape paragraph is exactly such a paragraph -- a per-paragraph field would need that skip rule changed, and with it every consumer's
 * paragraph indexing. A separate list cannot move any other intent's preservation (canonicalBlock never reads it).
 *
 * Each row carries `previous`: the text of the body element IMMEDIATELY BEFORE the shape's own body paragraph (null when that element is not a paragraph, or the shape paragraph is
 * first) -- the anchor relation insert_shape's retry identity is built on. The anchor is found by TEXT, never by wrapper identity (a body element wrapper is new per call).
 * A failed read is 'UNOBSERVABLE:<reason>', never a guessed value. Rows are primitives only.
 *
 * PLACEMENT (GH #309 part 3, LIVE-measured: test/live/docx-gh309-floating-probe.live.cjs): GetWrappingStyle() does NOT exist on this DS; the PROJECTION of ApiDrawing.ToJSON(false, false) -- the
 * same call docx-images-read.cjs projects -- carries drawingType ('inline' | 'anchor'), wrapType, behindDoc and positionH / positionV (relativeFrom + align or posOffset in EMU), and matched the
 * stored wp:anchor / wp:inline, wrap element, behindDoc and position for every measured variant. An INLINE shape reads wrapType 'through' and default positions: only drawingType tells it apart.
 */
function readShapesInto(doc, docInfo) {
  const unobservable = e => 'UNOBSERVABLE:' + String((e && e.message) || e).slice(0, 80);
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  const textOf = p => { const t = String(p.GetText({ Numbering: false })); let n = t.length; while (n > 0 && (t.charCodeAt(n - 1) === 10 || t.charCodeAt(n - 1) === 13)) n -= 1; return t.slice(0, n); };
  try {
    if (!doc || typeof doc.GetAllShapes !== 'function') { docInfo.shapes = 'UNOBSERVABLE:no-GetAllShapes'; return; }
    const rows = [];
    const count = doc.GetElementsCount();
    let previous = null;
    let previousIsParagraph = false;
    for (let i = 0; i < count; i += 1) {
      const el = doc.GetElement(i);
      const isParagraph = !!el && typeof el.GetClassType === 'function' && el.GetClassType() === 'paragraph';
      if (isParagraph && typeof el.GetAllDrawingObjects === 'function') {
        let objs = [];
        try { objs = el.GetAllDrawingObjects() || []; } catch { objs = []; }
        for (let k = 0; k < objs.length; k += 1) {
          const d = objs[k];
          let cls = null;
          try { cls = d.GetClassType(); } catch { cls = null; }
          if (cls !== 'shape') continue;
          const row = { index: rows.length, text: null, cx: null, cy: null, paragraph: null, previous: previousIsParagraph ? previous : null,
            drawingType: null, wrapType: null, behindDoc: null, hRel: null, hAlign: null, hOffset: null, vRel: null, vOffset: null };
          try { row.paragraph = textOf(el); } catch (e) { row.paragraph = unobservable(e); }
          try { const c = d.GetDocContent(); const first = c ? c.GetElement(0) : null; row.text = first ? textOf(first) : ''; } catch (e) { row.text = unobservable(e); }
          try { row.cx = num(d.GetWidth()); row.cy = num(d.GetHeight()); } catch (e) { row.cx = unobservable(e); row.cy = row.cx; }
          try {
            const j = JSON.parse(d.ToJSON(false, false));
            const str = v => (typeof v === 'string' ? v : null);
            const ph = j.positionH || {};
            const pv = j.positionV || {};
            row.drawingType = str(j.drawingType); row.wrapType = str(j.wrapType); row.behindDoc = typeof j.behindDoc === 'boolean' ? j.behindDoc : null;
            row.hRel = str(ph.relativeFrom); row.hAlign = str(ph.align); row.hOffset = num(ph.posOffset);
            row.vRel = str(pv.relativeFrom); row.vAlign = str(pv.align); row.vOffset = num(pv.posOffset);
          } catch (e) { const u = unobservable(e); row.drawingType = u; row.wrapType = u; row.behindDoc = u; row.hRel = u; row.hAlign = u; row.hOffset = u; row.vRel = u; row.vAlign = u; row.vOffset = u; }
          rows.push(row);
        }
      }
      if (isParagraph) { try { previous = textOf(el); previousIsParagraph = true; } catch { previous = null; previousIsParagraph = false; } } else { previous = null; previousIsParagraph = false; }
    }
    docInfo.shapes = rows;
  } catch (e) { docInfo.shapes = unobservable(e); }
}

const SHAPES_READ_CALL = `(${readShapesInto.toString()})(doc, docInfo);`;

module.exports = { readShapesInto, SHAPES_READ_CALL };
