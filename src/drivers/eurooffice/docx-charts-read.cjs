'use strict';

/**
 * GH #323 (insert_chart): the PAGE-SIDE native-chart read of the DOCX driver's readLive. `readChartsInto` is injected into the callCommand body as its own SOURCE (`CHARTS_READ_CALL`),
 * so it MUST stay self-contained: no outer variable, no require, no regex literal with a backslash, no backtick / dollar-brace (the body is a template literal) -- the SAME contract
 * docx-shapes-read.cjs states for SHAPES_READ_CALL.
 *
 * LIVE-measured (test/live/docx-gh323-native-chart-probe.live.cjs, -chart-readback-and-types-probe, -chart-tojson-series-probe; fresh reopen of the stored document):
 *   - a chart is a paragraph-level drawing object of class 'chart' (doc.GetAllCharts() lists it, GetAllImages() does NOT); GetChartType() reads back the type string it was created with;
 *     GetWidth() / GetHeight() read back the requested EMU (height within rounding: 3000000 -> 2999999.88).
 *   - the Word API's chart SERIES objects have NO GetName / GetValues / GetCatFormula (TypeError: not a function) -- the per-series getters the XLSX side uses do not exist here.
 *   - the chart's OWN public ToJSON() carries graphic.chart.plotArea.charts[0].ser[i] with tx.strRef.strCache.pt (the series name), cat.strRef.strCache.pt (the categories) and
 *     val.numRef.numCache.pt (the values), in the same session and after a fresh reopen; no base64 in it.
 * The series and categories ride as JSON STRINGS (seriesJson / categoriesJson): a nested array this deep in the callCommand result is the measured transport-escape class (a per-run array in
 * read_headers_footers wiped the WHOLE result) -- a string is a primitive, inflated Node-side by inventory.cjs.
 * SCOPE: body-level paragraphs only (a chart inside a table cell, a header/footer or a text box is not listed here).
 * Each row carries `previous`: the text of the body element IMMEDIATELY BEFORE the chart's own body paragraph (null when that element is not a paragraph, or the chart paragraph is first),
 * the SAME anchor relation docx-shapes-read.cjs gives insert_shape. A failed read is 'UNOBSERVABLE:<reason>', never a guessed value. Rows are primitives only.
 */
function readChartsInto(doc, docInfo) {
  const unobservable = e => 'UNOBSERVABLE:' + String((e && e.message) || e).slice(0, 80);
  const num = v => (typeof v === 'number' && isFinite(v) ? v : null);
  const textOf = p => { const t = String(p.GetText({ Numbering: false })); let n = t.length; while (n > 0 && (t.charCodeAt(n - 1) === 10 || t.charCodeAt(n - 1) === 13)) n -= 1; return t.slice(0, n); };
  // a cache's points in idx order; a missing idx stays null (never shifted), a non-numeric value stays its own string
  const points = (cache, count, asNumber) => {
    const out = [];
    for (let i = 0; i < count; i += 1) out.push(null);
    const pts = cache && Array.isArray(cache.pt) ? cache.pt : [];
    for (let i = 0; i < pts.length; i += 1) {
      const p = pts[i];
      if (!p || typeof p.idx !== 'number' || p.idx < 0 || p.idx >= count) continue;
      const v = p.v;
      if (asNumber) { const n = typeof v === 'number' ? v : Number(v); out[p.idx] = isFinite(n) ? n : String(v); } else out[p.idx] = v === null || v === undefined ? null : String(v);
    }
    return out;
  };
  const countOf = cache => (cache && typeof cache.ptCount === 'number' ? cache.ptCount : (cache && Array.isArray(cache.pt) ? cache.pt.length : 0));
  const project = ch => {
    const j = JSON.parse(ch.ToJSON(false, false));
    const pa = j && j.graphic && j.graphic.chart && j.graphic.chart.plotArea;
    const charts = pa && Array.isArray(pa.charts) ? pa.charts : null;
    if (!charts || charts.length !== 1) throw new Error('plot-area-charts-not-one');
    const sers = Array.isArray(charts[0].ser) ? charts[0].ser : [];
    let categories = null;
    const series = sers.map(s => {
      const tx = s.tx && s.tx.strRef ? s.tx.strRef.strCache : null;
      const cat = s.cat && s.cat.strRef ? s.cat.strRef.strCache : null;
      const val = s.val && s.val.numRef ? s.val.numRef.numCache : null;
      const cats = points(cat, countOf(cat), false);
      if (categories === null) categories = cats;
      else if (JSON.stringify(categories) !== JSON.stringify(cats)) throw new Error('series-categories-differ');
      const names = points(tx, countOf(tx), false);
      return { name: names.length ? names[0] : null, values: points(val, countOf(val), true) };
    });
    return { seriesJson: JSON.stringify(series), categoriesJson: JSON.stringify(categories || []) };
  };
  try {
    if (!doc || typeof doc.GetAllCharts !== 'function') { docInfo.charts = 'UNOBSERVABLE:no-GetAllCharts'; return; }
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
          if (cls !== 'chart') continue;
          const row = { index: rows.length, type: null, cx: null, cy: null, paragraph: null, previous: previousIsParagraph ? previous : null, seriesJson: null, categoriesJson: null };
          try { row.paragraph = textOf(el); } catch (e) { row.paragraph = unobservable(e); }
          try { const t = d.GetChartType(); row.type = typeof t === 'string' ? t : null; } catch (e) { row.type = unobservable(e); }
          try { row.cx = num(d.GetWidth()); row.cy = num(d.GetHeight()); } catch (e) { row.cx = unobservable(e); row.cy = row.cx; }
          try { const pr = project(d); row.seriesJson = pr.seriesJson; row.categoriesJson = pr.categoriesJson; } catch (e) { row.seriesJson = unobservable(e); row.categoriesJson = row.seriesJson; }
          rows.push(row);
        }
      }
      if (isParagraph) { try { previous = textOf(el); previousIsParagraph = true; } catch { previous = null; previousIsParagraph = false; } } else { previous = null; previousIsParagraph = false; }
    }
    docInfo.charts = rows;
  } catch (e) { docInfo.charts = unobservable(e); }
}

const CHARTS_READ_CALL = `(${readChartsInto.toString()})(doc, docInfo);`;

module.exports = { readChartsInto, CHARTS_READ_CALL };
