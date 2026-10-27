Api.GetEuroOfficeChartReadLayer = (function (install) {
  /* Euro-Office chart read layer, euro-office-chart-read-v1/9.3.5.0/1. Own implementation (M12, D2a = C), Apache-2.0.
     Reads the chart model fields the public setters of this runtime write; see README.md for the contract. */
  var LAYER = { layer: "euro-office-chart-read-v1", runtime: "9.3.5.0", revision: 1 };
  // The series class is not in scope here: take its prototype from the object this runtime's own GetSeries constructs
  // (the constructor only stores its two arguments; the stub chart answers the one existence check).
  var series = ApiChart.prototype.GetSeries.call({ Chart: { getSeriesByIdx: function () { return true; } } }, 0);
  install(ApiChart.prototype, Object.getPrototypeOf(series), ApiDrawing.prototype, "", LAYER, Asc.c_oAscChartLegendShowSettings);
  return function () { return { layer: LAYER.layer, runtime: LAYER.runtime, revision: LAYER.revision }; };
})(/*EO-CHART-READ-INSTALL-BEGIN*/function (chartProto, seriesProto, drawingProto, prefix, layer, legendTable) {
  var EMU_PER_MM = 36000;
  var LEGEND_NAMES = ["left", "top", "right", "bottom"];

  function rootOf(apiChart) {
    var space = apiChart && apiChart.Chart;
    return space && space.chart ? space.chart : null;
  }
  function textOfTitle(holder) {
    var title = holder ? holder.title : null;
    if (!title || typeof title.getDocContent !== "function") return null;
    var content = title.getDocContent();
    return content ? content.GetText() : null;
  }
  function legendName(code) {
    for (var i = 0; i < LEGEND_NAMES.length; i += 1) {
      if (legendTable[LEGEND_NAMES[i]] === code) return LEGEND_NAMES[i];
    }
    return null;
  }
  function axisOf(apiChart, horizontal) {
    var root = rootOf(apiChart);
    var area = root ? root.plotArea : null;
    if (!area) return null;
    return horizontal ? area.getHorizontalAxis() : area.getVerticalAxis();
  }
  function formulaOf(source) {
    if (!source) return null;
    var ref = source.strRef || source.numRef;
    var text = ref ? ref.f : null;
    if (typeof text !== "string" || text.length === 0) return null;
    return text.charAt(0) === "=" ? text : "=" + text;
  }
  function seriesModel(apiSeries) {
    return typeof apiSeries.private_GetSeries === "function" ? apiSeries.private_GetSeries() : null;
  }
  function define(proto, name, fn) {
    fn.euroOfficeChartRead = layer;
    proto[prefix + name] = fn;
  }

  define(chartProto, "GetLegendPos", function () {
    var root = rootOf(this);
    if (!root) return null;
    return root.legend ? legendName(root.legend.legendPos) : "none";
  });
  define(chartProto, "GetHorAxisTitle", function () {
    var axis = axisOf(this, true);
    return axis ? textOfTitle(axis) : null;
  });
  define(chartProto, "GetVerAxisTitle", function () {
    var axis = axisOf(this, false);
    return axis ? textOfTitle(axis) : null;
  });
  define(chartProto, "GetDataLabels", function () {
    var root = rootOf(this);
    var first = root && root.plotArea && root.plotArea.charts ? root.plotArea.charts[0] : null;
    var labels = first ? first.dLbls : null;
    return {
      showSerName: !!(labels && labels.showSerName),
      showCatName: !!(labels && labels.showCatName),
      showVal: !!(labels && labels.showVal),
      showPercent: !!(labels && labels.showPercent)
    };
  });
  function seriesFormula(apiSeries, pick) {
    var model = seriesModel(apiSeries);
    if (!model) return null;
    var formula = formulaOf(pick(model));
    return formula === null ? "" : formula;
  }
  define(seriesProto, "GetName", function () { return seriesFormula(this, function (m) { return m.tx; }); });
  define(seriesProto, "GetValues", function () {
    var model = seriesModel(this);
    // a series without a values reference (e.g. a scatter series, which keeps yVal) is not readable here: refuse loudly
    if (model && !model.val) throw new TypeError("GetValues: this series has no values reference");
    return seriesFormula(this, function (m) { return m.val; });
  });
  define(seriesProto, "GetXValues", function () { return seriesFormula(this, function (m) { return m.xVal; }); });
  define(seriesProto, "GetCatFormula", function () { return seriesFormula(this, function (m) { return m.cat; }); });
  define(drawingProto, "GetPosition", function () {
    var anchor = this.Drawing && this.Drawing.drawingBase ? this.Drawing.drawingBase.from : null;
    if (!anchor) return null;
    return {
      fromCol: anchor.col,
      colOffset: anchor.colOff * EMU_PER_MM,
      fromRow: anchor.row,
      rowOffset: anchor.rowOff * EMU_PER_MM
    };
  });
}/*EO-CHART-READ-INSTALL-END*/),
