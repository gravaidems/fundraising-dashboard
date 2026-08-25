/*
 * CABViz — SVG chart primitives and the ⓘ provenance bubble.
 *
 * The gate: nothing in here draws a number or a mark without a provenance object.
 * requireProvenance() throws rather than degrading, so a missing bubble is a loud
 * failure during development instead of a silent omission in front of staff.
 *
 * Charts are hand-rolled SVG. No charting library is vendored: the workbook needs
 * five mark types, and owning the markup is what lets every mark carry its source.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CABViz = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // The provenance gate
  // ---------------------------------------------------------------------------

  var REQUIRED_FIELDS = ['tab', 'dataRange', 'columnsUsed', 'rowsAvailable', 'rowsUsed', 'transform'];

  function requireProvenance(p, label) {
    if (!p) throw new Error('CABViz: "' + label + '" was rendered without a provenance object.');
    REQUIRED_FIELDS.forEach(function (f) {
      if (p[f] === undefined || p[f] === null) {
        throw new Error('CABViz: "' + label + '" provenance is missing required field "' + f + '".');
      }
    });
    if (!Array.isArray(p.columnsUsed) || !p.columnsUsed.length) {
      throw new Error('CABViz: "' + label + '" provenance lists no columns.');
    }
    p.columnsUsed.forEach(function (c) {
      if (!c.letter) throw new Error('CABViz: "' + label + '" column "' + c.name + '" has no Excel letter.');
    });
    (p.rowsExcluded || []).forEach(function (e) {
      if (!e.reason) throw new Error('CABViz: "' + label + '" has an excluded-row bucket with no reason.');
    });
    return p;
  }

  // ---------------------------------------------------------------------------
  // Formatting
  // ---------------------------------------------------------------------------

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  var fmt = {
    currency: function (v) {
      if (v == null || !isFinite(v)) return '—';
      return '$' + Math.round(v).toLocaleString('en-US');
    },
    currencyPrecise: function (v) {
      if (v == null || !isFinite(v)) return '—';
      return '$' + v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    },
    compact: function (v) {
      if (v == null || !isFinite(v)) return '—';
      var a = Math.abs(v), s = v < 0 ? '-' : '';
      if (a >= 1e6) return s + '$' + (a / 1e6).toFixed(a >= 1e7 ? 0 : 1) + 'M';
      if (a >= 1e3) return s + '$' + (a / 1e3).toFixed(a >= 1e4 ? 0 : 1) + 'k';
      return s + '$' + Math.round(a);
    },
    number: function (v) {
      if (v == null || !isFinite(v)) return '—';
      return Math.round(v).toLocaleString('en-US');
    },
    percent: function (v, dp) {
      if (v == null || !isFinite(v)) return '—';
      return (v * 100).toFixed(dp == null ? 1 : dp) + '%';
    },
    ratio: function (v) {
      if (v == null || !isFinite(v)) return '—';
      return v.toFixed(2) + '×';
    },
    date: function (d) {
      if (!(d instanceof Date)) return '—';
      return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getUTCMonth()] +
        ' ' + d.getUTCDate() + ', ' + d.getUTCFullYear();
    },
    auto: function (v, kind) {
      if (kind === 'currency') return fmt.currency(v);
      if (kind === 'percent') return fmt.percent(v);
      if (kind === 'ratio') return fmt.ratio(v);
      return fmt.number(v);
    }
  };

  // ---------------------------------------------------------------------------
  // The ⓘ bubble
  // ---------------------------------------------------------------------------

  function bubbleHTML(p, label) {
    requireProvenance(p, label || 'unnamed');
    var h = [];
    h.push('<div class="prov">');
    h.push('<div class="prov-hd">' + esc(label || 'How this is calculated') + '</div>');

    h.push('<dl class="prov-dl">');
    h.push('<dt>Source</dt><dd>' + esc(p.tab) +
      (p.headerRow != null ? ' <span class="prov-dim">· header row ' + esc(p.headerRow) + '</span>' : '') + '</dd>');
    h.push('<dt>Cells</dt><dd><code>' + esc(p.dataRange) + '</code></dd>');

    h.push('<dt>Columns</dt><dd><ul class="prov-cols">');
    p.columnsUsed.forEach(function (c) {
      h.push('<li><span class="prov-letter">' + esc(c.letter) + '</span> ' + esc(c.name) +
        (c.role ? ' <span class="prov-dim">— ' + esc(c.role) + '</span>' : '') + '</li>');
    });
    h.push('</ul></dd>');

    h.push('<dt>Rows</dt><dd><strong>' + esc(fmt.number(p.rowsUsed)) + '</strong> used of ' +
      esc(fmt.number(p.rowsAvailable)) + ' available</dd>');

    var excl = (p.rowsExcluded || []).filter(function (e) { return e.count > 0; });
    if (excl.length) {
      h.push('<dt>Excluded</dt><dd><ul class="prov-excl">');
      excl.forEach(function (e) {
        h.push('<li><span class="prov-count">' + esc(fmt.number(e.count)) + '</span> ' + esc(e.reason) + '</li>');
      });
      h.push('</ul></dd>');
    }

    h.push('<dt>Transform</dt><dd><code class="prov-transform">' + esc(p.transform) + '</code></dd>');
    h.push('</dl>');

    if (p.notes && p.notes.length) {
      h.push('<div class="prov-notes">');
      p.notes.forEach(function (n) { h.push('<p>' + esc(n) + '</p>'); });
      h.push('</div>');
    }
    h.push('</div>');
    return h.join('');
  }

  // ---------------------------------------------------------------------------
  // Scales and ticks
  // ---------------------------------------------------------------------------

  function niceTicks(min, max, count) {
    if (min === max) { min = Math.min(0, min); max = max || 1; }
    var span = max - min;
    if (span <= 0) span = Math.abs(max) || 1;
    var raw = span / (count || 5);
    var mag = Math.pow(10, Math.floor(Math.log(raw) / Math.LN10));
    var norm = raw / mag;
    var step = (norm >= 7.5 ? 10 : norm >= 3.5 ? 5 : norm >= 1.5 ? 2 : 1) * mag;
    var lo = Math.floor(min / step) * step;
    var hi = Math.ceil(max / step) * step;
    var ticks = [];
    for (var v = lo; v <= hi + step / 1e6; v += step) ticks.push(Math.abs(v) < step / 1e6 ? 0 : v);
    return { ticks: ticks, lo: lo, hi: hi };
  }

  var M = { top: 16, right: 18, bottom: 34, left: 60 };

  function frame(w, h, m) {
    m = m || M;
    return {
      w: w, h: h, m: m,
      iw: w - m.left - m.right,
      ih: h - m.top - m.bottom
    };
  }

  function axes(f, scale, xLabels, opts) {
    opts = opts || {};
    var s = [];
    // Recessive gridlines and baseline; ticks in muted ink.
    scale.ticks.forEach(function (t) {
      var y = f.m.top + f.ih - ((t - scale.lo) / (scale.hi - scale.lo)) * f.ih;
      s.push('<line class="v-grid" x1="' + f.m.left + '" x2="' + (f.m.left + f.iw) + '" y1="' + y.toFixed(1) + '" y2="' + y.toFixed(1) + '"/>');
      s.push('<text class="v-tick v-tick-y" x="' + (f.m.left - 8) + '" y="' + (y + 3.5).toFixed(1) + '">' +
        esc(opts.yFormat ? opts.yFormat(t) : fmt.number(t)) + '</text>');
    });
    s.push('<line class="v-axis" x1="' + f.m.left + '" x2="' + (f.m.left + f.iw) + '" y1="' + (f.m.top + f.ih) + '" y2="' + (f.m.top + f.ih) + '"/>');

    var n = xLabels.length;
    // Thin by real pixel spacing rather than by a count heuristic: measure roughly
    // how wide each label renders, then keep one only if it clears the last kept one.
    // A 380px small multiple fits far fewer "Aug 2026"s than a 760px chart.
    function labelWidth(lb) { return String(lb).length * 6.1; }

    var placed = [];
    for (var i = 0; i < n; i++) {
      var x = n === 1 ? f.m.left + f.iw / 2 : f.m.left + (i / (n - 1)) * f.iw;
      if (opts.band) x = f.m.left + (i + 0.5) * (f.iw / n);
      placed.push({ i: i, x: x, w: labelWidth(xLabels[i]) });
    }

    // One extent function, used for both the reserved labels and the candidates —
    // computing it two slightly different ways is what let a pair overlap.
    function anchorOf(i) {
      if (opts.band) return 'middle';
      return i === n - 1 ? 'end' : i === 0 ? 'start' : 'middle';
    }
    function extent(i) {
      var p = placed[i], a = anchorOf(i);
      var left = a === 'start' ? p.x : a === 'end' ? p.x - p.w : p.x - p.w / 2;
      return [left, left + p.w];
    }

    var GAP = 8;
    var keep = {};
    if (n) keep[0] = true;

    function clears(i) {
      var c = extent(i);
      var idxs = Object.keys(keep).map(Number);
      for (var k = 0; k < idxs.length; k++) {
        var o = extent(idxs[k]);
        if (c[0] < o[1] + GAP && c[1] + GAP > o[0]) return false;
      }
      return true;
    }

    // The first and last labels anchor the axis, so reserve the last one too — but
    // drop it rather than let it collide with the first on a very narrow chart.
    if (n > 1 && clears(n - 1)) keep[n - 1] = true;
    for (var step = 1; step < n - 1; step++) {
      if (clears(step)) keep[step] = true;
    }

    Object.keys(keep).map(Number).sort(function (a, b) { return a - b; }).forEach(function (i) {
      var p = placed[i];
      s.push('<text class="v-tick v-tick-x" text-anchor="' + anchorOf(i) + '" x="' + p.x.toFixed(1) +
        '" y="' + (f.m.top + f.ih + 18) + '">' + esc(xLabels[i]) + '</text>');
    });
    return s.join('');
  }

  function yPos(f, scale, v) {
    return f.m.top + f.ih - ((v - scale.lo) / (scale.hi - scale.lo)) * f.ih;
  }
  function xPos(f, i, n, band) {
    if (band) return f.m.left + (i + 0.5) * (f.iw / n);
    return n === 1 ? f.m.left + f.iw / 2 : f.m.left + (i / (n - 1)) * f.iw;
  }

  // ---------------------------------------------------------------------------
  // Charts
  //
  // One y-scale per chart, always. Two measures of different magnitude become two
  // charts sharing an x-axis (small multiples) — never a second y-axis.
  // ---------------------------------------------------------------------------

  var CHART_W = 760, CHART_H = 260;

  /** Multi-series line chart with a crosshair hover layer. */
  function lineChart(o) {
    requireProvenance(o.provenance, o.label);
    var f = frame(o.width || CHART_W, o.height || CHART_H, o.margin);
    var series = o.series.filter(function (s) {
      return s.values.some(function (v) { return v != null && isFinite(v); });
    });
    var all = [];
    series.forEach(function (s) { s.values.forEach(function (v) { if (v != null && isFinite(v)) all.push(v); }); });
    if (!all.length) return { svg: '', empty: true };

    var lo = Math.min.apply(null, all), hi = Math.max.apply(null, all);
    if (o.zeroBased !== false) lo = Math.min(0, lo);
    var scale = niceTicks(lo, hi, 5);
    var n = o.labels.length;
    var s = ['<svg class="v-svg" viewBox="0 0 ' + f.w + ' ' + f.h + '" role="img" aria-label="' + esc(o.label) + '">'];
    s.push(axes(f, scale, o.labels, { yFormat: o.yFormat }));

    series.forEach(function (ser, si) {
      var d = '', started = false;
      ser.values.forEach(function (v, i) {
        if (v == null || !isFinite(v)) { started = false; return; }
        var x = xPos(f, i, n).toFixed(1), y = yPos(f, scale, v).toFixed(1);
        d += (started ? 'L' : 'M') + x + ' ' + y + ' ';
        started = true;
      });
      s.push('<path class="v-line" d="' + d + '" stroke="' + esc(ser.color) + '"' +
        (ser.dashed ? ' stroke-dasharray="6 5"' : '') + '/>');
      // Markers get a 2px surface ring so overlapping points stay separable.
      ser.values.forEach(function (v, i) {
        if (v == null || !isFinite(v)) return;
        s.push('<circle class="v-dot" cx="' + xPos(f, i, n).toFixed(1) + '" cy="' + yPos(f, scale, v).toFixed(1) +
          '" r="4.5" fill="' + esc(ser.color) + '"/>');
      });
    });

    s.push(hoverLayer(f, n, o.labels, series, o, false));
    s.push('</svg>');
    return { svg: s.join(''), empty: false, series: series };
  }

  /** Stacked area. A 2px surface-colored stroke separates adjacent bands. */
  function stackedArea(o) {
    requireProvenance(o.provenance, o.label);
    var f = frame(o.width || CHART_W, o.height || CHART_H, o.margin);
    var n = o.labels.length;
    if (!n || !o.series.length) return { svg: '', empty: true };

    var totals = [];
    for (var i = 0; i < n; i++) {
      totals[i] = o.series.reduce(function (t, ser) { return t + (ser.values[i] || 0); }, 0);
    }
    var hi = Math.max.apply(null, totals.concat([0]));
    if (hi <= 0) return { svg: '', empty: true };
    var scale = niceTicks(0, hi, 5);
    var s = ['<svg class="v-svg" viewBox="0 0 ' + f.w + ' ' + f.h + '" role="img" aria-label="' + esc(o.label) + '">'];
    s.push(axes(f, scale, o.labels, { yFormat: o.yFormat }));

    var base = new Array(n).fill(0);
    o.series.forEach(function (ser) {
      var top = base.map(function (b, i) { return b + (ser.values[i] || 0); });
      var d = '';
      top.forEach(function (v, i) { d += (i ? 'L' : 'M') + xPos(f, i, n).toFixed(1) + ' ' + yPos(f, scale, v).toFixed(1) + ' '; });
      for (var k = n - 1; k >= 0; k--) d += 'L' + xPos(f, k, n).toFixed(1) + ' ' + yPos(f, scale, base[k]).toFixed(1) + ' ';
      d += 'Z';
      s.push('<path class="v-band" d="' + d + '" fill="' + esc(ser.color) + '"/>');
      base = top;
    });

    s.push(hoverLayer(f, n, o.labels, o.series, o, true));
    s.push('</svg>');
    return { svg: s.join(''), empty: false };
  }

  /** Vertical bars with 4px rounded tops anchored to the baseline. */
  function barChart(o) {
    requireProvenance(o.provenance, o.label);
    var f = frame(o.width || CHART_W, o.height || (o.compact ? 150 : CHART_H), o.margin);
    var n = o.labels.length;
    var vals = o.values.filter(function (v) { return v != null && isFinite(v); });
    if (!vals.length) return { svg: '', empty: true };
    var scale = niceTicks(Math.min(0, Math.min.apply(null, vals)), Math.max.apply(null, vals), 4);
    var bandW = f.iw / n;
    var barW = Math.max(3, Math.min(46, bandW - 10));
    var s = ['<svg class="v-svg" viewBox="0 0 ' + f.w + ' ' + f.h + '" role="img" aria-label="' + esc(o.label) + '">'];
    s.push(axes(f, scale, o.labels, { yFormat: o.yFormat, band: true }));

    var zeroY = yPos(f, scale, 0);
    o.values.forEach(function (v, i) {
      if (v == null || !isFinite(v)) return;
      var x = f.m.left + i * bandW + (bandW - barW) / 2;
      var y = yPos(f, scale, v);
      var top = Math.min(y, zeroY), hgt = Math.abs(zeroY - y);
      var r = Math.min(4, hgt, barW / 2);
      var color = o.colorFor ? o.colorFor(i, v) : o.color;
      var d = v >= 0
        ? 'M' + x + ' ' + (top + hgt) + ' L' + x + ' ' + (top + r) + ' Q' + x + ' ' + top + ' ' + (x + r) + ' ' + top +
          ' L' + (x + barW - r) + ' ' + top + ' Q' + (x + barW) + ' ' + top + ' ' + (x + barW) + ' ' + (top + r) +
          ' L' + (x + barW) + ' ' + (top + hgt) + ' Z'
        : 'M' + x + ' ' + top + ' L' + x + ' ' + (top + hgt - r) + ' Q' + x + ' ' + (top + hgt) + ' ' + (x + r) + ' ' + (top + hgt) +
          ' L' + (x + barW - r) + ' ' + (top + hgt) + ' Q' + (x + barW) + ' ' + (top + hgt) + ' ' + (x + barW) + ' ' + (top + hgt - r) +
          ' L' + (x + barW) + ' ' + top + ' Z';
      s.push('<path class="v-bar" d="' + d + '" fill="' + esc(color) + '"' +
        ' data-i="' + i + '" data-tip="' + esc(o.labels[i] + ' · ' + (o.tipFormat ? o.tipFormat(v, i) : fmt.number(v))) + '"/>');
    });
    s.push('</svg>');
    return { svg: s.join(''), empty: false };
  }

  /** Scatter — used for the P2P panel, where every send is plotted individually. */
  function scatterChart(o) {
    requireProvenance(o.provenance, o.label);
    var f = frame(o.width || CHART_W, o.height || 240, o.margin);
    var pts = o.points.filter(function (p) { return p.y != null && isFinite(p.y); });
    if (!pts.length) return { svg: '', empty: true };
    var ys = pts.map(function (p) { return p.y; });
    var scale = niceTicks(Math.min(0, Math.min.apply(null, ys)), Math.max.apply(null, ys), 4);
    var n = o.labels.length;
    var s = ['<svg class="v-svg" viewBox="0 0 ' + f.w + ' ' + f.h + '" role="img" aria-label="' + esc(o.label) + '">'];
    s.push(axes(f, scale, o.labels, { yFormat: o.yFormat, band: true }));
    pts.forEach(function (p) {
      s.push('<circle class="v-pt" cx="' + xPos(f, p.i, n, true).toFixed(1) + '" cy="' + yPos(f, scale, p.y).toFixed(1) +
        '" r="5.5" fill="' + esc(p.color || o.color) + '" data-tip="' + esc(p.tip || '') + '"/>');
    });
    s.push('</svg>');
    return { svg: s.join(''), empty: false };
  }

  /**
   * Shared crosshair hover layer for line and area charts: one invisible column
   * per x position, with a hit target wider than the mark.
   */
  function hoverLayer(f, n, labels, series, o, stacked) {
    var s = ['<g class="v-hover">'];
    var bandW = f.iw / Math.max(1, n - (stacked ? 1 : 1));
    for (var i = 0; i < n; i++) {
      var cx = xPos(f, i, n);
      var rows = series.map(function (ser) {
        var v = ser.values[i];
        return { name: ser.name, color: ser.color, text: v == null || !isFinite(v) ? '—' : (o.tipFormat ? o.tipFormat(v) : fmt.number(v)) };
      });
      var payload = JSON.stringify({ title: labels[i], rows: rows });
      s.push('<rect class="v-hit" x="' + (cx - bandW / 2).toFixed(1) + '" y="' + f.m.top + '" width="' + bandW.toFixed(1) +
        '" height="' + f.ih + '" data-x="' + cx.toFixed(1) + '" data-payload="' + esc(payload) + '"/>');
    }
    s.push('<line class="v-cross" x1="0" x2="0" y1="' + f.m.top + '" y2="' + (f.m.top + f.ih) + '" style="display:none"/>');
    s.push('</g>');
    return s.join('');
  }

  function legend(series) {
    if (!series || series.length < 2) return '';   // one series needs no legend — the title names it
    return '<div class="v-legend">' + series.map(function (s) {
      return '<span class="v-key"><i style="background:' + esc(s.color) + '"></i>' + esc(s.name) + '</span>';
    }).join('') + '</div>';
  }

  return {
    requireProvenance: requireProvenance,
    bubbleHTML: bubbleHTML,
    esc: esc, fmt: fmt, niceTicks: niceTicks,
    lineChart: lineChart, stackedArea: stackedArea, barChart: barChart, scatterChart: scatterChart,
    legend: legend
  };
});
