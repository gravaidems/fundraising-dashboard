/*
 * CABApp — drop zone, filters, panels, and the Data Sources report.
 *
 * This is the only layer that touches the DOM. CABParse and CABMetrics stay
 * DOM-free so they port to React unchanged.
 */
(function () {
  'use strict';

  var V = window.CABViz;
  var esc = V.esc, fmt = V.fmt;

  // Chart colors are CSS custom properties, so the light/dark swap is free —
  // inline SVG resolves var() against the document just like any other element.
  var SERIES = ['var(--series-1)', 'var(--series-2)', 'var(--series-3)', 'var(--series-4)',
                'var(--series-5)', 'var(--series-6)', 'var(--series-7)', 'var(--series-8)'];

  var state = {
    data: null, metrics: null, fileName: null,
    filters: { from: null, to: null, ask: [], subtype: [], audience: [], sender: [] },
    breakdownDim: 'audience',
    tab: 'overview',
    scenario: null,
    bubbles: {},      // id -> provenance, populated fresh on every render
    bubbleSeq: 0
  };

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }

  // ---------------------------------------------------------------------------
  // Bubble registry — every ⓘ button resolves through here.
  // ---------------------------------------------------------------------------

  function bubble(prov, label) {
    V.requireProvenance(prov, label);   // throws before anything renders
    var id = 'b' + (++state.bubbleSeq);
    state.bubbles[id] = { prov: prov, label: label };
    return '<button class="info" data-bubble="' + id + '" aria-label="How &quot;' + esc(label) +
      '&quot; is calculated" title="Hover for the source; click to keep it open">i</button>';
  }

  /**
   * Bubbles open on hover and close when the pointer leaves both the button and the
   * popover. Click pins one open so a long bubble can be scrolled and read, and
   * keyboard focus opens it too — hover alone would strand keyboard users.
   */
  var HOVER_IN = 90, HOVER_OUT = 220;
  var hoverTimer = null, closeTimer = null;

  function positionBubble(btn, pop) {
    var r = btn.getBoundingClientRect();
    var vw = document.documentElement.clientWidth;
    var vh = document.documentElement.clientHeight;
    var w = pop.offsetWidth, hgt = pop.offsetHeight;

    var left = r.left + window.scrollX - 12;
    if (left + w > window.scrollX + vw - 12) left = window.scrollX + vw - w - 12;
    if (left < window.scrollX + 8) left = window.scrollX + 8;

    // Flip above the button when there is not room below it.
    var below = r.bottom + window.scrollY + 8;
    var top = below;
    if (r.bottom + hgt + 16 > vh && r.top - hgt - 8 > 0) top = r.top + window.scrollY - hgt - 8;

    pop.style.top = top + 'px';
    pop.style.left = left + 'px';
  }

  function openBubble(btn, pinned) {
    if (state.openBtn === btn && $('.prov-pop')) {
      if (pinned) { state.pinned = true; $('.prov-pop').classList.add('prov-pinned'); }
      return;
    }
    closeBubble(true);
    var rec = state.bubbles[btn.getAttribute('data-bubble')];
    if (!rec) return;
    var pop = document.createElement('div');
    pop.className = 'prov-pop' + (pinned ? ' prov-pinned' : '');
    pop.innerHTML = V.bubbleHTML(rec.prov, rec.label);
    if (pinned) pop.innerHTML += '<button class="prov-close" aria-label="Close">Close</button>';
    document.body.appendChild(pop);

    // Keep an unpinned bubble alive while the pointer is inside it.
    pop.addEventListener('mouseenter', function () { clearTimeout(closeTimer); });
    pop.addEventListener('mouseleave', function () { if (!state.pinned) scheduleClose(); });

    positionBubble(btn, pop);
    btn.classList.add('info-on');
    state.openBtn = btn;
    state.pinned = !!pinned;
  }

  function closeBubble(force) {
    if (state.pinned && !force) return;
    clearTimeout(hoverTimer); clearTimeout(closeTimer);
    var p = $('.prov-pop');
    if (p) p.parentNode.removeChild(p);
    if (state.openBtn) { state.openBtn.classList.remove('info-on'); state.openBtn = null; }
    state.pinned = false;
  }

  function scheduleClose() {
    clearTimeout(closeTimer);
    closeTimer = setTimeout(function () { closeBubble(); }, HOVER_OUT);
  }

  // ---------------------------------------------------------------------------
  // Panel scaffolding
  // ---------------------------------------------------------------------------

  function panel(o) {
    var h = ['<section class="panel' + (o.wide ? ' panel-wide' : '') + '">'];
    h.push('<header class="panel-hd"><div>');
    h.push('<h2>' + esc(o.title) + (o.prov ? bubble(o.prov, o.title) : '') + '</h2>');
    if (o.subtitle) h.push('<p class="panel-sub">' + esc(o.subtitle) + '</p>');
    h.push('</div>');
    if (o.controls) h.push('<div class="panel-ctl">' + o.controls + '</div>');
    h.push('</header>');
    (o.caveats || []).forEach(function (c) {
      h.push('<p class="caveat"><span class="caveat-i" aria-hidden="true">!</span>' + esc(c) + '</p>');
    });
    h.push('<div class="panel-body">' + o.body + '</div>');
    h.push('</section>');
    return h.join('');
  }

  function emptyState(title, detail) {
    return '<div class="empty"><strong>' + esc(title) + '</strong><p>' + esc(detail) + '</p></div>';
  }

  // ---------------------------------------------------------------------------
  // Filters
  // ---------------------------------------------------------------------------

  function isoDate(d) { return d ? d.toISOString().slice(0, 10) : ''; }

  function renderFilters() {
    var facets = state.metrics.facets();
    var f = state.filters;
    var h = ['<div class="filters">'];
    h.push('<div class="filter"><label for="f-from">From</label><input type="date" id="f-from" value="' + isoDate(f.from) + '"></div>');
    h.push('<div class="filter"><label for="f-to">To</label><input type="date" id="f-to" value="' + isoDate(f.to) + '"></div>');
    h.push('<div class="filter filter-presets">' +
      '<button class="chip" data-range="6">Last 6 months</button>' +
      '<button class="chip" data-range="3">Last 3 months</button>' +
      '<button class="chip" data-range="all">All data</button></div>');

    [['ask', 'Ask'], ['audience', 'Audience'], ['subtype', 'Subtype'], ['sender', 'Sender']].forEach(function (pair) {
      var key = pair[0], label = pair[1], vals = facets[key] || [];
      if (!vals.length) return;
      var sel = f[key];
      h.push('<div class="filter"><label for="f-' + key + '">' + esc(label) + '</label>' +
        '<select id="f-' + key + '" class="f-multi" data-key="' + key + '" multiple size="1">');
      vals.forEach(function (v) {
        h.push('<option value="' + esc(v) + '"' + (sel.indexOf(v) !== -1 ? ' selected' : '') + '>' + esc(v) + '</option>');
      });
      h.push('</select><span class="f-count">' + (sel.length ? sel.length + ' selected' : 'all') + '</span></div>');
    });

    // The Client filter appears only when there is more than one client to choose
    // between — a one-item dropdown is noise.
    if ((facets.clients || []).length > 1) {
      h.push('<div class="filter"><label>Client</label><span class="f-note">' + esc(facets.clients.join(', ')) + '</span></div>');
    }
    h.push('<button class="chip chip-reset" id="f-reset">Reset filters</button>');
    h.push('</div>');
    return h.join('');
  }

  // ---------------------------------------------------------------------------
  // KPI row
  // ---------------------------------------------------------------------------

  function renderKPIs() {
    var kpis = state.metrics.kpis(state.filters);
    var h = ['<div class="kpis">'];
    kpis.forEach(function (k) {
      var m = k.metric;
      var txt = m.value == null ? '—' : fmt.auto(m.value, m.format);
      h.push('<div class="kpi">');
      h.push('<div class="kpi-lb">' + esc(k.label) + bubble(m.provenance, k.label) + '</div>');
      h.push('<div class="kpi-v' + (m.value == null ? ' kpi-v-empty' : '') + '">' + esc(txt) + '</div>');
      if (k.id === 'K5' && m.goalTotal) {
        h.push('<div class="kpi-sub">' + esc(fmt.compact(m.actualTotal)) + ' of ' + esc(fmt.compact(m.goalTotal)) + ' goal</div>');
      } else if (m.emptyReason) {
        h.push('<div class="kpi-sub">' + esc(m.emptyReason) + '</div>');
      } else {
        h.push('<div class="kpi-sub">' + esc(fmt.number(m.provenance.rowsUsed)) + ' rows</div>');
      }
      h.push('</div>');
    });
    h.push('</div>');
    return h.join('');
  }

  // ---------------------------------------------------------------------------
  // Panel 1 — actual vs goal. Three series, three separate bubbles.
  // ---------------------------------------------------------------------------

  function renderActualVsGoal() {
    var p = state.metrics.panelActualVsGoal(state.filters, state.scenario);
    var series = [
      { name: 'Actual (computed from sends)', color: SERIES[0], values: p.series.actual },
      { name: 'Goal (Digital Goals tab)', color: SERIES[1], values: p.series.goal, dashed: true }
    ];
    if (p.scenario && p.scenario.usable) {
      series.push({ name: 'Projection — ' + p.scenario.scenario.replace(/^Digital Projections\s*/, ''),
                    color: SERIES[2], values: p.series.projection, dashed: true });
    }

    var chart = V.lineChart({
      label: 'Raised vs goal by month', labels: p.labels, series: series,
      provenance: p.provenanceActual, yFormat: fmt.compact, tipFormat: fmt.currency
    });

    var opts = p.scenarios.map(function (s) {
      var nm = s.name.replace(/^Digital Projections\s*/, '');
      return '<option value="' + esc(s.name) + '"' + (state.scenario === s.name ? ' selected' : '') +
        (s.usable ? '' : ' disabled') + '>' + esc(nm) + (s.usable ? '' : ' — unusable') + '</option>';
    }).join('');
    var controls = '<label class="sel-lb" for="p1-scenario">Projection</label>' +
      '<select id="p1-scenario"><option value="">None</option>' + opts + '</select>';

    var body = [];
    body.push('<div class="chart">' + chart.svg + '</div>');
    body.push(V.legend(series));

    body.push('<div class="srcs">');
    body.push('<span class="src"><i style="background:' + SERIES[0] + '"></i>Actual' + bubble(p.provenanceActual, 'Actual (computed from send-level data)') + '</span>');
    body.push('<span class="src"><i style="background:' + SERIES[1] + '"></i>Goal' + bubble(p.provenanceGoal, 'Goal (April 2026 Updated Goals)') + '</span>');
    if (p.provenanceProjection) {
      body.push('<span class="src"><i style="background:' + SERIES[2] + '"></i>Projection' + bubble(p.provenanceProjection, 'Projection (' + p.scenario.scenario + ')') + '</span>');
    }
    body.push('</div>');

    // Table view — this is also the relief for the light-mode contrast warning.
    body.push('<table class="tbl"><thead><tr><th>Month</th><th class="num">Actual</th><th class="num">Goal</th><th class="num">Variance</th></tr></thead><tbody>');
    p.labels.forEach(function (lb, i) {
      var a = p.series.actual[i], g = p.series.goal[i], v = p.variance[i];
      body.push('<tr><td>' + esc(lb) + '</td><td class="num">' + esc(fmt.currency(a)) + '</td>' +
        '<td class="num">' + esc(g == null ? '—' : fmt.currency(g)) + '</td>' +
        '<td class="num ' + (v == null ? '' : v < 0 ? 'neg' : 'pos') + '">' + esc(v == null ? '—' : fmt.currency(v)) + '</td></tr>');
    });
    body.push('</tbody></table>');

    var caveats = [
      'The "actual" line is computed by this dashboard from send-level rows in Email Statistics and P2P Statistics. Neither monthly tab in this workbook contains actuals.',
      'It therefore covers Email and SMS P2P only, while the goal line covers all nine channels including Recurring, Ads, Website and Social. The two lines are not like-for-like, and the gap between them is wider than true performance.'
    ];
    var unusable = p.scenarios.filter(function (s) { return !s.usable; });
    if (unusable.length) {
      caveats.push('Digital Projections holds three scenario blocks. ' + unusable.length + ' of them cannot be plotted: ' +
        unusable.map(function (s) { return '"' + s.name.replace(/^Digital Projections\s*/, '') + '" — ' + s.reason; }).join('; ') + '.');
    }

    return panel({
      title: 'Raised vs goal by month', wide: true, controls: controls, caveats: caveats,
      subtitle: 'Computed actuals against the goals tab, with an optional projection scenario.',
      body: body.join('')
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 2 — channel mix
  // ---------------------------------------------------------------------------

  function renderChannelMix() {
    var p = state.metrics.panelChannelMix(state.filters);
    if (!p.channels.length) {
      return panel({ title: 'Channel mix over time', prov: p.provenance,
        body: emptyState('No channel values in range', 'The goals tab has no non-zero channel figures for the selected months.') });
    }
    var series = p.channels.map(function (c, i) { return { name: c.label, color: SERIES[i % SERIES.length], values: c.values }; });
    var chart = V.stackedArea({
      label: 'Channel mix over time', labels: p.labels, series: series,
      provenance: p.provenance, yFormat: fmt.compact, tipFormat: fmt.currency
    });
    var body = '<div class="chart">' + chart.svg + '</div>' + V.legend(series);
    if (p.emptyChannels.length) {
      body += '<p class="foot">Not shown (zero in every month in range): ' + esc(p.emptyChannels.join(', ')) + '.</p>';
    }
    return panel({
      title: 'Channel mix over time', prov: p.provenance, wide: true,
      subtitle: 'Goal mix by channel, from the Digital Goals tab.',
      caveats: ['This is the planned mix, not the achieved mix. The workbook carries no per-channel actuals.'],
      body: body
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 3 — email performance trend.
  // Revenue-per-1k and the three rates differ by two orders of magnitude, so they
  // are small multiples with independent scales, never a shared or second y-axis.
  // ---------------------------------------------------------------------------

  function renderEmailTrend() {
    var p = state.metrics.panelEmailTrend(state.filters);
    if (!p.labels.length) {
      return panel({ title: 'Email performance trend', prov: p.revenuePerK.provenance,
        body: emptyState('No email sends in range', 'Widen the date range or clear a filter.') });
    }
    var main = V.lineChart({
      label: 'Revenue per 1,000 recipients', labels: p.labels,
      series: [{ name: 'Revenue per 1,000 recipients', color: SERIES[0], values: p.revenuePerK.values }],
      provenance: p.revenuePerK.provenance,
      yFormat: function (v) { return '$' + Math.round(v); }, tipFormat: fmt.currencyPrecise
    });

    var body = ['<h3 class="sub-h">Revenue per 1,000 recipients' + bubble(p.revenuePerK.provenance, 'Revenue per 1,000 recipients') + '</h3>'];
    body.push('<div class="chart">' + main.svg + '</div>');

    body.push('<h3 class="sub-h">Engagement rates' + bubble(p.rates.provenance, 'Open / click / donate rates') + '</h3>');
    body.push('<p class="foot">Three separate scales: open rates run near 30%, donate rates near 0.1%. Plotted on one axis the smaller two would flatline, so each gets its own chart.</p>');
    body.push('<div class="smalls">');
    [['Open rate', p.rates.open, SERIES[0], 1], ['Click rate', p.rates.click, SERIES[2], 2], ['Donate rate', p.rates.donate, SERIES[4], 3]]
      .forEach(function (t) {
        if (p.missingColumns.length && t[0] === 'Open rate' && p.missingColumns.indexOf('openRate') !== -1) {
          body.push('<div class="small">' + emptyState(t[0] + ' unavailable', 'The Open Rate column was not found in this workbook.') + '</div>');
          return;
        }
        var c = V.lineChart({
          label: t[0], labels: p.labels, width: 380, height: 180,
          margin: { top: 14, right: 14, bottom: 30, left: 52 },
          series: [{ name: t[0], color: t[2], values: t[1] }],
          provenance: p.rates.provenance,
          yFormat: function (v) { return (v * 100).toFixed(v < 0.02 ? 2 : 1) + '%'; },
          tipFormat: function (v) { return fmt.percent(v, 2); }
        });
        body.push('<div class="small"><h4>' + esc(t[0]) + '</h4>' + (c.empty ? emptyState('No data', 'No values in range.') : c.svg) + '</div>');
      });
    body.push('</div>');

    return panel({
      title: 'Email performance trend', wide: true,
      subtitle: 'From Email Statistics — the densest and cleanest data in this workbook.',
      body: body.join('')
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 4 — list fatigue. Cadence and unsub rate share an x-axis as stacked
  // small multiples rather than sharing one plot with two y-scales.
  // ---------------------------------------------------------------------------

  function renderListFatigue() {
    var p = state.metrics.panelListFatigue(state.filters);
    if (!p.weeks.length) {
      return panel({ title: 'List fatigue', prov: p.provenance,
        body: emptyState('No sends in range', 'Widen the date range or clear a filter.') });
    }
    var bars = V.barChart({
      label: 'Sends per week', labels: p.labels, values: p.sends, color: SERIES[0],
      provenance: p.provenance, height: 150, compact: true,
      yFormat: function (v) { return String(Math.round(v)); },
      tipFormat: function (v) { return v + ' send' + (v === 1 ? '' : 's'); }
    });
    var line = V.lineChart({
      label: 'Unsub rate by week', labels: p.labels, height: 170,
      series: [{ name: 'Unsub rate', color: SERIES[7], values: p.unsubRate }],
      provenance: p.provenance,
      yFormat: function (v) { return (v * 100).toFixed(2) + '%'; },
      tipFormat: function (v) { return fmt.percent(v, 3); }
    });
    var body = '<h3 class="sub-h">Sends per week</h3><div class="chart">' + bars.svg + '</div>' +
      '<h3 class="sub-h">Unsub rate, same weeks</h3><div class="chart">' + line.svg + '</div>' +
      '<p class="foot">Two charts sharing an x-axis rather than one chart with two y-axes — a count and a rate on the same scale would misrepresent both.</p>';
    return panel({
      title: 'List fatigue', prov: p.provenance, wide: true,
      subtitle: 'Send cadence against unsubscribe rate, by ISO week.',
      body: body
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 5 — email breakdown. Every group row carries its own bubble.
  // ---------------------------------------------------------------------------

  function renderBreakdown() {
    var dim = state.breakdownDim;
    var p = state.metrics.panelEmailBreakdown(state.filters, dim);
    var tabs = [['audience', 'Audience'], ['ask', 'Ask'], ['subtype', 'Subtype'], ['sender', 'Sender']]
      .map(function (t) {
        return '<button class="tab' + (dim === t[0] ? ' tab-on' : '') + '" data-dim="' + t[0] + '">' + esc(t[1]) + '</button>';
      }).join('');

    var body = ['<div class="tabs">' + tabs + '</div>'];
    if (!p.rows.length) {
      body.push(emptyState('No sends in range', 'Widen the date range or clear a filter.'));
    } else {
      body.push('<table class="tbl tbl-sort"><thead><tr>');
      body.push('<th data-sort="key">' + esc(p.dimLabel) + '</th>');
      [['sends', 'Sends'], ['recipients', 'Recipients'], ['raised', 'Raised'], ['revPerK', 'Rev / 1k'],
       ['openRate', 'Open'], ['clickRate', 'Click'], ['donateRate', 'Donate'], ['unsubRate', 'Unsub']]
        .forEach(function (c) { body.push('<th class="num" data-sort="' + c[0] + '">' + esc(c[1]) + '</th>'); });
      body.push('<th class="src-col">Source</th></tr></thead><tbody>');
      p.rows.forEach(function (r) {
        body.push('<tr>');
        body.push('<td>' + esc(r.key) + '</td>');
        body.push('<td class="num">' + esc(fmt.number(r.sends)) + '</td>');
        body.push('<td class="num">' + esc(fmt.number(r.recipients)) + '</td>');
        body.push('<td class="num">' + esc(fmt.currency(r.raised)) + '</td>');
        body.push('<td class="num">' + esc(r.revPerK == null ? '—' : fmt.currencyPrecise(r.revPerK)) + '</td>');
        body.push('<td class="num">' + esc(fmt.percent(r.openRate)) + '</td>');
        body.push('<td class="num">' + esc(fmt.percent(r.clickRate, 2)) + '</td>');
        body.push('<td class="num">' + esc(fmt.percent(r.donateRate, 3)) + '</td>');
        body.push('<td class="num">' + esc(fmt.percent(r.unsubRate, 3)) + '</td>');
        body.push('<td class="src-col">' + bubble(r.provenance, p.dimLabel + ' = ' + r.key) + '</td>');
        body.push('</tr>');
      });
      body.push('</tbody></table>');
    }

    var caveats = dim === 'sender'
      ? ['Sender values appear exactly as written in the workbook. Near-duplicates such as "Booker HQ", "Cory Booker HQ" and "Booker HQ Finance Team" are shown separately and never merged automatically — they may be genuinely different senders.']
      : [];

    return panel({
      title: 'Email breakdown', prov: p.provenance, wide: true, caveats: caveats,
      subtitle: 'Performance by ' + p.dimLabel.toLowerCase() + '. Every row has its own source bubble.',
      body: body.join('')
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 6 — Ads. Empty state that explains itself and lights up on its own.
  // ---------------------------------------------------------------------------

  function renderAds() {
    var p = state.metrics.panelAds(state.filters);
    if (!p.hasRevenue) {
      var body = emptyState('No ads revenue data in source', p.emptyReason);
      body += '<table class="tbl"><thead><tr><th>Month</th><th>Placement</th><th>Goal</th>' +
        '<th class="num">Gross Spend</th><th class="num">NTL</th><th class="num">Gross CPA</th>' +
        '<th class="num">Lifetime Raised</th><th class="num">Lifetime ROAS</th></tr></thead><tbody>';
      p.rows.forEach(function (r) {
        body += '<tr><td>' + esc(r.month) + '</td><td>' + esc(r.placement) + '</td><td>' + esc(r.goal) + '</td>' +
          '<td class="num">' + esc(fmt.currencyPrecise(r.grossSpend)) + '</td>' +
          '<td class="num">' + esc(fmt.number(r.ntl)) + '</td>' +
          '<td class="num">' + esc(fmt.currencyPrecise(r.grossCpa)) + '</td>' +
          '<td class="num empty-cell">blank</td><td class="num empty-cell">blank</td></tr>';
      });
      body += '</tbody></table>';
      return panel({
        title: 'Ads CPA & ROAS by placement', prov: p.provenance,
        subtitle: 'Standard toplines vs finance-adjusted.',
        body: body
      });
    }

    var placements = {}, order = [];
    p.rows.forEach(function (r) { if (!placements[r.placement]) { placements[r.placement] = r; order.push(r.placement); } });
    var std = V.barChart({
      label: 'Lifetime ROAS by placement', labels: order,
      values: order.map(function (k) { return placements[k].roasStandard; }),
      color: SERIES[0], provenance: p.provenance, tipFormat: fmt.ratio,
      yFormat: function (v) { return v.toFixed(1) + '×'; }
    });
    var adj = V.barChart({
      label: 'Finance-adjusted ROAS by placement', labels: order,
      values: order.map(function (k) { return placements[k].roasFinanceAdj; }),
      color: SERIES[1], provenance: p.provenance, tipFormat: fmt.ratio,
      yFormat: function (v) { return v.toFixed(1) + '×'; }
    });
    return panel({
      title: 'Ads CPA & ROAS by placement', prov: p.provenance, wide: true,
      subtitle: 'Standard toplines vs finance-adjusted, side by side.',
      body: '<div class="smalls"><div class="small"><h4>Standard toplines</h4>' + std.svg + '</div>' +
            '<div class="small"><h4>Finance-adjusted</h4>' + adj.svg + '</div></div>'
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 7 — P2P. Small sample stated on the panel face, not just the bubble.
  // ---------------------------------------------------------------------------

  function renderP2P() {
    var p = state.metrics.panelP2P(state.filters);
    if (!p.rows.length) {
      return panel({ title: 'P2P performance', prov: p.provenance,
        body: emptyState('No P2P sends in range', 'All ' + p.totalAvailable + ' P2P sends fall outside the selected dates.') });
    }
    var labels = p.rows.map(function (r) { return fmt.date(r.date).replace(/, \d{4}$/, ''); });
    var pts = [], ltv = [];
    p.rows.forEach(function (r, i) {
      pts.push({ i: i, y: r.immediateRoas, color: SERIES[0],
        tip: labels[i] + ' · Immediate ROAS ' + fmt.ratio(r.immediateRoas) + ' · row ' + r._row });
      if (r.ltvRoas != null) ltv.push({ i: i, y: r.ltvRoas, color: SERIES[1],
        tip: labels[i] + ' · LTV ROAS ' + fmt.ratio(r.ltvRoas) + ' · row ' + r._row });
    });
    var roas = V.scatterChart({
      label: 'Immediate vs LTV ROAS by send', labels: labels, points: pts.concat(ltv),
      provenance: p.provenance, yFormat: function (v) { return v.toFixed(1) + '×'; }
    });
    var cpa = V.barChart({
      label: 'Gross CPA by send', labels: labels, values: p.rows.map(function (r) { return r.grossCpa; }),
      color: SERIES[2], provenance: p.provenance, height: 170,
      yFormat: function (v) { return '$' + Math.round(v); }, tipFormat: fmt.currencyPrecise
    });

    var body = ['<h3 class="sub-h">ROAS by send</h3><div class="chart">' + roas.svg + '</div>'];
    body.push(V.legend([{ name: 'Immediate ROAS', color: SERIES[0] }, { name: 'LTV ROAS from NTL', color: SERIES[1] }]));
    body.push('<h3 class="sub-h">Gross CPA by send</h3><div class="chart">' + cpa.svg + '</div>');
    body.push('<table class="tbl"><thead><tr><th>Date</th><th>Topic</th><th>Audience</th>' +
      '<th class="num">Recipients</th><th class="num">Spend</th><th class="num">Raised</th>' +
      '<th class="num">Imm. ROAS</th><th class="num">LTV ROAS</th><th class="num">NTL</th></tr></thead><tbody>');
    p.rows.forEach(function (r) {
      body.push('<tr><td>' + esc(fmt.date(r.date)) + '</td><td>' + esc(r.topic || '—') + '</td><td>' + esc(r.audience || '—') + '</td>' +
        '<td class="num">' + esc(fmt.number(r.recipients)) + '</td>' +
        '<td class="num">' + esc(fmt.currency(r.grossSpend)) + '</td>' +
        '<td class="num">' + esc(fmt.currency(r.immediateRaised)) + '</td>' +
        '<td class="num">' + esc(fmt.ratio(r.immediateRoas)) + '</td>' +
        '<td class="num">' + esc(fmt.ratio(r.ltvRoas)) + '</td>' +
        '<td class="num">' + esc(fmt.number(r.ntl)) + '</td></tr>');
    });
    body.push('</tbody></table>');

    return panel({
      title: 'P2P performance', prov: p.provenance, wide: true,
      subtitle: 'Immediate and lifetime ROAS, CPA and new-to-list, per send.',
      caveats: ['Small sample: this workbook holds ' + p.totalAvailable + ' P2P sends in total' +
        (p.n !== p.totalAvailable ? ' (' + p.n + ' in the current range)' : '') +
        '. Individual sends move the averages a lot — read this as indicative, not conclusive.'],
      body: body.join('')
    });
  }

  // ---------------------------------------------------------------------------
  // Panel 8 — high-dollar aggregates. Aggregates only, by construction.
  // ---------------------------------------------------------------------------

  function renderHighDollar() {
    var p = state.metrics.panelHighDollar();
    if (!p.available) {
      return panel({ title: 'High-dollar donations', prov: p.provenance, body: emptyState('Not available', p.emptyReason) });
    }
    var bands = V.barChart({
      label: 'Donations by amount band', labels: p.byBand.map(function (b) { return b.key; }),
      values: p.byBand.map(function (b) { return b.sum; }), color: SERIES[0],
      provenance: p.provenance, height: 200,
      yFormat: fmt.compact,
      tipFormat: function (v, i) { return fmt.currency(v) + ' from ' + p.byBand[i].count + ' gift' + (p.byBand[i].count === 1 ? '' : 's'); }
    });

    var body = ['<p class="pii-note"><strong>Aggregates only.</strong> No individual donor data appears in this dashboard. ' +
      'Donor name, email, receipt ID and fundraising page are discarded while the file is being read and never enter the page.</p>'];
    body.push('<div class="hd-top"><div class="hd-fig"><span class="hd-n">' + esc(fmt.number(p.total.count)) +
      '</span><span class="hd-l">donations</span></div><div class="hd-fig"><span class="hd-n">' +
      esc(fmt.currency(p.total.sum)) + '</span><span class="hd-l">total</span></div></div>');
    body.push('<h3 class="sub-h">By amount band</h3><div class="chart">' + bands.svg + '</div>');

    body.push('<div class="smalls">');
    [['By category', p.byCategory], ['By month', p.byMonth]].forEach(function (t) {
      var h = '<div class="small"><h4>' + esc(t[0]) + '</h4><table class="tbl"><thead><tr><th>' +
        esc(t[0].replace('By ', '')) + '</th><th class="num">Gifts</th><th class="num">Total</th></tr></thead><tbody>';
      t[1].forEach(function (r) {
        h += '<tr><td>' + esc(r.key) + '</td><td class="num">' + esc(fmt.number(r.count)) +
          '</td><td class="num">' + esc(fmt.currency(r.sum)) + '</td></tr>';
      });
      h += '</tbody></table></div>';
      body.push(h);
    });
    body.push('</div>');

    return panel({
      title: 'High-dollar donations', prov: p.provenance, wide: true,
      subtitle: 'Counts and volume by category, month and amount band.',
      caveats: p.smallSample
        ? ['Small sample: ' + p.total.count + ' donation rows. The tab is pre-formatted down to row 500 and two of its columns are filled to the bottom, which makes it look far larger than it is.']
        : [],
      body: body.join('')
    });
  }

  // ---------------------------------------------------------------------------
  // Raw sheet view — the workbook as it looks in Excel, with real row numbers and
  // column letters, so a staffer can check any figure without opening the file.
  // ---------------------------------------------------------------------------

  function sheetReport(name) {
    return (state.data.parseReport.tabs || []).filter(function (t) { return t.tab === name; })[0] || null;
  }

  function rawTableHTML(sheetName, opts) {
    opts = opts || {};
    var rep = sheetReport(sheetName);
    var raw = rep && rep.raw;
    var h = ['<section class="raw" data-raw="' + esc(sheetName) + '">'];
    h.push('<header class="raw-hd"><div><h3>Sheet data — <code>' + esc(sheetName) + '</code></h3>');

    if (!raw || !raw.rows.length) {
      h.push('<p class="panel-sub">Nothing readable in this sheet.</p></div></header>');
      h.push(emptyState('No rows to show', opts.emptyNote || 'This sheet has no populated cells the dashboard could read.'));
      h.push('</section>');
      return h.join('');
    }

    h.push('<p class="panel-sub">' + esc(raw.rows.length) + ' row' + (raw.rows.length === 1 ? '' : 's') +
      ' × ' + esc(raw.columns.length) + ' columns' +
      (raw.headerRow ? ', header on row ' + esc(raw.headerRow) : ', no header row found') +
      '. Values are shown exactly as the sheet formats them.</p></div>');
    h.push('<input type="search" class="raw-find" placeholder="Filter rows…" aria-label="Filter rows in ' + esc(sheetName) + '">');
    h.push('</header>');

    if (raw.omitted && raw.omitted.length) {
      h.push('<p class="pii-note"><strong>' + esc(raw.omitted.length) + ' columns are deliberately not shown:</strong> ' +
        esc(raw.omitted.join(', ')) + '. These identify individual donors and are discarded while the file is read, ' +
        'so they are absent from this table for the same reason they are absent from the charts.</p>');
    }
    if (opts.note) h.push('<p class="foot">' + esc(opts.note) + '</p>');

    h.push('<div class="raw-wrap"><table class="tbl raw-tbl"><thead>');
    h.push('<tr class="raw-letters"><th class="raw-rn">#</th>');
    raw.columns.forEach(function (c) { h.push('<th>' + esc(c.letter) + '</th>'); });
    h.push('</tr><tr><th class="raw-rn"></th>');
    raw.columns.forEach(function (c) { h.push('<th>' + esc(c.name || '—') + '</th>'); });
    h.push('</tr></thead><tbody>');
    raw.rows.forEach(function (row) {
      h.push('<tr><td class="raw-rn">' + esc(row.r) + '</td>');
      row.cells.forEach(function (v) {
        var cls = /^#(VALUE|REF|N\/A|DIV|NAME|NULL|NUM)/.test(v) ? ' class="raw-err"' : '';
        h.push('<td' + cls + '>' + esc(v) + '</td>');
      });
      h.push('</tr>');
    });
    h.push('</tbody></table></div>');
    if (raw.truncated) h.push('<p class="foot">Showing the first ' + esc(raw.rows.length) + ' rows.</p>');
    h.push('<p class="raw-count foot"></p>');
    h.push('</section>');
    return h.join('');
  }

  // ---------------------------------------------------------------------------
  // Per-sheet tabs
  // ---------------------------------------------------------------------------

  function sheetHeader(sheetName, blurb) {
    var rep = sheetReport(sheetName);
    var h = ['<div class="sheet-hd">'];
    h.push('<div><h2>' + esc(sheetName) + '</h2><p class="panel-sub">' + esc(blurb) + '</p></div>');
    if (rep) {
      h.push('<dl class="sheet-facts">');
      h.push('<div><dt>Status</dt><dd><span class="st st-' + rep.status + '">' +
        esc(STATUS_LABEL[rep.status] || rep.status) + '</span></dd></div>');
      h.push('<div><dt>Header row</dt><dd>' + (rep.headerRow == null ? '—' : esc(rep.headerRow)) + '</dd></div>');
      h.push('<div><dt>Rows used</dt><dd>' + esc(fmt.number(rep.rowsLoaded || 0)) + '</dd></div>');
      h.push('<div><dt>Cells</dt><dd><code>' + esc(rep.dataRange || '—') + '</code></dd></div>');
      if (rep.dateRange) h.push('<div><dt>Dates</dt><dd>' + esc(rep.dateRange) + '</dd></div>');
      h.push('</dl>');
    }
    h.push('</div>');
    return h.join('');
  }

  function renderOverview() {
    var h = [renderKPIs()];
    h.push('<div class="grid">');
    h.push(renderActualVsGoal());
    h.push(renderWorkbookMap());
    h.push('</div>');
    return h.join('');
  }

  /** A directory of the workbook, so the tab bar is not the only way in. */
  function renderWorkbookMap() {
    var h = ['<section class="panel panel-wide" id="wbmap"><header class="panel-hd"><div>' +
      '<h2>What is in this workbook</h2><p class="panel-sub">Every sheet, and where to find it here.</p>' +
      '</div></header><div class="panel-body"><div class="map">'];
    TABS.filter(function (t) { return t.sheet; }).forEach(function (t) {
      var rep = sheetReport(t.sheet);
      if (!rep) return;
      h.push('<button class="map-card" data-goto="' + esc(t.id) + '">');
      h.push('<span class="map-nm">' + esc(t.sheet) + '</span>');
      h.push('<span class="st st-' + rep.status + '">' + esc(STATUS_LABEL[rep.status] || rep.status) + '</span>');
      h.push('<span class="map-n">' + esc(fmt.number(rep.rowsLoaded || (rep.raw ? rep.raw.rows.length : 0))) + ' rows</span>');
      h.push('<span class="map-d">' + esc(t.blurb || '') + '</span>');
      h.push('</button>');
    });
    h.push('</div></div></section>');
    return h.join('');
  }

  function renderEmailTab() {
    var dow = state.metrics.panelEmailByDayOfWeek(state.filters);
    var act = state.metrics.panelEmailActions(state.filters);

    var dowChart = V.barChart({
      label: 'Revenue per 1,000 by day of week', labels: dow.rows.map(function (r) { return r.key; }),
      values: dow.rows.map(function (r) { return r.revPerK; }), color: SERIES[0],
      provenance: dow.provenance, height: 190,
      yFormat: function (v) { return '$' + v.toFixed(1); }, tipFormat: fmt.currencyPrecise
    });
    var dowPanel = panel({
      title: 'Performance by day of week', prov: dow.provenance, wide: true,
      subtitle: 'From the Day of Week column, as the sheet records it.',
      body: '<div class="chart">' + dowChart.svg + '</div>' +
        '<table class="tbl"><thead><tr><th>Day</th><th class="num">Sends</th><th class="num">Recipients</th>' +
        '<th class="num">Raised</th><th class="num">Rev / 1k</th><th class="num">Open</th><th class="num">Unsub</th></tr></thead><tbody>' +
        dow.rows.map(function (r) {
          return '<tr><td>' + esc(r.key) + '</td><td class="num">' + esc(fmt.number(r.sends)) +
            '</td><td class="num">' + esc(fmt.number(r.recipients)) + '</td><td class="num">' + esc(fmt.currency(r.raised)) +
            '</td><td class="num">' + esc(r.revPerK == null ? '—' : fmt.currencyPrecise(r.revPerK)) +
            '</td><td class="num">' + esc(fmt.percent(r.openRate)) + '</td><td class="num">' + esc(fmt.percent(r.unsubRate, 3)) + '</td></tr>';
        }).join('') + '</tbody></table>'
    });

    var actChart = V.lineChart({
      label: 'Actions per month', labels: act.labels,
      series: [{ name: 'Actions', color: SERIES[2], values: act.actions }],
      provenance: act.provenance, yFormat: fmt.number, tipFormat: fmt.number
    });
    var avgChart = V.lineChart({
      label: 'Average gift', labels: act.labels, width: 380, height: 180,
      margin: { top: 14, right: 14, bottom: 30, left: 52 },
      series: [{ name: 'Average gift', color: SERIES[4], values: act.avgGift }],
      provenance: act.provenance,
      yFormat: function (v) { return '$' + Math.round(v); }, tipFormat: fmt.currencyPrecise
    });
    var rateChart = V.lineChart({
      label: 'Action rate', labels: act.labels, width: 380, height: 180,
      margin: { top: 14, right: 14, bottom: 30, left: 52 },
      series: [{ name: 'Action rate', color: SERIES[6], values: act.actionRate }],
      provenance: act.provenance,
      yFormat: function (v) { return (v * 100).toFixed(2) + '%'; },
      tipFormat: function (v) { return fmt.percent(v, 3); }
    });
    var actPanel = panel({
      title: 'Actions and average gift', prov: act.provenance, wide: true,
      subtitle: 'The advocacy and gift-size columns, which the revenue charts do not use.',
      body: '<h3 class="sub-h">Actions per month</h3><div class="chart">' + actChart.svg + '</div>' +
        '<div class="smalls"><div class="small"><h4>Average gift</h4>' + avgChart.svg + '</div>' +
        '<div class="small"><h4>Action rate</h4>' + rateChart.svg + '</div></div>'
    });

    return sheetHeader('Email Statistics',
      'Send-level email performance — the densest and cleanest data in the workbook.') +
      '<div class="grid">' + renderEmailTrend() + renderListFatigue() + dowPanel + renderBreakdown() + actPanel + '</div>' +
      rawTableHTML('Email Statistics', { note: 'One row per send. Subject Line, Mailing Topic, Send Time and Label are shown here in full — the charts above summarise the numeric columns only.' });
  }

  function renderP2PTab() {
    return sheetHeader('P2P Statistics', 'Peer-to-peer and SMS sends, with immediate and lifetime returns.') +
      '<div class="grid">' + renderP2P() + '</div>' +
      rawTableHTML('P2P Statistics', { note: 'Every column from the sheet, including Goal, Source Code, Client and the NTL breakdown.' });
  }

  function renderAdsTab() {
    return sheetHeader('Ads Report - Finance-Adjusted', 'Paid media by placement, standard toplines against finance-adjusted.') +
      '<div class="grid">' + renderAds() + '</div>' +
      rawTableHTML('Ads Report - Finance-Adjusted', { note: 'Columns H and I are the Standard Toplines pair; J and K are the Finance-Adjusted pair. Row 3 gives both pairs the same two names, so the dashboard tells them apart by position.' });
  }

  function renderGoalsTab() {
    var g = state.data.monthlyGoals;
    var rows = g.rows || [];
    var chan = state.metrics.panelChannelMix(state.filters);
    var totals = {};
    state.data.channelKeys.forEach(function (k) {
      totals[k] = rows.reduce(function (t, r) { return t + (r[k] || 0); }, 0);
    });
    var labelOf = { email: 'Email', recurring: 'Recurring', ads: 'Ads', website: 'Website', tandem: 'Tandem',
                    smsBroadcast: 'SMS Broadcast', smsP2p: 'SMS P2P', social: 'Social', allOther: 'All Other' };
    var present = state.data.channelKeys.filter(function (k) { return totals[k] !== 0; });
    var mixChart = V.barChart({
      label: 'Total goal by channel', labels: present.map(function (k) { return labelOf[k]; }),
      values: present.map(function (k) { return totals[k]; }), color: SERIES[0],
      provenance: chan.provenance, height: 200, yFormat: fmt.compact, tipFormat: fmt.currency
    });
    var totalPanel = panel({
      title: 'Total goal by channel', prov: chan.provenance, wide: true,
      subtitle: 'Every channel column summed across the months in range.',
      caveats: ['These are targets. The workbook records no per-channel actuals to compare them against.'],
      body: '<div class="chart">' + mixChart.svg + '</div>' +
        (present.length < state.data.channelKeys.length
          ? '<p class="foot">Zero in every month in range: ' +
            esc(state.data.channelKeys.filter(function (k) { return present.indexOf(k) === -1; })
              .map(function (k) { return labelOf[k]; }).join(', ')) + '.</p>' : '')
    });
    var gt = state.metrics.panelGoalTotals(state.filters);
    var gtSeries = [
      { name: 'Gross Raised', color: SERIES[0], values: gt.grossRaised },
      { name: 'Gross Spend', color: SERIES[1], values: gt.grossSpend },
      { name: 'Net Raised', color: SERIES[2], values: gt.netRaised }
    ];
    var gtChart = V.lineChart({
      label: 'Goal totals by month', labels: gt.labels, series: gtSeries,
      provenance: gt.provenance, yFormat: fmt.compact, tipFormat: fmt.currency
    });
    var gtPanel = panel({
      title: 'Goal totals by month', prov: gt.provenance, wide: true,
      subtitle: 'The Gross Raised, Gross Spend and Net Raised columns (L, M and N).',
      body: '<div class="chart">' + gtChart.svg + '</div>' + V.legend(gtSeries) +
        '<table class="tbl"><thead><tr><th>Month</th><th class="num">Gross Raised</th>' +
        '<th class="num">Gross Spend</th><th class="num">Net Raised</th></tr></thead><tbody>' +
        gt.labels.map(function (lb, i) {
          return '<tr><td>' + esc(lb) + '</td><td class="num">' + esc(fmt.currency(gt.grossRaised[i])) +
            '</td><td class="num">' + esc(fmt.currency(gt.grossSpend[i])) +
            '</td><td class="num">' + esc(fmt.currency(gt.netRaised[i])) + '</td></tr>';
        }).join('') + '</tbody></table>'
    });

    return sheetHeader('April 2026 Updated Goals', 'Titled "Digital Goals" in cell A3 — monthly targets for 2026, by channel.') +
      '<div class="grid">' + renderChannelMix() + gtPanel + totalPanel + '</div>' +
      rawTableHTML('April 2026 Updated Goals', { note: 'Row 18 is a Totals row. It appears here because the sheet contains it, but it is excluded from every calculation — including it would roughly double each figure.' });
  }

  function renderProjectionsTab() {
    var scenarios = state.metrics.panelProjections();
    var labelOf = { email: 'Email', recurring: 'Recurring', ads: 'Ads', website: 'Website', tandem: 'Tandem',
                    smsBroadcast: 'SMS Broadcast', smsP2p: 'SMS P2P', social: 'Social', allOther: 'All Other' };
    var body = scenarios.map(function (sc) {
      var name = sc.scenario.replace(/^Digital Projections\s*/, '');
      if (!sc.usable) {
        return panel({
          title: name, prov: sc.provenance,
          subtitle: 'Block starting at row ' + sc.headerRow + '.',
          body: emptyState('This scenario cannot be charted', sc.unusableReason.charAt(0).toUpperCase() + sc.unusableReason.slice(1) + '.')
        });
      }
      var series = sc.channels.map(function (c, i) {
        return { name: labelOf[c.key] || c.key, color: SERIES[i % SERIES.length], values: c.values };
      });
      var chart = V.stackedArea({
        label: name, labels: sc.labels, series: series, provenance: sc.provenance,
        yFormat: fmt.compact, tipFormat: fmt.currency
      });
      return panel({
        title: name, prov: sc.provenance, wide: true,
        subtitle: 'Block starting at row ' + sc.headerRow + ' · ' + sc.labels.length + ' months.',
        body: (chart.empty ? emptyState('No non-zero values', 'Every channel in this block is zero.')
                           : '<div class="chart">' + chart.svg + '</div>' + V.legend(series))
      });
    }).join('');

    return sheetHeader('Digital Projections', 'Three scenario blocks stacked in one sheet — Low, Medium and High Investment.') +
      '<p class="caveat"><span class="caveat-i" aria-hidden="true">!</span>' +
      'This sheet is not one table. It holds three separate projection blocks, each with its own header and Totals row. ' +
      'The dashboard reads them separately; treating them as one table would merge three different scenarios.</p>' +
      '<div class="grid">' + body + '</div>' +
      rawTableHTML('Digital Projections', { note: 'All three blocks in sequence. Header rows appear at 2, 20 and 74; Totals rows at 16, 70 and 124.' });
  }

  function renderHighDollarTab() {
    var p = state.metrics.panelHighDollar();
    var extra = '';
    if (p.available) {
      var hd = state.data.highDollarAgg;
      extra = panel({
        title: 'Claim status', prov: p.provenance, wide: true,
        subtitle: 'The Finance Claim and Authentic Added columns, which the amount charts do not use.',
        body: '<div class="smalls">' +
          [['Finance Claim', p.claims.financeClaim], ['Authentic Added', p.claims.authenticAdded]].map(function (pair) {
            return '<div class="small"><h4>' + esc(pair[0]) + '</h4>' +
              '<table class="tbl"><thead><tr><th>Value</th><th class="num">Gifts</th><th class="num">Total</th></tr></thead><tbody>' +
              pair[1].map(function (c) {
                return '<tr><td>' + esc(c.key) + '</td><td class="num">' + esc(fmt.number(c.count)) +
                  '</td><td class="num">' + esc(fmt.currency(c.sum)) + '</td></tr>';
              }).join('') + '</tbody></table></div>';
          }).join('') + '</div>'
      });
    }
    return sheetHeader('High-Dollar Donations', 'Donation-level records. Shown as aggregates only — donor identity never enters the page.') +
      '<div class="grid">' + renderHighDollar() + extra + '</div>' +
      rawTableHTML('High-Dollar Donations', {
        note: 'The sheet is pre-formatted to row 500 and the Finance Claim / Authentic Added columns are filled to the bottom, which is why hundreds of rows appear here with only those two cells populated. Only the 22 rows with a date and amount are real donations.'
      });
  }

  function renderSimpleSheetTab(sheetName, blurb, groupings, note) {
    var idMap = { 'Email Sending CalendarTracker': 'emailCalendar', 'P2P Calendar': 'p2pCalendar', 'Partner Toolkits': 'partnerToolkits' };
    var p = state.metrics.panelSimpleTab(idMap[sheetName], groupings);
    var body = '';
    if (p.available && p.groups.length) {
      body = panel({
        title: 'Breakdown', prov: p.provenance, wide: true,
        subtitle: 'Counts across ' + esc(fmt.number(p.total)) + ' rows.',
        body: '<div class="smalls">' + p.groups.map(function (g) {
          return '<div class="small"><h4>' + esc(g.label) + '</h4><table class="tbl"><thead><tr><th>' +
            esc(g.label) + '</th><th class="num">Rows</th></tr></thead><tbody>' +
            g.counts.map(function (c) {
              return '<tr><td>' + esc(c.key) + '</td><td class="num">' + esc(fmt.number(c.count)) + '</td></tr>';
            }).join('') + '</tbody></table></div>';
        }).join('') + '</div>'
      });
    }
    return sheetHeader(sheetName, blurb) + (body ? '<div class="grid">' + body + '</div>' : '') +
      rawTableHTML(sheetName, { note: note });
  }

  function renderRawOnlyTab(sheetName, blurb, why) {
    return sheetHeader(sheetName, blurb) +
      '<p class="caveat"><span class="caveat-i" aria-hidden="true">!</span>' + esc(why) + '</p>' +
      rawTableHTML(sheetName, { emptyNote: why });
  }

  // ---------------------------------------------------------------------------
  // Data Sources — the first thing to check when a number looks wrong.
  // ---------------------------------------------------------------------------

  var STATUS_LABEL = { parsed: 'Parsed', skipped: 'Skipped', failed: 'Failed', missing: 'Not found' };

  function renderDataSources() {
    var rep = state.data.parseReport;
    var warned = rep.warnings.length;
    var h = ['<section class="panel panel-wide" id="sources">'];
    h.push('<header class="panel-hd"><div><h2>Data sources</h2>' +
      '<p class="panel-sub">Every tab in the workbook, what was read from it, and anything that looked wrong.</p></div></header>');

    if (warned) {
      h.push('<div class="warns"><h3>' + warned + ' warning' + (warned === 1 ? '' : 's') + '</h3><ul>');
      rep.warnings.forEach(function (w) { h.push('<li>' + esc(w.text) + '</li>'); });
      h.push('</ul></div>');
    } else {
      h.push('<p class="ok-note">No warnings. Every expected tab and column was found by exact name.</p>');
    }

    h.push('<table class="tbl tbl-src"><thead><tr><th>Tab</th><th>Status</th><th class="num">Header row</th>' +
      '<th class="num">Rows loaded</th><th>Cells read</th><th>Dates covered</th><th>Notes</th></tr></thead><tbody>');
    rep.tabs.forEach(function (t) {
      var cls = 'st st-' + t.status;
      h.push('<tr>');
      h.push('<td><strong>' + esc(t.tab) + '</strong></td>');
      h.push('<td><span class="' + cls + '">' + esc(STATUS_LABEL[t.status] || t.status) + '</span></td>');
      h.push('<td class="num">' + (t.headerRow == null ? '—' : esc(t.headerRow) +
        (t.hintHeaderRow && t.headerRow !== t.hintHeaderRow ? ' <span class="shift">(expected ' + esc(t.hintHeaderRow) + ')</span>' : '')) + '</td>');
      h.push('<td class="num">' + esc(fmt.number(t.rowsLoaded || 0)) + '</td>');
      h.push('<td><code>' + esc(t.dataRange || '—') + '</code></td>');
      h.push('<td>' + esc(t.dateRange || '—') + '</td>');

      var notes = [];
      (t.notes || []).forEach(function (n) { notes.push(esc(n)); });
      (t.blocks || []).length > 1 && notes.push('<strong>' + t.blocks.length + ' stacked blocks:</strong> ' +
        t.blocks.map(function (b) { return esc(b.title || 'row ' + b.headerRow) + ' (' + b.rows + ' rows, ' + esc(b.dataRange || 'empty') + ')'; }).join('; '));
      (t.rowsExcluded || []).filter(function (e) { return e.count > 0; }).forEach(function (e) {
        notes.push('Excluded ' + esc(fmt.number(e.count)) + ' row(s): ' + esc(e.reason) + '.');
      });
      (t.rejections || []).forEach(function (r) {
        notes.push(esc(r.count) + ' unreadable value(s) in "' + esc(r.column) + '" (' + esc(r.reason) + ').');
      });
      if ((t.unmatchedColumns || []).length) {
        notes.push('<strong>Columns present but not used:</strong> ' + t.unmatchedColumns.map(function (c) {
          return esc(c.name || '(unnamed)') + ' (' + esc(c.letter) + ')';
        }).join(', ') + '. Add them to the parser to pick them up.');
      }
      (t.warnings || []).forEach(function (w) { if (w.text) notes.push('<em>' + esc(w.text) + '</em>'); });
      h.push('<td class="notes">' + (notes.length ? notes.join('<br>') : '—') + '</td>');
      h.push('</tr>');
    });
    h.push('</tbody></table>');
    h.push('<p class="foot">Row counts are found by walking the data, never by reading the sheet’s stated size — several tabs in this workbook are pre-formatted hundreds of rows past their last real value.</p>');
    h.push('</section>');
    return h.join('');
  }

  // ---------------------------------------------------------------------------
  // Tabs — one per sheet in the workbook, plus an overview for the figures that
  // combine several sheets, plus the data-sources report.
  // ---------------------------------------------------------------------------

  var TABS = [
    { id: 'overview', label: 'Overview', filters: true, render: renderOverview },
    { id: 'email', label: 'Email Statistics', sheet: 'Email Statistics', filters: true,
      blurb: 'Send-level email performance.',
      render: renderEmailTab },
    { id: 'p2p', label: 'P2P Statistics', sheet: 'P2P Statistics', filters: true,
      blurb: 'Peer-to-peer and SMS sends.',
      render: renderP2PTab },
    { id: 'ads', label: 'Ads Report', sheet: 'Ads Report - Finance-Adjusted',
      blurb: 'Paid media by placement.',
      render: renderAdsTab },
    { id: 'goals', label: 'Goals', sheet: 'April 2026 Updated Goals', filters: true,
      blurb: 'Monthly targets by channel.',
      render: renderGoalsTab },
    { id: 'projections', label: 'Projections', sheet: 'Digital Projections',
      blurb: 'Three stacked scenario blocks.',
      render: renderProjectionsTab },
    { id: 'highdollar', label: 'High-Dollar', sheet: 'High-Dollar Donations',
      blurb: 'Donation-level records, aggregated.',
      render: renderHighDollarTab },
    { id: 'emailcal', label: 'Email Calendar', sheet: 'Email Sending CalendarTracker',
      blurb: 'What was scheduled to send.',
      render: function () {
        return renderSimpleSheetTab('Email Sending CalendarTracker',
          'The email sending plan — what was scheduled, by whom, and where it got to.',
          [{ key: 'status', label: 'Status' }, { key: 'ask', label: 'Ask' }, { key: 'sender', label: 'Sender' }],
          'An operations tab: it records the plan, not the results. Link to draft and Notes are shown in full.');
      } },
    { id: 'p2pcal', label: 'P2P Calendar', sheet: 'P2P Calendar',
      blurb: 'Scheduled P2P sends.',
      render: function () {
        return renderSimpleSheetTab('P2P Calendar',
          'The P2P sending plan.',
          [{ key: 'status', label: 'Status' }, { key: 'type', label: 'Type' }],
          'Column C is headed "Topic \n(Link copy here)" with a line break inside the header cell — the dashboard normalises that when matching.');
      } },
    { id: 'toolkits', label: 'Partner Toolkits', sheet: 'Partner Toolkits',
      blurb: 'Partner reference list.',
      render: function () {
        return renderSimpleSheetTab('Partner Toolkits',
          'Reference list of partner toolkits and where they live.',
          [{ key: 'office', label: 'Office/Org' }],
          'A reference list, not performance data.');
      } },
    { id: 'fb', label: 'FB Audience', sheet: 'FB Audience Report',
      blurb: 'Broken — formula error.',
      render: function () {
        return renderRawOnlyTab('FB Audience Report', 'Facebook audience toplines.',
          'This sheet cannot be read: it has no header row, and cell C4 holds #VALUE! from a broken QUERY formula (B14 is #REF!). The raw cells are shown below so you can see the state it is in. Repair the formula in the workbook and this tab will start working.');
      } },
    { id: 'cover', label: 'Digital Report', sheet: 'Digital Report',
      blurb: 'Formatted cover page.',
      render: function () {
        return renderRawOnlyTab('Digital Report', 'The workbook cover page.',
          'This is a formatted cover sheet rather than a data table, so nothing is charted from it. Its cells are shown below.');
      } },
    { id: 'paidmedia', label: 'Paid Media Report', sheet: 'Paid Media Report',
      blurb: 'Metric definitions.',
      render: function () {
        return renderRawOnlyTab('Paid Media Report', 'Definitions of the paid-media metrics.',
          'This sheet defines terms rather than holding data, so nothing is charted from it. It is worth reading — it is where the metric definitions live.');
      } },
    { id: 'sources', label: 'Data sources', render: renderDataSources }
  ];

  function activeTab() {
    return TABS.filter(function (t) { return t.id === state.tab; })[0] || TABS[0];
  }

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------

  function render() {
    state.bubbles = {};
    state.bubbleSeq = 0;
    closeBubble(true);

    var rep = state.data.parseReport;
    var parsed = rep.tabs.filter(function (t) { return t.status === 'parsed'; }).length;
    var tab = activeTab();
    var h = [];

    h.push('<header class="top">');
    h.push('<div class="top-l"><h1>CAB Digital Report</h1>' +
      '<p class="top-sub">' + esc(state.fileName) + ' · ' + parsed + ' of ' + rep.tabs.length + ' sheets parsed · ' +
      esc(fmt.number(state.data.emailSends.rows.length)) + ' email sends, ' +
      esc(fmt.number(state.data.p2pSends.rows.length)) + ' P2P sends</p></div>');
    h.push('<div class="top-r">' +
      '<button class="chip" id="btn-theme" title="Toggle light and dark">Theme</button>' +
      '<button class="chip" id="btn-new">Load another file</button></div>');
    h.push('</header>');
    h.push('<p class="privacy">This file was read inside your browser. Nothing was uploaded, and the page makes no network requests. Hover any <span class="info-static" aria-hidden="true">i</span> to see exactly which cells a figure came from.</p>');

    h.push('<nav class="tabbar" aria-label="Workbook sheets">');
    TABS.forEach(function (t) {
      var srep = t.sheet ? sheetReport(t.sheet) : null;
      var bad = srep && (srep.status === 'failed' || srep.status === 'missing');
      var warnCount = t.id === 'sources' ? rep.warnings.length : 0;
      h.push('<button class="tbtn' + (t.id === state.tab ? ' tbtn-on' : '') + (bad ? ' tbtn-bad' : '') +
        '" data-tab="' + esc(t.id) + '">' + esc(t.label) +
        (warnCount ? ' <span class="badge">' + warnCount + '</span>' : '') +
        (bad ? ' <span class="badge badge-bad">!</span>' : '') + '</button>');
    });
    h.push('</nav>');

    if (tab.filters) h.push(renderFilters());
    h.push('<div class="tabpage">' + tab.render() + '</div>');

    $('#app').innerHTML = h.join('');
    $('#drop').style.display = 'none';
    $('#app').style.display = 'block';
    wire();
  }

  // ---------------------------------------------------------------------------
  // Events
  // ---------------------------------------------------------------------------

  function monthsAgo(n) {
    var ex = state.metrics.facets().dateExtent;
    var end = ex ? ex.max : new Date();
    return new Date(Date.UTC(end.getUTCFullYear(), end.getUTCMonth() - n + 1, 1));
  }

  function goTab(id) {
    state.tab = id;
    render();
    window.scrollTo(0, 0);
  }

  function wire() {
    $$('.tbtn[data-tab]').forEach(function (b) {
      b.onclick = function () { goTab(this.getAttribute('data-tab')); };
    });
    $$('.map-card[data-goto]').forEach(function (b) {
      b.onclick = function () { goTab(this.getAttribute('data-goto')); };
    });

    var from = $('#f-from'), to = $('#f-to');
    if (from) from.onchange = function () { state.filters.from = this.value ? new Date(this.value + 'T00:00:00Z') : null; render(); };
    if (to) to.onchange = function () { state.filters.to = this.value ? new Date(this.value + 'T23:59:59Z') : null; render(); };

    $$('.chip[data-range]').forEach(function (b) {
      b.onclick = function () {
        var r = this.getAttribute('data-range');
        var ex = state.metrics.facets().dateExtent;
        if (r === 'all') { state.filters.from = null; state.filters.to = null; }
        else { state.filters.from = monthsAgo(+r); state.filters.to = ex ? ex.max : null; }
        render();
      };
    });

    $$('.f-multi').forEach(function (sel) {
      sel.onchange = function () {
        state.filters[this.getAttribute('data-key')] =
          Array.prototype.filter.call(this.options, function (o) { return o.selected; })
            .map(function (o) { return o.value; });
        render();
      };
    });

    var reset = $('#f-reset');
    if (reset) reset.onclick = function () {
      state.filters = { from: monthsAgo(6), to: state.metrics.facets().dateExtent.max, ask: [], subtype: [], audience: [], sender: [] };
      render();
    };

    $$('.tab[data-dim]').forEach(function (b) {
      b.onclick = function () { state.breakdownDim = this.getAttribute('data-dim'); render(); };
    });

    var sc = $('#p1-scenario');
    if (sc) sc.onchange = function () { state.scenario = this.value || null; render(); };

    var nb = $('#btn-new');
    if (nb) nb.onclick = function () {
      state.data = null; state.metrics = null;
      $('#app').style.display = 'none';
      $('#drop').style.display = 'flex';
      $('#file').value = '';
      $('#drop-msg').textContent = '';
    };

    var tb = $('#btn-theme');
    if (tb) tb.onclick = function () {
      var cur = document.documentElement.getAttribute('data-theme');
      var next = cur === 'dark' ? 'light' : cur === 'light' ? '' : 'dark';
      if (next) document.documentElement.setAttribute('data-theme', next);
      else document.documentElement.removeAttribute('data-theme');
    };

    wireRawFilter();
    wireSort();
  }

  /** Plain-text row filter over a raw sheet table. */
  function wireRawFilter() {
    $$('.raw').forEach(function (sec) {
      var input = $('.raw-find', sec);
      var count = $('.raw-count', sec);
      if (!input) return;
      var rows = $$('.raw-tbl tbody tr', sec);
      input.oninput = function () {
        var q = this.value.trim().toLowerCase();
        var shown = 0;
        rows.forEach(function (tr) {
          var hit = !q || tr.textContent.toLowerCase().indexOf(q) !== -1;
          tr.style.display = hit ? '' : 'none';
          if (hit) shown++;
        });
        count.textContent = q ? shown + ' of ' + rows.length + ' rows match “' + this.value.trim() + '”' : '';
      };
    });
  }

  /** Sortable columns on the breakdown table. */
  function wireSort() {
    $$('.tbl-sort th[data-sort]').forEach(function (th) {
      th.onclick = function () {
        var tbl = th.closest('table');
        var idx = Array.prototype.indexOf.call(th.parentNode.children, th);
        var asc = th.getAttribute('data-dir') !== 'asc';
        $$('th', tbl).forEach(function (o) { o.removeAttribute('data-dir'); });
        th.setAttribute('data-dir', asc ? 'asc' : 'desc');
        var tb = $('tbody', tbl);
        var rows = $$('tr', tb);
        rows.sort(function (a, b) {
          var x = a.children[idx].textContent.trim(), y = b.children[idx].textContent.trim();
          var nx = parseFloat(x.replace(/[^0-9.\-]/g, '')), ny = parseFloat(y.replace(/[^0-9.\-]/g, ''));
          var cmp = (!isNaN(nx) && !isNaN(ny)) ? nx - ny : x.localeCompare(y);
          return asc ? cmp : -cmp;
        });
        rows.forEach(function (r) { tb.appendChild(r); });
      };
    });
  }

  // Bubble interaction, delegated so it survives every re-render.
  document.addEventListener('mouseover', function (e) {
    var btn = e.target.closest && e.target.closest('.info');
    if (!btn) return;
    clearTimeout(closeTimer);
    clearTimeout(hoverTimer);
    if (state.pinned) return;
    hoverTimer = setTimeout(function () { openBubble(btn, false); }, HOVER_IN);
  });

  document.addEventListener('mouseout', function (e) {
    var btn = e.target.closest && e.target.closest('.info');
    if (!btn) return;
    clearTimeout(hoverTimer);
    if (!state.pinned) scheduleClose();
  });

  document.addEventListener('click', function (e) {
    var btn = e.target.closest && e.target.closest('.info');
    if (btn) {
      e.preventDefault();
      if (state.pinned && state.openBtn === btn) closeBubble(true);
      else openBubble(btn, true);
      return;
    }
    if (e.target.closest && e.target.closest('.prov-close')) { closeBubble(true); return; }
    if (!e.target.closest || !e.target.closest('.prov-pop')) closeBubble(true);
  });

  // Keyboard: focus opens, Escape closes. Hover alone would strand keyboard users.
  // Gate on :focus-visible — a mouse click also fires focusin, and without this the
  // focus handler opens the bubble a beat before the click handler toggles it shut.
  document.addEventListener('focusin', function (e) {
    var btn = e.target.closest && e.target.closest('.info');
    if (!btn) return;
    var keyboard = true;
    try { keyboard = btn.matches(':focus-visible'); } catch (err) { /* older engines: keep the default */ }
    if (keyboard) openBubble(btn, true);
  });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeBubble(true); });
  window.addEventListener('resize', function () { closeBubble(true); });
  window.addEventListener('scroll', function () {
    var pop = $('.prov-pop');
    if (pop && state.openBtn) positionBubble(state.openBtn, pop);
  }, { passive: true });

  // Chart tooltips.
  var tip = null;
  function showTip(html, x, y) {
    if (!tip) { tip = document.createElement('div'); tip.className = 'v-tip'; document.body.appendChild(tip); }
    tip.innerHTML = html;
    tip.style.display = 'block';
    var w = tip.offsetWidth;
    var left = x + 14;
    if (left + w > window.scrollX + document.documentElement.clientWidth - 8) left = x - w - 14;
    tip.style.left = left + 'px';
    tip.style.top = (y + 14) + 'px';
  }
  function hideTip() { if (tip) tip.style.display = 'none'; }

  document.addEventListener('mouseover', function (e) {
    var t = e.target;
    if (!t.getAttribute) return;
    var simple = t.getAttribute('data-tip');
    if (simple) { showTip(esc(simple), e.pageX, e.pageY); return; }
    var payload = t.getAttribute('data-payload');
    if (payload) {
      var d = JSON.parse(payload);
      var h = '<div class="v-tip-t">' + esc(d.title) + '</div>';
      d.rows.forEach(function (r) {
        h += '<div class="v-tip-r"><i style="background:' + esc(r.color) + '"></i>' +
          '<span>' + esc(r.name) + '</span><b>' + esc(r.text) + '</b></div>';
      });
      showTip(h, e.pageX, e.pageY);
      var svg = t.ownerSVGElement;
      var cross = svg && svg.querySelector('.v-cross');
      if (cross) {
        var cx = t.getAttribute('data-x');
        cross.setAttribute('x1', cx); cross.setAttribute('x2', cx);
        cross.style.display = '';
      }
    }
  });
  document.addEventListener('mousemove', function (e) {
    if (tip && tip.style.display === 'block') {
      var w = tip.offsetWidth, left = e.pageX + 14;
      if (left + w > window.scrollX + document.documentElement.clientWidth - 8) left = e.pageX - w - 14;
      tip.style.left = left + 'px'; tip.style.top = (e.pageY + 14) + 'px';
    }
  });
  document.addEventListener('mouseout', function (e) {
    if (!e.target.getAttribute) return;
    if (e.target.getAttribute('data-tip') || e.target.getAttribute('data-payload')) {
      hideTip();
      var svg = e.target.ownerSVGElement;
      var cross = svg && svg.querySelector('.v-cross');
      if (cross) cross.style.display = 'none';
    }
  });

  // ---------------------------------------------------------------------------
  // File loading — everything stays in this tab.
  // ---------------------------------------------------------------------------

  function loadBuffer(buf, name) {
    var msg = $('#drop-msg');
    msg.className = 'drop-msg';
    msg.textContent = 'Reading ' + name + '…';
    try {
      var data = window.CABParse.createParser(window.XLSX).parseWorkbook(new Uint8Array(buf));
      var parsedTabs = data.parseReport.tabs.filter(function (t) { return t.status === 'parsed'; });
      if (!parsedTabs.length) {
        msg.className = 'drop-msg err';
        msg.innerHTML = '<strong>No readable tabs in that workbook.</strong><br>' +
          'Found: ' + esc(data.parseReport.sheetNames.join(', ') || '(no sheets)') +
          '.<br>None of them match the tabs this dashboard expects — check that this is the CAB Digital Report file.';
        return;
      }
      state.data = data;
      state.metrics = window.CABMetrics.create(data);
      state.fileName = name;
      var ex = state.metrics.facets().dateExtent;
      state.filters = { from: ex ? monthsAgo(6) : null, to: ex ? ex.max : null, ask: [], subtype: [], audience: [], sender: [] };
      state.scenario = null;
      state.tab = 'overview';
      render();
    } catch (err) {
      msg.className = 'drop-msg err';
      msg.innerHTML = '<strong>Could not read that file.</strong><br>' + esc(err && err.message ? err.message : String(err)) +
        '<br>Make sure it is an .xlsx workbook, not a .csv, .numbers or password-protected file.';
      if (window.console && console.error) console.error('CAB dashboard load failed:', err);
    }
  }

  function readFile(file) {
    if (!file) return;
    var msg = $('#drop-msg');
    if (!/\.(xlsx|xlsm|xls)$/i.test(file.name)) {
      msg.className = 'drop-msg err';
      msg.textContent = '"' + file.name + '" is not an Excel workbook. Drop an .xlsx file.';
      return;
    }
    var fr = new FileReader();
    fr.onload = function () { loadBuffer(fr.result, file.name); };
    fr.onerror = function () {
      msg.className = 'drop-msg err';
      msg.textContent = 'The browser could not read that file.';
    };
    fr.readAsArrayBuffer(file);
  }

  function initDrop() {
    var dz = $('#drop');
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drop-on'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); if (ev === 'dragleave' && dz.contains(e.relatedTarget)) return; dz.classList.remove('drop-on'); });
    });
    dz.addEventListener('drop', function (e) { readFile(e.dataTransfer.files && e.dataTransfer.files[0]); });
    $('#file').addEventListener('change', function () { readFile(this.files[0]); });
    $('#pick').addEventListener('click', function () { $('#file').click(); });
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) { e.preventDefault(); });
  }

  window.CABApp = { init: initDrop, _state: state, _loadBuffer: loadBuffer, _render: render };
})();
