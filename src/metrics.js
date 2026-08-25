/*
 * CABMetrics — every number the dashboard shows, each returned together with the
 * provenance that explains it.
 *
 * ZERO DOM DEPENDENCIES, same as CABParse.
 *
 * The contract: a metric function returns {value, provenance} (or {series, provenance}).
 * Value and provenance are built in the SAME function and returned as one object, so
 * they cannot drift apart. There is no path that produces a number without its origin.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CABMetrics = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  function monthKey(d) {
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
  }
  function monthLabel(key) {
    var p = key.split('-');
    var names = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return names[+p[1] - 1] + ' ' + p[0];
  }
  /** ISO week start (Monday), used by the list-fatigue panel. */
  function weekKey(d) {
    var t = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
    var dow = (t.getUTCDay() + 6) % 7;
    t.setUTCDate(t.getUTCDate() - dow);
    return t.toISOString().slice(0, 10);
  }
  function sum(rows, key) {
    var t = 0;
    for (var i = 0; i < rows.length; i++) if (typeof rows[i][key] === 'number') t += rows[i][key];
    return t;
  }

  /**
   * Build the columnsUsed array for a provenance object.
   * Every entry carries the Excel letter — internal camelCase keys never reach the UI.
   */
  function cols(table, entries) {
    return entries.map(function (e) {
      return {
        name: (table.columnNames && table.columnNames[e.key]) || e.label || e.key,
        letter: (table.columnLetters && table.columnLetters[e.key]) || '?',
        role: e.role || 'input'
      };
    });
  }

  /**
   * Any fuzzy or alias match that happened at parse time on a column this metric
   * uses becomes a note here — which is how a parse-time warning reaches every
   * downstream bubble rather than living only in the Data Sources panel.
   */
  function inheritedWarnings(table, keys) {
    var out = [];
    (table.warnings || []).forEach(function (w) {
      if (!w.column || keys.indexOf(w.column) === -1) return;
      out.push(w.text);
    });
    (table.rejections || []).forEach(function (r) {
      out.push(r.count + ' value(s) in "' + r.column + '" could not be read (' + r.reason + ') and were treated as blank.');
    });
    return out;
  }

  function prov(o) {
    return {
      tab: o.tab, headerRow: o.headerRow, dataRange: o.dataRange || '(no rows)',
      columnsUsed: o.columnsUsed || [], rowsAvailable: o.rowsAvailable || 0,
      rowsUsed: o.rowsUsed || 0, rowsExcluded: o.rowsExcluded || [],
      transform: o.transform, notes: o.notes || []
    };
  }

  /** Merge exclusion buckets, keeping every reason string intact. */
  function mergeExcluded() {
    var out = [], idx = {};
    for (var a = 0; a < arguments.length; a++) {
      (arguments[a] || []).forEach(function (e) {
        if (!e || !e.reason) return;
        if (!idx[e.reason]) { idx[e.reason] = { reason: e.reason, count: 0 }; out.push(idx[e.reason]); }
        idx[e.reason].count += e.count;
      });
    }
    return out;
  }

  // ---------------------------------------------------------------------------

  function create(data) {
    var email = data.emailSends;
    var p2p = data.p2pSends;
    var ads = data.adsMonthly;
    var goals = data.monthlyGoals;

    /** Apply the global filters to a send-level table, itemizing every exclusion. */
    function applyFilters(table, f, dimKeys) {
      var kept = [], excluded = {};
      function drop(reason) { excluded[reason] = (excluded[reason] || 0) + 1; }
      (table.rows || []).forEach(function (row) {
        if (f.from && row.date < f.from) { drop('outside selected date range'); return; }
        if (f.to && row.date > f.to) { drop('outside selected date range'); return; }
        for (var i = 0; i < (dimKeys || []).length; i++) {
          var k = dimKeys[i];
          var sel = f[k];
          if (sel && sel.length && sel.indexOf(row[k] || '(blank)') === -1) {
            drop('excluded by the ' + k + ' filter');
            return;
          }
        }
        kept.push(row);
      });
      var list = Object.keys(excluded).map(function (r) { return { reason: r, count: excluded[r] }; });
      return { rows: kept, excluded: list };
    }

    var EMAIL_DIMS = ['ask', 'subtype', 'audience', 'sender'];

    function emailScope(f) { return applyFilters(email, f, EMAIL_DIMS); }
    function p2pScope(f) { return applyFilters(p2p, f, ['audience']); }

    // -------------------------------------------------------------------------
    // KPI row
    // -------------------------------------------------------------------------

    function kpiGrossRaised(f) {
      var e = emailScope(f), p = p2pScope(f);
      var v = sum(e.rows, 'raised') + sum(p.rows, 'immediateRaised');
      return {
        value: v, format: 'currency',
        provenance: prov({
          tab: email.tab + ' + ' + p2p.tab,
          headerRow: email.headerRow + ' and ' + p2p.headerRow,
          dataRange: [email.dataRange, p2p.dataRange].filter(Boolean).join('  +  '),
          columnsUsed: cols(email, [{ key: 'raised', label: 'Raised', role: 'summed' }])
            .concat(cols(p2p, [{ key: 'immediateRaised', label: 'Immediate Raised', role: 'summed' }])),
          rowsAvailable: email.rows.length + p2p.rows.length,
          rowsUsed: e.rows.length + p.rows.length,
          rowsExcluded: mergeExcluded(e.excluded, p.excluded, email.rowsExcluded, p2p.rowsExcluded),
          transform: 'SUM(Raised) from Email Statistics + SUM(Immediate Raised) from P2P Statistics',
          notes: inheritedWarnings(email, ['raised']).concat(inheritedWarnings(p2p, ['immediateRaised']))
        })
      };
    }

    function kpiGrossSpend(f) {
      var p = p2pScope(f);
      // The Ads tab holds 3 rows from Sep 2021 and is outside any 2026 filter window,
      // so it contributes nothing today — but it is still summed and still disclosed.
      var adsRows = (ads.rows || []).filter(function (r) {
        if (f.from && r.month < f.from) return false;
        if (f.to && r.month > f.to) return false;
        return true;
      });
      var v = sum(p.rows, 'grossSpend') + sum(adsRows, 'grossSpend');
      return {
        value: v, format: 'currency',
        provenance: prov({
          tab: p2p.tab + ' + ' + ads.tab,
          headerRow: p2p.headerRow + ' and ' + ads.headerRow,
          dataRange: [p2p.dataRange, ads.dataRange].filter(Boolean).join('  +  '),
          columnsUsed: cols(p2p, [{ key: 'grossSpend', label: 'Gross Spend', role: 'summed' }])
            .concat(cols(ads, [{ key: 'grossSpend', label: 'Gross Spend', role: 'summed' }])),
          rowsAvailable: p2p.rows.length + ads.rows.length,
          rowsUsed: p.rows.length + adsRows.length,
          rowsExcluded: mergeExcluded(p.excluded,
            [{ reason: 'Ads rows outside selected date range', count: ads.rows.length - adsRows.length }],
            p2p.rowsExcluded, ads.rowsExcluded),
          transform: 'SUM(Gross Spend) from P2P Statistics + SUM(Gross Spend) from Ads Report',
          notes: ['Email Statistics has no spend column, so email production cost is not included in this figure.']
            .concat(inheritedWarnings(p2p, ['grossSpend']), inheritedWarnings(ads, ['grossSpend']))
        })
      };
    }

    function kpiNetRaised(f) {
      var g = kpiGrossRaised(f), s = kpiGrossSpend(f);
      return {
        value: g.value - s.value, format: 'currency',
        provenance: prov({
          tab: 'Email Statistics + P2P Statistics + Ads Report',
          headerRow: '5, 5 and 3',
          dataRange: [email.dataRange, p2p.dataRange, ads.dataRange].filter(Boolean).join('  +  '),
          columnsUsed: g.provenance.columnsUsed.concat(s.provenance.columnsUsed),
          rowsAvailable: g.provenance.rowsAvailable + ads.rows.length,
          rowsUsed: g.provenance.rowsUsed + s.provenance.rowsUsed,
          rowsExcluded: mergeExcluded(g.provenance.rowsExcluded, s.provenance.rowsExcluded),
          transform: 'Gross Raised − Gross Spend',
          notes: ['Computed from send-level data, not read from the Net Raised column on the goals tab — that column holds targets, not actuals.']
            .concat(s.provenance.notes)
        })
      };
    }

    function kpiRoas(f) {
      var g = kpiGrossRaised(f), s = kpiGrossSpend(f);
      return {
        value: s.value > 0 ? g.value / s.value : null, format: 'ratio',
        emptyReason: s.value > 0 ? null : 'No spend recorded in the selected range, so ROAS cannot be computed.',
        provenance: prov({
          tab: 'Email Statistics + P2P Statistics + Ads Report',
          headerRow: '5, 5 and 3',
          dataRange: [email.dataRange, p2p.dataRange, ads.dataRange].filter(Boolean).join('  +  '),
          columnsUsed: g.provenance.columnsUsed.concat(s.provenance.columnsUsed),
          rowsAvailable: g.provenance.rowsAvailable + ads.rows.length,
          rowsUsed: g.provenance.rowsUsed + s.provenance.rowsUsed,
          rowsExcluded: mergeExcluded(g.provenance.rowsExcluded, s.provenance.rowsExcluded),
          transform: 'Gross Raised / Gross Spend',
          notes: ['Spend covers P2P and Ads only, so this ratio flatters channels whose cost is not tracked in this workbook.']
            .concat(s.provenance.notes)
        })
      };
    }

    function kpiPctToGoal(f) {
      var g = kpiGrossRaised(f);
      var months = {};
      (goals.rows || []).forEach(function (r) {
        if (!(r.month instanceof Date)) return;
        if (f.from && r.month < startOfMonth(f.from)) return;
        if (f.to && r.month > f.to) return;
        months[monthKey(r.month)] = r;
      });
      var used = Object.keys(months);
      var goalTotal = used.reduce(function (t, k) { return t + (months[k].grossRaised || 0); }, 0);
      return {
        value: goalTotal > 0 ? g.value / goalTotal : null, format: 'percent',
        goalTotal: goalTotal, actualTotal: g.value,
        emptyReason: goalTotal > 0 ? null : 'No goal rows fall inside the selected date range.',
        provenance: prov({
          tab: goals.tab + ' (compared against Email Statistics + P2P Statistics)',
          headerRow: goals.headerRow,
          dataRange: goals.dataRange,
          columnsUsed: cols(goals, [{ key: 'grossRaised', label: 'Gross Raised', role: 'denominator (goal)' },
                                    { key: 'month', label: 'Month', role: 'date filter' }])
            .concat(g.provenance.columnsUsed.map(function (c) {
              return { name: c.name, letter: c.letter, role: 'numerator (actual)' };
            })),
          rowsAvailable: goals.rows.length,
          rowsUsed: used.length,
          rowsExcluded: mergeExcluded(
            [{ reason: 'goal months outside selected date range', count: goals.rows.length - used.length }],
            goals.rowsExcluded),
          transform: 'SUM(Raised)+SUM(Immediate Raised) / SUM(Gross Raised on ' + goals.tab + ') for the months in range',
          notes: [
            'Actuals here cover Email and SMS P2P only. The goal figure is the ALL-channel monthly goal, which includes Recurring, Ads, Website, Tandem, Social and All Other — channels this workbook carries no actuals for.',
            'That makes this percentage a floor, not a verdict on performance.'
          ].concat(inheritedWarnings(goals, ['grossRaised', 'month']))
        })
      };
    }

    function startOfMonth(d) { return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1)); }

    // -------------------------------------------------------------------------
    // Panel 1 — actual vs goal by month. Three series, each separately bubbled,
    // because it mixes three tabs and is the panel most likely to be misread.
    // -------------------------------------------------------------------------

    function panelActualVsGoal(f, scenarioName) {
      var e = emailScope(f), p = p2pScope(f);
      var actual = {};
      e.rows.forEach(function (r) { var k = monthKey(r.date); actual[k] = (actual[k] || 0) + (r.raised || 0); });
      p.rows.forEach(function (r) { var k = monthKey(r.date); actual[k] = (actual[k] || 0) + (r.immediateRaised || 0); });

      var goalRows = (goals.rows || []).filter(function (r) {
        if (!(r.month instanceof Date)) return false;
        if (f.from && r.month < startOfMonth(f.from)) return false;
        if (f.to && r.month > f.to) return false;
        return true;
      });
      var goal = {};
      goalRows.forEach(function (r) { goal[monthKey(r.month)] = r.grossRaised || 0; });

      var scenario = null, projSeries = {}, projRows = [];
      (data.projectionScenarios || []).forEach(function (s) {
        if (s.scenario === scenarioName) scenario = s;
      });
      if (scenario && scenario.usable) {
        scenario.rows.forEach(function (r) {
          if (!(r.month instanceof Date)) return;
          if (f.from && r.month < startOfMonth(f.from)) return;
          if (f.to && r.month > f.to) return;
          projSeries[monthKey(r.month)] = r.grossRaised != null ? r.grossRaised
            : data.channelKeys.reduce(function (t, k) { return t + (typeof r[k] === 'number' ? r[k] : 0); }, 0);
          projRows.push(r);
        });
      }

      var keys = {};
      [actual, goal, projSeries].forEach(function (o) { Object.keys(o).forEach(function (k) { keys[k] = true; }); });
      var months = Object.keys(keys).sort();

      return {
        months: months,
        labels: months.map(monthLabel),
        series: {
          actual: months.map(function (k) { return actual[k] || 0; }),
          goal: months.map(function (k) { return goal[k] != null ? goal[k] : null; }),
          projection: months.map(function (k) { return projSeries[k] != null ? projSeries[k] : null; })
        },
        variance: months.map(function (k) {
          if (goal[k] == null) return null;
          return (actual[k] || 0) - goal[k];
        }),
        scenario: scenario,
        scenarios: (data.projectionScenarios || []).map(function (s) {
          return { name: s.scenario, usable: s.usable, reason: s.unusableReason };
        }),
        provenanceActual: prov({
          tab: email.tab + ' + ' + p2p.tab,
          headerRow: email.headerRow + ' and ' + p2p.headerRow,
          dataRange: [email.dataRange, p2p.dataRange].filter(Boolean).join('  +  '),
          columnsUsed: cols(email, [{ key: 'date', label: 'Date', role: 'grouped by month' },
                                    { key: 'raised', label: 'Raised', role: 'summed' }])
            .concat(cols(p2p, [{ key: 'date', label: 'Date', role: 'grouped by month' },
                               { key: 'immediateRaised', label: 'Immediate Raised', role: 'summed' }])),
          rowsAvailable: email.rows.length + p2p.rows.length,
          rowsUsed: e.rows.length + p.rows.length,
          rowsExcluded: mergeExcluded(e.excluded, p.excluded, email.rowsExcluded, p2p.rowsExcluded),
          transform: 'SUM(Raised) + SUM(Immediate Raised), grouped by calendar month',
          notes: ['This series is COMPUTED from send-level rows. Neither monthly tab in this workbook contains actuals.',
                  'It covers Email and SMS P2P only — no other channel has an actuals source here.']
            .concat(inheritedWarnings(email, ['raised', 'date']), inheritedWarnings(p2p, ['immediateRaised', 'date']))
        }),
        provenanceGoal: prov({
          tab: goals.tab,
          headerRow: goals.headerRow,
          dataRange: goals.dataRange,
          columnsUsed: cols(goals, [{ key: 'month', label: 'Month', role: 'x axis' },
                                    { key: 'grossRaised', label: 'Gross Raised', role: 'plotted' }]),
          rowsAvailable: goals.rows.length,
          rowsUsed: goalRows.length,
          rowsExcluded: mergeExcluded(
            [{ reason: 'goal months outside selected date range', count: goals.rows.length - goalRows.length }],
            goals.rowsExcluded),
          transform: 'Gross Raised (column ' + (goals.columnLetters.grossRaised || '?') + ') read directly, one row per month',
          notes: ['This tab is titled "Digital Goals" in cell A3 — these are TARGETS across all nine channels, not actuals.']
            .concat(inheritedWarnings(goals, ['grossRaised', 'month']))
        }),
        provenanceProjection: scenario ? prov({
          tab: scenario.tab,
          headerRow: scenario.headerRow,
          dataRange: scenario.dataRange,
          columnsUsed: cols(scenario, [{ key: 'month', label: 'Month', role: 'x axis' },
                                       { key: 'grossRaised', label: 'Gross Raised', role: 'plotted' }]),
          rowsAvailable: scenario.rowsAvailable,
          rowsUsed: projRows.length,
          rowsExcluded: mergeExcluded(
            [{ reason: 'projection months outside selected date range', count: scenario.rowsAvailable - projRows.length }],
            scenario.rowsExcluded),
          transform: scenario.usable
            ? 'Gross Raised read directly from the "' + scenario.scenario + '" block, one row per month'
            : 'not plotted',
          notes: (scenario.usable ? [] : ['This scenario is not plotted: ' + scenario.unusableReason + '.'])
            .concat(['The Digital Projections tab holds three stacked scenario blocks; this is the block starting at row ' + scenario.headerRow + '.'])
        }) : null
      };
    }

    // -------------------------------------------------------------------------
    // Panel 2 — channel mix over time (goal mix; no actuals source per channel)
    // -------------------------------------------------------------------------

    function panelChannelMix(f) {
      var rows = (goals.rows || []).filter(function (r) {
        if (!(r.month instanceof Date)) return false;
        if (f.from && r.month < startOfMonth(f.from)) return false;
        if (f.to && r.month > f.to) return false;
        return true;
      });
      var keys = data.channelKeys;
      var labels = { email: 'Email', recurring: 'Recurring', ads: 'Ads', website: 'Website',
                     tandem: 'Tandem', smsBroadcast: 'SMS Broadcast', smsP2p: 'SMS P2P',
                     social: 'Social', allOther: 'All Other' };
      var present = keys.filter(function (k) {
        return rows.some(function (r) { return typeof r[k] === 'number' && r[k] !== 0; });
      });
      return {
        months: rows.map(function (r) { return monthKey(r.month); }),
        labels: rows.map(function (r) { return monthLabel(monthKey(r.month)); }),
        channels: present.map(function (k) {
          return { key: k, label: labels[k], values: rows.map(function (r) { return r[k] || 0; }) };
        }),
        emptyChannels: keys.filter(function (k) { return present.indexOf(k) === -1; }).map(function (k) { return labels[k]; }),
        provenance: prov({
          tab: goals.tab, headerRow: goals.headerRow, dataRange: goals.dataRange,
          columnsUsed: cols(goals, [{ key: 'month', label: 'Month', role: 'x axis' }].concat(
            present.map(function (k) { return { key: k, label: labels[k], role: 'stacked band' }; }))),
          rowsAvailable: goals.rows.length, rowsUsed: rows.length,
          rowsExcluded: mergeExcluded(
            [{ reason: 'months outside selected date range', count: goals.rows.length - rows.length }],
            goals.rowsExcluded),
          transform: 'Each channel column read directly per month and stacked. No division or rescaling.',
          notes: ['This is the GOAL mix from the "Digital Goals" tab, not achieved mix — the workbook has no per-channel actuals.']
            .concat(inheritedWarnings(goals, ['month'].concat(present)))
        })
      };
    }

    /** Goal totals by month — the Gross Raised / Gross Spend / Net Raised columns. */
    function panelGoalTotals(f) {
      var rows = (goals.rows || []).filter(function (r) {
        if (!(r.month instanceof Date)) return false;
        if (f.from && r.month < startOfMonth(f.from)) return false;
        if (f.to && r.month > f.to) return false;
        return true;
      });
      return {
        labels: rows.map(function (r) { return monthLabel(monthKey(r.month)); }),
        grossRaised: rows.map(function (r) { return r.grossRaised; }),
        grossSpend: rows.map(function (r) { return r.grossSpend; }),
        netRaised: rows.map(function (r) { return r.netRaised; }),
        rows: rows,
        provenance: prov({
          tab: goals.tab, headerRow: goals.headerRow, dataRange: goals.dataRange,
          columnsUsed: cols(goals, [{ key: 'month', label: 'Month', role: 'x axis' },
                                    { key: 'grossRaised', label: 'Gross Raised', role: 'plotted' },
                                    { key: 'grossSpend', label: 'Gross Spend', role: 'plotted' },
                                    { key: 'netRaised', label: 'Net Raised', role: 'plotted' }]),
          rowsAvailable: goals.rows.length, rowsUsed: rows.length,
          rowsExcluded: mergeExcluded(
            [{ reason: 'months outside selected date range', count: goals.rows.length - rows.length }],
            goals.rowsExcluded),
          transform: 'Columns L, M and N read directly, one row per month. Net Raised is the sheet\'s own value, not Gross minus Spend.',
          notes: ['All three are targets. Compare them with the computed actuals on the Overview tab, bearing in mind the actuals cover Email and SMS P2P only.']
            .concat(inheritedWarnings(goals, ['grossRaised', 'grossSpend', 'netRaised', 'month']))
        })
      };
    }

    // -------------------------------------------------------------------------
    // Panel 3 — email performance trend
    // -------------------------------------------------------------------------

    function panelEmailTrend(f) {
      var e = emailScope(f);
      var buckets = {}, order = [];
      e.rows.forEach(function (r) {
        var k = monthKey(r.date);
        if (!buckets[k]) { buckets[k] = { raised: 0, recipients: 0, opens: 0, clicks: 0, donors: 0, sends: 0 }; order.push(k); }
        var b = buckets[k];
        b.sends++;
        b.raised += r.raised || 0;
        b.recipients += r.recipients || 0;
        // Rates are per-send fractions; weight them by Recipients to aggregate honestly.
        if (typeof r.openRate === 'number' && typeof r.recipients === 'number') b.opens += r.openRate * r.recipients;
        if (typeof r.clickRate === 'number' && typeof r.recipients === 'number') b.clicks += r.clickRate * r.recipients;
        b.donors += r.donors || 0;
      });
      order.sort();
      var zeroRecipients = e.rows.filter(function (r) { return !r.recipients; }).length;

      // Notes are scoped per metric: the rate-weighting caveat belongs to the rates
      // chart, not to revenue-per-1k, which never touches a rate column.
      var sharedNotes = inheritedWarnings(email, ['raised', 'recipients', 'openRate', 'clickRate', 'donors', 'date']);
      var revNotes = ['Sends with zero recipients are excluded from this ratio rather than counted as $0 — they would otherwise drag the average down without representing a real send.']
        .concat(sharedNotes);
      var rateNotes = ['Open and click rates are averaged WEIGHTED BY RECIPIENTS, not averaged per send — a 5,000-recipient send and a 500,000-recipient send would otherwise count equally.',
        'Donate rate is recomputed from Donors ÷ Recipients rather than read from the Donate Rate column, so it stays consistent with the weighting above. The sheet\'s own Donate Rate and Unsub Rate columns are shown unaltered in the sheet table below.']
        .concat(sharedNotes);
      var missing = [];
      ['openRate', 'clickRate', 'donateRate'].forEach(function (k) {
        if (!email.columnLetters[k]) missing.push(k);
      });

      var baseProv = {
        tab: email.tab, headerRow: email.headerRow, dataRange: email.dataRange,
        rowsAvailable: email.rows.length, rowsUsed: e.rows.length,
        rowsExcluded: mergeExcluded(e.excluded, email.rowsExcluded,
          zeroRecipients ? [{ reason: 'Recipients blank or zero (excluded from per-1k and rate maths)', count: zeroRecipients }] : [])
      };

      return {
        months: order, labels: order.map(monthLabel),
        missingColumns: missing,
        revenuePerK: {
          values: order.map(function (k) {
            var b = buckets[k];
            return b.recipients > 0 ? (b.raised / b.recipients) * 1000 : null;
          }),
          provenance: prov(Object.assign({}, baseProv, {
            columnsUsed: cols(email, [{ key: 'date', label: 'Date', role: 'grouped by month' },
                                      { key: 'raised', label: 'Raised', role: 'numerator' },
                                      { key: 'recipients', label: 'Recipients', role: 'denominator' }]),
            transform: 'SUM(Raised) / SUM(Recipients) × 1000, per month',
            notes: revNotes
          }))
        },
        rates: {
          open: order.map(function (k) { return buckets[k].recipients > 0 ? buckets[k].opens / buckets[k].recipients : null; }),
          click: order.map(function (k) { return buckets[k].recipients > 0 ? buckets[k].clicks / buckets[k].recipients : null; }),
          donate: order.map(function (k) { return buckets[k].recipients > 0 ? buckets[k].donors / buckets[k].recipients : null; }),
          provenance: prov(Object.assign({}, baseProv, {
            columnsUsed: cols(email, [{ key: 'date', label: 'Date', role: 'grouped by month' },
                                      { key: 'openRate', label: 'Open Rate', role: 'weighted average' },
                                      { key: 'clickRate', label: 'Click Rate', role: 'weighted average' },
                                      { key: 'donors', label: 'Donors', role: 'numerator for donate rate' },
                                      { key: 'recipients', label: 'Recipients', role: 'weight / denominator' }]),
            transform: 'SUM(rate × Recipients) / SUM(Recipients) per month; donate rate = SUM(Donors) / SUM(Recipients)',
            recomputed: true,
            notes: rateNotes.concat(missing.length
              ? ['Column(s) not found in this workbook: ' + missing.join(', ') + '. Those lines are omitted.'] : [])
          }))
        }
      };
    }

    // -------------------------------------------------------------------------
    // Panel 4 — list fatigue: send cadence against unsub rate
    // -------------------------------------------------------------------------

    function panelListFatigue(f) {
      var e = emailScope(f);
      var buckets = {}, order = [];
      e.rows.forEach(function (r) {
        var k = weekKey(r.date);
        if (!buckets[k]) { buckets[k] = { sends: 0, unsubs: 0, recipients: 0 }; order.push(k); }
        buckets[k].sends++;
        buckets[k].unsubs += r.unsubs || 0;
        buckets[k].recipients += r.recipients || 0;
      });
      order.sort();
      return {
        weeks: order,
        labels: order.map(function (k) {
          var d = new Date(k + 'T00:00:00Z');
          return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getUTCMonth()] + ' ' + d.getUTCDate();
        }),
        sends: order.map(function (k) { return buckets[k].sends; }),
        unsubRate: order.map(function (k) {
          return buckets[k].recipients > 0 ? buckets[k].unsubs / buckets[k].recipients : null;
        }),
        provenance: prov({
          tab: email.tab, headerRow: email.headerRow, dataRange: email.dataRange,
          columnsUsed: cols(email, [{ key: 'date', label: 'Date', role: 'grouped into ISO weeks' },
                                    { key: 'unsubs', label: 'Unsubs', role: 'numerator' },
                                    { key: 'recipients', label: 'Recipients', role: 'denominator' }]),
          rowsAvailable: email.rows.length, rowsUsed: e.rows.length,
          rowsExcluded: mergeExcluded(e.excluded, email.rowsExcluded),
          transform: 'Sends per ISO week (count of rows) plotted against SUM(Unsubs) / SUM(Recipients) for the same week',
          notes: ['Weeks start Monday. A week with no sends produces no point rather than a zero.']
            .concat(inheritedWarnings(email, ['unsubs', 'recipients', 'date']))
        })
      };
    }

    // -------------------------------------------------------------------------
    // Panel 5 — email breakdown by dimension. Every group row carries its own bubble.
    // -------------------------------------------------------------------------

    var DIM_LABELS = { audience: 'Audience', ask: 'Ask', subtype: 'Subtype', sender: 'Sender' };

    function panelEmailBreakdown(f, dim) {
      var e = emailScope(f);
      var groups = {}, order = [];
      e.rows.forEach(function (r) {
        var k = r[dim] || '(blank)';
        if (!groups[k]) { groups[k] = { rows: [], raised: 0, recipients: 0, donors: 0, unsubs: 0, opens: 0, clicks: 0 }; order.push(k); }
        var g = groups[k];
        g.rows.push(r);
        g.raised += r.raised || 0;
        g.recipients += r.recipients || 0;
        g.donors += r.donors || 0;
        g.unsubs += r.unsubs || 0;
        if (typeof r.openRate === 'number' && typeof r.recipients === 'number') g.opens += r.openRate * r.recipients;
        if (typeof r.clickRate === 'number' && typeof r.recipients === 'number') g.clicks += r.clickRate * r.recipients;
      });

      var rows = order.map(function (k) {
        var g = groups[k];
        return {
          key: k, sends: g.rows.length, recipients: g.recipients, raised: g.raised,
          revPerK: g.recipients > 0 ? (g.raised / g.recipients) * 1000 : null,
          openRate: g.recipients > 0 ? g.opens / g.recipients : null,
          clickRate: g.recipients > 0 ? g.clicks / g.recipients : null,
          donateRate: g.recipients > 0 ? g.donors / g.recipients : null,
          unsubRate: g.recipients > 0 ? g.unsubs / g.recipients : null,
          // Each group row gets provenance scoped to that group's own rows.
          provenance: prov({
            tab: email.tab, headerRow: email.headerRow,
            dataRange: rangeOfRows(email, g.rows),
            columnsUsed: cols(email, [{ key: dim, label: DIM_LABELS[dim], role: 'grouped by' },
                                      { key: 'recipients', label: 'Recipients', role: 'denominator' },
                                      { key: 'raised', label: 'Raised', role: 'numerator' },
                                      { key: 'donors', label: 'Donors', role: 'donate rate numerator' },
                                      { key: 'unsubs', label: 'Unsubs', role: 'unsub rate numerator' }]),
            rowsAvailable: email.rows.length, rowsUsed: g.rows.length,
            rowsExcluded: mergeExcluded(
              [{ reason: 'belongs to a different ' + DIM_LABELS[dim] + ' group', count: e.rows.length - g.rows.length }],
              e.excluded, email.rowsExcluded),
            transform: 'Rows where ' + DIM_LABELS[dim] + ' = "' + k + '": SUM(Raised)/SUM(Recipients)×1000; rates weighted by Recipients',
            notes: ['Excel rows for this group: ' + rowNumbers(g.rows) + '.']
              .concat(inheritedWarnings(email, [dim, 'raised', 'recipients', 'donors', 'unsubs']))
          })
        };
      });
      rows.sort(function (a, b) { return b.raised - a.raised; });

      return {
        dim: dim, dimLabel: DIM_LABELS[dim], rows: rows,
        provenance: prov({
          tab: email.tab, headerRow: email.headerRow, dataRange: email.dataRange,
          columnsUsed: cols(email, [{ key: dim, label: DIM_LABELS[dim], role: 'grouped by' },
                                    { key: 'raised', label: 'Raised', role: 'summed' },
                                    { key: 'recipients', label: 'Recipients', role: 'summed' }]),
          rowsAvailable: email.rows.length, rowsUsed: e.rows.length,
          rowsExcluded: mergeExcluded(e.excluded, email.rowsExcluded),
          transform: 'Group by ' + DIM_LABELS[dim] + ', then SUM(Raised)/SUM(Recipients)×1000 per group',
          notes: (dim === 'sender'
            ? ['Sender values are shown exactly as they appear in the workbook. Near-duplicates such as "Booker HQ" and "Booker HQ Finance Team" are NOT merged automatically — they may be legitimately distinct senders.']
            : []).concat(inheritedWarnings(email, [dim, 'raised', 'recipients']))
        })
      };
    }

    /** Resolved Excel range covering a specific subset of rows. */
    function rangeOfRows(table, rows) {
      if (!rows.length) return '(no rows)';
      var nums = rows.map(function (r) { return r._row; });
      var letters = Object.keys(table.columnLetters || {}).map(function (k) { return table.columnLetters[k]; });
      var minL = letters.slice().sort()[0] || 'A';
      var maxL = letters.slice().sort(function (a, b) { return a.length - b.length || (a < b ? -1 : 1); }).pop() || 'A';
      return minL + Math.min.apply(null, nums) + ':' + maxL + Math.max.apply(null, nums) + ' (' + rows.length + ' non-contiguous rows)';
    }

    function rowNumbers(rows) {
      var nums = rows.map(function (r) { return r._row; }).sort(function (a, b) { return a - b; });
      if (nums.length <= 12) return nums.join(', ');
      return nums.slice(0, 10).join(', ') + ', … (' + (nums.length - 10) + ' more)';
    }

    // -------------------------------------------------------------------------
    // Panel 6 — Ads. Ships as an explanatory empty state and lights up on its own
    // if revenue data ever appears.
    // -------------------------------------------------------------------------

    function panelAds(f) {
      var rows = ads.rows || [];
      var hasRevenue = rows.some(function (r) {
        return ['lifetimeRaisedStandard', 'lifetimeRoasStandard', 'lifetimeRaisedFinanceAdj', 'lifetimeRoasFinanceAdj']
          .some(function (k) { return typeof r[k] === 'number'; });
      });
      var monthsSeen = {};
      rows.forEach(function (r) { if (r.month instanceof Date) monthsSeen[monthLabel(monthKey(r.month))] = true; });
      var monthList = Object.keys(monthsSeen);

      var emptyReason = null;
      if (!ads.available || !rows.length) {
        emptyReason = 'The Ads Report tab produced no rows.';
      } else if (!hasRevenue) {
        emptyReason = 'No ads revenue data in source (' + rows.length + ' row' + (rows.length === 1 ? '' : 's') +
          ' found, all ' + monthList.join(' / ') + '; columns ' +
          [ads.columnLetters.lifetimeRaisedStandard, ads.columnLetters.lifetimeRoasStandard,
           ads.columnLetters.lifetimeRaisedFinanceAdj, ads.columnLetters.lifetimeRoasFinanceAdj].join('/') +
          ' are blank). This panel fills in automatically once revenue is populated.';
      }

      return {
        hasRevenue: hasRevenue, emptyReason: emptyReason,
        rows: rows.map(function (r) {
          return {
            placement: r.placement, month: r.month ? monthLabel(monthKey(r.month)) : '(no month)',
            goal: r.goal, grossSpend: r.grossSpend, ntl: r.ntl, grossCpa: r.grossCpa,
            raisedStandard: r.lifetimeRaisedStandard, roasStandard: r.lifetimeRoasStandard,
            raisedFinanceAdj: r.lifetimeRaisedFinanceAdj, roasFinanceAdj: r.lifetimeRoasFinanceAdj
          };
        }),
        provenance: prov({
          tab: ads.tab, headerRow: ads.headerRow, dataRange: ads.dataRange,
          columnsUsed: cols(ads, [{ key: 'month', label: 'Month', role: 'grouped by' },
                                  { key: 'placement', label: 'Placement', role: 'grouped by' },
                                  { key: 'goal', label: 'Goal', role: 'listed' },
                                  { key: 'grossSpend', label: 'Gross Spend', role: 'plotted' },
                                  { key: 'ntl', label: 'NTL', role: 'plotted' },
                                  { key: 'grossCpa', label: 'Gross CPA', role: 'plotted' },
                                  { key: 'lifetimeRaisedStandard', label: 'Lifetime Raised (Standard Toplines)', role: 'plotted' },
                                  { key: 'lifetimeRoasStandard', label: 'Lifetime ROAS (Standard Toplines)', role: 'plotted' },
                                  { key: 'lifetimeRaisedFinanceAdj', label: 'Lifetime Raised (Finance-Adjusted)', role: 'plotted' },
                                  { key: 'lifetimeRoasFinanceAdj', label: 'Lifetime ROAS (Finance-Adjusted)', role: 'plotted' }]),
          rowsAvailable: rows.length, rowsUsed: rows.length,
          rowsExcluded: ads.rowsExcluded,
          transform: 'Values read directly per placement. No aggregation — there are too few rows to aggregate.',
          notes: ['Row 3 of this tab repeats the names "Lifetime Raised" and "Lifetime ROAS" twice, under the merged row-1 group labels "Standard Toplines" (H:I) and "Finance-Adjusted" (J:K). These four columns are matched BY POSITION, not by name, so the two pairs cannot be swapped.',
                  'This tab is not affected by the date filter — its rows sit far outside any 2026 window and would otherwise vanish without explanation.']
            .concat(inheritedWarnings(ads, ['grossSpend', 'ntl', 'grossCpa', 'month', 'placement']))
        })
      };
    }

    // -------------------------------------------------------------------------
    // Panel 7 — P2P. Small sample; the count is stated on the panel face.
    // -------------------------------------------------------------------------

    function panelP2P(f) {
      var p = p2pScope(f);
      return {
        n: p.rows.length,
        totalAvailable: p2p.rows.length,
        smallSample: p2p.rows.length < 30,
        rows: p.rows.map(function (r) {
          return {
            date: r.date, topic: r.topic, audience: r.audience, recipients: r.recipients,
            grossSpend: r.grossSpend, immediateRaised: r.immediateRaised,
            immediateRoas: r.immediateRoas, ltvRoas: r.ltvRoasFromNtl,
            grossCpa: r.grossCpa, ntl: r.ntl, _row: r._row
          };
        }),
        provenance: prov({
          tab: p2p.tab, headerRow: p2p.headerRow, dataRange: p2p.dataRange,
          columnsUsed: cols(p2p, [{ key: 'date', label: 'Date', role: 'x axis' },
                                  { key: 'client', label: 'Client', role: 'listed' },
                                  { key: 'goal', label: 'Goal', role: 'listed' },
                                  { key: 'topic', label: 'Topic', role: 'listed' },
                                  { key: 'audience', label: 'Audience', role: 'listed' },
                                  { key: 'recipients', label: 'Recipients', role: 'listed' },
                                  { key: 'grossSpend', label: 'Gross Spend', role: 'plotted' },
                                  { key: 'immediateRaised', label: 'Immediate Raised', role: 'listed' },
                                  { key: 'immediateRoas', label: 'Immediate ROAS', role: 'plotted' },
                                  { key: 'clickRate', label: 'Click Rate', role: 'listed' },
                                  { key: 'donors', label: 'Donors', role: 'listed' },
                                  { key: 'avgGift', label: 'Average Gift', role: 'listed' },
                                  { key: 'donateRate', label: 'Donate Rate', role: 'listed' },
                                  { key: 'ntl', label: 'NTL', role: 'plotted' },
                                  { key: 'grossCpa', label: 'Gross CPA', role: 'plotted' },
                                  { key: 'ltvFromNtl', label: 'LTV from NTL', role: 'listed' },
                                  { key: 'ltvRoasFromNtl', label: 'LTV ROAS from NTL', role: 'plotted' },
                                  { key: 'immediateRaisedFromNtl', label: 'Immediate Raised from NTL', role: 'listed' },
                                  { key: 'unsubs', label: 'Unsubs', role: 'listed' },
                                  { key: 'unsubRate', label: 'Unsub Rate', role: 'listed' },
                                  { key: 'sourceCode', label: 'Source Code', role: 'listed' }]),
          rowsAvailable: p2p.rows.length, rowsUsed: p.rows.length,
          rowsExcluded: mergeExcluded(p.excluded, p2p.rowsExcluded),
          transform: 'Values read directly per send. No aggregation — every row is plotted individually.',
          notes: ['Only ' + p2p.rows.length + ' P2P sends exist in this workbook, covering ' + (p2p.entry && p2p.entry.dateRange) + '. Treat any trend here as indicative, not conclusive.',
                  'Audience values in this tab carry a leading space in the source and are trimmed at parse time.']
            .concat(inheritedWarnings(p2p, ['immediateRoas', 'ltvRoasFromNtl', 'grossCpa', 'ntl', 'grossSpend', 'date']))
        })
      };
    }

    // -------------------------------------------------------------------------
    // Panel 8 — High-dollar aggregates. Aggregates only, by construction.
    // -------------------------------------------------------------------------

    function panelHighDollar() {
      var hd = data.highDollarAgg;
      if (!hd || !hd.available || !hd.agg) {
        return { available: false, emptyReason: 'The High-Dollar Donations tab could not be read.', provenance: prov({
          tab: 'High-Dollar Donations', headerRow: null, dataRange: null, columnsUsed: [],
          rowsAvailable: 0, rowsUsed: 0, rowsExcluded: [], transform: 'n/a', notes: [] }) };
      }
      var agg = hd.agg;
      function toList(obj) {
        return Object.keys(obj).map(function (k) { return { key: k, count: obj[k].count, sum: obj[k].sum }; })
          .sort(function (a, b) { return b.sum - a.sum; });
      }
      return {
        available: true,
        total: agg.total,
        smallSample: agg.total.count < 50,
        byCategory: toList(agg.byCategory),
        byBand: agg.bands.map(function (b) {
          var e = agg.byBand[b] || { count: 0, sum: 0 };
          return { key: b, count: e.count, sum: e.sum };
        }),
        byMonth: Object.keys(agg.byMonth).sort().map(function (k) {
          return { key: k === '(no date)' ? k : monthLabel(k), count: agg.byMonth[k].count, sum: agg.byMonth[k].sum };
        }),
        claims: { financeClaim: toList(agg.byFinanceClaim), authenticAdded: toList(agg.byAuthenticAdded) },
        droppedColumns: hd.droppedColumns,
        provenance: prov({
          tab: hd.tab, headerRow: hd.headerRow, dataRange: hd.dataRange,
          columnsUsed: [
            { name: 'Category', letter: hd.columnLetters.category || 'A', role: 'grouped by' },
            { name: 'date', letter: hd.columnLetters.date || 'C', role: 'grouped by month' },
            { name: 'amount', letter: hd.columnLetters.amount || 'D', role: 'summed and banded' },
            { name: 'Finance Claim', letter: hd.columnLetters.financeClaim || 'I', role: 'grouped by (claim status)' },
            { name: 'Authentic Added', letter: hd.columnLetters.authenticAdded || 'J', role: 'grouped by (claim status)' }
          ],
          rowsAvailable: hd.rowsAvailable, rowsUsed: agg.total.count,
          rowsExcluded: hd.rowsExcluded,
          transform: 'COUNT(rows) and SUM(amount), grouped by Category, by calendar month, and by amount band',
          notes: [
            'PII DROPPED AT PARSE TIME: ' + hd.droppedColumns.join(', ') + '. These columns are never read into dashboard state, so they cannot appear in a tooltip, an export, or a console dump. This is why no drill-through to individual donations exists.',
            'Only ' + hd.rowsAvailable + ' donation rows are present. The tab is pre-formatted down to row 500 and columns I and J are filled to the bottom, which makes it look much larger than it is — row counting is anchored on the date and amount columns instead.'
          ].concat(hd.warnings.map(function (w) { return w.text; }))
        })
      };
    }

    // -------------------------------------------------------------------------

    // -------------------------------------------------------------------------
    // Coverage panels — so every column in the workbook is reflected somewhere,
    // not just the ones that make the headline charts.
    // -------------------------------------------------------------------------

    var DOW_ORDER = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];

    /** Email performance by day of week — uses the Day of Week column directly. */
    function panelEmailByDayOfWeek(f) {
      var e = emailScope(f);
      var g = {};
      e.rows.forEach(function (r) {
        var k = r.dayOfWeek || (r.date ? DOW_ORDER[(r.date.getUTCDay() + 6) % 7] : '(unknown)');
        if (!g[k]) g[k] = { sends: 0, recipients: 0, raised: 0, opens: 0, unsubs: 0 };
        g[k].sends++;
        g[k].recipients += r.recipients || 0;
        g[k].raised += r.raised || 0;
        g[k].unsubs += r.unsubs || 0;
        if (typeof r.openRate === 'number' && typeof r.recipients === 'number') g[k].opens += r.openRate * r.recipients;
      });
      // The workbook writes abbreviated day names ("Wed"), so order on the first
      // three letters rather than requiring an exact match.
      function dowIndex(k) {
        var i = DOW_ORDER.map(function (d) { return d.slice(0, 3).toLowerCase(); })
          .indexOf(String(k).slice(0, 3).toLowerCase());
        return i === -1 ? 99 : i;
      }
      var keys = Object.keys(g).sort(function (a, b) { return dowIndex(a) - dowIndex(b); });
      return {
        labels: keys,
        rows: keys.map(function (k) {
          var b = g[k];
          return { key: k, sends: b.sends, recipients: b.recipients, raised: b.raised,
                   revPerK: b.recipients > 0 ? (b.raised / b.recipients) * 1000 : null,
                   openRate: b.recipients > 0 ? b.opens / b.recipients : null,
                   unsubRate: b.recipients > 0 ? b.unsubs / b.recipients : null };
        }),
        provenance: prov({
          tab: email.tab, headerRow: email.headerRow, dataRange: email.dataRange,
          columnsUsed: cols(email, [{ key: 'dayOfWeek', label: 'Day of Week', role: 'grouped by' },
                                    { key: 'raised', label: 'Raised', role: 'numerator' },
                                    { key: 'recipients', label: 'Recipients', role: 'denominator' },
                                    { key: 'openRate', label: 'Open Rate', role: 'weighted average' },
                                    { key: 'unsubs', label: 'Unsubs', role: 'unsub rate numerator' }]),
          rowsAvailable: email.rows.length, rowsUsed: e.rows.length,
          rowsExcluded: mergeExcluded(e.excluded, email.rowsExcluded),
          transform: 'Group by Day of Week, then SUM(Raised)/SUM(Recipients)×1000; rates weighted by Recipients',
          notes: ['Day of Week is read from the workbook column, not recomputed from Date, so it matches what the sheet says.']
            .concat(inheritedWarnings(email, ['dayOfWeek', 'raised', 'recipients']))
        })
      };
    }

    /** Actions and action rate — the advocacy side of Email Statistics. */
    function panelEmailActions(f) {
      var e = emailScope(f);
      var buckets = {}, order = [];
      e.rows.forEach(function (r) {
        var k = monthKey(r.date);
        if (!buckets[k]) { buckets[k] = { actions: 0, recipients: 0, donors: 0, avgSum: 0, avgN: 0 }; order.push(k); }
        buckets[k].actions += r.actions || 0;
        buckets[k].recipients += r.recipients || 0;
        buckets[k].donors += r.donors || 0;
        if (typeof r.average === 'number' && r.average > 0) { buckets[k].avgSum += r.average * (r.donors || 1); buckets[k].avgN += (r.donors || 1); }
      });
      order.sort();
      return {
        months: order, labels: order.map(monthLabel),
        actions: order.map(function (k) { return buckets[k].actions; }),
        actionRate: order.map(function (k) { return buckets[k].recipients > 0 ? buckets[k].actions / buckets[k].recipients : null; }),
        avgGift: order.map(function (k) { return buckets[k].avgN > 0 ? buckets[k].avgSum / buckets[k].avgN : null; }),
        provenance: prov({
          tab: email.tab, headerRow: email.headerRow, dataRange: email.dataRange,
          columnsUsed: cols(email, [{ key: 'date', label: 'Date', role: 'grouped by month' },
                                    { key: 'actions', label: 'Actions', role: 'summed' },
                                    { key: 'actionRate', label: 'Action Rate', role: 'recomputed from Actions / Recipients' },
                                    { key: 'average', label: 'Average', role: 'weighted by Donors' },
                                    { key: 'donors', label: 'Donors', role: 'weight' },
                                    { key: 'recipients', label: 'Recipients', role: 'denominator' }]),
          rowsAvailable: email.rows.length, rowsUsed: e.rows.length,
          rowsExcluded: mergeExcluded(e.excluded, email.rowsExcluded),
          transform: 'SUM(Actions) per month; action rate = SUM(Actions)/SUM(Recipients); average gift weighted by Donors',
          notes: ['Action rate is recomputed from Actions and Recipients rather than averaging the per-send Action Rate column, so months with uneven send sizes are not distorted.']
            .concat(inheritedWarnings(email, ['actions', 'actionRate', 'average', 'donors', 'recipients']))
        })
      };
    }

    /** A simple grouped count for the ops/reference tabs. */
    function groupCount(table, key, label) {
      var g = {}, order = [];
      (table.rows || []).forEach(function (r) {
        var k = (r[key] == null || r[key] === '') ? '(blank)' : String(r[key]);
        if (!g[k]) { g[k] = 0; order.push(k); }
        g[k]++;
      });
      return order.sort(function (a, b) { return g[b] - g[a]; }).map(function (k) { return { key: k, count: g[k] }; });
    }

    function panelSimpleTab(id, groupings) {
      var table = data[id];
      if (!table || !table.available) return { available: false };
      return {
        available: true,
        rows: table.rows,
        total: table.rows.length,
        groups: groupings.map(function (gr) {
          return { key: gr.key, label: gr.label, counts: groupCount(table, gr.key, gr.label) };
        }).filter(function (gr) { return gr.counts.length > 0; }),
        provenance: prov({
          tab: table.tab, headerRow: table.headerRow, dataRange: table.dataRange,
          columnsUsed: cols(table, groupings.map(function (gr) {
            return { key: gr.key, label: gr.label, role: 'grouped by' };
          })),
          rowsAvailable: table.rows.length, rowsUsed: table.rows.length,
          rowsExcluded: table.rowsExcluded,
          transform: 'COUNT(rows) grouped by ' + groupings.map(function (g) { return g.label; }).join(', '),
          notes: ['This is an operations tab. It is counted and listed, not charted — it records what was scheduled, not what it earned.']
            .concat(inheritedWarnings(table, groupings.map(function (g) { return g.key; })))
        })
      };
    }

    /** Projection scenarios, each with its own chartable series. */
    function panelProjections() {
      return (data.projectionScenarios || []).map(function (sc) {
        var rows = sc.rows.filter(function (r) { return r.month instanceof Date; });
        return {
          scenario: sc.scenario, usable: sc.usable, unusableReason: sc.unusableReason,
          headerRow: sc.headerRow, dataRange: sc.dataRange,
          labels: rows.map(function (r) { return monthLabel(monthKey(r.month)); }),
          channels: data.channelKeys.map(function (k) {
            return { key: k, values: rows.map(function (r) { return r[k] || 0; }) };
          }).filter(function (c) { return c.values.some(function (v) { return v !== 0; }); }),
          rows: rows,
          provenance: prov({
            tab: sc.tab, headerRow: sc.headerRow, dataRange: sc.dataRange,
            columnsUsed: cols(sc, [{ key: 'month', label: 'Month', role: 'x axis' }].concat(
              data.channelKeys.map(function (k) { return { key: k, label: k, role: 'plotted' }; }))),
            rowsAvailable: sc.rowsAvailable, rowsUsed: rows.length,
            rowsExcluded: sc.rowsExcluded,
            transform: 'Channel columns read directly for this block, one row per month',
            notes: (sc.usable ? [] : ['Not plotted: ' + sc.unusableReason + '.'])
              .concat(['This is one of three scenario blocks stacked in the Digital Projections tab, starting at row ' + sc.headerRow + '.'])
          })
        };
      });
    }

    function facets() {
      function distinct(rows, key) {
        var seen = {}, out = [];
        rows.forEach(function (r) { var v = r[key] || '(blank)'; if (!seen[v]) { seen[v] = true; out.push(v); } });
        return out.sort();
      }
      return {
        ask: distinct(email.rows || [], 'ask'),
        subtype: distinct(email.rows || [], 'subtype'),
        audience: distinct(email.rows || [], 'audience'),
        sender: distinct(email.rows || [], 'sender'),
        clients: distinct(p2p.rows || [], 'client'),
        dateExtent: (function () {
          var ds = (email.rows || []).concat(p2p.rows || [])
            .map(function (r) { return r.date; }).filter(function (d) { return d instanceof Date; });
          if (!ds.length) return null;
          return { min: new Date(Math.min.apply(null, ds)), max: new Date(Math.max.apply(null, ds)) };
        })()
      };
    }

    return {
      facets: facets,
      kpis: function (f) {
        return [
          { id: 'K1', label: 'Gross Raised', metric: kpiGrossRaised(f) },
          { id: 'K2', label: 'Gross Spend', metric: kpiGrossSpend(f) },
          { id: 'K3', label: 'Net Raised', metric: kpiNetRaised(f) },
          { id: 'K4', label: 'Blended ROAS', metric: kpiRoas(f) },
          { id: 'K5', label: '% to goal', metric: kpiPctToGoal(f) }
        ];
      },
      panelActualVsGoal: panelActualVsGoal,
      panelChannelMix: panelChannelMix,
      panelGoalTotals: panelGoalTotals,
      panelEmailTrend: panelEmailTrend,
      panelListFatigue: panelListFatigue,
      panelEmailBreakdown: panelEmailBreakdown,
      panelAds: panelAds,
      panelP2P: panelP2P,
      panelHighDollar: panelHighDollar,
      panelEmailByDayOfWeek: panelEmailByDayOfWeek,
      panelEmailActions: panelEmailActions,
      panelProjections: panelProjections,
      panelSimpleTab: panelSimpleTab,
      _util: { monthKey: monthKey, monthLabel: monthLabel, weekKey: weekKey }
    };
  }

  return { create: create };
});
