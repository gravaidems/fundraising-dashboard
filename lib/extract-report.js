/* =====================================================================
   extract-report.js — the single extractor for both dashboard tabs.

   Used in two places, so there is only ever one implementation of "what
   the numbers mean":
     - the browser, when someone drops a new .xlsx onto the page
     - node, at build time, to bake the snapshot into the page

   Everything is discovered from labels and header text rather than fixed
   row and column numbers, so added channels, added placement rows, added
   metric columns and moved columns are all picked up. Anything that cannot
   be found is reported as a problem instead of being guessed at.

   Each extractor returns {payload, findings, problems}:
     findings  what was located, to show in the validation report
     problems  {level: "error"|"warn"|"flag", text}

   The three levels differ in what they cost the reader:
     error  the section cannot be trusted at all — it is not rendered
     warn   something was tolerated (a renamed tab, a month with no rows)
     flag   the section renders, but one stored column disagrees with its
            own components. Carries `sections` (the ids of the sections that
            print the offending cells) and `cells` (specific cell refs), so
            the page can pin the note to the chart or table it concerns
            instead of dropping the tab.
   ===================================================================== */
(function (root, factory) {
  const api = factory(root.XlsxReader || (typeof require === "function" ? require("./xlsx-reader.js") : null));
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.ExtractReport = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function (XlsxReader) {
  "use strict";

  const numToCol = XlsxReader.numToCol, colToNum = XlsxReader.colToNum;

  const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase().replace(/:$/, "");
  const slug = s => norm(s).replace(/[^a-z0-9]+/g, "_").replace(/^_|_$/g, "");
  const MAX_SCAN = 200;   // rows to scan when hunting for an anchor label

  /* Find the first row in `col` whose text matches `re`. */
  function findRow(sheet, col, re, limit) {
    const max = Math.min(sheet.maxRow || MAX_SCAN, limit || MAX_SCAN);
    for (let r = 1; r <= max; r++) {
      const t = sheet.txt(col + r);
      if (t && re.test(norm(t))) return r;
    }
    return 0;
  }
  /* Value in `valCol` on the row whose `labelCol` text matches `re`. */
  function labelled(sheet, labelCol, valCol, re) {
    const r = findRow(sheet, labelCol, re);
    return r ? {value: sheet.val(valCol + r), num: sheet.num(valCol + r),
                txt: sheet.txt(valCol + r), cell: valCol + r, row: r} : null;
  }
  /* Header text -> column letter, across a row. */
  function headerMap(sheet, row, fromCol, toCol) {
    const map = {}, order = [];
    for (let c = colToNum(fromCol); c <= colToNum(toCol); c++) {
      const col = numToCol(c), t = sheet.txt(col + row);
      if (!t) continue;
      const k = norm(t);
      if (!(k in map)) { map[k] = col; order.push({key: k, label: t.trim(), col: col}); }
    }
    return {map: map, order: order};
  }

  /* ===================================================================
     DIGITAL REPORT
     =================================================================== */
  function extractDigital(sheet, opts) {
    const findings = [], problems = [];
    const err = t => problems.push({level: "error", text: t});
    const warn = t => problems.push({level: "warn", text: t});

    /* --- anchors --- */
    const hdrRow = findRow(sheet, "B", /^quarterly toplines$/);
    if (!hdrRow) {
      err("Could not find the “Quarterly Toplines” heading in column B, so this tab could not be read.");
      return {payload: null, findings: findings, problems: problems};
    }
    const monthRow = hdrRow, labelRow = hdrRow + 1;
    findings.push("Found “Quarterly Toplines” at B" + hdrRow + "; column headers on row " + labelRow + ".");

    /* --- period column triples: Goal / Actual / % to Goal --- */
    const periods = [];
    for (let c = 2; c <= Math.max(sheet.maxCol, 20); c++) {
      const g = numToCol(c);
      if (norm(sheet.txt(g + labelRow)) !== "goal") continue;
      const a = numToCol(c + 1), p = numToCol(c + 2);
      if (norm(sheet.txt(a + labelRow)) !== "actual") continue;
      if (!/^%/.test(norm(sheet.txt(p + labelRow)) || "")) continue;
      const mLabel = sheet.txt(g + monthRow);
      periods.push({goalCol: g, actualCol: a, pctCol: p, monthLabel: mLabel});
    }
    if (!periods.length) {
      err("Row " + labelRow + " has no Goal / Actual / % to Goal column group, so this tab could not be read.");
      return {payload: null, findings: findings, problems: problems};
    }
    // the first group is the quarter total; the rest are months, named from the row above
    const quarterP = periods[0];
    const monthPs = periods.slice(1).filter(p => p.monthLabel);
    if (periods.length > 1 && monthPs.length !== periods.length - 1) {
      warn("A month column group has no month name on row " + monthRow + " and was skipped.");
    }
    if (!monthPs.length) err("No month column groups were found next to the quarterly group.");
    const monthKeys = monthPs.map(p => slug(p.monthLabel));
    findings.push("Found " + monthPs.length + " month column group(s): " +
      monthPs.map(p => p.monthLabel.trim() + " (" + p.goalCol + "/" + p.actualCol + "/" + p.pctCol + ")").join(", ") + ".");

    /* --- channel rows: from under the header down to the Overall block --- */
    const channelRows = [];
    for (let r = labelRow + 1; r <= sheet.maxRow; r++) {
      const t = sheet.txt("B" + r);
      if (!t) break;
      if (/^overall$/.test(norm(t))) break;
      channelRows.push(r);
    }
    if (!channelRows.length) err("No channel rows were found under row " + labelRow + ".");
    findings.push("Found " + channelRows.length + " channel row(s): rows " +
      channelRows[0] + "–" + channelRows[channelRows.length - 1] + ".");

    /* --- helper columns behind the quarterly goal (=SUM(Q14:S14)) --- */
    let helperCols = null;
    const f0 = channelRows.length ? sheet.formula(quarterP.goalCol + channelRows[0]) : null;
    const sumM = f0 && f0.match(/^=\s*SUM\(\s*\$?([A-Z]+)\$?\d+\s*:\s*\$?([A-Z]+)\$?\d+\s*\)/i);
    if (sumM) {
      const a = colToNum(sumM[1]), b = colToNum(sumM[2]);
      helperCols = [];
      for (let c = a; c <= b; c++) helperCols.push(numToCol(c));
      if (helperCols.length !== monthPs.length) {
        warn("The quarterly goal sums " + helperCols.length + " helper column(s) but there are " +
             monthPs.length + " months; helper provenance may be off.");
      }
      findings.push("Quarterly goal is =SUM(" + sumM[1] + ":" + sumM[2] + "); helper columns " +
        helperCols.join(", ") + ".");
    } else {
      warn("Could not read the quarterly goal formula, so hidden helper columns are not shown.");
    }

    /* --- dynamic goal toggles, e.g. N9 "July:" with a boolean in O9 --- */
    const toggles = {};
    const dgRow = findRow(sheet, "N", /display dynamic goals/) ||
                  (function () { // the label may sit in another column
                    for (let c = 6; c <= Math.max(sheet.maxCol, 20); c++) {
                      const r = findRow(sheet, numToCol(c), /display dynamic goals/);
                      if (r) return r;
                    }
                    return 0;
                  })();
    monthPs.forEach(p => {
      const key = slug(p.monthLabel);
      let found = null;
      for (let c = 2; c <= Math.max(sheet.maxCol, 22) && !found; c++) {
        const col = numToCol(c);
        for (let r = 1; r < labelRow; r++) {
          const t = sheet.txt(col + r);
          if (t && norm(t) === norm(p.monthLabel)) {
            const vc = numToCol(c + 1);
            if (sheet.raw(vc + r)) { found = {cell: vc + r, value: sheet.bool(vc + r)}; break; }
          }
        }
      }
      toggles[key] = found || {cell: null, value: false};
    });
    const anyToggleOn = Object.keys(toggles).some(k => toggles[k].value);
    if (anyToggleOn) {
      warn("A “Display Dynamic Goals” toggle is switched on, which makes the quarterly goal use " +
           "actuals in place of goals for that month. Goal figures reflect that.");
    }
    if (dgRow) findings.push("Dynamic goal toggles read from row(s) near " + dgRow + ".");

    /* --- build channel entries --- */
    function periodCells(row) {
      const out = {};
      out.quarter = mkPeriod(row, quarterP, "Quarter");
      monthPs.forEach(p => out[slug(p.monthLabel)] = mkPeriod(row, p, p.monthLabel.trim()));
      return out;
    }
    function mkPeriod(row, p, label) {
      const g = p.goalCol + row, a = p.actualCol + row, pc = p.pctCol + row;
      return {
        label: label,
        goal: sheet.num(g), goal_cell: g,
        actual: sheet.num(a), actual_cell: a,
        pct: sheet.num(pc), pct_cell: pc, pct_raw: sheet.val(pc),
        goal_formula: sheet.formula(g), actual_formula: sheet.formula(a), pct_formula: sheet.formula(pc)
      };
    }

    const channels = channelRows.map(row => {
      const entry = {
        name: sheet.txt("B" + row), name_cell: "B" + row, row: row,
        periods: periodCells(row), helpers: {}
      };
      if (helperCols) {
        monthPs.forEach((p, i) => {
          const col = helperCols[i];
          if (!col) return;
          entry.helpers[slug(p.monthLabel)] =
            {value: sheet.num(col + row), cell: col + row, formula: sheet.formula(col + row)};
        });
      }
      const q = entry.periods.quarter;
      entry.dormant = !(q.goal || q.actual);
      return entry;
    });

    /* --- overall block --- */
    const overallHdr = findRow(sheet, "B", /^overall$/);
    const overall = {};
    const overallKeys = [];
    if (!overallHdr) {
      err("Could not find the “Overall” block in column B.");
    } else {
      for (let r = overallHdr + 1; r <= sheet.maxRow; r++) {
        const t = sheet.txt("B" + r);
        if (!t) break;
        const key = slug(t);
        overall[key] = {label: t.trim(), label_cell: "B" + r, row: r, periods: periodCells(r)};
        overallKeys.push(key);
      }
      findings.push("Found “Overall” block at B" + overallHdr + " with " + overallKeys.length +
        " row(s): " + overallKeys.map(k => overall[k].label).join(", ") + ".");
    }
    if (!overall.total_raised) {
      err("The Overall block has no “Total Raised” row, which the headline stats depend on.");
    }

    /* --- months with no actuals reported anywhere are pending, not zero --- */
    const pending = monthKeys.filter(k => channels.every(c => c.periods[k].actual == null));

    /* --- cumulative series, computed here rather than read --- */
    const cumulative = {};
    channels.forEach(ch => {
      const ca = [], cg = [];
      let runA = 0, runG = 0;
      monthKeys.forEach((k, i) => {
        const p = ch.periods[k];
        runG += p.goal || 0;
        cg.push({month: p.label, value: runG,
                 cells: monthKeys.slice(0, i + 1).map(m => ch.periods[m].goal_cell)});
        if (p.actual == null) ca.push({month: p.label, value: null, pending: true});
        else {
          runA += p.actual;
          ca.push({month: p.label, value: runA, pending: false,
                   cells: monthKeys.slice(0, i + 1).map(m => ch.periods[m].actual_cell)});
        }
      });
      cumulative[ch.name] = {actual: ca, goal: cg};
    });

    /* --- header metadata --- */
    const client = labelled(sheet, "B", "C", /^client$/);
    const lead = labelled(sheet, "B", "C", /^email lead$/);
    const year = labelled(sheet, "B", "C", /^year$/);
    const quarter = labelled(sheet, "B", "C", /^quarter$/);
    if (!year || year.num == null) warn("Could not read the report year.");
    if (!quarter || quarter.num == null) warn("Could not read the report quarter.");

    const titleRow = findRow(sheet, "B", /toplines$/, hdrRow - 1);

    /* --- source tab per column, from the IMPORTRANGE formulas --- */
    const SRC_NAMES = ["Import - Quarterly Goals", "Quarterly Fundraising Toplines",
                       "Import - Monthly Goals", "Monthly Fundraising Toplines"];
    const sources = {};
    if (channels.length) {
      const first = channels[0];
      Object.keys(first.periods).forEach(pk => {
        ["goal", "actual"].forEach(field => {
          const f = first.periods[pk][field + "_formula"] || "";
          const hit = SRC_NAMES.filter(n => f.indexOf(n) >= 0)[0];
          if (hit) sources[pk + "_" + field] = hit;
        });
      });
    }

    const payload = {
      meta: {
        title: titleRow ? sheet.txt("B" + titleRow) : "Digital Fundraising Toplines",
        client: {value: client ? client.txt : null, cell: client ? client.cell : null},
        email_lead: {value: lead ? lead.txt : null, cell: lead ? lead.cell : null},
        year: {value: year ? year.num : null, cell: year ? year.cell : null},
        quarter: {value: quarter ? quarter.num : null, cell: quarter ? quarter.cell : null},
        dynamic_goals: Object.assign({label: dgRow ? "Display Dynamic Goals:" : null}, toggles),
        source_workbook: opts.workbookName,
        extracted_at: opts.extractedAt,
        sheet: sheet.name
      },
      channels: channels,
      overall: overall,
      overall_keys: overallKeys,
      cumulative: cumulative,
      month_keys: monthKeys,
      pending_months: pending,
      sources: sources
    };
    return {payload: payload, findings: findings, problems: problems};
  }

  /* ===================================================================
     PAID MEDIA REPORT
     =================================================================== */
  function metricKind(key, label) {
    const n = norm(label || key);
    if (/roas/.test(n)) return "roas";
    if (/rate$/.test(n)) return "rate";
    if (/cpa|cost per|average/.test(n)) return "usd2";
    if (/^ntl$|names|count|recipients|donors|unsubs?$|actions?$/.test(n)) return "count";
    return "usd";
  }

  /* Excel's day-number date system, 1900-based (with the historical Feb 1900
     leap-year bug baked in, same as every spreadsheet program). Returns
     "YYYY-MM-DD" or null. Only used for display — nothing is compared against
     this value, since the workbook's own Day of Week / Month text columns are
     read directly. */
  function serialToDate(n) {
    if (n == null || !isFinite(n)) return null;
    const d = new Date(Math.round((n - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }

  function extractPaid(sheet, opts) {
    const findings = [], problems = [];
    const err = t => problems.push({level: "error", text: t});
    const warn = t => problems.push({level: "warn", text: t});

    const qHdr = findRow(sheet, "B", /^quarterly toplines$/);
    const mHdr = findRow(sheet, "B", /^monthly toplines$/);
    const dHdr = findRow(sheet, "B", /^month$/);
    if (!qHdr || !mHdr || !dHdr) {
      err("Could not find the " +
        [!qHdr && "“Quarterly Toplines”", !mHdr && "“Monthly Toplines”", !dHdr && "detail table"]
          .filter(Boolean).join(" and ") + " heading(s) in column B, so this tab could not be read.");
      return {payload: null, findings: findings, problems: problems};
    }
    findings.push("Headings found at B" + qHdr + " (quarterly), B" + mHdr + " (monthly), B" + dHdr + " (detail).");

    /* --- metric columns, discovered from the detail header row --- */
    const lastCol = numToCol(Math.max(sheet.maxCol, colToNum("N")));
    const hm = headerMap(sheet, dHdr, "F", lastCol);
    const metrics = hm.order.filter(h => h.key !== "helper" && !/^helper/.test(h.key));
    if (!metrics.length) {
      err("No metric columns were found on the detail header row " + dHdr + ".");
      return {payload: null, findings: findings, problems: problems};
    }
    const metricOrder = [], metricLabels = {}, metricKinds = {}, metricColumns = {};
    metrics.forEach(h => {
      const key = slug(h.label);
      metricOrder.push(key);
      metricLabels[key] = h.label;
      metricKinds[key] = metricKind(key, h.label);
      metricColumns[key] = h.col;
    });
    findings.push("Found " + metricOrder.length + " metric column(s): " +
      metricOrder.map(k => metricLabels[k] + " (" + metricColumns[k] + ")").join(", ") + ".");
    // the quarterly/monthly blocks should use the same columns; warn if not
    [["quarterly", qHdr], ["monthly", mHdr]].forEach(([what, row]) => {
      const other = headerMap(sheet, row, "F", lastCol);
      const mismatched = metricOrder.filter(k =>
        other.map[norm(metricLabels[k])] && other.map[norm(metricLabels[k])] !== metricColumns[k]);
      if (mismatched.length) {
        warn("The " + what + " table puts " + mismatched.map(k => metricLabels[k]).join(", ") +
             " in a different column from the detail table; figures are read from each table's own header.");
      }
    });

    function metricsAt(row, colsFor) {
      const out = {};
      metricOrder.forEach(key => {
        const col = (colsFor && colsFor[key]) || metricColumns[key];
        const ref = col + row;
        out[key] = {label: metricLabels[key], kind: metricKinds[key],
                    value: sheet.num(ref), cell: ref, formula: sheet.formula(ref)};
      });
      return out;
    }
    // per-block column maps, so a shifted column in one table doesn't corrupt another
    function colsForRow(headerRow) {
      const other = headerMap(sheet, headerRow, "F", lastCol), cols = {};
      metricOrder.forEach(k => { cols[k] = other.map[norm(metricLabels[k])] || metricColumns[k]; });
      return cols;
    }
    const qCols = colsForRow(qHdr), mCols = colsForRow(mHdr), dCols = colsForRow(dHdr);

    /* --- quarterly rows: numeric quarter in column E --- */
    const qLabelCol = (headerMap(sheet, qHdr, "B", "E").map["quarter"]) || "E";
    const quarters = [];
    for (let r = qHdr + 1; r <= sheet.maxRow; r++) {
      const n = sheet.num(qLabelCol + r);
      if (n == null) break;
      // `label` is derived for display ("Q1"); the cell holds the number, so the
      // provenance claim is quarter_cell -> quarter, not label_cell -> label.
      quarters.push({quarter: Math.round(n), quarter_cell: qLabelCol + r,
                     label: "Q" + Math.round(n), label_derived: true,
                     row: r, metrics: metricsAt(r, qCols)});
    }
    if (!quarters.length) err("No quarterly rows were found under B" + qHdr + ".");

    /* --- monthly rows: month name in the same label column --- */
    const mLabelCol = (headerMap(sheet, mHdr, "B", "E").map["month"]) || "E";
    const months = [];
    for (let r = mHdr + 1; r <= sheet.maxRow; r++) {
      const t = sheet.txt(mLabelCol + r);
      if (!t) break;
      months.push({label: t, label_cell: mLabelCol + r, row: r, metrics: metricsAt(r, mCols)});
    }
    if (!months.length) err("No monthly rows were found under B" + mHdr + ".");
    findings.push("Read " + quarters.length + " quarterly row(s) and " + months.length + " monthly row(s).");

    /* --- detail rows: dimension columns from the header, then scan down --- */
    const dimHdr = headerMap(sheet, dHdr, "B", "E");
    const dimCol = {
      month: dimHdr.map["month"] || "B",
      channel: dimHdr.map["channel"] || "C",
      placement: dimHdr.map["placement"] || "D",
      objective: dimHdr.map["goal"] || dimHdr.map["objective"] || "E"
    };
    const objectiveHeader = sheet.txt(dimCol.objective + dHdr);
    const detail = [];
    for (let r = dHdr + 1; r <= sheet.maxRow; r++) {
      const m = sheet.txt(dimCol.month + r);
      if (!m) break;
      detail.push({
        month: m, month_cell: dimCol.month + r,
        channel: sheet.txt(dimCol.channel + r), channel_cell: dimCol.channel + r,
        placement: sheet.txt(dimCol.placement + r), placement_cell: dimCol.placement + r,
        objective: sheet.txt(dimCol.objective + r), objective_cell: dimCol.objective + r,
        row: r, metrics: metricsAt(r, dCols)
      });
    }
    if (!detail.length) err("No detail rows were found under the header on row " + dHdr + ".");
    else findings.push("Read " + detail.length + " detail row(s): rows " + detail[0].row + "–" +
      detail[detail.length - 1].row + ".");
    const missingDims = detail.filter(d => !d.channel || !d.placement).length;
    if (missingDims) warn(missingDims + " detail row(s) are missing a channel or placement name.");

    /* --- definitions key: "Term:" with prose beside it --- */
    const definitions = [];
    for (let r = 1; r < qHdr; r++) {
      for (let c = colToNum("F"); c <= Math.min(colToNum("H"), Math.max(sheet.maxCol, 8)); c++) {
        const tc = numToCol(c), vc = numToCol(c + 1);
        const term = sheet.txt(tc + r), text = sheet.txt(vc + r);
        if (term && text && /:$/.test(term.trim()) && text.length > 12) {
          definitions.push({term: term.trim().replace(/:$/, ""), text: text,
                            term_cell: tc + r, text_cell: vc + r});
          break;
        }
      }
    }
    if (definitions.length) findings.push("Read " + definitions.length + " metric definition(s) from the key.");
    else warn("No metric definitions key was found; info bubbles will omit the definitions.");

    /* --- additive vs ratio metrics --- */
    const additive = metricOrder.filter(k => metricKinds[k] === "usd" || metricKinds[k] === "count");
    const ratios = metricOrder.filter(k => additive.indexOf(k) < 0);

    /* --- year-to-date rollup: sums for additive, recomputed ratios --- */
    const ytd = {quarters: quarters.map(q => q.label), metrics: {}};
    additive.forEach(key => {
      let tot = 0; const refs = [];
      quarters.forEach(q => {
        const m = q.metrics[key];
        if (m && m.value != null) { tot += m.value; refs.push(m.cell); }
      });
      ytd.metrics[key] = {label: metricLabels[key], kind: metricKinds[key],
                          value: tot, cells: refs, computed: true};
    });
    // ratio components by naming convention; anything unresolvable is reported
    function componentsFor(key) {
      if (/_cpa$|^cpa$/.test(key)) {
        const n = additive.filter(k => /spend/.test(k))[0], d = additive.filter(k => /ntl|name/.test(k))[0];
        return n && d ? [n, d] : null;
      }
      if (/_roas$|^roas$/.test(key)) {
        const n = key.replace(/_?roas$/, "") ? additive.filter(k =>
          k.indexOf(key.replace(/_?roas$/, "")) === 0 && /rais/.test(k))[0] : null;
        const d = additive.filter(k => /spend/.test(k))[0];
        return n && d ? [n, d] : null;
      }
      return null;
    }
    ratios.forEach(key => {
      const comp = componentsFor(key);
      if (!comp) {
        warn("Could not work out how to combine “" + metricLabels[key] + "” across quarters, " +
             "so it is omitted from the year-to-date stats rather than averaged.");
        return;
      }
      const [nk, dk] = comp;
      const nv = ytd.metrics[nk].value, dv = ytd.metrics[dk].value;
      ytd.metrics[key] = {
        label: metricLabels[key], kind: metricKinds[key],
        value: dv ? nv / dv : null, computed: true,
        numerator: nk, denominator: dk,
        cells: ytd.metrics[nk].cells.concat(ytd.metrics[dk].cells)
      };
    });

    /* --- cumulative series --- */
    const accumKeys = ["gross_spend", "ntl", "total_raised"].filter(k => metricOrder.indexOf(k) >= 0);
    if (accumKeys.length < 3) {
      const fallback = additive.slice(0, 3);
      warn("Expected Gross Spend, NTL and Total Raised for the accumulation charts; using " +
           fallback.map(k => metricLabels[k]).join(", ") + " instead.");
      while (accumKeys.length < Math.min(3, fallback.length)) {
        const k = fallback[accumKeys.length];
        if (accumKeys.indexOf(k) < 0) accumKeys.push(k);
        else break;
      }
    }
    function cumulate(rows) {
      const out = {};
      accumKeys.forEach(key => {
        let run = 0; const series = [];
        rows.forEach(row => {
          const v = row.metrics[key] ? row.metrics[key].value : null;
          if (v != null) run += v;
          // `value` is a running total computed here; `step_cell` is the single
          // cell behind `step`, so neither field misattributes the other.
          series.push({label: row.label, value: run, computed_running_total: true,
                       step: v, step_cell: row.metrics[key] ? row.metrics[key].cell : null});
        });
        out[key] = series;
      });
      return out;
    }

    const year = labelled(sheet, "B", "C", /^year$/);
    const quarter = labelled(sheet, "B", "C", /^quarter$/);
    const client = labelled(sheet, "B", "C", /^client$/);
    const lead = labelled(sheet, "B", "C", /^paid media lead$/);
    if (!quarter || quarter.num == null) warn("Could not read which quarter this tab is set to.");

    const payload = {
      meta: {
        title: sheet.txt("B" + (findRow(sheet, "B", /paid media report/) || 2)),
        sheet: sheet.name,
        source_workbook: opts.workbookName,
        extracted_at: opts.extractedAt,
        year: {value: year ? year.num : null, cell: year ? year.cell : null},
        quarter: {value: quarter ? quarter.num : null, cell: quarter ? quarter.cell : null},
        client: {value: client ? client.txt : null, cell: client ? client.cell : null},
        lead: {value: lead ? lead.txt : null, cell: lead ? lead.cell : null}
      },
      metric_order: metricOrder,
      metric_labels: metricLabels,
      metric_kinds: metricKinds,
      metric_columns: metricColumns,
      additive: additive,
      accum_keys: accumKeys,
      objective_header: objectiveHeader,
      definitions: definitions,
      quarters: quarters,
      months: months,
      detail: detail,
      ytd: ytd,
      cumulative: {quarterly: cumulate(quarters), monthly: cumulate(months)},
      channels: uniqSorted(detail.map(d => d.channel)),
      objectives: uniqSorted(detail.map(d => d.objective))
    };
    return {payload: payload, findings: findings, problems: problems};
  }

  function uniqSorted(arr) {
    const seen = {}, out = [];
    arr.forEach(v => { if (v && !seen[v]) { seen[v] = 1; out.push(v); } });
    return out.sort((a, b) => a.localeCompare(b));
  }

  /* ===================================================================
     P2P STATISTICS
     The sheet is a single flat detail table (one row per date × topic ×
     audience segment) — there is no separate quarterly/monthly rollup block
     like Paid Media has, so month and quarter-to-date totals are computed
     here from the detail rows rather than read from a stored cell.
     =================================================================== */
  const P2P_DIM_KEYS = ["client", "year", "quarter", "month", "day of week", "date", "goal", "topic", "audience"];

  /* The header row is found by content (has both a "Topic" and an
     "Audience" column), not by a fixed row number, so it survives inserted
     rows the same way the other two extractors do. */
  function findP2PHeaderRow(sheet) {
    const max = Math.min(sheet.maxRow || MAX_SCAN, MAX_SCAN);
    for (let r = 1; r <= max; r++) {
      let hasTopic = false, hasAudience = false;
      for (let c = 1; c <= Math.max(sheet.maxCol, 1); c++) {
        const t = norm(sheet.txt(numToCol(c) + r));
        if (t === "topic") hasTopic = true;
        else if (t === "audience") hasAudience = true;
      }
      if (hasTopic && hasAudience) return r;
    }
    return 0;
  }

  /* Which additive metrics combine into each ratio metric, keyed by column
     label rather than position. "Click Rate" has no stored click count, so
     it cannot be recombined across rows and is reported as such rather than
     silently averaged. */
  const P2P_RATIO_COMPONENTS = {
    gross_cpa: ["gross_spend", "ntl"],
    immediate_roas: ["immediate_raised", "gross_spend"],
    average_gift: ["immediate_raised", "donors"],
    donate_rate: ["donors", "recipients"],
    unsub_rate: ["unsubs", "recipients"],
    ltv_roas_from_ntl: ["ltv_from_ntl", "gross_spend"]
  };

  function extractP2P(sheet, opts) {
    const findings = [], problems = [];
    const err = t => problems.push({level: "error", text: t});
    const warn = t => problems.push({level: "warn", text: t});

    const hdrRow = findP2PHeaderRow(sheet);
    if (!hdrRow) {
      err("Could not find a header row with both a “Topic” and an “Audience” column, so this tab could not be read.");
      return {payload: null, findings: findings, problems: problems};
    }
    const lastCol = numToCol(Math.max(sheet.maxCol, colToNum("T")));
    const hm = headerMap(sheet, hdrRow, "A", lastCol);
    findings.push("Header row found at row " + hdrRow + ".");

    const dimCol = {};
    P2P_DIM_KEYS.forEach(k => { if (hm.map[k]) dimCol[k.replace(/ /g, "_")] = hm.map[k]; });
    ["month", "topic", "audience"].forEach(k => {
      if (!dimCol[k]) err("Could not find a “" + k[0].toUpperCase() + k.slice(1) + "” column on row " + hdrRow + ".");
    });
    if (problems.some(p => p.level === "error")) return {payload: null, findings: findings, problems: problems};
    const objectiveHeader = dimCol.goal ? sheet.txt(dimCol.goal + hdrRow) : null;

    /* A column counts as a metric only if every non-blank cell under it parses
       as a number — a text column like "Source Code" would otherwise become a
       metric whose value is always null. */
    function columnIsNumeric(col) {
      let sawValue = false;
      for (let r = hdrRow + 1; r <= sheet.maxRow; r++) {
        const t = sheet.txt(col + r);
        if (t == null) continue;
        sawValue = true;
        if (sheet.num(col + r) == null) return false;
      }
      return sawValue;
    }

    /* --- metric columns: everything on the header row that isn't a dimension --- */
    const candidates = hm.order.filter(h => P2P_DIM_KEYS.indexOf(h.key) < 0 && !/^helper/.test(h.key));
    const metrics = candidates.filter(h => columnIsNumeric(h.col));
    const textCols = candidates.filter(h => columnIsNumeric(h.col) === false);
    if (textCols.length) {
      findings.push("Column(s) " + textCols.map(h => "“" + h.label + "” (" + h.col + ")").join(", ") +
        " hold text rather than numbers, so they are not treated as metrics.");
    }
    if (!metrics.length) {
      err("No metric columns were found on the header row " + hdrRow + ".");
      return {payload: null, findings: findings, problems: problems};
    }
    const metricOrder = [], metricLabels = {}, metricKinds = {}, metricColumns = {};
    metrics.forEach(h => {
      const key = slug(h.label);
      metricOrder.push(key);
      metricLabels[key] = h.label;
      metricKinds[key] = metricKind(key, h.label);
      metricColumns[key] = h.col;
    });
    findings.push("Found " + metricOrder.length + " metric column(s): " +
      metricOrder.map(k => metricLabels[k] + " (" + metricColumns[k] + ")").join(", ") + ".");

    function metricsAt(row) {
      const out = {};
      metricOrder.forEach(key => {
        const ref = metricColumns[key] + row;
        out[key] = {label: metricLabels[key], kind: metricKinds[key],
                    value: sheet.num(ref), cell: ref, formula: sheet.formula(ref)};
      });
      return out;
    }

    /* --- detail rows: one per date × topic × audience, until both go blank --- */
    const detail = [];
    for (let r = hdrRow + 1; r <= sheet.maxRow; r++) {
      const month = sheet.txt(dimCol.month + r);
      const topic = sheet.txt(dimCol.topic + r);
      if (!month && !topic) break;
      const dateSerial = dimCol.date ? sheet.num(dimCol.date + r) : null;
      detail.push({
        row: r,
        client: dimCol.client ? sheet.txt(dimCol.client + r) : null,
        month: month, month_cell: dimCol.month + r,
        day_of_week: dimCol.day_of_week ? sheet.txt(dimCol.day_of_week + r) : null,
        date_serial: dateSerial, date_cell: dimCol.date ? dimCol.date + r : null,
        date: serialToDate(dateSerial),
        objective: dimCol.goal ? sheet.txt(dimCol.goal + r) : null,
        objective_cell: dimCol.goal ? dimCol.goal + r : null,
        topic: topic, topic_cell: dimCol.topic + r,
        audience: sheet.txt(dimCol.audience + r), audience_cell: dimCol.audience + r,
        metrics: metricsAt(r)
      });
    }
    if (!detail.length) err("No detail rows were found under the header on row " + hdrRow + ".");
    else findings.push("Read " + detail.length + " detail row(s): rows " + detail[0].row + "–" +
      detail[detail.length - 1].row + ".");
    const missingDims = detail.filter(d => !d.topic || !d.audience).length;
    if (missingDims) warn(missingDims + " detail row(s) are missing a topic or audience name.");

    /* --- additive vs ratio metrics, and how ratios recombine --- */
    const additive = metricOrder.filter(k => metricKinds[k] === "usd" || metricKinds[k] === "count");
    const ratios = metricOrder.filter(k => additive.indexOf(k) < 0);
    function componentsFor(key) {
      const c = P2P_RATIO_COMPONENTS[key];
      return (c && additive.indexOf(c[0]) >= 0 && additive.indexOf(c[1]) >= 0) ? c : null;
    }
    // Some ratio metrics (e.g. Click Rate) have no stored numerator on this
    // tab, so they can't be recombined across rows — they're just omitted
    // from monthly/total rollups rather than averaged. That's an expected,
    // permanent property of this tab's layout, not a data problem worth
    // surfacing as a warning; the UI explains it in-context (see the info
    // bubble on the monthly toplines table) instead.
    const unresolvable = ratios.filter(k => !componentsFor(k));

    function sumRows(rows, key) {
      let t = 0, any = false; const refs = [];
      rows.forEach(r => { const v = r.metrics[key].value; if (v != null) { t += v; any = true; refs.push(r.metrics[key].cell); } });
      return {value: any ? t : null, cells: refs};
    }
    /* Sums additive metrics across `rows`, then recomputes each ratio from
       the summed components — never averages a per-row ratio. */
    function rollup(rows, label) {
      const out = {label: label, rows: rows.map(r => r.row), metrics: {}};
      additive.forEach(key => {
        const s = sumRows(rows, key);
        out.metrics[key] = {label: metricLabels[key], kind: metricKinds[key], value: s.value,
                             cells: s.cells, computed: true};
      });
      ratios.forEach(key => {
        const comp = componentsFor(key);
        if (!comp) {
          out.metrics[key] = {label: metricLabels[key], kind: metricKinds[key], value: null,
                               computed: true, unresolvable: true};
          return;
        }
        const [nk, dk] = comp;
        const nv = out.metrics[nk].value, dv = out.metrics[dk].value;
        out.metrics[key] = {
          label: metricLabels[key], kind: metricKinds[key],
          value: dv ? nv / dv : null, computed: true,
          numerator: nk, denominator: dk,
          cells: out.metrics[nk].cells.concat(out.metrics[dk].cells)
        };
      });
      return out;
    }

    /* --- month rollups, in the order months first appear --- */
    const monthKeys = [], monthLabelOf = {};
    detail.forEach(d => {
      const k = slug(d.month);
      if (k && monthLabelOf[k] == null) { monthKeys.push(k); monthLabelOf[k] = d.month.trim(); }
    });
    const months = monthKeys.map(k => {
      const rows = detail.filter(d => slug(d.month) === k);
      const m = rollup(rows, monthLabelOf[k]);
      m.key = k;
      return m;
    });
    if (months.length) findings.push("Computed " + months.length + " monthly rollup(s) from the detail rows: " +
      months.map(m => m.label).join(", ") + ".");

    /* --- quarter-to-date totals across every detail row --- */
    const totals = rollup(detail, "Quarter to date");

    /* --- categories --- */
    const audiences = uniqSorted(detail.map(d => d.audience));
    const topics = uniqSorted(detail.map(d => d.topic));
    const objectives = uniqSorted(detail.map(d => d.objective));

    /* --- cumulative series: running totals, month by month --- */
    const accumKeys = ["gross_spend", "immediate_raised", "donors"].filter(k => metricOrder.indexOf(k) >= 0);
    function cumulateByMonth(rowsForMonth) {
      const out = {};
      accumKeys.forEach(key => {
        let run = 0; const series = [];
        monthKeys.forEach((mk, i) => {
          const rows = rowsForMonth[i];
          // A month with zero rows for this category is a real zero unless
          // the whole sheet has nothing for that month yet, in which case it
          // is pending rather than zero (same convention as the Digital tab).
          if (!rows.length && !detail.some(d => slug(d.month) === mk)) {
            series.push({month: mk, label: monthLabelOf[mk], value: null, pending: true});
            return;
          }
          // Once the month is known to have data somewhere on the sheet, a
          // category with no rows in it is a real, reportable zero — not a
          // missing figure — so `step` is 0 rather than null.
          const s = sumRows(rows, key);
          const stepVal = s.value == null ? 0 : s.value;
          run += stepVal;
          series.push({month: mk, label: monthLabelOf[mk], value: run, step: stepVal,
                       step_cells: s.cells, computed_running_total: true});
        });
        out[key] = series;
      });
      return out;
    }
    const byMonth = mk => detail.filter(d => slug(d.month) === mk);
    const cumulative = {
      overall: cumulateByMonth(monthKeys.map(byMonth)),
      by_audience: {}, by_topic: {}
    };
    audiences.forEach(a => {
      cumulative.by_audience[a] = cumulateByMonth(monthKeys.map(mk =>
        byMonth(mk).filter(d => d.audience === a)));
    });
    topics.forEach(tp => {
      cumulative.by_topic[tp] = cumulateByMonth(monthKeys.map(mk =>
        byMonth(mk).filter(d => d.topic === tp)));
    });

    const year = labelled(sheet, "A", "B", /^year$/);
    const quarter = labelled(sheet, "A", "B", /^quarter$/);
    if (!year || year.num == null) warn("Could not read the report year.");
    if (!quarter || quarter.num == null) warn("Could not read which quarter this tab is set to.");

    const payload = {
      meta: {
        title: "P2P Statistics",
        sheet: sheet.name,
        source_workbook: opts.workbookName,
        extracted_at: opts.extractedAt,
        year: {value: year ? year.num : null, cell: year ? year.cell : null},
        quarter: {value: quarter ? quarter.num : null, cell: quarter ? quarter.cell : null}
      },
      metric_order: metricOrder,
      metric_labels: metricLabels,
      metric_kinds: metricKinds,
      metric_columns: metricColumns,
      additive: additive,
      ratios: ratios,
      unresolvable_ratios: unresolvable,
      // {ratioKey: [numeratorKey, denominatorKey] | null} — lets any client
      // (this page's leaderboard totals row, in particular) recombine a
      // ratio over an arbitrary filtered subset of rows without re-deriving
      // the naming convention itself.
      ratio_components: ratios.reduce((m, k) => { m[k] = componentsFor(k); return m; }, {}),
      objective_header: objectiveHeader,
      month_keys: monthKeys,
      months: months,
      totals: totals,
      detail: detail,
      audiences: audiences,
      topics: topics,
      objectives: objectives,
      accum_keys: accumKeys,
      cumulative: cumulative
    };
    return {payload: payload, findings: findings, problems: problems};
  }

  const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July",
                        "August", "September", "October", "November", "December"];

  /* Excel serial -> {year, month, quarter}, built on serialToDate so the
     1900 leap-year handling lives in exactly one place. */
  function ymFromSerial(serial) {
    const iso = serialToDate(serial);
    if (!iso) return null;
    const y = parseInt(iso.slice(0, 4), 10), m = parseInt(iso.slice(5, 7), 10);
    return {year: y, month: m, quarter: Math.floor((m - 1) / 3) + 1};
  }

  /* ===================================================================
     EMAIL STATISTICS
     Same shape as P2P Statistics: a single flat detail table (one row per
     email send), no stored monthly/quarterly rollup, so every rollup here
     is computed from the detail rows. Unlike P2P, this sheet's own "Month"
     column is a bare 1-12 integer rather than a "July 2026"-style label, so
     a display month label is derived from the Date column instead of being
     read verbatim — the raw integer is kept too, under month_num.
     =================================================================== */
  const EMAIL_DIM_MAP = [
    ["year", "year"], ["quarter", "quarter"], ["month", "month_num"],
    ["day of week", "day_of_week"], ["date", "date"], ["send time (et)", "send_time"],
    ["ask", "ask"], ["subtype", "subtype"], ["mailing topic", "topic"],
    ["sender", "sender"], ["subject line", "subject_line"], ["audience", "audience"],
    ["label", "label"]
  ];
  const EMAIL_DIM_KEYS = EMAIL_DIM_MAP.map(p => p[0]);

  /* Which additive metrics recombine into each ratio metric. Open Rate and
     Click Rate have no stored opens/clicks count on this sheet (same
     situation as P2P's Click Rate), so they are deliberately left out of
     this map rather than guessed at. */
  const EMAIL_RATIO_COMPONENTS = {
    average: ["raised", "donors"],
    donate_rate: ["donors", "recipients"],
    action_rate: ["actions", "recipients"],
    unsub_rate: ["unsubs", "recipients"]
  };

  function findEmailHeaderRow(sheet) {
    const max = Math.min(sheet.maxRow || MAX_SCAN, MAX_SCAN);
    for (let r = 1; r <= max; r++) {
      let hasAudience = false, hasTopic = false;
      for (let c = 1; c <= Math.max(sheet.maxCol, 1); c++) {
        const t = norm(sheet.txt(numToCol(c) + r));
        if (t === "audience") hasAudience = true;
        else if (t === "mailing topic") hasTopic = true;
      }
      if (hasAudience && hasTopic) return r;
    }
    return 0;
  }

  function extractEmailStats(sheet, opts) {
    const findings = [], problems = [];
    const err = t => problems.push({level: "error", text: t});
    const warn = t => problems.push({level: "warn", text: t});

    const hdrRow = findEmailHeaderRow(sheet);
    if (!hdrRow) {
      err("Could not find a header row with both an “Audience” and a “Mailing Topic” column, so this tab could not be read.");
      return {payload: null, findings: findings, problems: problems};
    }
    const lastCol = numToCol(Math.max(sheet.maxCol, colToNum("X")));
    const hm = headerMap(sheet, hdrRow, "A", lastCol);
    findings.push("Header row found at row " + hdrRow + ".");

    const dimCol = {};
    EMAIL_DIM_MAP.forEach(function (pair) { if (hm.map[pair[0]]) dimCol[pair[1]] = hm.map[pair[0]]; });
    if (!dimCol.date) err("Could not find a “Date” column on row " + hdrRow + ".");
    if (!dimCol.topic) err("Could not find a “Mailing Topic” column on row " + hdrRow + ".");
    if (!dimCol.audience) err("Could not find an “Audience” column on row " + hdrRow + ".");
    if (!dimCol.label) err("Could not find a “Label” column on row " + hdrRow + ".");
    if (problems.some(p => p.level === "error")) return {payload: null, findings: findings, problems: problems};

    function columnIsNumeric(col) {
      let sawValue = false;
      for (let r = hdrRow + 1; r <= sheet.maxRow; r++) {
        const t = sheet.txt(col + r);
        if (t == null) continue;
        sawValue = true;
        if (sheet.num(col + r) == null) return false;
      }
      return sawValue;
    }

    const candidates = hm.order.filter(h => EMAIL_DIM_KEYS.indexOf(h.key) < 0 && !/^helper/.test(h.key));
    const metrics = candidates.filter(h => columnIsNumeric(h.col));
    const textCols = candidates.filter(h => columnIsNumeric(h.col) === false);
    if (textCols.length) {
      findings.push("Column(s) " + textCols.map(h => "“" + h.label + "” (" + h.col + ")").join(", ") +
        " hold text rather than numbers, so they are not treated as metrics.");
    }
    if (!metrics.length) {
      err("No metric columns were found on the header row " + hdrRow + ".");
      return {payload: null, findings: findings, problems: problems};
    }
    const metricOrder = [], metricLabels = {}, metricKinds = {}, metricColumns = {};
    metrics.forEach(h => {
      const key = slug(h.label);
      metricOrder.push(key);
      metricLabels[key] = h.label;
      metricKinds[key] = metricKind(key, h.label);
      metricColumns[key] = h.col;
    });
    findings.push("Found " + metricOrder.length + " metric column(s): " +
      metricOrder.map(k => metricLabels[k] + " (" + metricColumns[k] + ")").join(", ") + ".");

    function metricsAt(row) {
      const out = {};
      metricOrder.forEach(key => {
        const ref = metricColumns[key] + row;
        out[key] = {label: metricLabels[key], kind: metricKinds[key],
                    value: sheet.num(ref), cell: ref, formula: sheet.formula(ref)};
      });
      return out;
    }

    /* --- detail rows: one per email send, until label and date both go blank --- */
    const detail = [];
    for (let r = hdrRow + 1; r <= sheet.maxRow; r++) {
      const label = sheet.txt(dimCol.label + r);
      const dateSerial = dimCol.date ? sheet.num(dimCol.date + r) : null;
      if (!label && dateSerial == null) break;
      const ym = dateSerial != null ? ymFromSerial(dateSerial) : null;
      const monthLabel = ym ? MONTH_NAMES[ym.month - 1] + " " + ym.year : null;
      detail.push({
        row: r,
        year: dimCol.year ? sheet.num(dimCol.year + r) : (ym ? ym.year : null),
        quarter: dimCol.quarter ? sheet.num(dimCol.quarter + r) : (ym ? ym.quarter : null),
        month: monthLabel, month_derived: true,
        month_num: dimCol.month_num ? sheet.num(dimCol.month_num + r) : (ym ? ym.month : null),
        month_num_cell: dimCol.month_num ? dimCol.month_num + r : null,
        day_of_week: dimCol.day_of_week ? sheet.txt(dimCol.day_of_week + r) : null,
        date_serial: dateSerial, date_cell: dimCol.date ? dimCol.date + r : null,
        date: serialToDate(dateSerial),
        send_time: dimCol.send_time ? sheet.txt(dimCol.send_time + r) : null,
        ask: dimCol.ask ? sheet.txt(dimCol.ask + r) : null,
        subtype: dimCol.subtype ? sheet.txt(dimCol.subtype + r) : null,
        topic: sheet.txt(dimCol.topic + r), topic_cell: dimCol.topic + r,
        sender: dimCol.sender ? sheet.txt(dimCol.sender + r) : null,
        subject_line: dimCol.subject_line ? sheet.txt(dimCol.subject_line + r) : null,
        audience: sheet.txt(dimCol.audience + r), audience_cell: dimCol.audience + r,
        label: label, label_cell: dimCol.label + r,
        metrics: metricsAt(r)
      });
    }
    if (!detail.length) err("No detail rows were found under the header on row " + hdrRow + ".");
    else findings.push("Read " + detail.length + " detail row(s): rows " + detail[0].row + "–" +
      detail[detail.length - 1].row + ".");
    const missingDims = detail.filter(d => !d.topic || !d.audience).length;
    if (missingDims) warn(missingDims + " detail row(s) are missing a mailing topic or audience name.");

    /* --- additive vs ratio metrics, and how ratios recombine --- */
    const additive = metricOrder.filter(k => metricKinds[k] === "usd" || metricKinds[k] === "count");
    const ratios = metricOrder.filter(k => additive.indexOf(k) < 0);
    function componentsFor(key) {
      const c = EMAIL_RATIO_COMPONENTS[key];
      return (c && additive.indexOf(c[0]) >= 0 && additive.indexOf(c[1]) >= 0) ? c : null;
    }
    // Some ratio metrics (e.g. Click Rate) have no stored numerator on this
    // tab, so they can't be recombined across rows — they're just omitted
    // from monthly/total rollups rather than averaged. That's an expected,
    // permanent property of this tab's layout, not a data problem worth
    // surfacing as a warning; the UI explains it in-context (see the info
    // bubble on the monthly toplines table) instead.
    const unresolvable = ratios.filter(k => !componentsFor(k));

    function sumRows(rows, key) {
      let t = 0, any = false; const refs = [];
      rows.forEach(r => { const v = r.metrics[key].value; if (v != null) { t += v; any = true; refs.push(r.metrics[key].cell); } });
      return {value: any ? t : null, cells: refs};
    }
    /* Sums additive metrics across `rows`, then recomputes each ratio from
       the summed components — never averages a per-row ratio. */
    function rollup(rows, label) {
      const out = {label: label, rows: rows.map(r => r.row), metrics: {}};
      additive.forEach(key => {
        const s = sumRows(rows, key);
        out.metrics[key] = {label: metricLabels[key], kind: metricKinds[key], value: s.value,
                             cells: s.cells, computed: true};
      });
      ratios.forEach(key => {
        const comp = componentsFor(key);
        if (!comp) {
          out.metrics[key] = {label: metricLabels[key], kind: metricKinds[key], value: null,
                               computed: true, unresolvable: true};
          return;
        }
        const nk = comp[0], dk = comp[1];
        const nv = out.metrics[nk].value, dv = out.metrics[dk].value;
        out.metrics[key] = {
          label: metricLabels[key], kind: metricKinds[key],
          value: dv ? nv / dv : null, computed: true,
          numerator: nk, denominator: dk,
          cells: out.metrics[nk].cells.concat(out.metrics[dk].cells)
        };
      });
      return out;
    }

    /* --- month rollups, in the order months first appear --- */
    const monthKeys = [], monthLabelOf = {};
    detail.forEach(d => {
      const k = d.month ? slug(d.month) : null;
      if (k && monthLabelOf[k] == null) { monthKeys.push(k); monthLabelOf[k] = d.month.trim(); }
    });
    const months = monthKeys.map(k => {
      const rows = detail.filter(d => d.month && slug(d.month) === k);
      const m = rollup(rows, monthLabelOf[k]);
      m.key = k;
      return m;
    });
    if (months.length) findings.push("Computed " + months.length + " monthly rollup(s) from the detail rows: " +
      months.map(m => m.label).join(", ") + ".");

    /* --- quarter-to-date totals across every detail row --- */
    const totals = rollup(detail, "Quarter to date");

    /* --- categories --- */
    const audiences = uniqSorted(detail.map(d => d.audience));
    const topics = uniqSorted(detail.map(d => d.topic));

    /* --- cumulative series: running totals, month by month --- */
    const accumKeys = ["raised", "recipients", "donors"].filter(k => metricOrder.indexOf(k) >= 0);
    function cumulateByMonth(rowsForMonth) {
      const out = {};
      accumKeys.forEach(key => {
        let run = 0; const series = [];
        monthKeys.forEach((mk, i) => {
          const rows = rowsForMonth[i];
          if (!rows.length && !detail.some(d => d.month && slug(d.month) === mk)) {
            series.push({month: mk, label: monthLabelOf[mk], value: null, pending: true});
            return;
          }
          const s = sumRows(rows, key);
          const stepVal = s.value == null ? 0 : s.value;
          run += stepVal;
          series.push({month: mk, label: monthLabelOf[mk], value: run, step: stepVal,
                       step_cells: s.cells, computed_running_total: true});
        });
        out[key] = series;
      });
      return out;
    }
    const byMonth = mk => detail.filter(d => d.month && slug(d.month) === mk);
    const cumulative = {
      overall: cumulateByMonth(monthKeys.map(byMonth)),
      by_audience: {}, by_topic: {}
    };
    audiences.forEach(a => {
      cumulative.by_audience[a] = cumulateByMonth(monthKeys.map(mk =>
        byMonth(mk).filter(d => d.audience === a)));
    });
    topics.forEach(tp => {
      cumulative.by_topic[tp] = cumulateByMonth(monthKeys.map(mk =>
        byMonth(mk).filter(d => d.topic === tp)));
    });

    const year = labelled(sheet, "A", "B", /^year$/);
    const quarter = labelled(sheet, "A", "B", /^quarter$/);
    if (!year || year.num == null) warn("Could not read the report year.");
    if (!quarter || quarter.num == null) warn("Could not read which quarter this tab is set to.");

    const payload = {
      meta: {
        title: "Email Statistics",
        sheet: sheet.name,
        source_workbook: opts.workbookName,
        extracted_at: opts.extractedAt,
        year: {value: year ? year.num : null, cell: year ? year.cell : null},
        quarter: {value: quarter ? quarter.num : null, cell: quarter ? quarter.cell : null}
      },
      metric_order: metricOrder,
      metric_labels: metricLabels,
      metric_kinds: metricKinds,
      metric_columns: metricColumns,
      additive: additive,
      ratios: ratios,
      unresolvable_ratios: unresolvable,
      ratio_components: ratios.reduce((m, k) => { m[k] = componentsFor(k); return m; }, {}),
      month_keys: monthKeys,
      months: months,
      totals: totals,
      detail: detail,
      audiences: audiences,
      topics: topics,
      accum_keys: accumKeys,
      cumulative: cumulative
    };
    return {payload: payload, findings: findings, problems: problems};
  }

  /* ===================================================================
     HIGH-DOLLAR DONATIONS
     One row per donation. Unlike every other tab, there is no "Year:" /
     "Quarter:" label block anywhere in the header rows (row 1 is blank,
     row 2 is a prose note, row 3 is blank, row 4 is the header) — year,
     quarter and month are derived here from each row's `date` instead of
     being read from a labelled cell. Only `amount` is additive; there are
     no ratio metrics on this sheet.
     =================================================================== */
  const HD_DIM_MAP = [
    ["category", "category"], ["receipt_id", "receipt_id"], ["date", "date"],
    ["fundraising_page", "fundraising_page"], ["first", "first"], ["last", "last"],
    ["email", "email"], ["finance claim", "finance_claim"], ["authentic added", "authentic_added"]
  ];
  const HD_DIM_KEYS = HD_DIM_MAP.map(p => p[0]);

  function findHDHeaderRow(sheet) {
    const max = Math.min(sheet.maxRow || MAX_SCAN, MAX_SCAN);
    for (let r = 1; r <= max; r++) {
      let hasReceipt = false, hasAmount = false;
      for (let c = 1; c <= Math.max(sheet.maxCol, 1); c++) {
        const t = norm(sheet.txt(numToCol(c) + r));
        if (t === "receipt_id") hasReceipt = true;
        else if (t === "amount") hasAmount = true;
      }
      if (hasReceipt && hasAmount) return r;
    }
    return 0;
  }

  function extractHighDollar(sheet, opts) {
    const findings = [], problems = [];
    const err = t => problems.push({level: "error", text: t});
    const warn = t => problems.push({level: "warn", text: t});

    const hdrRow = findHDHeaderRow(sheet);
    if (!hdrRow) {
      err("Could not find a header row with both a “receipt_id” and an “amount” column, so this tab could not be read.");
      return {payload: null, findings: findings, problems: problems};
    }
    const lastCol = numToCol(Math.max(sheet.maxCol, colToNum("J")));
    const hm = headerMap(sheet, hdrRow, "A", lastCol);
    findings.push("Header row found at row " + hdrRow + ".");

    const dimCol = {};
    HD_DIM_MAP.forEach(function (pair) { if (hm.map[pair[0]]) dimCol[pair[1]] = hm.map[pair[0]]; });
    if (!dimCol.receipt_id) err("Could not find a “receipt_id” column on row " + hdrRow + ".");
    if (!dimCol.date) err("Could not find a “date” column on row " + hdrRow + ".");
    if (problems.some(p => p.level === "error")) return {payload: null, findings: findings, problems: problems};

    function columnIsNumeric(col) {
      let sawValue = false;
      for (let r = hdrRow + 1; r <= sheet.maxRow; r++) {
        const t = sheet.txt(col + r);
        if (t == null) continue;
        sawValue = true;
        if (sheet.num(col + r) == null) return false;
      }
      return sawValue;
    }
    const candidates = hm.order.filter(h => HD_DIM_KEYS.indexOf(h.key) < 0 && !/^helper/.test(h.key));
    const metrics = candidates.filter(h => columnIsNumeric(h.col));
    if (!metrics.length) {
      err("No metric columns were found on the header row " + hdrRow + " (expected an “amount” column).");
      return {payload: null, findings: findings, problems: problems};
    }
    const metricOrder = [], metricLabels = {}, metricKinds = {}, metricColumns = {};
    metrics.forEach(h => {
      const key = slug(h.label);
      metricOrder.push(key);
      metricLabels[key] = h.label;
      metricKinds[key] = metricKind(key, h.label);
      metricColumns[key] = h.col;
    });
    findings.push("Found " + metricOrder.length + " metric column(s): " +
      metricOrder.map(k => metricLabels[k] + " (" + metricColumns[k] + ")").join(", ") + ".");

    function metricsAt(row) {
      const out = {};
      metricOrder.forEach(key => {
        const ref = metricColumns[key] + row;
        out[key] = {label: metricLabels[key], kind: metricKinds[key],
                    value: sheet.num(ref), cell: ref, formula: sheet.formula(ref)};
      });
      return out;
    }

    /* --- detail rows: one per donation, until receipt_id and date both go blank --- */
    const detail = [];
    for (let r = hdrRow + 1; r <= sheet.maxRow; r++) {
      const receiptId = sheet.txt(dimCol.receipt_id + r);
      const dateSerial = dimCol.date ? sheet.num(dimCol.date + r) : null;
      if (!receiptId && dateSerial == null) break;
      const ym = dateSerial != null ? ymFromSerial(dateSerial) : null;
      detail.push({
        row: r,
        category: dimCol.category ? sheet.txt(dimCol.category + r) : null,
        category_cell: dimCol.category ? dimCol.category + r : null,
        receipt_id: receiptId, receipt_id_cell: dimCol.receipt_id + r,
        date_serial: dateSerial, date_cell: dimCol.date ? dimCol.date + r : null,
        date: serialToDate(dateSerial),
        year: ym ? ym.year : null, quarter: ym ? ym.quarter : null,
        month: ym ? MONTH_NAMES[ym.month - 1] + " " + ym.year : null, month_derived: true,
        fundraising_page: dimCol.fundraising_page ? sheet.txt(dimCol.fundraising_page + r) : null,
        first: dimCol.first ? sheet.txt(dimCol.first + r) : null,
        last: dimCol.last ? sheet.txt(dimCol.last + r) : null,
        email: dimCol.email ? sheet.txt(dimCol.email + r) : null,
        finance_claim: dimCol.finance_claim ? sheet.bool(dimCol.finance_claim + r) : null,
        finance_claim_cell: dimCol.finance_claim ? dimCol.finance_claim + r : null,
        authentic_added: dimCol.authentic_added ? sheet.bool(dimCol.authentic_added + r) : null,
        authentic_added_cell: dimCol.authentic_added ? dimCol.authentic_added + r : null,
        metrics: metricsAt(r)
      });
    }
    if (!detail.length) err("No detail rows were found under the header on row " + hdrRow + ".");
    else findings.push("Read " + detail.length + " detail row(s): rows " + detail[0].row + "–" +
      detail[detail.length - 1].row + ".");
    const missingDate = detail.filter(d => d.date_serial == null).length;
    if (missingDate) warn(missingDate + " detail row(s) have no date, so they are excluded from monthly rollups.");

    const additive = metricOrder.filter(k => metricKinds[k] === "usd" || metricKinds[k] === "count");
    const ratios = [];

    function sumRows(rows, key) {
      let t = 0, any = false; const refs = [];
      rows.forEach(r => { const v = r.metrics[key].value; if (v != null) { t += v; any = true; refs.push(r.metrics[key].cell); } });
      return {value: any ? t : null, cells: refs};
    }
    function rollup(rows, label) {
      const out = {label: label, rows: rows.map(r => r.row), count: rows.length, metrics: {}};
      additive.forEach(key => {
        const s = sumRows(rows, key);
        out.metrics[key] = {label: metricLabels[key], kind: metricKinds[key], value: s.value,
                             cells: s.cells, computed: true};
      });
      return out;
    }

    /* --- month rollups, in the order months first appear (rows with no date are skipped) --- */
    const monthKeys = [], monthLabelOf = {};
    detail.forEach(d => {
      if (!d.month) return;
      const k = slug(d.month);
      if (k && monthLabelOf[k] == null) { monthKeys.push(k); monthLabelOf[k] = d.month.trim(); }
    });
    const months = monthKeys.map(k => {
      const rows = detail.filter(d => d.month && slug(d.month) === k);
      const m = rollup(rows, monthLabelOf[k]);
      m.key = k;
      return m;
    });
    if (months.length) findings.push("Computed " + months.length + " monthly rollup(s) from the detail rows: " +
      months.map(m => m.label).join(", ") + ".");

    const totals = rollup(detail, "Quarter to date");

    /* --- categories --- */
    const categories = uniqSorted(detail.map(d => d.category));

    const accumKeys = metricOrder.slice();
    function cumulateByMonth(rowsForMonth) {
      const out = {};
      accumKeys.forEach(key => {
        let run = 0; const series = [];
        monthKeys.forEach((mk, i) => {
          const rows = rowsForMonth[i];
          if (!rows.length && !detail.some(d => d.month && slug(d.month) === mk)) {
            series.push({month: mk, label: monthLabelOf[mk], value: null, pending: true});
            return;
          }
          const s = sumRows(rows, key);
          const stepVal = s.value == null ? 0 : s.value;
          run += stepVal;
          series.push({month: mk, label: monthLabelOf[mk], value: run, step: stepVal,
                       step_cells: s.cells, computed_running_total: true});
        });
        out[key] = series;
      });
      return out;
    }
    const byMonth = mk => detail.filter(d => d.month && slug(d.month) === mk);
    const cumulative = {
      overall: cumulateByMonth(monthKeys.map(byMonth)),
      by_category: {}
    };
    categories.forEach(c => {
      cumulative.by_category[c] = cumulateByMonth(monthKeys.map(mk =>
        byMonth(mk).filter(d => d.category === c)));
    });

    /* --- year/quarter: no labelled cell exists on this sheet for either, so
       both are derived from the dates on the detail rows. A span across
       more than one quarter is reported rather than guessed at, since
       picking one would misrepresent the others. --- */
    const dated = detail.filter(d => d.year != null);
    const qkeys = uniqSorted(dated.map(d => d.year + "-Q" + d.quarter));
    let derivedYear = null, derivedQuarter = null;
    if (qkeys.length === 1) {
      derivedYear = dated[0].year; derivedQuarter = dated[0].quarter;
      findings.push("This tab has no “Year:”/“Quarter:” label cells; derived Q" + derivedQuarter + " " +
        derivedYear + " from the “date” column on every detail row.");
    } else if (qkeys.length > 1) {
      warn("Detail rows span more than one quarter (" + qkeys.join(", ") +
           "), and this tab has no “Year:”/“Quarter:” label cells, so no single quarter can be named for it.");
    } else {
      warn("Could not derive a year or quarter: no detail row has a readable date.");
    }

    const payload = {
      meta: {
        title: "High-Dollar Donations",
        sheet: sheet.name,
        source_workbook: opts.workbookName,
        extracted_at: opts.extractedAt,
        year: {value: derivedYear, cell: null, derived: true},
        quarter: {value: derivedQuarter, cell: null, derived: true}
      },
      metric_order: metricOrder,
      metric_labels: metricLabels,
      metric_kinds: metricKinds,
      metric_columns: metricColumns,
      additive: additive,
      ratios: ratios,
      unresolvable_ratios: [],
      ratio_components: {},
      month_keys: monthKeys,
      months: months,
      totals: totals,
      detail: detail,
      categories: categories,
      accum_keys: accumKeys,
      cumulative: cumulative
    };
    return {payload: payload, findings: findings, problems: problems};
  }

  /* ===================================================================
     RECONCILIATION — the checks that decide whether to trust a section.
     These are the same identities the offline verifiers assert, run again
     on whatever was just dropped in.
     =================================================================== */
  function reconcile(dig, paid, p2p, email, hd) {
    const problems = [], findings = [];
    const near = (a, b, tol) => a != null && b != null && Math.abs(a - b) <= (tol == null ? 0.02 : tol);

    /* A ratio-identity failure is a FLAG, not an error. The rows parsed
       cleanly; a single stored ratio cell just disagrees with the components
       sitting next to it. Hiding the whole tab over that throws away four
       sound sections to suppress one suspect column, so instead the tab
       renders and a note is pinned to the chart or table that actually shows
       the stored cell. Rollup sum mismatches stay errors: there the page's own
       computed totals would visibly contradict the workbook, and a note under
       a chart cannot make those numbers safe to read.

       `sections` are the ids of the <section> elements that display the
       offending cells — not every section on the tab, only the ones reading
       the stored value. Rollups recomputed from components are unaffected and
       deliberately carry no note. */
    const MAX_CELLS = 6;
    const flag = (text, sections, cells) => problems.push({
      level: "flag",
      text: text,
      sections: sections,
      cells: (cells || []).slice(0, MAX_CELLS),
      moreCells: Math.max(0, (cells || []).length - MAX_CELLS)
    });
    /* enough precision to see the disagreement, without printing float noise */
    const shortNum = v => v == null ? "—"
      : Math.abs(v) >= 1 ? v.toFixed(2)
      : v.toFixed(6).replace(/0+$/, "").replace(/\.$/, "");
    const badCell = (cellRef, stored, expected) =>
      (cellRef || "?") + " shows " + shortNum(stored) + ", components give " + shortNum(expected);

    if (dig) {
      let bad = 0, checked = 0;
      const pctCells = [];
      dig.channels.forEach(c => {
        Object.keys(c.periods).forEach(pk => {
          const p = c.periods[pk];
          if (!p.goal || p.actual == null || p.pct == null) return;
          checked++;
          const want = p.actual / p.goal;
          if (!near(p.pct, want, 1e-6)) { bad++; pctCells.push(badCell(p.pct_cell, p.pct, want)); }
        });
      });
      // stored percentages drive the hero cards and the channel bars; the
      // monthly chart recomputes its own percentages and is unaffected
      if (bad) flag(bad + " of " + checked + " “% to goal” cells on the Digital Report do not equal actual ÷ goal.",
        ["s-hero", "s-channel"], pctCells);
      else if (checked) findings.push("All " + checked + " “% to goal” cells equal actual ÷ goal.");

      const tr = dig.overall && dig.overall.total_raised;
      if (tr) {
        ["quarter"].concat(dig.month_keys).forEach(pk => {
          const stored = tr.periods[pk] && tr.periods[pk].actual;
          if (stored == null) return;
          let sum = 0, any = false;
          dig.channels.forEach(c => {
            const v = c.periods[pk].actual;
            if (v != null) { sum += v; any = true; }
          });
          if (!any) return;
          if (!near(sum, stored, 0.05)) {
            problems.push({level: "error", text: "Digital Report channel rows sum to " +
              sum.toFixed(2) + " for " + pk + ", but the Overall row (" +
              tr.periods[pk].actual_cell + ") says " + stored.toFixed(2) + "."});
          }
        });
        findings.push("Channel rows reconcile with the Overall row.");
      }
    }

    if (paid) {
      // Every row group on this tab is displayed from its stored cells, so a
      // mismatch is routed to whichever groups it actually turned up in.
      const groups = [
        [paid.quarters, ["pm-hero", "pm-quarter"]],
        [paid.months, ["pm-month"]],
        [paid.detail, ["pm-detail"]]
      ];
      let cpaBad = 0, roasBad = 0, cpaN = 0, roasN = 0;
      const cpaCells = [], roasCells = [], cpaSec = [], roasSec = [];
      const mark = (list, secs) => secs.forEach(s => { if (list.indexOf(s) < 0) list.push(s); });
      groups.forEach(([rows, secs]) => {
        rows.forEach(r => {
          const M = r.metrics, sp = M.gross_spend && M.gross_spend.value,
                ntl = M.ntl && M.ntl.value, cpa = M.gross_cpa && M.gross_cpa.value;
          if (sp != null && ntl) {
            cpaN++;
            const want = sp / ntl;
            if (!near(cpa, want, Math.max(0.02, Math.abs(want) * 1e-6))) {
              cpaBad++; cpaCells.push(badCell(M.gross_cpa.cell, cpa, want)); mark(cpaSec, secs);
            }
          }
          [["immediate_roas", "immediate_raised"], ["lifetime_roas", "lifetime_raised"],
           ["total_roas", "total_raised"]].forEach(([rk, ak]) => {
            if (!M[rk] || !M[ak] || !sp) return;
            roasN++;
            const want = M[ak].value / sp;
            if (!near(M[rk].value, want, 1e-6)) {
              roasBad++; roasCells.push(badCell(M[rk].cell, M[rk].value, want)); mark(roasSec, secs);
            }
          });
        });
      });
      if (cpaBad) flag(cpaBad + " of " + cpaN + " Gross CPA cells on the Paid Media tab do not equal spend ÷ NTL.",
        cpaSec, cpaCells);
      if (roasBad) flag(roasBad + " of " + roasN + " ROAS cells on the Paid Media tab do not equal raised ÷ spend.",
        roasSec, roasCells);
      if (!cpaBad && !roasBad && (cpaN || roasN)) {
        findings.push("All " + cpaN + " CPA and " + roasN + " ROAS cells match their components.");
      }

      paid.months.forEach(m => {
        const kids = paid.detail.filter(d => String(d.month).trim() === String(m.label).trim());
        if (!kids.length) {
          problems.push({level: "warn", text: "No detail rows match the month “" +
            String(m.label).trim() + "”, so the placement breakdown will not include it."});
          return;
        }
        paid.additive.forEach(key => {
          const stored = m.metrics[key] && m.metrics[key].value;
          if (stored == null) return;
          const sum = kids.reduce((a, d) => a + ((d.metrics[key] && d.metrics[key].value) || 0), 0);
          if (!near(sum, stored, key === "ntl" ? 0.5 : 0.05)) {
            problems.push({level: "error", text: "Paid Media detail rows for " +
              String(m.label).trim() + " sum to " + sum.toFixed(2) + " for " + paid.metric_labels[key] +
              ", but the monthly row (" + m.metrics[key].cell + ") says " + stored.toFixed(2) + "."});
          }
        });
      });

      const qsel = paid.meta.quarter.value;
      const q = paid.quarters.filter(x => x.quarter === qsel)[0];
      if (q) {
        paid.additive.forEach(key => {
          const stored = q.metrics[key] && q.metrics[key].value;
          if (stored == null) return;
          const sum = paid.months.reduce((a, m) => a + ((m.metrics[key] && m.metrics[key].value) || 0), 0);
          if (!near(sum, stored, key === "ntl" ? 0.5 : 0.05)) {
            problems.push({level: "error", text: "Paid Media monthly rows sum to " + sum.toFixed(2) +
              " for " + paid.metric_labels[key] + ", but the Q" + qsel + " row (" + q.metrics[key].cell +
              ") says " + stored.toFixed(2) + "."});
          }
        });
        findings.push("Detail rows reconcile with the monthly rows, and the months with Q" + qsel + ".");
      } else if (qsel != null) {
        problems.push({level: "warn", text: "The tab is set to Q" + qsel +
          " but there is no Q" + qsel + " row in the quarterly table."});
      }
    }

    if (p2p) {
      // Per-row ratio identities: these cells are stored in the workbook
      // (not computed by this page), so a mismatch is a real workbook issue.
      const checks = [
        ["gross_cpa", "gross_spend", "ntl", "Gross CPA", "spend ÷ NTL"],
        ["immediate_roas", "immediate_raised", "gross_spend", "Immediate ROAS", "raised ÷ spend"],
        ["average_gift", "immediate_raised", "donors", "Average Gift", "raised ÷ donors"],
        ["donate_rate", "donors", "recipients", "Donate Rate", "donors ÷ recipients"],
        ["unsub_rate", "unsubs", "recipients", "Unsub Rate", "unsubs ÷ recipients"],
        ["ltv_roas_from_ntl", "ltv_from_ntl", "gross_spend", "LTV ROAS from NTL", "LTV from NTL ÷ spend"]
      ].filter(c => p2p.metric_order.indexOf(c[0]) >= 0);
      checks.forEach(([rk, nk, dk, label, desc]) => {
        let bad = 0, n = 0;
        const cells = [];
        p2p.detail.forEach(r => {
          const M = r.metrics, num = M[nk] && M[nk].value, den = M[dk] && M[dk].value, ratio = M[rk] && M[rk].value;
          if (num == null || !den) return;
          n++;
          const want = num / den;
          if (!near(ratio, want, Math.max(0.02, Math.abs(want) * 1e-6))) {
            bad++; cells.push(badCell(M[rk].cell, ratio, want));
          }
        });
        // Only the row-by-row leaderboard prints these stored cells. The
        // toplines and monthly charts rebuild every ratio from summed
        // components, so they stay correct and stay unannotated.
        if (bad) flag(bad + " of " + n + " " + label + " cells on the P2P Statistics tab do not equal " +
          desc + ".", ["p2-detail"], cells);
        else if (n) findings.push("All " + n + " " + label + " cells on the P2P Statistics tab equal " + desc + ".");
      });
    }

    if (email) {
      // Only the ratios that actually have stored, resolvable components on
      // this sheet get an identity check here — Open Rate and Click Rate
      // have no stored opens/clicks count, so there is no real numerator to
      // check them against, and they are left out on purpose.
      const checks = [
        ["average", "raised", "donors", "Average", "raised ÷ donors"],
        ["donate_rate", "donors", "recipients", "Donate Rate", "donors ÷ recipients"],
        ["action_rate", "actions", "recipients", "Action Rate", "actions ÷ recipients"],
        ["unsub_rate", "unsubs", "recipients", "Unsub Rate", "unsubs ÷ recipients"]
      ].filter(c => email.metric_order.indexOf(c[0]) >= 0);
      checks.forEach(([rk, nk, dk, label, desc]) => {
        let bad = 0, n = 0;
        const cells = [];
        email.detail.forEach(r => {
          const M = r.metrics, num = M[nk] && M[nk].value, den = M[dk] && M[dk].value, ratio = M[rk] && M[rk].value;
          if (num == null || !den) return;
          n++;
          const want = num / den;
          if (!near(ratio, want, Math.max(0.02, Math.abs(want) * 1e-6))) {
            bad++; cells.push(badCell(M[rk].cell, ratio, want));
          }
        });
        // as on P2P: the leaderboard is the only place these stored cells surface
        if (bad) flag(bad + " of " + n + " " + label + " cells on the Email Statistics tab do not equal " +
          desc + ".", ["es-detail"], cells);
        else if (n) findings.push("All " + n + " " + label + " cells on the Email Statistics tab equal " + desc + ".");
      });
    }

    // High-Dollar Donations has a single additive metric and no ratios, so
    // there is no arithmetic identity to check here — its detail-sums-to-
    // rollup checks live in verify_high_dollar.py instead, mirroring how
    // P2P's rollup checks live in verify_p2p.py rather than in this function.

    return {problems: problems, findings: findings};
  }

  /* ===================================================================
     TOP LEVEL
     =================================================================== */
  const DIGITAL_SHEET = /^digital report$/;
  const PAID_SHEET = /^paid media report$/;
  const P2P_SHEET = /^p2p statistics$/;
  const EMAIL_SHEET = /^email statistics$/;
  const HD_SHEET = /^high[- ]dollar donations$/;

  function pickSheet(book, re, label, problems) {
    const exact = book.sheetNames.filter(n => re.test(norm(n)))[0];
    if (exact) return book.sheets[exact];
    // tolerate a renamed tab if only one plausibly matches
    const words = label.toLowerCase().split(" ");
    const loose = book.sheetNames.filter(n => words.every(w => norm(n).indexOf(w) >= 0));
    if (loose.length === 1) {
      problems.push({level: "warn", text: "No tab named “" + label + "”; using “" + loose[0] +
        "” instead, which looks like a renamed version."});
      return book.sheets[loose[0]];
    }
    problems.push({level: "error", text: "No “" + label + "” tab found in this workbook. " +
      "Tabs present: " + book.sheetNames.join(", ") + "."});
    return null;
  }

  /* book: result of XlsxReader.read; returns the full page payload + report */
  function extractAll(book, opts) {
    opts = opts || {};
    const o = {
      workbookName: opts.workbookName || "workbook.xlsx",
      extractedAt: opts.extractedAt || new Date().toISOString().replace(/\.\d+Z$/, "")
    };
    const problems = [], findings = [];
    findings.push("Workbook has " + book.sheetNames.length + " tab(s).");

    const digProblems = [], paidProblems = [], p2pProblems = [], emailProblems = [], hdProblems = [];
    const ds = pickSheet(book, DIGITAL_SHEET, "Digital Report", digProblems);
    const ps = pickSheet(book, PAID_SHEET, "Paid Media Report", paidProblems);
    const p2s = pickSheet(book, P2P_SHEET, "P2P Statistics", p2pProblems);
    const es = pickSheet(book, EMAIL_SHEET, "Email Statistics", emailProblems);
    const hs = pickSheet(book, HD_SHEET, "High-Dollar Donations", hdProblems);

    let dig = null, paid = null, p2p = null, email = null, hd = null;
    if (ds) {
      const r = extractDigital(ds, o);
      dig = r.payload;
      r.findings.forEach(f => findings.push("Digital Report: " + f));
      r.problems.forEach(p => digProblems.push({level: p.level, text: "Digital Report: " + p.text}));
    }
    if (ps) {
      const r = extractPaid(ps, o);
      paid = r.payload;
      r.findings.forEach(f => findings.push("Paid Media: " + f));
      r.problems.forEach(p => paidProblems.push({level: p.level, text: "Paid Media: " + p.text}));
    }
    if (p2s) {
      const r = extractP2P(p2s, o);
      p2p = r.payload;
      r.findings.forEach(f => findings.push("P2P Statistics: " + f));
      r.problems.forEach(p => p2pProblems.push({level: p.level, text: "P2P Statistics: " + p.text}));
    }
    if (es) {
      const r = extractEmailStats(es, o);
      email = r.payload;
      r.findings.forEach(f => findings.push("Email Statistics: " + f));
      r.problems.forEach(p => emailProblems.push({level: p.level, text: "Email Statistics: " + p.text}));
    }
    if (hs) {
      const r = extractHighDollar(hs, o);
      hd = r.payload;
      r.findings.forEach(f => findings.push("High-Dollar Donations: " + f));
      r.problems.forEach(p => hdProblems.push({level: p.level, text: "High-Dollar Donations: " + p.text}));
    }
    // Reconciliation results are attributed to the tab they concern, so one
    // broken tab never suppresses a good one. Errors hide their tab; flags
    // (level "flag") do not — they ride along and get pinned to the sections
    // named in p.sections once the tab renders.
    const rec = reconcile(dig, paid, p2p, email, hd);
    rec.problems.forEach(p => {
      if (/paid media/i.test(p.text)) paidProblems.push(p);
      else if (/digital report/i.test(p.text)) digProblems.push(p);
      else if (/email statistics/i.test(p.text)) emailProblems.push(p);
      else if (/high-dollar donations/i.test(p.text)) hdProblems.push(p);
      else if (/p2p statistics/i.test(p.text)) p2pProblems.push(p);
      else problems.push(p);
    });
    rec.findings.forEach(f => findings.push(f));

    digProblems.forEach(p => { p.tab = "digital"; problems.push(p); });
    paidProblems.forEach(p => { p.tab = "paid"; problems.push(p); });
    p2pProblems.forEach(p => { p.tab = "p2p"; problems.push(p); });
    emailProblems.forEach(p => { p.tab = "email"; problems.push(p); });
    hdProblems.forEach(p => { p.tab = "hd"; problems.push(p); });

    const count = (list, lvl) => list.filter(p => p.level === lvl).length;
    return {
      digital: dig, paid: paid, p2p: p2p, email: email, highDollar: hd,
      report: {
        findings: findings,
        problems: problems,
        errors: count(problems, "error"),
        warnings: count(problems, "warn"),
        flagged: count(problems, "flag"),
        digitalErrors: count(digProblems, "error"),
        paidErrors: count(paidProblems, "error"),
        p2pErrors: count(p2pProblems, "error"),
        emailErrors: count(emailProblems, "error"),
        hdErrors: count(hdProblems, "error"),
        workbookName: o.workbookName,
        extractedAt: o.extractedAt,
        sheetNames: book.sheetNames
      }
    };
  }

  return {
    extractAll: extractAll,
    extractDigital: extractDigital,
    extractPaid: extractPaid,
    extractP2P: extractP2P,
    extractEmailStats: extractEmailStats,
    extractHighDollar: extractHighDollar,
    reconcile: reconcile,
    _helpers: {norm: norm, slug: slug, findRow: findRow, headerMap: headerMap}
  };
});
