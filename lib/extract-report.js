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
     problems  {level: "error"|"warn", text} — errors mean the section
               cannot be trusted and the page will say so
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
    if (/cpa|cost per/.test(n)) return "usd2";
    if (/^ntl$|names|count/.test(n)) return "count";
    return "usd";
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
     RECONCILIATION — the checks that decide whether to trust a section.
     These are the same identities the offline verifiers assert, run again
     on whatever was just dropped in.
     =================================================================== */
  function reconcile(dig, paid) {
    const problems = [], findings = [];
    const near = (a, b, tol) => a != null && b != null && Math.abs(a - b) <= (tol == null ? 0.02 : tol);

    if (dig) {
      let bad = 0, checked = 0;
      dig.channels.forEach(c => {
        Object.keys(c.periods).forEach(pk => {
          const p = c.periods[pk];
          if (!p.goal || p.actual == null || p.pct == null) return;
          checked++;
          if (!near(p.pct, p.actual / p.goal, 1e-6)) bad++;
        });
      });
      if (bad) problems.push({level: "error",
        text: bad + " of " + checked + " “% to goal” cells on the Digital Report do not equal actual ÷ goal."});
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
      const rows = paid.quarters.concat(paid.months).concat(paid.detail);
      let cpaBad = 0, roasBad = 0, cpaN = 0, roasN = 0;
      rows.forEach(r => {
        const M = r.metrics, sp = M.gross_spend && M.gross_spend.value,
              ntl = M.ntl && M.ntl.value, cpa = M.gross_cpa && M.gross_cpa.value;
        if (sp != null && ntl) { cpaN++; if (!near(cpa, sp / ntl, Math.max(0.02, Math.abs(sp / ntl) * 1e-6))) cpaBad++; }
        [["immediate_roas", "immediate_raised"], ["lifetime_roas", "lifetime_raised"],
         ["total_roas", "total_raised"]].forEach(([rk, ak]) => {
          if (!M[rk] || !M[ak] || !sp) return;
          roasN++;
          if (!near(M[rk].value, M[ak].value / sp, 1e-6)) roasBad++;
        });
      });
      if (cpaBad) problems.push({level: "error",
        text: cpaBad + " of " + cpaN + " Gross CPA cells on the Paid Media tab do not equal spend ÷ NTL."});
      if (roasBad) problems.push({level: "error",
        text: roasBad + " of " + roasN + " ROAS cells on the Paid Media tab do not equal raised ÷ spend."});
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
    return {problems: problems, findings: findings};
  }

  /* ===================================================================
     TOP LEVEL
     =================================================================== */
  const DIGITAL_SHEET = /^digital report$/;
  const PAID_SHEET = /^paid media report$/;

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

    const digProblems = [], paidProblems = [];
    const ds = pickSheet(book, DIGITAL_SHEET, "Digital Report", digProblems);
    const ps = pickSheet(book, PAID_SHEET, "Paid Media Report", paidProblems);

    let dig = null, paid = null;
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
    // Reconciliation failures are attributed to the tab they concern, so one
    // broken tab never suppresses a good one.
    const rec = reconcile(dig, paid);
    rec.problems.forEach(p => {
      if (/paid media/i.test(p.text)) paidProblems.push(p);
      else if (/digital report/i.test(p.text)) digProblems.push(p);
      else problems.push(p);
    });
    rec.findings.forEach(f => findings.push(f));

    digProblems.forEach(p => { p.tab = "digital"; problems.push(p); });
    paidProblems.forEach(p => { p.tab = "paid"; problems.push(p); });

    const count = (list, lvl) => list.filter(p => p.level === lvl).length;
    return {
      digital: dig, paid: paid,
      report: {
        findings: findings,
        problems: problems,
        errors: count(problems, "error"),
        warnings: count(problems, "warn"),
        digitalErrors: count(digProblems, "error"),
        paidErrors: count(paidProblems, "error"),
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
    reconcile: reconcile,
    _helpers: {norm: norm, slug: slug, findRow: findRow, headerMap: headerMap}
  };
});
