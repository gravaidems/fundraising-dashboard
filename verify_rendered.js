/*
 * Render digital-report.html in jsdom, scrape every figure the page actually
 * displays, and cross-check each one against the source workbook values held in
 * digital-report-data.json.
 *
 * This is the last gate: extract/verify prove the JSON matches the workbook,
 * and this proves the rendered page matches the JSON.
 *
 * Usage: node verify_rendered.js [digital-report.html] [digital-report-data.json]
 */
const fs = require("fs");
const path = require("path");
const {JSDOM} = require(path.join(process.env.JSDOM_PATH || "/tmp/vfy/node_modules", "jsdom"));

const htmlPath = process.argv[2] || "digital-report.html";
const dataPath = process.argv[3] || "digital-report-data.json";

const D = JSON.parse(fs.readFileSync(dataPath, "utf8"));
const MONTHS = D.month_keys;
const PENDING = new Set(D.pending_months || []);

const results = [];
const check = (name, ok, detail) => results.push({ok, name, detail});

// Parse "$1,234.56" / "12,345" / "20.89%" back into a number.
const parseNum = s => {
  if (s == null) return null;
  const t = String(s).replace(/[^0-9.\-]/g, "");
  if (t === "" || t === "-" || t === ".") return null;
  const n = parseFloat(t);
  return isNaN(n) ? null : n;
};
const close = (a, b, tol) => a != null && b != null && Math.abs(a - b) <= (tol == null ? 0.02 : tol);

/* The page ships with no data, so the workbook is loaded through the same path a
   reader uses: the page's own parser, driven by its own load function. Every
   check below therefore tests the live parse, not a baked-in snapshot. */
const wbPath = process.argv[5] || "EXTERNAL_ CAB Digital Report.xlsx";

(async function main() {
// jsdom logs "Not implemented" for a few window APIs the page guards anyway;
// keep those out of the check output, but surface real script errors.
const vc = new (require(path.join(process.env.JSDOM_PATH || "/tmp/vfy/node_modules", "jsdom")).VirtualConsole)();
vc.on("jsdomError", e => { if (!/Not implemented/.test(e.message)) console.error(e.message); });
const dom = new JSDOM(fs.readFileSync(htmlPath, "utf8"),
  {runScripts: "dangerously", pretendToBeVisual: true, virtualConsole: vc});
const doc = dom.window.document;
const txt = el => (el ? el.textContent.replace(/\s+/g, " ").trim() : "");

if (!fs.existsSync(wbPath)) {
  console.error("workbook not found: " + wbPath);
  process.exit(2);
}
// jsdom has no DecompressionStream; hand the reader node's zlib instead
dom.window.XlsxReader.setInflate(b => new Uint8Array(require("zlib").inflateRawSync(Buffer.from(b))));

check("page shows nothing until a workbook is loaded",
  !doc.getElementById("emptyState").hidden &&
  doc.getElementById("pgHeader").hidden &&
  [...doc.querySelectorAll(".panel")].every(p => p.hidden),
  "empty state hidden=" + doc.getElementById("emptyState").hidden +
  ", header hidden=" + doc.getElementById("pgHeader").hidden +
  ", visible panels=" + [...doc.querySelectorAll(".panel")].filter(p => !p.hidden).length);
const shipped = fs.readFileSync(htmlPath, "utf8");
check("no workbook data is baked into the page file",
  /let DATA = null;/.test(shipped) && /let PM = null;/.test(shipped) &&
  !/"source_workbook"|"month_keys"|"metric_order"/.test(shipped),
  "the shipped page carries a serialised payload");
check("the shipped page names no client",
  !new RegExp(String(D.meta.client.value || "\\u0000").replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
    .test(shipped),
  "the client name appears in the shipped file");
check("the empty state explains what to do",
  /drop|choose/i.test(txt(doc.getElementById("emptyState"))),
  "empty state text: " + txt(doc.getElementById("emptyState")).slice(0, 120));

const wbBuf = fs.readFileSync(wbPath);
await dom.window.__dashboard.load({
  name: path.basename(wbPath), size: wbBuf.length,
  arrayBuffer: async () => wbBuf.buffer.slice(wbBuf.byteOffset, wbBuf.byteOffset + wbBuf.byteLength)
});
check("loading a workbook reveals the report",
  doc.getElementById("emptyState").hidden && !doc.getElementById("pgHeader").hidden,
  "empty state hidden=" + doc.getElementById("emptyState").hidden);

// ---------- 0. the page rendered at all -------------------------------------
check("page has no script errors and rendered heroes",
  doc.querySelectorAll("#heroes .hero").length === 4,
  "found " + doc.querySelectorAll("#heroes .hero").length + " hero cards, expected 4");
check("month chart rendered", doc.querySelectorAll("#chMonth svg").length === 1,
  "svg count " + doc.querySelectorAll("#chMonth svg").length);
check("channel bars rendered", doc.querySelectorAll("#chChannel .cbar").length === D.channels.length,
  "rows " + doc.querySelectorAll("#chChannel .cbar").length + " vs channels " + D.channels.length);
const activeCount = D.channels.filter(c => !c.dormant).length;
check("accumulation small multiples rendered", doc.querySelectorAll("#chAccum .sm").length === activeCount,
  "charts " + doc.querySelectorAll("#chAccum .sm").length + " vs active channels " + activeCount);

// ---------- 1. header metadata ---------------------------------------------
const metaTxt = txt(doc.getElementById("pgMeta"));
[["client", D.meta.client.value], ["email lead", D.meta.email_lead.value]].forEach(([k, v]) => {
  check("header shows " + k, metaTxt.includes(String(v)), k + " = " + v + " not found in: " + metaTxt);
});

// ---------- 2. hero stats ---------------------------------------------------
const q = k => D.overall[k].periods.quarter;
const heroes = [...doc.querySelectorAll("#heroes .hero")].map(h => ({
  label: txt(h.querySelector(".lbl")).replace(/\s*i$/, "").trim(),
  big: parseNum(txt(h.querySelector(".big"))),
  pct: parseNum(txt(h.querySelector(".pct")))
}));
const heroMap = {
  "Total raised": {v: q("total_raised").actual, pct: q("total_raised").pct},
  "Total donations": {v: q("total_donations").actual, pct: null},
  "Average donation": {v: q("average_donation").actual, pct: null},
  "Total finance": {v: q("total_finance").actual, pct: null}
};
Object.keys(heroMap).forEach(label => {
  const h = heroes.find(x => x.label.toLowerCase() === label.toLowerCase());
  if (!h) { check("hero card '" + label + "' present", false, "not rendered"); return; }
  const exp = heroMap[label].v;
  // Total raised and donations render without cents, so tolerance is 1 unit.
  const tol = (label === "Total raised" || label === "Total donations") ? 1 : 0.02;
  check("hero '" + label + "' value", close(h.big, exp, tol),
    "displayed " + h.big + " vs workbook " + exp);
  if (heroMap[label].pct != null) {
    check("hero '" + label + "' % to goal", close(h.pct, heroMap[label].pct * 100, 0.1),
      "displayed " + h.pct + "% vs workbook " + (heroMap[label].pct * 100).toFixed(4) + "%");
  }
});

// ---------- 3. month table vs workbook column sums --------------------------
const chSum = (period, field) => {
  let t = 0, any = false;
  D.channels.forEach(c => { const v = c.periods[period][field]; if (v != null) { t += v; any = true; } });
  return any ? t : null;
};
const monthRows = [...doc.querySelectorAll("#tbMonth tbody tr")];
check("month table has one row per month", monthRows.length === MONTHS.length,
  "rows " + monthRows.length + " vs months " + MONTHS.length);
MONTHS.forEach((m, i) => {
  const tr = monthRows[i];
  if (!tr) return;
  const td = [...tr.children].map(x => txt(x));
  const label = D.channels[0].periods[m].label;
  check("month table row " + label + " label", td[0] === label, "got '" + td[0] + "'");
  check("month table " + label + " goal", close(parseNum(td[1]), chSum(m, "goal")),
    "displayed " + td[1] + " vs channel sum " + chSum(m, "goal"));
  if (PENDING.has(m)) {
    check("month table " + label + " actual marked pending", /not reported/i.test(td[2]),
      "expected 'not reported', got '" + td[2] + "'");
    check("month table " + label + " pct suppressed", !/\d/.test(td[3]),
      "expected no number, got '" + td[3] + "'");
  } else {
    const exp = chSum(m, "actual");
    check("month table " + label + " actual", close(parseNum(td[2]), exp),
      "displayed " + td[2] + " vs channel sum " + exp);
    check("month table " + label + " pct", close(parseNum(td[3]), (exp / chSum(m, "goal")) * 100, 0.01),
      "displayed " + td[3] + " vs computed " + ((exp / chSum(m, "goal")) * 100).toFixed(4));
    // and against the workbook's own Overall row for that month
    const ow = D.overall.total_raised.periods[m];
    check("month " + label + " matches workbook Overall row " + ow.actual_cell,
      close(exp, ow.actual), "channel sum " + exp + " vs " + ow.actual_cell + " = " + ow.actual);
  }
});
// footer = quarter
const mFoot = [...doc.querySelectorAll("#tbMonth tfoot td")].map(x => txt(x));
check("month table footer goal = C23", close(parseNum(mFoot[1]), q("total_raised").goal),
  "displayed " + mFoot[1] + " vs C23 " + q("total_raised").goal);
check("month table footer actual = D23", close(parseNum(mFoot[2]), q("total_raised").actual),
  "displayed " + mFoot[2] + " vs D23 " + q("total_raised").actual);

// ---------- 4. channel table vs workbook rows ------------------------------
const chRows = [...doc.querySelectorAll("#tbChannel tbody tr")];
check("channel table has one row per channel", chRows.length === D.channels.length,
  "rows " + chRows.length);
chRows.forEach(tr => {
  const td = [...tr.children].map(x => txt(x));
  const c = D.channels.find(x => x.name === td[0]);
  if (!c) { check("channel table row '" + td[0] + "' maps to workbook", false, "unknown channel"); return; }
  const p = c.periods.quarter;
  check(c.name + " table goal = " + p.goal_cell, close(parseNum(td[1]), p.goal || 0),
    "displayed " + td[1] + " vs " + p.goal);
  check(c.name + " table actual = " + p.actual_cell, close(parseNum(td[2]), p.actual || 0),
    "displayed " + td[2] + " vs " + p.actual);
  if (p.pct == null) {
    check(c.name + " table pct is dash", !/\d/.test(td[3]), "expected dash, got '" + td[3] + "'");
  } else {
    check(c.name + " table pct = " + p.pct_cell, close(parseNum(td[3]), p.pct * 100, 0.01),
      "displayed " + td[3] + " vs " + (p.pct * 100).toFixed(4));
  }
  check(c.name + " table cites its cells",
    td[4].includes(p.goal_cell) && td[4].includes(p.actual_cell) && td[4].includes(p.pct_cell),
    "cell refs '" + td[4] + "' missing one of " + [p.goal_cell, p.actual_cell, p.pct_cell].join("/"));
});
// channel table sorted by pct descending, dormant last
const order = chRows.map(tr => txt(tr.children[0]));
const pctOf = n => { const c = D.channels.find(x => x.name === n); return c.dormant ? -1 : (c.periods.quarter.pct || 0); };
let sorted = true;
for (let i = 1; i < order.length; i++) if (pctOf(order[i]) > pctOf(order[i - 1]) + 1e-9) sorted = false;
check("channel table sorted by % to goal, dormant last", sorted, "order: " + order.join(", "));

// ---------- 5. accumulation table = running sums ---------------------------
const accRows = [...doc.querySelectorAll("#tbAccum tbody tr")];
check("accumulation table has one row per active channel", accRows.length === activeCount,
  "rows " + accRows.length + " vs " + activeCount);
accRows.forEach(tr => {
  const td = [...tr.children].map(x => txt(x));
  const c = D.channels.find(x => x.name === td[0]);
  if (!c) { check("accum row '" + td[0] + "' maps to workbook", false, "unknown"); return; }
  let runA = 0, runG = 0;
  MONTHS.forEach((m, i) => {
    const a = c.periods[m].actual, g = c.periods[m].goal || 0;
    runG += g;
    const shownA = td[1 + i], shownG = td[1 + MONTHS.length + i];
    if (a == null) {
      check(c.name + " cum actual " + m + " pending", !/\d/.test(shownA),
        "expected dash, got '" + shownA + "'");
    } else {
      runA += a;
      check(c.name + " cum actual " + m, close(parseNum(shownA), runA),
        "displayed " + shownA + " vs running sum " + runA.toFixed(2));
    }
    check(c.name + " cum goal " + m, close(parseNum(shownG), runG),
      "displayed " + shownG + " vs running sum " + runG.toFixed(2));
  });
  // final cumulative goal must equal the quarterly goal cell
  check(c.name + " final cum goal = " + c.periods.quarter.goal_cell,
    close(runG, c.periods.quarter.goal || 0),
    "cumulative " + runG.toFixed(2) + " vs " + c.periods.quarter.goal);
});

// ---------- 6. info bubbles present and cite cells -------------------------
const bubbles = [...doc.querySelectorAll("button.info")];
check("info bubble on every hero card",
  [...doc.querySelectorAll("#heroes .hero")].every(h => h.querySelector("button.info")),
  "one or more hero cards lack an info bubble");
["s-month", "s-channel", "s-accum"].forEach(id => {
  check("info bubble on section " + id,
    !!doc.querySelector("#" + id + " .sechead button.info"), "missing");
});
check("info bubble on every channel bar",
  [...doc.querySelectorAll("#chChannel .cbar")].every(r => r.querySelector("button.info")),
  "one or more channel bars lack an info bubble");
check("info bubble on every accumulation chart",
  [...doc.querySelectorAll("#chAccum .sm")].every(r => r.querySelector("button.info")),
  "one or more small multiples lack an info bubble");
check("total info bubbles >= 4 heroes + 3 sections + 8 channels + active charts",
  bubbles.length >= 4 + 3 + D.channels.length + activeCount,
  "found " + bubbles.length);

// Open every bubble and confirm it renders content naming at least one real
// cell. Cell refs are marked up as <code> elements, so inspect those rather
// than the concatenated textContent (where "Value" + "D25" glues together).
const CELL_RE = /^[A-S]\d{1,2}(:[A-S]\d{1,2})?$/;
const bubbleProblems = [];
bubbles.forEach(b => {
  b.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
  const pop = doc.querySelector(".pop");
  const label = b.getAttribute("aria-label") || "(unlabelled)";
  if (!pop) { bubbleProblems.push(label + ": did not open"); return; }
  const codes = [...pop.querySelectorAll("code")].map(c => txt(c));
  const refs = codes.filter(c => CELL_RE.test(c));
  const body = txt(pop);
  if (!refs.length) bubbleProblems.push(label + ": cites no cell (codes: " + codes.join(",") + ")");
  else if (body.length < 40) bubbleProblems.push(label + ": body too short (" + body.length + " chars)");
  else if (!/Calculation|÷|\+|SUM|running sum/i.test(body)) bubbleProblems.push(label + ": no calculation shown");
  b.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
});
check("every info bubble opens, cites a cell, and shows a calculation",
  bubbleProblems.length === 0,
  bubbleProblems.length + " of " + bubbles.length + " problems:\n        " + bubbleProblems.join("\n        "));

// ---------- 6b. hover tooltips ---------------------------------------------
// Every chart must expose hover targets, and each tooltip's numbers must match
// the workbook. Tooltips are rendered into a .tip element on mouseenter.
const fireHover = el => el.dispatchEvent(new dom.window.MouseEvent("mouseenter", {bubbles: true, clientX: 40, clientY: 40}));
const fireOut = el => el.dispatchEvent(new dom.window.MouseEvent("mouseleave", {bubbles: true}));
const tipNow = () => doc.querySelector(".tip");

const monthHits = [...doc.querySelectorAll("#chMonth [data-tip-id]")];
check("month chart has one hover target per month", monthHits.length === MONTHS.length,
  "targets " + monthHits.length + " vs months " + MONTHS.length);
monthHits.forEach((h, i) => {
  fireHover(h);
  const tip = tipNow();
  const m = MONTHS[i], label = D.channels[0].periods[m].label;
  if (!tip || tip.style.display === "none") {
    check("month tooltip " + label + " shows on hover", false, "tooltip not displayed");
    return;
  }
  const vals = [...tip.querySelectorAll("dd")].map(x => txt(x));
  const keys = [...tip.querySelectorAll("dt")].map(x => txt(x));
  const get = k => { const i2 = keys.indexOf(k); return i2 < 0 ? null : vals[i2]; };
  check("month tooltip " + label + " title", txt(tip.querySelector(".th")).includes(label),
    "title '" + txt(tip.querySelector(".th")) + "'");
  check("month tooltip " + label + " goal", close(parseNum(get("Goal")), chSum(m, "goal")),
    "tooltip " + get("Goal") + " vs " + chSum(m, "goal"));
  if (PENDING.has(m)) {
    check("month tooltip " + label + " actual pending", /not yet reported/i.test(get("Actual") || ""),
      "got '" + get("Actual") + "'");
  } else {
    check("month tooltip " + label + " actual", close(parseNum(get("Actual")), chSum(m, "actual")),
      "tooltip " + get("Actual") + " vs " + chSum(m, "actual"));
    const expPct = chSum(m, "actual") / chSum(m, "goal");
    check("month tooltip " + label + " % to goal", close(parseNum(get("% to goal")), expPct * 100, 0.06),
      "tooltip " + get("% to goal") + " vs " + (expPct * 100).toFixed(3));
    const expGap = chSum(m, "actual") - chSum(m, "goal");
    check("month tooltip " + label + " gap", close(Math.abs(parseNum(get("Gap to goal"))), Math.abs(expGap), 0.02),
      "tooltip " + get("Gap to goal") + " vs " + expGap.toFixed(2));
  }
  check("month tooltip " + label + " cites source cells",
    /[A-S]\d{1,2}:[A-S]\d{1,2}/.test(txt(tip.querySelector(".tf"))),
    "footer '" + txt(tip.querySelector(".tf")) + "'");
  fireOut(h);
});

const chanHits = [...doc.querySelectorAll("#chChannel [data-tip-id]")];
check("channel chart has one hover target per channel", chanHits.length === D.channels.length,
  "targets " + chanHits.length);
chanHits.forEach(h => {
  const name = txt(h.querySelector(".nm span"));
  const c = D.channels.find(x => x.name === name);
  fireHover(h);
  const tip = tipNow();
  if (!c || !tip) { check("channel tooltip for '" + name + "'", false, "missing"); return; }
  const keys = [...tip.querySelectorAll("dt")].map(x => txt(x));
  const vals = [...tip.querySelectorAll("dd")].map(x => txt(x));
  const get = k => { const i2 = keys.indexOf(k); return i2 < 0 ? null : vals[i2]; };
  const q2 = c.periods.quarter;
  check("channel tooltip " + name + " goal", close(parseNum(get("Goal")), q2.goal || 0),
    "tooltip " + get("Goal") + " vs " + q2.goal);
  check("channel tooltip " + name + " actual", close(parseNum(get("Actual")), q2.actual || 0),
    "tooltip " + get("Actual") + " vs " + q2.actual);
  if (q2.pct != null) {
    check("channel tooltip " + name + " % to goal", close(parseNum(get("% to goal")), q2.pct * 100, 0.06),
      "tooltip " + get("% to goal") + " vs " + (q2.pct * 100).toFixed(3));
  }
  // monthly breakdown rows must match the monthly cells
  MONTHS.forEach(m => {
    const p = c.periods[m], shown = get(p.label);
    if (shown == null) { check(name + " tooltip lists " + p.label, false, "row missing"); return; }
    if (p.actual == null) {
      check(name + " tooltip " + p.label + " pending", /not reported/i.test(shown), "got '" + shown + "'");
    } else {
      check(name + " tooltip " + p.label + " actual", close(parseNum(shown.split("/")[0]), p.actual),
        "tooltip '" + shown + "' vs " + p.actual);
    }
  });
  fireOut(h);
});

const accHits = [...doc.querySelectorAll("#chAccum [data-tip-id]")];
check("accumulation charts have a hover target per month per chart",
  accHits.length === activeCount * MONTHS.length,
  "targets " + accHits.length + " vs expected " + activeCount * MONTHS.length);
// spot-check every accumulation tooltip against recomputed running sums
[...doc.querySelectorAll("#chAccum .sm")].forEach(sm => {
  const name = txt(sm.querySelector("h3 span"));
  const c = D.channels.find(x => x.name === name);
  const hits = [...sm.querySelectorAll("[data-tip-id]")];
  if (!c) { check("accum chart '" + name + "' maps to workbook", false, "unknown channel"); return; }
  let runA = 0, runG = 0;
  MONTHS.forEach((m, i) => {
    const p = c.periods[m];
    runG += p.goal || 0;
    if (p.actual != null) runA += p.actual;
    const h = hits[i];
    if (!h) { check(name + " accum hover target " + m, false, "missing"); return; }
    fireHover(h);
    const tip = tipNow();
    const keys = [...tip.querySelectorAll("dt")].map(x => txt(x));
    const vals = [...tip.querySelectorAll("dd")].map(x => txt(x));
    const get = k => { const i2 = keys.indexOf(k); return i2 < 0 ? null : vals[i2]; };
    check(name + " accum tooltip " + p.label + " cumulative goal",
      close(parseNum(get("Cumulative goal")), runG),
      "tooltip " + get("Cumulative goal") + " vs running sum " + runG.toFixed(2));
    if (p.actual == null) {
      check(name + " accum tooltip " + p.label + " actual pending",
        /not yet reported/i.test(get("Cumulative actual") || ""), "got '" + get("Cumulative actual") + "'");
    } else {
      check(name + " accum tooltip " + p.label + " cumulative actual",
        close(parseNum(get("Cumulative actual")), runA),
        "tooltip " + get("Cumulative actual") + " vs running sum " + runA.toFixed(2));
      check(name + " accum tooltip " + p.label + " pace",
        close(parseNum(get("Pace to date")), (runA / runG) * 100, 0.06),
        "tooltip " + get("Pace to date") + " vs " + ((runA / runG) * 100).toFixed(3));
    }
    fireOut(h);
  });
});

// ---------- 6c. requested styling changes ----------------------------------
const rawHtml = fs.readFileSync(htmlPath, "utf8");
check("goal bars in the month chart use the new --goal-bar color",
  /--goal-bar:/.test(rawHtml) && /fill="var\(--goal-bar\)"/.test(rawHtml),
  "token or fill reference missing");
check("month chart goal bars no longer use the old grey --goal fill",
  !/fill="var\(--goal\)" opacity/.test(rawHtml), "old grey goal fill still present");
check("--goal-bar defined for both light and dark",
  (rawHtml.match(/--goal-bar:/g) || []).length >= 2,
  "found " + (rawHtml.match(/--goal-bar:/g) || []).length + " definitions");

const trackWidths = [...doc.querySelectorAll("#chChannel .cbar .track")].map(t => t.style.width);
check("every channel goal bar is the same length",
  trackWidths.length === D.channels.length && new Set(trackWidths).size === 1 && trackWidths[0] === "100%",
  "widths: " + [...new Set(trackWidths)].join(", "));
// fills must still encode % to goal
[...doc.querySelectorAll("#chChannel .cbar")].forEach(row => {
  const name = txt(row.querySelector(".nm span"));
  const c = D.channels.find(x => x.name === name);
  const fill = row.querySelector(".fill");
  if (!c || !c.periods.quarter.goal) return;
  const expected = Math.min(100, (c.periods.quarter.actual / c.periods.quarter.goal) * 100);
  check(name + " fill width encodes % to goal",
    close(parseNum(fill && fill.style.width), expected, 0.5),
    "fill " + (fill && fill.style.width) + " vs expected " + expected.toFixed(2) + "%");
});

check("accumulation grid is two charts per row",
  /\.smgrid\{[^}]*grid-template-columns:repeat\(2,minmax\(0,1fr\)\)/.test(rawHtml),
  "smgrid is not fixed at 2 columns");
check("accumulation grid collapses to one column on narrow screens",
  /max-width:760px\)\{\s*\.smgrid[^{]*\{grid-template-columns:minmax\(0,1fr\)\}/.test(rawHtml),
  "no single-column fallback found");
const accVb = [...doc.querySelectorAll("#chAccum svg")].map(s => s.getAttribute("viewBox"));
check("accumulation charts use the enlarged canvas",
  accVb.length > 0 && accVb.every(v => v === "0 0 520 210"),
  "viewBoxes: " + [...new Set(accVb)].join(", "));

// ---------- 7. self-containment --------------------------------------------
const raw = fs.readFileSync(htmlPath, "utf8");
const ext = raw.match(/(?:src|href)\s*=\s*["'](?!#)([a-z]+:)?\/\//gi) || [];
check("no external resource references", ext.length === 0, "found: " + ext.join(", "));
check("no localStorage/sessionStorage use", !/localStorage|sessionStorage/.test(raw), "found browser storage use");

// ---------- 8. September is never shown as $0 ------------------------------
const overviewTxt = txt(doc.getElementById("p-overview"));
check("page states September is not yet reported",
  /not yet reported|pending|not reported/i.test(overviewTxt), "no pending language found");

// ==========================================================================
// PAID MEDIA TAB
// ==========================================================================
const pmPath = process.argv[4] || "paid-media-data.json";
if (fs.existsSync(pmPath)) {
  const P = JSON.parse(fs.readFileSync(pmPath, "utf8"));
  const LBL = P.metric_labels, KND = P.metric_kinds;
  const QSEL = P.meta.quarter.value;
  const roasOf = v => v == null ? null : v;

  // tooltip readers
  const tipPairs = () => {
    const tip = tipNow();
    if (!tip) return null;
    const keys = [...tip.querySelectorAll("dt")].map(x => txt(x));
    const vals = [...tip.querySelectorAll("dd")].map(x => txt(x));
    return {
      title: txt(tip.querySelector(".th")),
      foot: txt(tip.querySelector(".tf")),
      get: k => { const i = keys.indexOf(k); return i < 0 ? null : vals[i]; }
    };
  };

  // ---------- tab wiring ----------
  check("Paid Media tab button exists", !!doc.getElementById("t-paid"), "missing");
  check("Paid Media panel exists", !!doc.getElementById("p-paid"), "missing");
  check("Paid Media panel starts hidden", doc.getElementById("p-paid").hasAttribute("hidden"),
    "panel should be hidden until its tab is selected");

  // ---------- scope banner ----------
  const scope = txt(doc.getElementById("pmScope"));
  check("scope banner names the year-to-date span", P.ytd.quarters.every(q => scope.includes(q)),
    "banner: " + scope.slice(0, 160));
  check("scope banner warns that monthly and detail cover only the selected quarter",
    scope.includes("Q" + QSEL) && /only/i.test(scope), "banner: " + scope.slice(0, 200));

  // ---------- 1. year-to-date heroes ----------
  const sumQ = k => P.quarters.reduce((a, q) => a + (q.metrics[k].value || 0), 0);
  const ytdExp = {
    gross_spend: sumQ("gross_spend"), ntl: sumQ("ntl"),
    immediate_raised: sumQ("immediate_raised"), lifetime_raised: sumQ("lifetime_raised"),
    total_raised: sumQ("total_raised")
  };
  ytdExp.gross_cpa = ytdExp.gross_spend / ytdExp.ntl;
  ytdExp.immediate_roas = ytdExp.immediate_raised / ytdExp.gross_spend;
  ytdExp.lifetime_roas = ytdExp.lifetime_raised / ytdExp.gross_spend;
  ytdExp.total_roas = ytdExp.total_raised / ytdExp.gross_spend;

  const pmHeroes = [...doc.querySelectorAll("#pmHeroes .hero")];
  check("paid media renders six year-to-date hero cards", pmHeroes.length === 6,
    "found " + pmHeroes.length);
  const heroExpect = [
    ["Gross spend", "gross_spend", 1], ["NTL added", "ntl", 1], ["Gross CPA", "gross_cpa", 0.01],
    ["Immediate raised", "immediate_raised", 1], ["Lifetime raised", "lifetime_raised", 1],
    ["Total raised", "total_raised", 1]
  ];
  heroExpect.forEach(([label, key, tol]) => {
    const h = pmHeroes.find(x => txt(x.querySelector(".lbl")).replace(/\s*i$/, "").trim() === label);
    if (!h) { check("YTD hero '" + label + "' present", false, "not rendered"); return; }
    check("YTD hero '" + label + "' value", close(parseNum(txt(h.querySelector(".big"))), ytdExp[key], tol),
      "displayed " + txt(h.querySelector(".big")) + " vs expected " + ytdExp[key]);
  });
  // ROAS sub-lines on the three Raised cards
  [["Immediate raised", "immediate_roas"], ["Lifetime raised", "lifetime_roas"],
   ["Total raised", "total_roas"]].forEach(([label, rk]) => {
    const h = pmHeroes.find(x => txt(x.querySelector(".lbl")).replace(/\s*i$/, "").trim() === label);
    if (!h) return;
    check("YTD hero '" + label + "' shows its ROAS",
      close(parseNum(txt(h.querySelector(".of"))), ytdExp[rk], 0.01),
      "sub-line '" + txt(h.querySelector(".of")) + "' vs expected " + ytdExp[rk].toFixed(4));
  });
  // the crucial one: ratios recomputed, not averaged
  const naiveRoas = P.quarters.reduce((a, q) => a + q.metrics.total_roas.value, 0) / P.quarters.length;
  const shownRoas = parseNum(txt(pmHeroes.find(x =>
    txt(x.querySelector(".lbl")).replace(/\s*i$/, "").trim() === "Total raised").querySelector(".of")));
  check("YTD Total ROAS is recomputed from sums, not averaged across quarters",
    close(shownRoas, ytdExp.total_roas, 0.01) && Math.abs(shownRoas - naiveRoas) > 0.005,
    "displayed " + shownRoas + "; recomputed " + ytdExp.total_roas.toFixed(4) +
    "; naive average " + naiveRoas.toFixed(4));

  // ---------- 2 & 3. topline bar charts ----------
  [["pmChQSpend", P.quarters, "quarter"], ["pmChMSpend", P.months, "month"]].forEach(([id, rows, grain]) => {
    const hits = [...doc.querySelectorAll("#" + id + " [data-tip-id]")];
    check(grain + " money chart has one hover target per " + grain, hits.length === rows.length,
      "targets " + hits.length + " vs rows " + rows.length);
    hits.forEach((h, i) => {
      const r = rows[i], M = r.metrics;
      fireHover(h);
      const T = tipPairs();
      if (!T) { check(grain + " money tooltip " + r.label, false, "no tooltip"); return; }
      check(grain + " money tooltip " + r.label.trim() + " spend",
        close(parseNum(T.get("Gross spend")), M.gross_spend.value, 1),
        "tooltip " + T.get("Gross spend") + " vs " + M.gross_spend.cell + " = " + M.gross_spend.value);
      check(grain + " money tooltip " + r.label.trim() + " total raised",
        close(parseNum(T.get("Total raised")), M.total_raised.value, 1),
        "tooltip " + T.get("Total raised") + " vs " + M.total_raised.cell + " = " + M.total_raised.value);
      check(grain + " money tooltip " + r.label.trim() + " total ROAS",
        close(parseNum(T.get("Total ROAS")), M.total_roas.value, 0.01),
        "tooltip " + T.get("Total ROAS") + " vs " + M.total_roas.cell + " = " + M.total_roas.value);
      check(grain + " money tooltip " + r.label.trim() + " cites its row cells",
        T.foot.includes(M.gross_spend.cell) && T.foot.includes(M.total_roas.cell),
        "foot '" + T.foot + "'");
      fireOut(h);
    });
  });
  [["pmChQNtl", P.quarters, "quarter"], ["pmChMNtl", P.months, "month"]].forEach(([id, rows, grain]) => {
    const hits = [...doc.querySelectorAll("#" + id + " [data-tip-id]")];
    check(grain + " NTL chart has one hover target per " + grain, hits.length === rows.length,
      "targets " + hits.length);
    hits.forEach((h, i) => {
      const r = rows[i], M = r.metrics;
      fireHover(h);
      const T = tipPairs();
      if (!T) return;
      check(grain + " NTL tooltip " + r.label.trim(), close(parseNum(T.get("NTL added")), M.ntl.value, 0.5),
        "tooltip " + T.get("NTL added") + " vs " + M.ntl.cell + " = " + M.ntl.value);
      if (M.gross_cpa.value != null) {
        check(grain + " NTL tooltip " + r.label.trim() + " CPA",
          close(parseNum(T.get("Gross CPA")), M.gross_cpa.value, 0.01),
          "tooltip " + T.get("Gross CPA") + " vs " + M.gross_cpa.cell);
        // and CPA must equal spend / NTL
        check(grain + " " + r.label.trim() + " CPA equals spend over NTL",
          close(M.gross_cpa.value, M.gross_spend.value / M.ntl.value, 0.01),
          M.gross_cpa.value + " vs " + (M.gross_spend.value / M.ntl.value));
      }
      fireOut(h);
    });
  });

  // ---------- 4. accumulation ----------
  [["pmAccumQ", "quarterly"], ["pmAccumM", "monthly"]].forEach(([id, grain]) => {
    const charts = [...doc.querySelectorAll("#" + id + " .sm")];
    check(grain + " accumulation renders one chart per accumulated metric",
      charts.length === P.accum_keys.length, "charts " + charts.length + " vs " + P.accum_keys.length);
    charts.forEach((sm, ci) => {
      const key = P.accum_keys[ci];
      const series = P.cumulative[grain][key];
      check(grain + " accumulation chart " + ci + " is titled " + LBL[key],
        txt(sm.querySelector("h3 span")) === LBL[key], "got '" + txt(sm.querySelector("h3 span")) + "'");
      const hits = [...sm.querySelectorAll("[data-tip-id]")];
      check(grain + " " + LBL[key] + " has a hover target per point", hits.length === series.length,
        "targets " + hits.length + " vs points " + series.length);
      let run = 0;
      series.forEach((p, i) => {
        run += p.step || 0;
        const h = hits[i];
        if (!h) return;
        fireHover(h);
        const T = tipPairs();
        if (!T) { check(grain + " " + LBL[key] + " tooltip " + p.label, false, "no tooltip"); return; }
        const tol = KND[key] === "count" ? 0.5 : 1;
        check(grain + " " + LBL[key] + " cumulative at " + p.label.trim(),
          close(parseNum(T.get("Cumulative")), run, tol),
          "tooltip " + T.get("Cumulative") + " vs running sum " + run);
        check(grain + " " + LBL[key] + " step at " + p.label.trim(),
          close(parseNum(T.get(p.label.trim() + " alone")), p.step, tol),
          "tooltip " + T.get(p.label.trim() + " alone") + " vs " + p.step);
        fireOut(h);
      });
      // monthly final point must equal the selected quarter's cell
      if (grain === "monthly") {
        const q = P.quarters.find(x => x.quarter === QSEL);
        if (q) {
          check("monthly cumulative " + LBL[key] + " ends at the Q" + QSEL + " value in " + q.metrics[key].cell,
            close(series[series.length - 1].value, q.metrics[key].value, 0.5),
            "cumulative " + series[series.length - 1].value + " vs " + q.metrics[key].value);
        }
      }
    });
  });

  // ---------- 5a. stacked charts ----------
  const monthTotal = (key, mLabel) => P.detail
    .filter(d => d.month.trim() === mLabel)
    .reduce((a, d) => a + (d.metrics[key].value || 0), 0);
  [["pmChStackSpend", "gross_spend"], ["pmChStackRaised", "total_raised"]].forEach(([id, key]) => {
    const hits = [...doc.querySelectorAll("#" + id + " [data-tip-id]")];
    check("stacked " + LBL[key] + " has a hover target per month", hits.length === P.months.length,
      "targets " + hits.length);
    hits.forEach((h, i) => {
      const mLabel = P.months[i].label.trim();
      fireHover(h);
      const T = tipPairs();
      if (!T) return;
      const exp = monthTotal(key, mLabel);
      check("stacked " + LBL[key] + " " + mLabel + " total matches detail rows",
        close(parseNum(T.get("Total")), exp, 1),
        "tooltip " + T.get("Total") + " vs detail sum " + exp);
      // and that total must equal the monthly topline cell
      check("stacked " + LBL[key] + " " + mLabel + " matches monthly topline " + P.months[i].metrics[key].cell,
        close(exp, P.months[i].metrics[key].value, 0.5),
        "detail sum " + exp + " vs " + P.months[i].metrics[key].value);
      fireOut(h);
    });
  });

  // ---------- 5b. scatter ----------
  const scatterEligible = P.detail.filter(d =>
    d.metrics.gross_cpa.value != null && d.metrics.total_roas.value != null && d.metrics.gross_cpa.value > 0);
  const circles = [...doc.querySelectorAll("#pmChScatter circle")];
  check("scatter plots every placement with a defined CPA",
    circles.length === scatterEligible.length,
    "circles " + circles.length + " vs eligible rows " + scatterEligible.length);
  check("scatter excludes rows with zero NTL and says so",
    P.detail.length === scatterEligible.length ||
      /hidden/i.test(txt(doc.querySelector("#pmChScatter").parentNode.querySelector(".note"))),
    "note: " + txt(doc.querySelector("#pmChScatter").parentNode.querySelector(".note")));
  check("scatter marks the break-even line",
    /break even/i.test(doc.getElementById("pmChScatter").innerHTML), "no break-even reference found");
  // every circle tooltip must match its source row
  let scatterBad = 0;
  circles.forEach(c => {
    fireHover(c);
    const T = tipPairs();
    if (!T) { scatterBad++; return; }
    const row = P.detail.find(d =>
      T.title === d.placement + " — " + d.month.trim() &&
      close(parseNum(T.get("Gross spend")), d.metrics.gross_spend.value, 1));
    if (!row) { scatterBad++; fireOut(c); return; }
    if (!close(parseNum(T.get("Gross CPA")), row.metrics.gross_cpa.value, 0.01)) scatterBad++;
    else if (!close(parseNum(T.get("Total ROAS")), row.metrics.total_roas.value, 0.01)) scatterBad++;
    fireOut(c);
  });
  check("every scatter bubble's tooltip matches its source row", scatterBad === 0,
    scatterBad + " of " + circles.length + " bubbles mismatched");

  // ---------- 5c. leaderboard ----------
  const lbRows = () => [...doc.querySelectorAll("#pmLeader tbody tr")];
  check("leaderboard renders every detail row", lbRows().length === P.detail.length,
    "rows " + lbRows().length + " vs " + P.detail.length);
  // each row's numbers must match its source cells
  let lbBad = [];
  lbRows().forEach(tr => {
    const tds = [...tr.children].map(x => txt(x));
    const placement = tds[0], month = tds[1];
    // Placement names overlap ("True Blue Analytics" vs "True Blue Analytics P2P"),
    // so require an exact match and only fall back to a prefix for truncated text.
    const sameMonth = P.detail.filter(d => d.month.trim() === month);
    let row = sameMonth.find(d => d.placement === placement);
    if (!row && /…$/.test(placement)) {
      const stem = placement.replace(/…$/, "");
      const cands = sameMonth.filter(d => d.placement.startsWith(stem));
      if (cands.length === 1) row = cands[0];
    }
    if (!row) { lbBad.push("no source row for '" + placement + " / " + month + "'"); return; }
    const M = row.metrics;
    const pairs = [[3, "gross_spend", 1], [4, "ntl", 0.5], [5, "gross_cpa", 0.01],
                   [6, "immediate_raised", 1], [7, "lifetime_raised", 1],
                   [8, "total_raised", 1], [9, "total_roas", 0.01]];
    pairs.forEach(([idx, key, tol]) => {
      const shown = parseNum(tds[idx]), exp = M[key].value;
      if (exp == null) {
        if (/\d/.test(tds[idx])) lbBad.push(row.placement + " " + key + " should be blank, got '" + tds[idx] + "'");
      } else if (!close(shown, exp, tol)) {
        lbBad.push(row.placement + "/" + month + " " + key + ": shown " + tds[idx] + " vs " + M[key].cell + " = " + exp);
      }
    });
  });
  check("every leaderboard cell matches its workbook cell", lbBad.length === 0,
    lbBad.length + " mismatches:\n        " + lbBad.slice(0, 6).join("\n        "));

  // totals row: additive summed, ratios recomputed
  const foot = [...doc.querySelectorAll("#pmLeader tfoot td")].map(x => txt(x));
  const allSpend = P.detail.reduce((a, d) => a + (d.metrics.gross_spend.value || 0), 0);
  const allNtl = P.detail.reduce((a, d) => a + (d.metrics.ntl.value || 0), 0);
  const allRaised = P.detail.reduce((a, d) => a + (d.metrics.total_raised.value || 0), 0);
  check("leaderboard totals row sums spend", close(parseNum(foot[3]), allSpend, 1),
    "footer " + foot[3] + " vs " + allSpend);
  check("leaderboard totals row sums NTL", close(parseNum(foot[4]), allNtl, 0.5),
    "footer " + foot[4] + " vs " + allNtl);
  check("leaderboard totals row recomputes CPA from summed spend and NTL",
    close(parseNum(foot[5]), allSpend / allNtl, 0.01),
    "footer " + foot[5] + " vs " + (allSpend / allNtl).toFixed(4));
  check("leaderboard totals row recomputes ROAS from summed raised and spend",
    close(parseNum(foot[9]), allRaised / allSpend, 0.01),
    "footer " + foot[9] + " vs " + (allRaised / allSpend).toFixed(4));
  // the recomputed ratio must differ from the naive per-row average
  const naiveCpa = P.detail.filter(d => d.metrics.gross_cpa.value != null)
    .reduce((a, d, _, arr) => a + d.metrics.gross_cpa.value / arr.length, 0);
  check("leaderboard total CPA is not the average of per-row CPAs",
    Math.abs(parseNum(foot[5]) - naiveCpa) > 0.02,
    "footer " + foot[5] + " vs naive average " + naiveCpa.toFixed(2));

  // sorting: click Gross Spend twice and confirm descending then ascending
  const spendTh = [...doc.querySelectorAll("#pmLeader th")].find(t => t.getAttribute("data-sort") === "gross_spend");
  check("leaderboard headers are sortable", !!spendTh, "no sortable Gross Spend header");
  if (spendTh) {
    spendTh.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
    let vals = lbRows().map(tr => parseNum(txt(tr.children[3])));
    let desc = vals.every((v, i) => i === 0 || v <= vals[i - 1] + 0.01);
    check("clicking Gross Spend sorts descending", desc, "order: " + vals.slice(0, 6).join(", "));
    const th2 = [...doc.querySelectorAll("#pmLeader th")].find(t => t.getAttribute("data-sort") === "gross_spend");
    th2.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
    vals = lbRows().map(tr => parseNum(txt(tr.children[3])));
    check("clicking again sorts ascending", vals.every((v, i) => i === 0 || v >= vals[i - 1] - 0.01),
      "order: " + vals.slice(0, 6).join(", "));
    check("sorting does not change the row count", lbRows().length === P.detail.length,
      "rows " + lbRows().length);
  }

  // ---------- 5d. filters drive all three views ----------
  const chips = [...doc.querySelectorAll("#pmFilters .chip")];
  check("filter chips cover All plus each month and channel",
    chips.length === (P.months.length + 1) + (P.channels.length + 1),
    "chips " + chips.length + " vs expected " + ((P.months.length + 1) + (P.channels.length + 1)));
  const statTxt = () => txt(doc.getElementById("pmFilterStat"));
  check("unfiltered summary counts every detail row",
    statTxt().startsWith(String(P.detail.length) + " of " + P.detail.length),
    "stat: " + statTxt().slice(0, 90));

  const testChannel = P.channels[P.channels.length - 1];
  const chChip = chips.find(c => c.getAttribute("data-filter") === "channel" &&
    c.getAttribute("data-value") === testChannel);
  if (!chChip) {
    check("channel chip for " + testChannel + " exists", false, "missing");
  } else {
    chChip.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
    const expRows = P.detail.filter(d => d.channel === testChannel);
    check("filtering by " + testChannel + " updates the summary count",
      statTxt().startsWith(String(expRows.length) + " of " + P.detail.length),
      "stat: " + statTxt().slice(0, 90));
    check("filtering by " + testChannel + " updates the leaderboard",
      lbRows().length === expRows.length, "rows " + lbRows().length + " vs " + expRows.length);
    const expSpend = expRows.reduce((a, d) => a + (d.metrics.gross_spend.value || 0), 0);
    // grab only the first number after the word "spend", not the whole tail
    const spendShown = (statTxt().match(/spend\s*\$?([\d,]+(?:\.\d+)?)/) || [])[1];
    check("filtered summary spend matches the filtered rows",
      close(parseNum(spendShown), expSpend, 1),
      "stat spend '" + spendShown + "' vs " + expSpend);
    const expScatter = expRows.filter(d => d.metrics.gross_cpa.value != null && d.metrics.gross_cpa.value > 0);
    check("filtering by " + testChannel + " updates the scatter",
      doc.querySelectorAll("#pmChScatter circle").length === expScatter.length,
      "circles " + doc.querySelectorAll("#pmChScatter circle").length + " vs " + expScatter.length);
    check("filtering by " + testChannel + " updates the stacked charts",
      doc.querySelectorAll("#pmChStackSpend [data-tip-id]").length === P.months.length,
      "stack groups " + doc.querySelectorAll("#pmChStackSpend [data-tip-id]").length);
    // restore
    const allChip = chips.find(c => c.getAttribute("data-filter") === "channel" &&
      c.getAttribute("data-value") === "all");
    allChip.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
    check("clearing the channel filter restores every row", lbRows().length === P.detail.length,
      "rows " + lbRows().length);
  }
  // month filter narrows the stacked charts to one group
  const janChip = chips.find(c => c.getAttribute("data-filter") === "month" &&
    c.getAttribute("data-value") === P.months[0].label.trim());
  if (janChip) {
    janChip.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
    check("filtering to one month leaves one stack group",
      doc.querySelectorAll("#pmChStackSpend [data-tip-id]").length === 1,
      "groups " + doc.querySelectorAll("#pmChStackSpend [data-tip-id]").length);
    const expRows = P.detail.filter(d => d.month.trim() === P.months[0].label.trim());
    check("filtering to one month narrows the leaderboard", lbRows().length === expRows.length,
      "rows " + lbRows().length + " vs " + expRows.length);
    chips.find(c => c.getAttribute("data-filter") === "month" && c.getAttribute("data-value") === "all")
      .dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
  }

  // ---------- 6. info bubbles on the paid media tab ----------
  ["pm-hero", "pm-quarter", "pm-month", "pm-accum", "pm-detail"].forEach(id => {
    const sec = doc.getElementById(id);
    check("info bubble in section " + id, !!sec && !!sec.querySelector("button.info"), "missing");
  });
  check("info bubble on every year-to-date hero card",
    pmHeroes.every(h => h.querySelector("button.info")), "one or more lack a bubble");
  check("info bubble on every accumulation chart",
    [...doc.querySelectorAll("#pmAccumQ .sm, #pmAccumM .sm")].every(s => s.querySelector("button.info")),
    "one or more lack a bubble");
  ["pmChStackSpend", "pmChStackRaised", "pmChScatter", "pmLeader"].forEach(id => {
    const card = doc.getElementById(id).closest(".card");
    check("info bubble on the " + id + " view", !!card.querySelector(".sechead.sub button.info"), "missing");
  });
  // the objective-column caveat must be stated somewhere
  const pmBubbles = [...doc.querySelectorAll("#p-paid button.info")];
  let sawObjective = false, sawDedup = false;
  pmBubbles.forEach(b => {
    b.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
    const pop = doc.querySelector(".pop");
    if (pop) {
      const body = txt(pop);
      if (/Objective/i.test(body) && /Goal/.test(body)) sawObjective = true;
      if (/double counted|not counted twice|de-dup/i.test(body)) sawDedup = true;
    }
    b.dispatchEvent(new dom.window.MouseEvent("click", {bubbles: true}));
  });
  check("an info bubble explains that column E is an objective, not a numeric goal", sawObjective,
    "no bubble mentioned the Goal/Objective header discrepancy");
  check("an info bubble explains the Total Raised de-duplication", sawDedup,
    "no bubble explained why Total Raised is not Immediate + Lifetime");

  // ---------- 7. cross-tab period guard ----------
  check("the two tabs report different quarters and the page says so",
    P.meta.quarter.value === D.meta.quarter.value ||
      /different quarter/i.test(scope),
    "tabs cover Q" + P.meta.quarter.value + " and Q" + D.meta.quarter.value +
    " but the scope banner does not flag it");
} else {
  console.log("note: " + pmPath + " not found — paid media checks skipped");
}

// ---------- report ---------------------------------------------------------
const fails = results.filter(r => !r.ok);
console.log(results.length + " render checks run, " + (results.length - fails.length) + " passed, " + fails.length + " failed");
fails.forEach(f => console.log("  FAIL  " + f.name + "\n        " + f.detail));
if (!fails.length) console.log("rendered page matches the workbook on every displayed figure");
process.exit(fails.length ? 1 : 0);
})().catch(e => { console.error(e.stack || String(e)); process.exit(1); });
