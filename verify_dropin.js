/*
 * verify_dropin.js — exercises the drop-a-workbook path through the page's own
 * code, in jsdom, using the real .xlsx plus synthetic variants.
 *
 * What this proves:
 *   1. The parser inlined in the page reads the real workbook and reproduces
 *      the baked-in snapshot exactly (so the browser path and the build path
 *      cannot silently diverge).
 *   2. Added rows and added channels are picked up, not truncated.
 *   3. A structurally broken workbook is refused with the previous data intact.
 *   4. A workbook missing one tab hides that tab rather than showing stale
 *      numbers from the previous file.
 *   5. After a load, the rendered figures come from the new file.
 *
 * Usage: node verify_dropin.js [digital-report.html] [workbook.xlsx]
 */
const fs = require("fs");
const path = require("path");
const zlib = require("zlib");
const JSDOMMod = require(path.join(process.env.JSDOM_PATH || "/tmp/vfy/node_modules", "jsdom"));
const {JSDOM, VirtualConsole} = JSDOMMod;

const htmlPath = process.argv[2] || "digital-report.html";
const wbPath = process.argv[3] || "EXTERNAL_ CAB Digital Report.xlsx";

const results = [];
const check = (ok, name, detail) => results.push({ok, name, detail});
const close = (a, b, tol) => a != null && b != null && Math.abs(a - b) <= (tol == null ? 0.02 : tol);

/* ---------- a jsdom page with the inflate hook wired to node zlib ---------- */
function makePage() {
  const errors = [];
  const vc = new VirtualConsole();
  vc.on("jsdomError", e => errors.push(e.message));
  const dom = new JSDOM(fs.readFileSync(htmlPath, "utf8"),
    {runScripts: "dangerously", pretendToBeVisual: true, virtualConsole: vc});
  const w = dom.window;
  // the page's own reader, with the one platform primitive jsdom lacks
  w.XlsxReader.setInflate(bytes => new Uint8Array(zlib.inflateRawSync(Buffer.from(bytes))));
  return {dom, w, doc: w.document, errors};
}
const txt = el => (el ? el.textContent.replace(/\s+/g, " ").trim() : "");
const parseNum = s => {
  if (s == null) return null;
  const t = String(s).replace(/[^0-9.\-]/g, "");
  if (!t || t === "-" || t === ".") return null;
  const n = parseFloat(t);
  return isNaN(n) ? null : n;
};

/* a stand-in for the browser File object: only these three members are used */
const fakeFile = (name, buf) => ({
  name, size: buf.length,
  arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
});

/* ---------- minimal ZIP writer, to build synthetic workbooks ---------- */
function crc32(buf) {
  let c, table = crc32.t;
  if (!table) {
    table = crc32.t = [];
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function zipWrite(files) {
  // files: [{name, data:Buffer}] — stored with deflate, matching real xlsx
  const locals = [], centrals = [];
  let offset = 0;
  files.forEach(f => {
    const comp = zlib.deflateRawSync(f.data);
    const nameBuf = Buffer.from(f.name, "utf8");
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0, 6);
    lh.writeUInt16LE(8, 8); lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0, 12);
    lh.writeUInt32LE(crc32(f.data), 14);
    lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(f.data.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26); lh.writeUInt16LE(0, 28);
    locals.push(lh, nameBuf, comp);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0, 8); ch.writeUInt16LE(8, 10);
    ch.writeUInt32LE(crc32(f.data), 16);
    ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(f.data.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28);
    ch.writeUInt32LE(offset, 42);
    centrals.push(ch, nameBuf);
    offset += lh.length + nameBuf.length + comp.length;
  });
  const cd = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([Buffer.concat(locals), cd, eocd]);
}

/* Build a workbook from a simple {sheetName: {A1: value}} spec. */
function synthWorkbook(spec) {
  const names = Object.keys(spec);
  const sheetsXml = names.map((name, i) => {
    const cells = spec[name];
    const byRow = {};
    Object.keys(cells).forEach(ref => {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) return;
      (byRow[m[2]] = byRow[m[2]] || []).push({ref, col: m[1], v: cells[ref]});
    });
    const rows = Object.keys(byRow).sort((a, b) => a - b).map(r => {
      const cs = byRow[r].sort((a, b) => a.col.length - b.col.length || a.col.localeCompare(b.col))
        .map(c => {
          if (typeof c.v === "number") return '<c r="' + c.ref + '"><v>' + c.v + "</v></c>";
          if (typeof c.v === "boolean") return '<c r="' + c.ref + '" t="b"><v>' + (c.v ? 1 : 0) + "</v></c>";
          const s = String(c.v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          return '<c r="' + c.ref + '" t="inlineStr"><is><t>' + s + "</t></is></c>";
        }).join("");
      return '<row r="' + r + '">' + cs + "</row>";
    }).join("");
    return {
      path: "xl/worksheets/sheet" + (i + 1) + ".xml",
      xml: '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
           "<sheetData>" + rows + "</sheetData></worksheet>"
    };
  });
  const wb = '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
    names.map((n, i) => '<sheet name="' + n.replace(/&/g, "&amp;").replace(/"/g, "&quot;") +
      '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") +
    "</sheets></workbook>";
  const rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    names.map((n, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join("") +
    "</Relationships>";
  const ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
  return zipWrite([
    {name: "[Content_Types].xml", data: Buffer.from(ct)},
    {name: "xl/workbook.xml", data: Buffer.from(wb)},
    {name: "xl/_rels/workbook.xml.rels", data: Buffer.from(rels)}
  ].concat(sheetsXml.map(s => ({name: s.path, data: Buffer.from(s.xml)}))));
}

/* ---------- spec builders for synthetic tabs ---------- */
function digitalSpec(channels, months) {
  // months: [{label, goalCol, actualCol, pctCol}] laid out from G onward
  const s = {
    B4: "Digital Fundraising Toplines", B6: "Client:", C6: "Test Client",
    B7: "Email Lead:", C7: "Tester", B9: "Year:", C9: 2027, B10: "Quarter:", C10: 4,
    B12: "Quarterly Toplines", B13: "Channel", C13: "Goal", D13: "Actual", E13: "% to Goal"
  };
  months.forEach(m => {
    s[m.goalCol + 12] = m.label;
    s[m.goalCol + 13] = "Goal";
    s[m.actualCol + 13] = "Actual";
    s[m.pctCol + 13] = "% to Goal";
  });
  let r = 14;
  channels.forEach(ch => {
    s["B" + r] = ch.name;
    let qg = 0, qa = 0;
    months.forEach((m, i) => {
      const g = ch.goals[i], a = ch.actuals[i];
      s[m.goalCol + r] = g;
      if (a != null) s[m.actualCol + r] = a;
      s[m.pctCol + r] = a == null ? 0 : (g ? a / g : 0);
      qg += g; qa += (a || 0);
    });
    s["C" + r] = qg; s["D" + r] = qa; s["E" + r] = qg ? qa / qg : 0;
    r++;
  });
  const overallRow = r + 1;
  s["B" + overallRow] = "Overall";
  const tr = overallRow + 1;
  s["B" + tr] = "Total Raised";
  let tqg = 0, tqa = 0;
  months.forEach((m, i) => {
    const g = channels.reduce((a, c) => a + c.goals[i], 0);
    const av = channels.reduce((a, c) => a + (c.actuals[i] || 0), 0);
    const anyActual = channels.some(c => c.actuals[i] != null);
    s[m.goalCol + tr] = g;
    if (anyActual) s[m.actualCol + tr] = av;
    s[m.pctCol + tr] = g ? (anyActual ? av / g : 0) : 0;
    tqg += g; tqa += av;
  });
  s["C" + tr] = tqg; s["D" + tr] = tqa; s["E" + tr] = tqg ? tqa / tqg : 0;
  s["B" + (tr + 1)] = "Total Donations";
  s["D" + (tr + 1)] = 100;
  months.forEach((m, i) => {
    if (channels.some(c => c.actuals[i] != null)) s[m.actualCol + (tr + 1)] = 50;
  });
  return s;
}

function paidSpec(detailRows, months) {
  const METRICS = ["Gross Spend", "NTL", "Gross CPA", "Immediate Raised", "Immediate ROAS",
                   "Lifetime Raised", "Lifetime ROAS", "Total Raised", "Total ROAS"];
  const COLS = ["F", "G", "H", "I", "J", "K", "L", "M", "N"];
  const s = {
    B2: "Paid Media Report", B4: "Year:", C4: 2027, B5: "Quarter:", C5: 4,
    B7: "Client:", C7: "Test Client",
    F4: "NTL:", G4: "Count of new-to-list names",
    F5: "CPA:", G5: "Cost per name added to the email list",
    F6: "ROAS:", G6: "Return on ad spend",
    B11: "Quarterly Toplines", E11: "Quarter",
    B17: "Monthly Toplines", E17: "Month",
    B22: "Month", C22: "Channel", D22: "Placement", E22: "Goal"
  };
  METRICS.forEach((m, i) => { s[COLS[i] + 11] = m; s[COLS[i] + 17] = m; s[COLS[i] + 22] = m; });

  const mk = (spend, ntl, imm, life, tot) => [
    spend, ntl, ntl ? spend / ntl : null, imm, spend ? imm / spend : 0,
    life, spend ? life / spend : 0, tot, spend ? tot / spend : 0];

  // detail rows
  let r = 23;
  const perMonth = {};
  detailRows.forEach(d => {
    s["B" + r] = d.month; s["C" + r] = d.channel; s["D" + r] = d.placement; s["E" + r] = d.objective;
    const vals = mk(d.spend, d.ntl, d.imm, d.life, d.tot);
    vals.forEach((v, i) => { if (v != null) s[COLS[i] + r] = v; });
    const p = perMonth[d.month] = perMonth[d.month] ||
      {spend: 0, ntl: 0, imm: 0, life: 0, tot: 0};
    p.spend += d.spend; p.ntl += d.ntl; p.imm += d.imm; p.life += d.life; p.tot += d.tot;
    r++;
  });
  // monthly rows
  let mr = 18, qs = {spend: 0, ntl: 0, imm: 0, life: 0, tot: 0};
  months.forEach(m => {
    const p = perMonth[m] || {spend: 0, ntl: 0, imm: 0, life: 0, tot: 0};
    s["E" + mr] = m;
    mk(p.spend, p.ntl, p.imm, p.life, p.tot).forEach((v, i) => { if (v != null) s[COLS[i] + mr] = v; });
    qs.spend += p.spend; qs.ntl += p.ntl; qs.imm += p.imm; qs.life += p.life; qs.tot += p.tot;
    mr++;
  });
  // one quarterly row, matching the tab's quarter
  s["E12"] = 4;
  mk(qs.spend, qs.ntl, qs.imm, qs.life, qs.tot).forEach((v, i) => { if (v != null) s[COLS[i] + 12] = v; });
  return s;
}

/* ===================================================================== */
(async function run() {
  const realBuf = fs.readFileSync(wbPath);

  /* ---- 0. the page shows nothing until a workbook is loaded ---- */
  {
    const {w, doc, errors} = makePage();
    check(errors.length === 0, "page loads with no script errors", errors.slice(0, 2).join(" | "));
    const st = () => w.__dashboard.state();
    check(st().DATA === null && st().PM === null, "no data is loaded on open",
      "DATA/PM were populated at load time");
    check(!doc.getElementById("emptyState").hidden, "the empty state is shown", "empty state hidden");
    check(doc.getElementById("pgHeader").hidden && doc.querySelector(".tabs").hidden &&
          doc.getElementById("pgFoot").hidden,
      "header, tabs and footer are hidden until a workbook is loaded",
      "header hidden=" + doc.getElementById("pgHeader").hidden +
      ", tabs hidden=" + doc.querySelector(".tabs").hidden);
    check([...doc.querySelectorAll(".panel")].every(p => p.hidden),
      "no tab panel is visible on open",
      [...doc.querySelectorAll(".panel")].filter(p => !p.hidden).length + " visible");
    check(doc.querySelectorAll("#heroes .hero").length === 0,
      "no figures are rendered on open",
      doc.querySelectorAll("#heroes .hero").length + " hero cards");
    check(txt(doc.getElementById("srcNow")).indexOf("No workbook") === 0,
      "the source bar says nothing is loaded", txt(doc.getElementById("srcNow")));
  }

  /* ---- 1. the page's parser matches the node build byte for byte ---- */
  {
    const {w} = makePage();
    // the snapshot JSONs are produced by build-data.js using these same libs;
    // this proves the browser path and the build path agree
    const expected = {
      digital: JSON.parse(fs.readFileSync("digital-report-data.json", "utf8")),
      paid: JSON.parse(fs.readFileSync("paid-media-data.json", "utf8"))
    };
    const baked = expected;
    const book = await w.XlsxReader.read(new Uint8Array(realBuf));
    const out = w.ExtractReport.extractAll(book, {
      workbookName: path.basename(wbPath),
      extractedAt: expected.digital.meta.extracted_at
    });
    check(out.report.errors === 0, "real workbook parses with no errors",
      out.report.problems.filter(p => p.level === "error").map(p => p.text).join(" | "));

    const diffs = [];
    (function walk(a, b, p) {
      if (p.endsWith(".extracted_at")) return;
      if (a && b && typeof a === "object" && typeof b === "object") {
        const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
        for (const k of keys) walk(a[k], b[k], p + "." + k);
        return;
      }
      if (typeof a === "number" && typeof b === "number") {
        if (Math.abs(a - b) > 1e-9 * Math.max(1, Math.abs(a))) diffs.push(p + ": " + a + " vs " + b);
      } else if (a !== b) {
        diffs.push(p + ": " + JSON.stringify(a) + " vs " + JSON.stringify(b));
      }
    })(baked.digital, out.digital, "digital");
    (function walk(a, b, p) {
      if (p.endsWith(".extracted_at")) return;
      if (a && b && typeof a === "object" && typeof b === "object") {
        const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
        for (const k of keys) walk(a[k], b[k], p + "." + k);
        return;
      }
      if (typeof a === "number" && typeof b === "number") {
        if (Math.abs(a - b) > 1e-9 * Math.max(1, Math.abs(a))) diffs.push(p + ": " + a + " vs " + b);
      } else if (a !== b) {
        diffs.push(p + ": " + JSON.stringify(a) + " vs " + JSON.stringify(b));
      }
    })(baked.paid, out.paid, "paid");
    check(diffs.length === 0, "in-browser parse matches the node build field for field",
      diffs.length + " difference(s): " + diffs.slice(0, 5).join(" | "));
  }

  /* ---- 2. loading the real workbook renders it, clearing puts it away ---- */
  {
    const {w, doc} = makePage();
    const expected = JSON.parse(fs.readFileSync("digital-report-data.json", "utf8"));
    await w.__dashboard.load(fakeFile(path.basename(wbPath), realBuf));
    const shown = parseNum(txt(doc.querySelector("#heroes .hero .big")));
    check(close(shown, expected.overall.total_raised.periods.quarter.actual, 1),
      "the headline figure comes from the loaded workbook",
      "rendered " + shown + " vs " + expected.overall.total_raised.periods.quarter.actual);
    check(txt(doc.getElementById("srcNow")).includes(path.basename(wbPath)),
      "source bar names the loaded file", txt(doc.getElementById("srcNow")).slice(0, 90));
    check(doc.getElementById("emptyState").hidden && !doc.getElementById("pgHeader").hidden,
      "the empty state gives way to the report", "empty state still visible");
    const rep = doc.getElementById("report");
    check(!rep.hidden, "a validation report is shown after loading", "report hidden");
    check(/reconcil/i.test(txt(rep)), "report says the workbook reconciled", txt(rep).slice(0, 120));
    check(doc.querySelectorAll("#pmLeader tbody tr").length === w.__dashboard.state().PM.detail.length,
      "leaderboard re-renders with the loaded row count",
      doc.querySelectorAll("#pmLeader tbody tr").length + " vs " + w.__dashboard.state().PM.detail.length);
    check(doc.querySelectorAll("button.info").length > 30,
      "info bubbles are rebuilt, not duplicated or lost",
      doc.querySelectorAll("button.info").length + " bubbles");
    // clearing returns to the empty state, leaving nothing on screen
    doc.getElementById("btnClear").dispatchEvent(new w.MouseEvent("click", {bubbles: true}));
    check(!doc.getElementById("emptyState").hidden &&
          doc.querySelectorAll("#heroes .hero").length === 0 &&
          w.__dashboard.state().DATA === null,
      "clearing puts the page back to the empty state",
      "empty hidden=" + doc.getElementById("emptyState").hidden +
      ", heroes=" + doc.querySelectorAll("#heroes .hero").length);
  }

  /* ---- 3. more rows and more channels than today are picked up ---- */
  {
    const {w, doc} = makePage();
    const months = [
      {label: "October", goalCol: "G", actualCol: "H", pctCol: "I"},
      {label: "November", goalCol: "J", actualCol: "K", pctCol: "L"},
      {label: "December", goalCol: "M", actualCol: "N", pctCol: "O"}
    ];
    // eleven channels where today there are eight, including a brand new one
    const channels = [];
    for (let i = 0; i < 11; i++) {
      channels.push({
        name: i === 10 ? "Connected TV" : "Channel " + (i + 1),
        goals: [1000 + i, 2000 + i, 3000 + i],
        actuals: [500 + i, 900 + i, null]        // last month unreported
      });
    }
    // thirty detail rows where today there are eighteen
    const detail = [];
    ["October", "November", "December"].forEach(m => {
      for (let i = 0; i < 10; i++) {
        detail.push({
          month: m, channel: ["Facebook", "SMS P2P", "Connected TV"][i % 3],
          placement: "Placement " + (i + 1), objective: "Donations",
          spend: 1000 + i * 10, ntl: 100 + i, imm: 1200 + i * 10, life: 300 + i, tot: 1400 + i * 10
        });
      }
    });
    const buf = synthWorkbook({
      "Digital Report": digitalSpec(channels, months),
      "Paid Media Report": paidSpec(detail, ["October", "November", "December"])
    });
    await w.__dashboard.load(fakeFile("grown.xlsx", buf));

    check(w.__dashboard.state().DATA && w.__dashboard.state().DATA.channels.length === 11, "all eleven channel rows were read",
      w.__dashboard.state().DATA ? w.__dashboard.state().DATA.channels.length + " channels" : "digital payload is null");
    check(doc.querySelectorAll("#chChannel .cbar").length === 11,
      "channel chart renders all eleven rows",
      doc.querySelectorAll("#chChannel .cbar").length + " bars");
    check(txt(doc.getElementById("chChannel")).includes("Connected TV"),
      "the newly added channel appears", "not found");
    check(w.__dashboard.state().PM && w.__dashboard.state().PM.detail.length === 30, "all thirty detail rows were read",
      w.__dashboard.state().PM ? w.__dashboard.state().PM.detail.length + " rows" : "paid payload is null");
    check(doc.querySelectorAll("#pmLeader tbody tr").length === 30,
      "leaderboard shows all thirty rows",
      doc.querySelectorAll("#pmLeader tbody tr").length + " rows");
    check(w.__dashboard.state().PM.channels.length === 3 && w.__dashboard.state().PM.channels.indexOf("Connected TV") >= 0,
      "a new paid channel is discovered and coloured", w.__dashboard.state().PM.channels.join(", "));
    check(w.__dashboard.state().DATA.month_keys.join(",") === "october,november,december",
      "the new quarter's months are used, not last quarter's",
      w.__dashboard.state().DATA.month_keys.join(","));
    check((w.__dashboard.state().DATA.pending_months || []).join(",") === "december",
      "the unreported month is treated as pending",
      "pending: " + (w.__dashboard.state().DATA.pending_months || []).join(","));
    check(!/September|July/.test(txt(doc.getElementById("p-overview"))),
      "no month names from the previous workbook survive the swap", "stale month names found");
    // spot-check a rendered figure against the synthetic input
    const expTotal = channels.reduce((a, c) => a + c.actuals[0] + c.actuals[1], 0);
    const heroBig = parseNum(txt(doc.querySelector("#heroes .hero .big")));
    check(close(heroBig, expTotal, 1), "headline total matches the synthetic workbook",
      "rendered " + heroBig + " vs expected " + expTotal);
  }

  /* ---- 4. a structurally broken workbook is refused ---- */
  {
    const {w, doc} = makePage();
    // load a good workbook first, so we can prove a failed load leaves it alone
    await w.__dashboard.load(fakeFile(path.basename(wbPath), realBuf));
    const goodTotal = txt(doc.querySelector("#heroes .hero .big"));
    check(goodTotal.length > 1, "a good workbook is loaded before the broken one", "nothing rendered");
    const buf = synthWorkbook({
      "Digital Report": {B1: "nothing useful here", B2: "no headings at all"},
      "Paid Media Report": {B1: "also empty"}
    });
    await w.__dashboard.load(fakeFile("broken.xlsx", buf));
    const rep = doc.getElementById("report");
    check(!rep.hidden && /not loaded/i.test(txt(rep)), "a broken workbook is refused",
      txt(rep).slice(0, 140));
    check(txt(doc.querySelector("#heroes .hero .big")) === goodTotal,
      "the previous data is left intact when a load is refused",
      "was " + goodTotal + ", now " + txt(doc.querySelector("#heroes .hero .big")));
    check(w.__dashboard.state().LOADED && w.__dashboard.state().LOADED.name === path.basename(wbPath),
      "a refused load does not become the active source",
      "active source is " + JSON.stringify(w.__dashboard.state().LOADED &&
        w.__dashboard.state().LOADED.name));
  }

  /* ---- 5. figures that no longer reconcile are reported, not rendered ---- */
  {
    const {w, doc} = makePage();
    const months = [{label: "October", goalCol: "G", actualCol: "H", pctCol: "I"}];
    const spec = digitalSpec([{name: "Email", goals: [1000], actuals: [400]}], months);
    // corrupt the Overall row so channels no longer sum to it
    const trRow = Object.keys(spec).filter(k => spec[k] === "Total Raised")[0].replace("B", "");
    spec["D" + trRow] = 999999;
    // pair it with a Paid Media tab that IS sound, so we can see per-tab handling
    const buf = synthWorkbook({
      "Digital Report": spec,
      "Paid Media Report": paidSpec([{
        month: "October", channel: "Facebook", placement: "Meta", objective: "Donations",
        spend: 100, ntl: 10, imm: 120, life: 30, tot: 140
      }], ["October"])
    });
    await w.__dashboard.load(fakeFile("mismatched.xlsx", buf));
    const rep = txt(doc.getElementById("report"));
    check(/do not|does not|says/i.test(rep) && /Overall/i.test(rep),
      "a workbook whose channels do not sum to its Overall row is reported",
      rep.slice(0, 200));
    check(w.__dashboard.state().DATA === null && doc.getElementById("t-overview").hidden,
      "the unreconciled tab is hidden rather than shown with bad numbers",
      "overview hidden=" + doc.getElementById("t-overview").hidden +
      ", DATA " + (w.__dashboard.state().DATA ? "set" : "null"));
    check(w.__dashboard.state().PM !== null && !doc.getElementById("t-paid").hidden,
      "the sound tab in the same workbook is still shown",
      "paid hidden=" + doc.getElementById("t-paid").hidden);
    check(/errors above are hidden/i.test(txt(doc.getElementById("report"))),
      "the report explains that failing sections were hidden",
      txt(doc.getElementById("report")).slice(0, 200));
  }

  /* ---- 6. a workbook missing a tab hides that tab ---- */
  {
    const {w, doc} = makePage();
    // load the full workbook first, so the Paid Media tab is populated and we
    // can prove the next load discards it rather than leaving it on screen
    await w.__dashboard.load(fakeFile(path.basename(wbPath), realBuf));
    check(!doc.getElementById("t-paid").hidden, "the paid tab is present before the swap",
      "paid tab already hidden");
    const months = [
      {label: "October", goalCol: "G", actualCol: "H", pctCol: "I"},
      {label: "November", goalCol: "J", actualCol: "K", pctCol: "L"}
    ];
    const buf = synthWorkbook({
      "Digital Report": digitalSpec(
        [{name: "Email", goals: [1000, 2000], actuals: [400, 900]}], months)
    });
    await w.__dashboard.load(fakeFile("digital-only.xlsx", buf));
    check(doc.getElementById("t-paid").hidden, "the missing Paid Media tab is hidden",
      "paid tab still shown");
    check(w.__dashboard.state().PM === null, "stale paid data is discarded, not left on screen",
      "PM still populated");
    check(!doc.getElementById("t-overview").hidden, "the tab that did parse stays available",
      "overview hidden too");
    check(/no .paid media report. tab/i.test(txt(doc.getElementById("report"))),
      "the report explains which tab was missing",
      txt(doc.getElementById("report")).slice(0, 160));
  }

  /* ---- 7. a non-xlsx file is rejected with a useful message ---- */
  {
    const {w, doc} = makePage();
    await w.__dashboard.load(fakeFile("notes.csv", Buffer.from("a,b,c\n1,2,3")));
    check(/not an \.xlsx/i.test(txt(doc.getElementById("report"))),
      "a .csv is rejected by extension with a clear message",
      txt(doc.getElementById("report")).slice(0, 120));
    await w.__dashboard.load(fakeFile("fake.xlsx", Buffer.from("this is not a zip at all")));
    check(/does not look like|ZIP/i.test(txt(doc.getElementById("report"))),
      "a mislabelled .xlsx is rejected on its signature",
      txt(doc.getElementById("report")).slice(0, 140));
  }

  /* ---- report ---- */
  const fails = results.filter(r => !r.ok);
  console.log(results.length + " drop-in checks run, " + (results.length - fails.length) +
    " passed, " + fails.length + " failed");
  fails.forEach(f => console.log("  FAIL  " + f.name + "\n        " + f.detail));
  if (!fails.length) console.log("the drop-in path parses, reconciles, re-renders and fails safe");
  process.exit(fails.length ? 1 : 0);
})().catch(e => { console.error(e.stack || String(e)); process.exit(1); });
