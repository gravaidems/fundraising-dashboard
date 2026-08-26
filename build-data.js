#!/usr/bin/env node
/*
 * build-data.js — produce the baked-in snapshot JSONs from a workbook, using
 * the SAME reader and extractor the page runs in the browser. There is only one
 * implementation of the extraction logic; this is just the node entry point.
 *
 * Usage:
 *   node build-data.js ["EXTERNAL_ CAB Digital Report.xlsx"]
 *        [-d digital-report-data.json] [-p paid-media-data.json] [-s p2p-data.json]
 *        [-e email-stats-data.json] [-hd high-dollar-data.json]
 */
const fs = require("fs");
const path = require("path");

const XlsxReader = require("./lib/xlsx-reader.js");
global.XlsxReader = XlsxReader;
const ExtractReport = require("./lib/extract-report.js");

function arg(flag, dflt) {
  const i = process.argv.indexOf(flag);
  return i > 0 && process.argv[i + 1] ? process.argv[i + 1] : dflt;
}

(async function main() {
  const flags = ["-d", "-p", "-s", "-e", "-hd"];
  const positional = process.argv.slice(2).filter(a => !a.startsWith("-") &&
    flags.indexOf(process.argv[process.argv.indexOf(a) - 1]) < 0);
  const wbPath = positional[0] || "EXTERNAL_ CAB Digital Report.xlsx";
  const outD = arg("-d", "digital-report-data.json");
  const outP = arg("-p", "paid-media-data.json");
  const outS = arg("-s", "p2p-data.json");
  const outE = arg("-e", "email-stats-data.json");
  const outHD = arg("-hd", "high-dollar-data.json");

  if (!fs.existsSync(wbPath)) {
    console.error("workbook not found: " + wbPath);
    process.exit(2);
  }
  const buf = fs.readFileSync(wbPath);
  const book = await XlsxReader.read(new Uint8Array(buf));
  const out = ExtractReport.extractAll(book, {
    workbookName: path.basename(wbPath),
    extractedAt: new Date().toISOString().replace(/\.\d+Z$/, "")
  });

  const rep = out.report;
  console.log("tabs: " + rep.sheetNames.length + " — " + rep.sheetNames.join(", "));
  rep.findings.forEach(f => console.log("  · " + f));
  rep.problems.forEach(p => console.log("  " + (p.level === "error" ? "ERROR" : "warn ") + ": " + p.text));

  if (rep.errors) {
    console.error("\n" + rep.errors + " error(s) — refusing to write snapshot files.");
    process.exit(1);
  }
  if (out.digital) {
    fs.writeFileSync(outD, JSON.stringify(out.digital, null, 2));
    console.log("wrote " + outD + " (" + out.digital.channels.length + " channels, " +
      out.digital.month_keys.length + " months)");
  }
  if (out.paid) {
    fs.writeFileSync(outP, JSON.stringify(out.paid, null, 2));
    console.log("wrote " + outP + " (" + out.paid.quarters.length + " quarters, " +
      out.paid.months.length + " months, " + out.paid.detail.length + " detail rows)");
  }
  if (out.p2p) {
    fs.writeFileSync(outS, JSON.stringify(out.p2p, null, 2));
    console.log("wrote " + outS + " (" + out.p2p.months.length + " months, " +
      out.p2p.detail.length + " detail rows, " + out.p2p.audiences.length + " audiences, " +
      out.p2p.topics.length + " topics)");
  }
  if (out.email) {
    fs.writeFileSync(outE, JSON.stringify(out.email, null, 2));
    console.log("wrote " + outE + " (" + out.email.months.length + " months, " +
      out.email.detail.length + " detail rows, " + out.email.audiences.length + " audiences, " +
      out.email.topics.length + " topics)");
  }
  if (out.highDollar) {
    fs.writeFileSync(outHD, JSON.stringify(out.highDollar, null, 2));
    console.log("wrote " + outHD + " (" + out.highDollar.months.length + " months, " +
      out.highDollar.detail.length + " detail rows, " + out.highDollar.categories.length + " categories)");
  }
  if (rep.warnings) console.log(rep.warnings + " warning(s) — see above.");
})().catch(e => { console.error(e.stack || String(e)); process.exit(1); });
