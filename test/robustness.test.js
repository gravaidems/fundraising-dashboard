/*
 * Robustness: what happens when staff edit the workbook.
 *
 * Every variant is built in a scratch copy — the original file is never written to.
 */
var fs = require('fs');
var path = require('path');
var os = require('os');
var XLSX = require('../vendor/xlsx.full.min.js');
var CABParse = require('../src/parse.js');

var ROOT = path.join(__dirname, '..');
var SRC = process.env.CAB_WORKBOOK || path.join(ROOT, 'EXTERNAL_ CAB Digital Report.xlsx');
var OUT = fs.mkdtempSync(path.join(os.tmpdir(), 'cab-robust-'));

var pass = 0, fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (detail ? '\n         ' + detail : '')); }
}

function load() { return XLSX.read(fs.readFileSync(SRC), { type: 'buffer' }); }
function parseWb(wb) {
  var buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  return CABParse.createParser(XLSX).parseWorkbook(buf);
}
function tabReport(data, name) {
  return data.parseReport.tabs.filter(function (t) { return t.tab === name; })[0];
}

console.log('\n== Appending rows ==');
(function () {
  var wb = load();
  var ws = wb.Sheets['Email Statistics'];
  // Clone 50 existing rows onto the end of the data, at rows 241-290.
  for (var i = 0; i < 50; i++) {
    var from = 6 + (i % 100), to = 241 + i;
    for (var c = 0; c <= 23; c++) {
      var L = XLSX.utils.encode_col(c);
      var src = ws[L + from];
      if (src) ws[L + to] = { t: src.t, v: src.v, w: src.w };
    }
    // Push the dates forward so the appended rows are genuinely new sends.
    ws['E' + to] = { t: 'n', v: 46265 + i };
  }
  ws['!ref'] = 'A1:X1181';
  var data = parseWb(wb);
  ok('50 appended rows are picked up with no code change', data.emailSends.rows.length === 285,
    'got ' + data.emailSends.rows.length);
  ok('resolved range grew to match', /A6:X290/.test(data.emailSends.dataRange), data.emailSends.dataRange);
})();

console.log('\n== Renamed column (typo) ==');
(function () {
  var wb = load();
  wb.Sheets['Email Statistics']['M5'] = { t: 's', v: 'Recipiants' };
  var data = parseWb(wb);
  var em = data.emailSends;
  ok('workbook still parses with the typo', em.rows.length === 235, 'got ' + em.rows.length);
  ok('Recipients still resolves to column M', em.columnLetters.recipients === 'M', em.columnLetters.recipients);
  var fuzzy = em.warnings.filter(function (w) { return w.kind === 'fuzzy' && w.column === 'recipients'; });
  ok('a fuzzy-match warning is raised', fuzzy.length === 1, JSON.stringify(em.warnings));
  ok('the warning names both the found and expected header',
    fuzzy.length && /Recipiants/.test(fuzzy[0].text) && /Recipients/.test(fuzzy[0].text), fuzzy[0] && fuzzy[0].text);
  ok('the warning reaches the top-level report',
    data.parseReport.warnings.some(function (w) { return /Recipiants/.test(w.text); }));
  fs.writeFileSync(path.join(OUT, 'typo.xlsx'), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
})();

console.log('\n== Deleted optional column ==');
(function () {
  var wb = load();
  var ws = wb.Sheets['Email Statistics'];
  delete ws['N5'];                                  // the Open Rate header
  for (var r = 6; r <= 240; r++) delete ws['N' + r];
  var data = parseWb(wb);
  var em = data.emailSends;
  ok('workbook still parses without Open Rate', em.rows.length === 235, 'got ' + em.rows.length);
  ok('Open Rate is reported as missing',
    (tabReport(data, 'Email Statistics').missingOptional || []).indexOf('Open Rate') !== -1,
    JSON.stringify(tabReport(data, 'Email Statistics').missingOptional));
  ok('other columns are unaffected', em.columnLetters.raised === 'P' && em.columnLetters.clickRate === 'O');
  ok('every openRate value is null, not NaN',
    em.rows.every(function (r2) { return r2.openRate === null; }));
})();

console.log('\n== Renamed tab ==');
(function () {
  var wb = load();
  var i = wb.SheetNames.indexOf('P2P Statistics');
  wb.SheetNames[i] = 'P2P Stats 2026';
  wb.Sheets['P2P Stats 2026'] = wb.Sheets['P2P Statistics'];
  delete wb.Sheets['P2P Statistics'];
  var data = parseWb(wb);
  var t = tabReport(data, 'P2P Statistics');
  ok('the missing tab is reported by name, not ignored', t && t.status === 'missing', JSON.stringify(t && t.status));
  ok('a warning names the renamed tab',
    data.parseReport.warnings.some(function (w) { return /P2P Statistics/.test(w.text) && /renamed/.test(w.text); }));
  ok('the new tab name is surfaced too',
    data.parseReport.tabs.some(function (x) { return x.tab === 'P2P Stats 2026'; }));
  ok('the rest of the dashboard still has data', data.emailSends.rows.length === 235);
})();

console.log('\n== Required column missing ==');
(function () {
  var wb = load();
  var ws = wb.Sheets['Email Statistics'];
  delete ws['E5'];
  for (var r = 6; r <= 240; r++) delete ws['E' + r];
  var data = parseWb(wb);
  var t = tabReport(data, 'Email Statistics');
  ok('the tab fails rather than guessing', t.status === 'failed', t.status);
  ok('the failure names the missing column',
    (t.missingRequired || []).indexOf('Date') !== -1, JSON.stringify(t.missingRequired));
  ok('other tabs still parse', data.p2pSends.rows.length === 11);
})();

console.log('\n== Text pasted into numeric cells ==');
(function () {
  var wb = load();
  var ws = wb.Sheets['Email Statistics'];
  ws['P6'] = { t: 's', v: '$1,234.56' };      // formatted currency
  ws['P7'] = { t: 's', v: '(500)' };          // parens-as-negative
  ws['P8'] = { t: 's', v: 'TBD' };            // genuinely unreadable
  ws['N9'] = { t: 's', v: '13.5%' };          // a rate typed with a percent sign
  ws['N10'] = { t: 'n', v: 45 };              // a rate typed as 45 meaning 45%
  var data = parseWb(wb);
  var em = data.emailSends;
  var byRow = {};
  em.rows.forEach(function (r) { byRow[r._row] = r; });
  ok('"$1,234.56" parses to 1234.56', byRow[6] && Math.abs(byRow[6].raised - 1234.56) < 1e-9, byRow[6] && byRow[6].raised);
  ok('"(500)" parses to -500', byRow[7] && byRow[7].raised === -500, byRow[7] && byRow[7].raised);
  ok('"TBD" becomes null with a counted reason, and the row survives',
    byRow[8] && byRow[8].raised === null && byRow[8].recipients > 0 &&
    em.rejections.some(function (r) { return r.column === 'Raised' && /not numeric/.test(r.reason); }),
    JSON.stringify(em.rejections));
  ok('the row count is unchanged — one bad cell does not discard a send',
    em.rows.length === 235, 'got ' + em.rows.length);
  ok('"13.5%" becomes 0.135, not 13.5', byRow[9] && Math.abs(byRow[9].openRate - 0.135) < 1e-9, byRow[9] && byRow[9].openRate);
  ok('a bare 45 in a rate column is flagged suspect, not charted',
    byRow[10] && byRow[10].openRate === null &&
    em.rejections.some(function (r) { return r.column === 'Open Rate' && /rate > 1/.test(r.reason); }),
    JSON.stringify(em.rejections.filter(function (r) { return r.column === 'Open Rate'; })));
})();

console.log('\n== Degenerate workbooks ==');
(function () {
  var empty = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(empty, XLSX.utils.aoa_to_sheet([[]]), 'Sheet1');
  var d1 = parseWb(empty);
  ok('an empty workbook produces zero parsed tabs, no crash',
    d1.parseReport.tabs.filter(function (t) { return t.status === 'parsed'; }).length === 0);
  ok('the unknown sheet is still listed', d1.parseReport.tabs.some(function (t) { return t.tab === 'Sheet1'; }));

  var wrong = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wrong, XLSX.utils.aoa_to_sheet([['Name', 'Qty'], ['Widget', 3]]), 'Inventory');
  var d2 = parseWb(wrong);
  ok('a completely unrelated workbook does not crash',
    d2.parseReport.tabs.filter(function (t) { return t.status === 'parsed'; }).length === 0);
  ok('every expected tab is reported missing',
    d2.parseReport.warnings.filter(function (w) { return w.kind === 'tab'; }).length >= 8);
})();

console.log('\n== Volume ==');
(function () {
  var wb = load();
  var ws = wb.Sheets['Email Statistics'];
  for (var i = 0; i < 470; i++) {
    var from = 6 + (i % 235), to = 241 + i;
    for (var c = 0; c <= 23; c++) {
      var L = XLSX.utils.encode_col(c);
      var src = ws[L + from];
      if (src) ws[L + to] = { t: src.t, v: src.v, w: src.w };
    }
    ws['E' + to] = { t: 'n', v: 46265 + i };
  }
  ws['!ref'] = 'A1:X1181';
  var t0 = Date.now();
  var data = parseWb(wb);
  var ms = Date.now() - t0;
  ok('3× data volume parses (705 rows)', data.emailSends.rows.length === 705, 'got ' + data.emailSends.rows.length);
  ok('parse stays under 3s at 3× volume', ms < 3000, ms + 'ms');
  fs.writeFileSync(path.join(OUT, 'volume.xlsx'), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
})();

console.log('\nvariants written to ' + OUT);
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
