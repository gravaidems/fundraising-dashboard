/*
 * Headless parser assertions.
 *
 * Running under Node also proves the DOM-free requirement: if CABParse ever touches
 * `document` or `window`, this throws.
 */
var fs = require('fs');
var path = require('path');
var XLSX = require('../vendor/xlsx.full.min.js');
var CABParse = require('../src/parse.js');

var WB = process.env.CAB_WORKBOOK ||
  path.join(__dirname, '..', 'EXTERNAL_ CAB Digital Report.xlsx');

var pass = 0, fail = 0;
function check(name, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + '\n         expected: ' + JSON.stringify(expected) + '\n         actual:   ' + JSON.stringify(actual)); }
}
function checkThat(name, cond, detail) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (detail ? '\n         ' + detail : '')); }
}

if (typeof document !== 'undefined') throw new Error('document should not exist under Node');

var parser = CABParse.createParser(XLSX);
var data = parser.parseWorkbook(fs.readFileSync(WB));

function iso(d) { return d.toISOString().slice(0, 10); }
function monthKey(d) { return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0'); }

console.log('\n== Email Statistics ==');
var em = data.emailSends;
check('emailSends.length', em.rows.length, 235);
check('email header row', em.headerRow, 5);
check('email resolved range', em.dataRange, 'A6:X240');
var edates = em.rows.map(function (r) { return r.date; });
check('email date range',
  [iso(new Date(Math.min.apply(null, edates))), iso(new Date(Math.max.apply(null, edates)))],
  ['2026-01-07', '2026-08-24']);
var byMonth = {};
em.rows.forEach(function (r) { var k = monthKey(r.date); byMonth[k] = (byMonth[k] || 0) + r.raised; });
check('email raised by month (rounded)',
  Object.keys(byMonth).sort().map(function (k) { return k + ':' + Math.round(byMonth[k]); }),
  ['2026-01:18760', '2026-02:17358', '2026-03:13478', '2026-04:13945',
   '2026-05:24553', '2026-06:30205', '2026-07:35778', '2026-08:42309']);
check('email column letters (raised/recipients)', [em.columnLetters.raised, em.columnLetters.recipients], ['P', 'M']);
check('coercion rejections on Email Statistics', em.rejections.length, 0);
checkThat('no rate > 1 in email rows',
  em.rows.every(function (r) { return ['openRate','clickRate','donateRate','unsubRate'].every(function (k) { return r[k] == null || r[k] <= 1; }); }));

console.log('\n== P2P Statistics ==');
var p2 = data.p2pSends;
check('p2pSends.length', p2.rows.length, 11);
var pdates = p2.rows.map(function (r) { return r.date; });
check('p2p date range',
  [iso(new Date(Math.min.apply(null, pdates))), iso(new Date(Math.max.apply(null, pdates)))],
  ['2026-07-23', '2026-08-24']);

console.log('\n== Ads Report ==');
var ads = data.adsMonthly;
check('adsMonthly.length', ads.rows.length, 3);
checkThat('all ads revenue fields null',
  ads.rows.every(function (r) {
    return r.lifetimeRaisedStandard == null && r.lifetimeRoasStandard == null &&
           r.lifetimeRaisedFinanceAdj == null && r.lifetimeRoasFinanceAdj == null;
  }));
check('ads positional letters H/I/J/K',
  [ads.columnLetters.lifetimeRaisedStandard, ads.columnLetters.lifetimeRoasStandard,
   ads.columnLetters.lifetimeRaisedFinanceAdj, ads.columnLetters.lifetimeRoasFinanceAdj],
  ['H', 'I', 'J', 'K']);

console.log('\n== Monthly Goals ==');
var g = data.monthlyGoals;
check('monthlyGoals.length (Totals row 18 excluded)', g.rows.length, 12);
check('goals header row', g.headerRow, 4);
var aug = g.rows.filter(function (r) { return monthKey(r.month) === '2026-08'; })[0];
check('goals Aug gross raised', aug && Math.round(aug.grossRaised * 100) / 100, 639615.65);
checkThat('goals excludes a totals row with a stated reason',
  g.rowsExcluded.some(function (e) { return /totals/i.test(e.reason); }),
  'rowsExcluded = ' + JSON.stringify(g.rowsExcluded));

console.log('\n== Digital Projections (multi-block) ==');
var sc = data.projectionScenarios;
check('projectionScenarios.length', sc.length, 3);
check('scenario titles', sc.map(function (s) { return s.scenario; }),
  ['Digital Projections (Low Investment)',
   'Digital Projections (Medium Investment)',
   'Digital Projections (High Investment)']);
check('scenario header rows', sc.map(function (s) { return s.headerRow; }), [2, 20, 74]);
check('medium scenario rows (Totals row 70 excluded)', sc[1].rows.length, 48);
check('scenario usability', sc.map(function (s) { return s.usable; }), [false, true, false]);
checkThat('unusable scenarios state a reason',
  sc.filter(function (s) { return !s.usable; }).every(function (s) { return !!s.unusableReason; }));

console.log('\n== High-Dollar (PII) ==');
var hd = data.highDollarAgg;
check('highDollarAgg.total.count', hd.agg.total.count, 22);
checkThat('high-dollar sum > 0', hd.agg.total.sum > 0);

console.log('\n== PII containment ==');
var EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
// Donor-identity keys that must not exist anywhere.
var BANNED_KEYS = ['first', 'last', 'receipt_id', 'receiptId', 'fundraising_page', 'fundraisingPage'];
// "email" is banned only under the High-Dollar subtree: on the monthly tabs it is the
// Email *channel* dollar column (a number), not a donor address.
var EMAIL_KEY_SCOPE = /highDollar/;
var piiKeyHits = [], piiValueHits = [], seen = new Set();
(function walk(node, pathStr) {
  if (node == null) return;
  if (typeof node === 'string') {
    if (EMAIL_RE.test(node)) piiValueHits.push(pathStr);
    return;
  }
  if (typeof node !== 'object') return;
  if (seen.has(node)) return;
  seen.add(node);
  if (node instanceof Date) return;
  if (Array.isArray(node)) { node.forEach(function (v, i) { walk(v, pathStr + '[' + i + ']'); }); return; }
  Object.keys(node).forEach(function (k) {
    if (BANNED_KEYS.indexOf(k) !== -1) piiKeyHits.push(pathStr + '.' + k);
    if (k === 'email' && EMAIL_KEY_SCOPE.test(pathStr)) piiKeyHits.push(pathStr + '.' + k);
    walk(node[k], pathStr + '.' + k);
  });
})(data, '$');
check('no PII keys anywhere in parser output', piiKeyHits, []);
check('no email-shaped strings anywhere in parser output', piiValueHits, []);

console.log('\n== Parse report ==');
var rep = data.parseReport;
check('all 12 tabs reported', rep.tabs.length, 12);
var fb = rep.tabs.filter(function (t) { return t.tab === 'FB Audience Report'; })[0];
check('FB Audience Report skipped, not crashed', fb && fb.status, 'skipped');
checkThat('FB Audience Report warns about #VALUE!',
  fb && /#VALUE!/.test(JSON.stringify(fb.notes)), JSON.stringify(fb && fb.notes));
checkThat('every excluded-row entry carries a reason',
  rep.tabs.every(function (t) {
    return (t.rowsExcluded || []).every(function (e) { return typeof e.reason === 'string' && e.reason.length > 0; });
  }));
checkThat('no dataRange uses the nominal sheet dimension',
  !/X1181|Y1000|K1001|J500|D985/.test(JSON.stringify(rep.tabs.map(function (t) { return t.dataRange; }))));

console.log('\n== Raw sheet capture ==');
(function () {
  var wb = XLSX.read(fs.readFileSync(WB), { type: 'buffer' });
  var PII = ['receipt_id', 'fundraising_page', 'first', 'last', 'email'];
  var problems = [], totalCols = 0;

  data.parseReport.tabs.forEach(function (t) {
    var ws = wb.Sheets[t.tab];
    if (!ws || !ws['!ref']) return;
    if (!t.raw) { problems.push(t.tab + ': no raw capture'); return; }

    // Every column the sheet actually names must be either captured or explicitly omitted.
    var hdrRow = t.raw.headerRow;
    if (!hdrRow) return;                       // headerless sheets are captured by letter only
    var rng = XLSX.utils.decode_range(ws['!ref']);
    var expected = [];
    for (var c = rng.s.c; c <= rng.e.c; c++) {
      var cell = ws[XLSX.utils.encode_cell({ r: hdrRow - 1, c: c })];
      var nm = cell && cell.v != null ? String(cell.v).replace(/[\r\n]+/g, ' ').trim() : '';
      if (nm) expected.push({ name: nm, letter: XLSX.utils.encode_col(c) });
    }
    var captured = {};
    t.raw.columns.forEach(function (c) { if (c.name) captured[c.letter] = c.name; });
    var omitted = (t.raw.omitted || []).join(' ');

    expected.forEach(function (e) {
      totalCols++;
      if (captured[e.letter]) return;
      if (omitted.indexOf('(' + e.letter + ')') !== -1) return;   // dropped on purpose
      problems.push(t.tab + ': column "' + e.name + '" (' + e.letter + ') is in the sheet but not in the raw table');
    });

    // And nothing PII may have survived into the capture. Scoped to the donor tab:
    // on the monthly tabs "Email" is the channel column, a dollar figure, not an address.
    if (t.tab === 'High-Dollar Donations') {
      Object.keys(captured).forEach(function (letter) {
        if (PII.indexOf(captured[letter].toLowerCase()) !== -1) {
          problems.push(t.tab + ': PII column "' + captured[letter] + '" (' + letter + ') reached the raw table');
        }
      });
    }
  });

  check('every named column in every sheet is captured or explicitly omitted', problems, []);
  checkThat('the audit actually covered the workbook (>= 100 columns)', totalCols >= 100, 'saw ' + totalCols);
  var hd = data.parseReport.tabs.filter(function (t) { return t.tab === 'High-Dollar Donations'; })[0];
  check('the five identity columns are the ones omitted', hd.raw.omitted.length, 5);
  checkThat('all 12 sheets have a raw capture',
    data.parseReport.tabs.filter(function (t) { return t.raw && t.raw.rows.length > 0; }).length === 12,
    data.parseReport.tabs.filter(function (t) { return !t.raw || !t.raw.rows.length; }).map(function (t) { return t.tab; }).join(', '));
})();

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
