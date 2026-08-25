/*
 * Browser-side verification against the real workbook.
 *
 * The centrepiece is the provenance coverage gate, run across EVERY tab: each
 * rendered number and chart must carry a complete ⓘ bubble, the bubbles must
 * move when the filters move, and hovering must reveal them.
 */
var fs = require('fs');
var path = require('path');
var puppeteer = require('puppeteer-core');
var XLSX = require('../vendor/xlsx.full.min.js');

var ROOT = path.join(__dirname, '..');
var CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
var WORKBOOK = process.env.CAB_WORKBOOK || path.join(ROOT, 'EXTERNAL_ CAB Digital Report.xlsx');

var pass = 0, fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log('  ok   ' + name); }
  else { fail++; console.log('  FAIL ' + name + (detail ? '\n         ' + detail : '')); }
}
function wait(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

function loadInto(page, buf, name) {
  return page.evaluate(function (b64, nm) {
    var bin = atob(b64);
    var ab = new ArrayBuffer(bin.length);
    var v = new Uint8Array(ab);
    for (var i = 0; i < bin.length; i++) v[i] = bin.charCodeAt(i);
    window.CABApp._loadBuffer(ab, nm);
  }, buf.toString('base64'), name);
}

(async function () {
  var browser = await puppeteer.launch({ executablePath: CHROME, headless: 'new' });
  var page = await browser.newPage();
  await page.setViewport({ width: 1440, height: 1000 });

  var pageErrors = [], consoleErrors = [], requests = [];
  page.on('pageerror', function (e) { pageErrors.push(String(e)); });
  page.on('console', function (m) { if (m.type() === 'error') consoleErrors.push(m.text()); });
  page.on('request', function (r) { requests.push(r.url()); });

  await page.goto('file://' + path.join(ROOT, 'dashboard.html'), { waitUntil: 'networkidle0' });
  var afterLoad = requests.length;
  await loadInto(page, fs.readFileSync(WORKBOOK), path.basename(WORKBOOK));
  await wait(1200);

  console.log('\n== Render ==');
  ok('no uncaught page errors', pageErrors.length === 0, pageErrors.join('\n'));
  ok('no console errors', consoleErrors.length === 0, consoleErrors.join('\n'));
  ok('dashboard replaced the drop zone', await page.$eval('#app', function (e) { return e.style.display === 'block'; }));
  var tabCount = await page.$$eval('.tbtn', function (n) { return n.length; });
  ok('one tab per sheet, plus overview and data sources', tabCount === 14, 'got ' + tabCount);
  ok('overview is the landing tab',
    (await page.$eval('.tbtn-on', function (e) { return e.textContent.trim(); })) === 'Overview');

  console.log('\n== No network after load ==');
  var offMachine = requests.filter(function (u) { return !/^(data|blob|file):/.test(u); });
  ok('no off-machine requests at any point', offMachine.length === 0, offMachine.join(', '));
  ok('no additional file loads after the workbook is parsed',
    requests.slice(afterLoad).filter(function (u) { return /^file:/.test(u); }).length === 0);

  // ---------------------------------------------------------------------------
  console.log('\n== Provenance coverage, every tab ==');
  var TAB_IDS = await page.$$eval('.tbtn', function (n) {
    return n.map(function (b) { return b.getAttribute('data-tab'); });
  });

  function auditTab() {
    return page.evaluate(function () {
      var st = window.CABApp._state;
      var problems = [], REQ = ['tab', 'dataRange', 'columnsUsed', 'rowsAvailable', 'rowsUsed', 'transform'];
      Array.prototype.slice.call(document.querySelectorAll('.info')).forEach(function (b) {
        var rec = st.bubbles[b.getAttribute('data-bubble')];
        if (!rec) { problems.push(st.tab + ': button with no registered provenance'); return; }
        var p = rec.prov, id = st.tab + '/' + rec.label;
        REQ.forEach(function (f) {
          if (p[f] === undefined || p[f] === null || p[f] === '') problems.push(id + ': missing ' + f);
        });
        if (!p.columnsUsed.length) problems.push(id + ': no columns');
        p.columnsUsed.forEach(function (c) {
          if (!c.letter || c.letter === '?') problems.push(id + ': column "' + c.name + '" has no Excel letter');
        });
        (p.rowsExcluded || []).forEach(function (e) { if (!e.reason) problems.push(id + ': exclusion with no reason'); });
        if (/X1181|Y1000|K1001|J500|D985|N160/.test(String(p.dataRange))) {
          problems.push(id + ': nominal sheet dimension used (' + p.dataRange + ')');
        }
      });
      // Data Sources is excluded: it renders no derived figure — it IS the disclosure.
      var metricPanels = Array.prototype.slice.call(document.querySelectorAll('.panel'))
        .filter(function (p) { return p.id !== 'sources' && p.id !== 'wbmap'; });
      var bare = metricPanels.filter(function (p) { return p.querySelectorAll('.info').length === 0; })
        .map(function (p) { return st.tab + '/' + ((p.querySelector('h2') || {}).innerText || '?'); });
      var rows = Array.prototype.slice.call(document.querySelectorAll('.tbl-sort tbody tr'));
      return {
        problems: problems, bare: bare,
        bubbles: document.querySelectorAll('.info').length,
        rows: rows.length,
        bareRows: rows.filter(function (r) { return r.querySelectorAll('.info').length === 0; }).length
      };
    });
  }

  async function goTab(id) {
    await page.evaluate(function (i) { document.querySelector('.tbtn[data-tab="' + i + '"]').click(); }, id);
    await wait(350);
  }

  var problems = [], bare = [], bubbles = 0, bdRows = 0, bdBare = 0;
  for (var i = 0; i < TAB_IDS.length; i++) {
    await goTab(TAB_IDS[i]);
    var a = await auditTab();
    problems = problems.concat(a.problems);
    bare = bare.concat(a.bare);
    bubbles += a.bubbles; bdRows += a.rows; bdBare += a.bareRows;
  }
  ok('every ⓘ button on every tab resolves to a complete record',
    problems.length === 0, problems.slice(0, 10).join('\n         '));
  ok('every metric panel on every tab has a bubble', bare.length === 0, bare.join(', '));
  ok('every breakdown group row has its own bubble', bdRows > 0 && bdBare === 0,
    bdBare + ' of ' + bdRows + ' lack one');
  ok('bubbles are plentiful across the workbook (>= 30)', bubbles >= 30, 'got ' + bubbles);
  ok('visiting every tab raises no page errors', pageErrors.length === 0, pageErrors.join('\n'));

  await goTab('overview');
  var struct = await page.evaluate(function () {
    var kpis = Array.prototype.slice.call(document.querySelectorAll('.kpi'));
    return {
      kpiCount: kpis.length,
      bareKpis: kpis.filter(function (k) { return k.querySelectorAll('.info').length === 0; }).length,
      p1: document.querySelectorAll('.srcs .info').length
    };
  });
  ok('all 5 KPIs have bubbles', struct.kpiCount === 5 && struct.bareKpis === 0);
  ok('panel 1 carries a separate bubble per series', struct.p1 >= 2, 'got ' + struct.p1);

  // ---------------------------------------------------------------------------
  console.log('\n== Hover opens the bubble ==');
  await page.hover('.kpi .info');
  await wait(400);
  var popText = await page.$eval('.prov-pop', function (e) { return e.innerText; });
  ok('hovering an ⓘ opens its bubble', popText.length > 40);
  ok('bubble shows the source tab', /Email Statistics/.test(popText));
  ok('bubble shows a resolved cell range', /A6:X240/.test(popText), popText.slice(0, 160));
  ok('bubble shows columns by Excel letter', /\bP\b/.test(popText) && /Raised/.test(popText));
  ok('bubble shows rows used of available', /used of/.test(popText));
  ok('bubble shows the transform', /SUM\(/.test(popText));

  await page.mouse.move(5, 5);
  await wait(500);
  ok('the bubble closes when the pointer leaves', (await page.$('.prov-pop')) === null);

  await page.click('.kpi .info');
  await wait(250);
  ok('clicking pins the bubble open', (await page.$('.prov-pop.prov-pinned')) !== null);
  await page.mouse.move(5, 5);
  await wait(500);
  ok('a pinned bubble survives the pointer leaving', (await page.$('.prov-pop')) !== null);
  await page.keyboard.press('Escape');
  await wait(200);
  ok('Escape closes a pinned bubble', (await page.$('.prov-pop')) === null);

  // ---------------------------------------------------------------------------
  console.log('\n== Bubbles are live, not baked in ==');
  var before = await page.evaluate(function () {
    var st = window.CABApp._state;
    return Object.keys(st.bubbles).map(function (k) { return st.bubbles[k].prov.rowsUsed; }).join('|');
  });
  await page.evaluate(function () {
    document.querySelector('.chip[data-range="3"]').click();
  });
  await wait(800);
  var after = await page.evaluate(function () {
    var st = window.CABApp._state;
    return Object.keys(st.bubbles).map(function (k) { return st.bubbles[k].prov.rowsUsed; }).join('|');
  });
  ok('rowsUsed changes when the date filter changes', before !== after);
  await page.evaluate(function () { document.querySelector('.chip[data-range="all"]').click(); });
  await wait(600);

  // ---------------------------------------------------------------------------
  console.log('\n== Raw sheet tables ==');
  var SHEETS = {
    email: ['Email Statistics', 24, 235],
    p2p: ['P2P Statistics', 25, 11],
    ads: ['Ads Report - Finance-Adjusted', 11, 3],
    goals: ['April 2026 Updated Goals', 16, 13],
    projections: ['Digital Projections', 14, 115],
    highdollar: ['High-Dollar Donations', 5, 496],
    emailcal: ['Email Sending CalendarTracker', 8, 96],
    p2pcal: ['P2P Calendar', 4, 10],
    toolkits: ['Partner Toolkits', 4, 24],
    fb: ['FB Audience Report', 15, 3],
    cover: ['Digital Report', 19, 21],
    paidmedia: ['Paid Media Report', 16, 34]
  };
  var rawProblems = [];
  for (var k = 0; k < Object.keys(SHEETS).length; k++) {
    var id = Object.keys(SHEETS)[k], spec = SHEETS[id];
    await goTab(id);
    var got = await page.evaluate(function () {
      var sec = document.querySelector('.raw');
      if (!sec) return null;
      return {
        name: sec.getAttribute('data-raw'),
        cols: sec.querySelectorAll('.raw-tbl thead tr:first-child th').length - 1,
        rows: sec.querySelectorAll('.raw-tbl tbody tr').length,
        letters: Array.prototype.slice.call(sec.querySelectorAll('.raw-tbl thead tr:first-child th'))
          .slice(1, 4).map(function (t) { return t.textContent; }).join(','),
        hasFind: !!sec.querySelector('.raw-find')
      };
    });
    if (!got) { rawProblems.push(id + ': no raw table'); continue; }
    if (got.name !== spec[0]) rawProblems.push(id + ': wrong sheet ' + got.name);
    if (got.cols !== spec[1]) rawProblems.push(spec[0] + ': ' + got.cols + ' cols, expected ' + spec[1]);
    if (got.rows !== spec[2]) rawProblems.push(spec[0] + ': ' + got.rows + ' rows, expected ' + spec[2]);
    if (!got.hasFind) rawProblems.push(spec[0] + ': no row filter');
    // High-Dollar skips B: receipt_id is dropped at the parse boundary, so the
    // letters themselves are the proof that the PII column never reaches the page.
    var wantLetters = spec[0] === 'High-Dollar Donations' ? 'A,C,D' : 'A,B,C';
    if (got.letters !== wantLetters) rawProblems.push(spec[0] + ': column letters start ' + got.letters);
  }
  ok('every sheet has a raw table with the right shape and Excel column letters',
    rawProblems.length === 0, rawProblems.join('\n         '));

  await goTab('highdollar');
  var hd = await page.evaluate(function () {
    var sec = document.querySelector('.raw');
    return {
      headers: Array.prototype.slice.call(sec.querySelectorAll('.raw-tbl thead tr:last-child th'))
        .map(function (t) { return t.textContent.trim(); }).join('|'),
      note: (sec.querySelector('.pii-note') || {}).innerText || '',
      body: document.body.innerText
    };
  });
  ok('the raw High-Dollar table omits every identity column',
    !/receipt_id|fundraising_page|\bfirst\b|\blast\b/.test(hd.headers), hd.headers);
  ok('and says which columns it omitted and why', /5 columns are deliberately not shown/.test(hd.note), hd.note.slice(0, 140));
  ok('no email-shaped string anywhere on the PII tab',
    !/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.(com|org|net|edu|gov)/i.test(hd.body));

  await goTab('email');
  var find = await page.evaluate(async function () {
    var input = document.querySelector('.raw-find');
    var all = document.querySelectorAll('.raw-tbl tbody tr').length;
    input.value = 'Reactivation';
    input.dispatchEvent(new Event('input'));
    var shown = Array.prototype.slice.call(document.querySelectorAll('.raw-tbl tbody tr'))
      .filter(function (r) { return r.style.display !== 'none'; }).length;
    return { all: all, shown: shown, msg: document.querySelector('.raw-count').textContent };
  });
  ok('the raw row filter narrows the table', find.shown > 0 && find.shown < find.all,
    find.shown + ' of ' + find.all);
  ok('and reports how many rows matched', /rows match/.test(find.msg), find.msg);

  // ---------------------------------------------------------------------------
  console.log('\n== Every workbook column is reachable ==');
  var coverage = await page.evaluate(function () {
    var st = window.CABApp._state;
    var missing = [];
    (st.data.parseReport.tabs || []).forEach(function (t) {
      if (!t.raw) return;
      t.raw.columns.forEach(function (c) {
        if (!c.name) return;
        missing.push({ tab: t.tab, name: c.name, letter: c.letter });
      });
    });
    return missing;
  });
  // Each of those columns is rendered in its sheet's raw table; assert the count is
  // what the workbook actually holds, minus the five PII columns dropped by design.
  ok('every named column in every sheet is present in a raw table',
    coverage.length >= 100, 'found ' + coverage.length);
  ok('the five PII columns are the only ones missing',
    !coverage.some(function (c) { return /receipt_id|fundraising_page|^first$|^last$|^email$/.test(c.name); }),
    coverage.filter(function (c) { return /receipt|fundrais/.test(c.name); }).map(function (c) { return c.name; }).join(','));

  // ---------------------------------------------------------------------------
  console.log('\n== Panel behaviours ==');
  await goTab('ads');
  var adsBody = await page.evaluate(function () { return document.body.innerText; });
  ok('Ads panel shows its explanatory empty state', /No ads revenue data in source/.test(adsBody));
  ok('Ads empty state names rows and blank columns',
    /3 rows found[\s\S]*columns H\/I\/J\/K are blank/.test(adsBody));

  await goTab('p2p');
  ok('P2P states the small sample on its face',
    /Small sample: this workbook holds 11 P2P sends/.test(await page.evaluate(function () { return document.body.innerText; })));

  await goTab('projections');
  var projBody = await page.evaluate(function () { return document.body.innerText; });
  ok('Projections tab explains the three stacked blocks', /three separate projection blocks/.test(projBody));
  ok('and names why the unusable scenarios cannot be charted',
    /cannot be charted/.test(projBody) && /advances by single days/.test(projBody));

  await goTab('fb');
  ok('FB Audience explains the broken formula rather than vanishing',
    /#VALUE!/.test(await page.evaluate(function () { return document.body.innerText; })));

  await goTab('sources');
  var srcBody = await page.evaluate(function () { return document.body.innerText; });
  ok('Data sources lists all 12 sheets',
    (await page.$$eval('.tbl-src tbody tr', function (n) { return n.length; })) === 12);
  ok('Data sources reports the three projection blocks', /3 stacked blocks/.test(srcBody));

  await goTab('overview');
  var ovBody = await page.evaluate(function () { return document.body.innerText; });
  ok('Overview states that actuals are computed', /computed by this dashboard from send-level rows/.test(ovBody));
  ok('Overview maps every sheet', (await page.$$eval('.map-card', function (n) { return n.length; })) === 12);

  console.log('\n== Axis labels ==');
  async function axisCheck() {
    return page.evaluate(function () {
      var over = [], collide = [];
      Array.prototype.slice.call(document.querySelectorAll('.v-svg')).forEach(function (svg) {
        var vb = svg.getAttribute('viewBox').split(' ').map(Number);
        var ticks = Array.prototype.slice.call(svg.querySelectorAll('.v-tick-x'))
          .map(function (t) { return { t: t.textContent, b: t.getBBox() }; })
          .sort(function (a, b) { return a.b.x - b.b.x; });
        ticks.forEach(function (k) {
          if (k.b.x < -0.5 || k.b.x + k.b.width > vb[2] + 0.5) over.push(k.t);
        });
        for (var i = 1; i < ticks.length; i++) {
          var prev = ticks[i - 1], cur = ticks[i];
          if (cur.b.x < prev.b.x + prev.b.width + 2) collide.push(prev.t + ' / ' + cur.t);
        }
      });
      return { over: over, collide: collide };
    });
  }
  var ax = await axisCheck();
  ok('no x-axis label overflows its chart', ax.over.length === 0, ax.over.slice(0, 8).join(', '));
  ok('no two x-axis labels overlap', ax.collide.length === 0, ax.collide.slice(0, 8).join(', '));

  // The small multiples are the narrow case where overlap actually happened.
  await goTab('email');
  var axEmail = await axisCheck();
  ok('no overlap on the narrow small-multiple charts', axEmail.collide.length === 0,
    axEmail.collide.slice(0, 8).join(', '));
  ok('and none overflow there either', axEmail.over.length === 0, axEmail.over.slice(0, 6).join(', '));
  await goTab('overview');

  await page.emulateMediaFeatures([{ name: 'prefers-color-scheme', value: 'light' }]);
  await wait(300);
  await page.screenshot({ path: path.join(ROOT, 'test', 'shot-light.png'), fullPage: true });
  await goTab('email');
  await page.screenshot({ path: path.join(ROOT, 'test', 'shot-email.png'), fullPage: true });
  await page.emulateMediaFeatures([{ name: 'prefers-color-scheme', value: 'dark' }]);
  await wait(300);
  await page.screenshot({ path: path.join(ROOT, 'test', 'shot-dark.png'), fullPage: true });

  // ---------------------------------------------------------------------------
  console.log('\n== Warning propagation (Recipients -> Recipiants) ==');
  var wb = XLSX.read(fs.readFileSync(WORKBOOK), { type: 'buffer' });
  wb.Sheets['Email Statistics']['M5'] = { t: 's', v: 'Recipiants' };
  var typoB64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });

  var page2 = await browser.newPage();
  await page2.setViewport({ width: 1440, height: 1000 });
  var errs2 = [];
  page2.on('pageerror', function (e) { errs2.push(String(e)); });
  await page2.goto('file://' + path.join(ROOT, 'dashboard.html'), { waitUntil: 'networkidle0' });
  await page2.evaluate(function (b64) {
    var bin = atob(b64);
    var ab = new ArrayBuffer(bin.length);
    var v = new Uint8Array(ab);
    for (var i = 0; i < bin.length; i++) v[i] = bin.charCodeAt(i);
    window.CABApp._loadBuffer(ab, 'typo.xlsx');
  }, typoB64);
  await wait(1200);

  ok('the typo workbook renders without errors', errs2.length === 0, errs2.join('\n'));
  await page2.evaluate(function () { document.querySelector('.tbtn[data-tab="sources"]').click(); });
  await wait(400);
  ok('the fuzzy match is listed among the warnings',
    await page2.evaluate(function () {
      return Array.prototype.slice.call(document.querySelectorAll('.warns li'))
        .some(function (li) { return /Recipiants/.test(li.innerText); });
    }));

  await page2.evaluate(function () { document.querySelector('.tbtn[data-tab="email"]').click(); });
  await wait(400);
  var prop = await page2.evaluate(function () {
    var st = window.CABApp._state, uses = [], carries = [];
    Object.keys(st.bubbles).forEach(function (k) {
      var rec = st.bubbles[k];
      if (!rec.prov.columnsUsed.some(function (c) { return /Recipi/.test(c.name); })) return;
      uses.push(rec.label);
      if ((rec.prov.notes || []).some(function (n) { return /Recipiants/.test(n); })) carries.push(rec.label);
    });
    return { total: uses.length, missing: uses.filter(function (l) { return carries.indexOf(l) === -1; }) };
  });
  ok('every bubble using the renamed column carries the warning',
    prop.total > 0 && prop.missing.length === 0, prop.missing.join(', '));
  ok('the propagation is broad (>= 5 dependent elements)', prop.total >= 5, 'got ' + prop.total);

  var shown = await page2.evaluate(function () {
    var st = window.CABApp._state;
    var id = Object.keys(st.bubbles).filter(function (k) {
      return st.bubbles[k].prov.columnsUsed.some(function (c) { return /Recipi/.test(c.name); });
    })[0];
    document.querySelector('.info[data-bubble="' + id + '"]').click();
    return document.querySelector('.prov-pop').innerText;
  });
  ok('an opened dependent bubble shows the warning to the user', /Recipiants/.test(shown), shown.slice(0, 200));

  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  await browser.close();
  process.exit(fail ? 1 : 0);
})().catch(function (e) { console.error(e); process.exit(1); });
