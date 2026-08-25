/*
 * Build: inline every dependency into a single self-contained dashboard.html.
 *
 * SheetJS is vendored from disk, never fetched at runtime — the dashboard has to
 * work offline, from file://, on a locked-down machine.
 */
var fs = require('fs');
var path = require('path');

var root = __dirname;
function read(p) { return fs.readFileSync(path.join(root, p), 'utf8'); }

function tag(src) { return '<script>\n' + src + '\n</script>\n'; }

// Pass the replacement as a FUNCTION. A plain string replacement would have "$&",
// "$'" and friends interpreted as patterns, and the minified SheetJS is full of
// them — that silently quadrupled the output the first time.
function inline(html, marker, file) {
  return html.replace(marker, function () { return tag(read(file)); });
}

var html = read('src/shell.html');
html = inline(html, '<!--XLSX-->',    'vendor/xlsx.full.min.js');
html = inline(html, '<!--PARSE-->',   'src/parse.js');
html = inline(html, '<!--METRICS-->', 'src/metrics.js');
html = inline(html, '<!--VIZ-->',     'src/viz.js');
html = inline(html, '<!--APP-->',     'src/app.js');

// The offline guarantee: OUR code must never reach the network at runtime. The
// vendored SheetJS contains a sheetjs.com URL as inert string data, so it is checked
// for the things that actually fetch rather than for the substring "https://".
var OURS = ['src/shell.html', 'src/parse.js', 'src/metrics.js', 'src/viz.js', 'src/app.js'];
var BANNED = [
  ['<script src', 'external script tag'],
  ['<link rel="stylesheet"', 'external stylesheet'],
  ['http://', 'plain-text URL'],
  ['https://', 'plain-text URL'],
  ['fetch(', 'network call'],
  ['XMLHttpRequest', 'network call'],
  ['navigator.sendBeacon', 'network call'],
  ['localStorage', 'browser storage'],
  ['sessionStorage', 'browser storage'],
  ['indexedDB', 'browser storage']
];
OURS.forEach(function (f) {
  var src = read(f);
  BANNED.forEach(function (b) {
    var i = src.indexOf(b[0]);
    if (i === -1) return;
    throw new Error('Build refused: ' + b[1] + ' (" ' + b[0] + '") in ' + f + '\\n  …' +
      src.slice(Math.max(0, i - 70), i + 90).replace(/\\n/g, ' ') + '…');
  });
});

// And the vendored parser must actually be inlined. Check for the source itself,
// not for a size threshold — html.length counts UTF-16 chars while the file is
// measured in UTF-8 bytes, and SheetJS embeds non-ASCII codepage tables.
var vendorHead = read('vendor/xlsx.full.min.js').slice(0, 300);
if (html.indexOf(vendorHead) === -1) {
  throw new Error('Build refused: the vendored SheetJS source is not present in the page.');
}
if (html.indexOf('<!--') !== -1 && /<!--(XLSX|PARSE|METRICS|VIZ|APP)-->/.test(html)) {
  throw new Error('Build refused: an inline marker was left unreplaced.');
}

fs.writeFileSync(path.join(root, 'dashboard.html'), html);
var kb = (Buffer.byteLength(html) / 1024).toFixed(0);
console.log('dashboard.html written — ' + kb + ' KB, fully self-contained.');
