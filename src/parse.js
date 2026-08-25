/*
 * CABParse — CAB Digital Report workbook parser.
 *
 * ZERO DOM DEPENDENCIES. This module must never touch `document`, `window`, or any
 * global other than the XLSX object handed to createParser(). It is the layer that
 * ports to the React app unchanged, and it is exercised headlessly under Node.
 *
 * Entry point:  createParser(XLSX).parseWorkbook(arrayBuffer) -> ParseResult
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CABParse = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // String helpers
  // ---------------------------------------------------------------------------

  /** Normalize a header for matching: lowercase, strip newlines/punctuation, collapse space. */
  function normalize(s) {
    return String(s == null ? '' : s)
      .replace(/[\r\n]+/g, ' ')
      .toLowerCase()
      .replace(/[^a-z0-9%\s]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function levenshtein(a, b) {
    if (a === b) return 0;
    if (!a.length) return b.length;
    if (!b.length) return a.length;
    var prev = new Array(b.length + 1);
    for (var j = 0; j <= b.length; j++) prev[j] = j;
    for (var i = 1; i <= a.length; i++) {
      var cur = [i];
      for (var k = 1; k <= b.length; k++) {
        cur[k] = Math.min(
          prev[k] + 1,
          cur[k - 1] + 1,
          prev[k - 1] + (a.charAt(i - 1) === b.charAt(k - 1) ? 0 : 1)
        );
      }
      prev = cur;
    }
    return prev[b.length];
  }

  /** Token-set similarity: |intersection| / |union| over whitespace tokens. */
  function tokenSetSim(a, b) {
    var A = a.split(' ').filter(Boolean);
    var B = b.split(' ').filter(Boolean);
    if (!A.length || !B.length) return 0;
    var setB = {};
    B.forEach(function (t) { setB[t] = true; });
    var inter = 0, seen = {};
    A.forEach(function (t) { if (setB[t] && !seen[t]) { inter++; seen[t] = true; } });
    var union = {};
    A.concat(B).forEach(function (t) { union[t] = true; });
    return inter / Object.keys(union).length;
  }

  var TOTALS_RE = /^(totals?|grand\s*total|sum|subtotal)$/i;

  // ---------------------------------------------------------------------------
  // Cell access + coercion
  //
  // Every coercion returns {ok:true, value} or {ok:false, reason}. Never a silent NaN.
  // ---------------------------------------------------------------------------

  function makeSheet(XLSX, ws, tabName) {
    return {
      name: tabName,
      cell: function (r, c) { return ws[XLSX.utils.encode_cell({ r: r, c: c })] || null; },
      letter: function (c) { return XLSX.utils.encode_col(c); },
      addr: function (r, c) { return XLSX.utils.encode_cell({ r: r, c: c }); },
      range: (function () {
        var ref = ws['!ref'];
        return ref ? XLSX.utils.decode_range(ref) : null;
      })()
    };
  }

  function isBlank(cell) {
    return !cell || cell.v == null || (typeof cell.v === 'string' && cell.v.trim() === '');
  }

  /**
   * Surface the error text so it becomes the rejection reason.
   *
   * Two storage forms must both be caught. Excel writes a typed error cell (t:'e').
   * Google Sheets exports — which this workbook is — write the error as a plain
   * STRING (t:'s', v:'#VALUE!'), so a t:'e' check alone silently misses them and
   * reports "not numeric" instead of naming the broken formula.
   */
  var ERROR_VALUE_RE = /^#(VALUE!|REF!|N\/A|DIV\/0!|NAME\?|NULL!|NUM!|GETTING_DATA)$/;

  function errorText(cell) {
    if (!cell) return null;
    if (cell.t === 'e') return typeof cell.w === 'string' ? cell.w : '#ERROR';
    if (typeof cell.v === 'string' && ERROR_VALUE_RE.test(cell.v.trim())) return cell.v.trim();
    return null;
  }

  function coerceText(cell) {
    var err = errorText(cell);
    if (err) return { ok: false, reason: 'Excel error ' + err };
    if (isBlank(cell)) return { ok: false, reason: 'blank' };
    return { ok: true, value: String(cell.v).replace(/[\r\n]+/g, ' ').trim() };
  }

  function coerceNumber(cell) {
    var err = errorText(cell);
    if (err) return { ok: false, reason: 'Excel error ' + err };
    if (isBlank(cell)) return { ok: false, reason: 'blank' };
    if (typeof cell.v === 'number') return { ok: true, value: cell.v, hadPercent: false };
    if (typeof cell.v === 'boolean') return { ok: false, reason: 'boolean in numeric column' };

    var raw = String(cell.v).trim();
    var neg = /^\(.*\)$/.test(raw);          // parens-as-negative
    var hadPercent = raw.indexOf('%') !== -1;
    var cleaned = raw.replace(/[()$,\s%]/g, '');
    if (cleaned === '' || cleaned === '-') return { ok: false, reason: 'blank' };
    if (!/^-?\d*\.?\d+(e[-+]?\d+)?$/i.test(cleaned)) {
      return { ok: false, reason: 'not numeric: "' + raw.slice(0, 24) + '"' };
    }
    var n = parseFloat(cleaned);
    if (!isFinite(n)) return { ok: false, reason: 'not finite: "' + raw.slice(0, 24) + '"' };
    if (neg) n = -n;
    // A staffer typing "13.5%" must not become 1350%.
    if (hadPercent) n = n / 100;
    return { ok: true, value: n, hadPercent: hadPercent };
  }

  /**
   * Rates are stored as decimal fractions (0.135 = 13.5%). A bare number > 1 is
   * suspect rather than silently charted; a value that carried a literal % has
   * already been divided by 100 and is fine.
   */
  function coerceRate(cell) {
    var r = coerceNumber(cell);
    if (!r.ok) return r;
    if (r.value > 1 && !r.hadPercent) {
      return { ok: false, reason: 'rate > 1 (' + r.value + ') — expected a decimal fraction', suspect: true };
    }
    return { ok: true, value: r.value };
  }

  var MONTH_NAMES = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

  /** Excel serial (1900 system, with the 1900 leap-year bug baked in) -> UTC Date. */
  function serialToDate(n) {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000);
  }

  function coerceDate(cell) {
    var err = errorText(cell);
    if (err) return { ok: false, reason: 'Excel error ' + err };
    if (isBlank(cell)) return { ok: false, reason: 'blank' };

    if (cell.v instanceof Date) return { ok: true, value: cell.v };
    if (typeof cell.v === 'number') {
      if (cell.v < 1 || cell.v > 2958465) {
        return { ok: false, reason: 'date serial out of range: ' + cell.v };
      }
      return { ok: true, value: serialToDate(cell.v) };
    }

    var raw = String(cell.v).trim();
    // "Jan 2026" / "January 2026" -> first of month
    var m = raw.match(/^([a-z]{3,9})\.?\s+(\d{4})$/i);
    if (m) {
      var mi = MONTH_NAMES.indexOf(m[1].slice(0, 3).toLowerCase());
      if (mi >= 0) return { ok: true, value: new Date(Date.UTC(+m[2], mi, 1)) };
    }
    var parsed = Date.parse(raw);
    if (!isNaN(parsed)) return { ok: true, value: new Date(parsed) };
    return { ok: false, reason: 'unparseable date: "' + raw.slice(0, 24) + '"' };
  }

  var COERCERS = { text: coerceText, number: coerceNumber, rate: coerceRate, date: coerceDate };

  // ---------------------------------------------------------------------------
  // Header detection
  //
  // Row numbers from the spec are HINTS ONLY. Every header row is found by scoring,
  // then compared against the hint so a shift can be reported rather than assumed.
  // ---------------------------------------------------------------------------

  function readHeaderRow(sheet, r, maxCol) {
    var cells = [];
    for (var c = 0; c <= maxCol; c++) {
      var cell = sheet.cell(r, c);
      var txt = (cell && cell.t !== 'e' && cell.v != null) ? String(cell.v).trim() : '';
      cells.push({ col: c, letter: sheet.letter(c), raw: txt, norm: normalize(txt) });
    }
    return cells;
  }

  /** Does this row look like data rather than labels? Used to confirm a header guess. */
  function looksLikeData(sheet, r, maxCol) {
    var nonEmpty = 0, numeric = 0;
    for (var c = 0; c <= maxCol; c++) {
      var cell = sheet.cell(r, c);
      if (isBlank(cell) && !errorText(cell)) continue;
      nonEmpty++;
      if ((cell && typeof cell.v === 'number') || errorText(cell)) numeric++;
    }
    return nonEmpty >= 2 && numeric >= 1;
  }

  function scoreHeaderRow(sheet, r, expected, maxCol) {
    var cells = readHeaderRow(sheet, r, maxCol);
    var strings = 0, matches = 0, matchedKeys = {};
    cells.forEach(function (h) {
      if (!h.norm) return;
      strings++;
      for (var i = 0; i < expected.length; i++) {
        var spec = expected[i];
        if (matchedKeys[spec.key]) continue;
        if (h.norm === normalize(spec.label)) { matches++; matchedKeys[spec.key] = true; return; }
      }
    });
    if (TOTALS_RE.test(cells[0] && cells[0].raw)) return { score: -1, matches: 0, cells: cells };
    var dataBelow = looksLikeData(sheet, r + 1, maxCol);
    var score = matches * 10 + strings + (dataBelow ? 5 : 0);
    return { score: score, matches: matches, strings: strings, dataBelow: dataBelow, cells: cells };
  }

  /** Minimum matches for a row to qualify as this tab's header. */
  function headerThreshold(expected) {
    return Math.max(2, Math.ceil(expected.length * 0.3));
  }

  function findHeaderRow(sheet, expected, maxCol, fromRow, toRow) {
    var best = null;
    for (var r = fromRow; r <= toRow; r++) {
      var s = scoreHeaderRow(sheet, r, expected, maxCol);
      if (s.matches < headerThreshold(expected)) continue;
      if (!best || s.score > best.score) { best = s; best.row = r; }
    }
    return best;
  }

  // ---------------------------------------------------------------------------
  // Column matching cascade: exact -> alias -> fuzzy.
  // Any fuzzy hit raises a warning that propagates into every dependent bubble.
  // ---------------------------------------------------------------------------

  function matchColumns(headerCells, expected, forcedByLetter) {
    var map = {};            // key -> {col, letter, foundName, how}
    var warnings = [];
    var used = {};           // col index -> key that claimed it
    var duplicates = [];

    function claim(spec, h, how) {
      map[spec.key] = { col: h.col, letter: h.letter, foundName: h.raw || spec.label, how: how };
      used[h.col] = spec.key;
    }

    // Stage 0 — positional overrides. The Ads tab repeats "Lifetime Raised"/"Lifetime ROAS"
    // under merged row-1 group labels, so those four columns bypass name matching entirely.
    expected.forEach(function (spec) {
      if (!forcedByLetter || !forcedByLetter[spec.key]) return;
      var letter = forcedByLetter[spec.key];
      for (var i = 0; i < headerCells.length; i++) {
        if (headerCells[i].letter === letter) {
          claim(spec, headerCells[i], 'position');
          return;
        }
      }
    });

    // Stage 1 — exact match on the normalized header. First (leftmost) wins; a later
    // duplicate of the same name is recorded rather than silently overwriting.
    expected.forEach(function (spec) {
      if (map[spec.key]) return;
      var want = normalize(spec.label);
      for (var i = 0; i < headerCells.length; i++) {
        var h = headerCells[i];
        if (!h.norm || used[h.col]) continue;
        if (h.norm === want) { claim(spec, h, 'exact'); return; }
      }
    });

    // Stage 2 — alias table.
    expected.forEach(function (spec) {
      if (map[spec.key] || !spec.aliases) return;
      var wants = spec.aliases.map(normalize);
      for (var i = 0; i < headerCells.length; i++) {
        var h = headerCells[i];
        if (!h.norm || used[h.col]) continue;
        if (wants.indexOf(h.norm) !== -1) {
          claim(spec, h, 'alias');
          warnings.push({
            kind: 'alias',
            column: spec.key,
            text: 'Column "' + h.raw + '" (' + h.letter + ') matched expected "' + spec.label + '" by alias.'
          });
          return;
        }
      }
    });

    // Stage 3 — fuzzy: Levenshtein <= 2, or token-set similarity >= 0.85.
    expected.forEach(function (spec) {
      if (map[spec.key]) return;
      var want = normalize(spec.label);
      var best = null;
      for (var i = 0; i < headerCells.length; i++) {
        var h = headerCells[i];
        if (!h.norm || used[h.col]) continue;
        var dist = levenshtein(h.norm, want);
        var sim = tokenSetSim(h.norm, want);
        if (dist <= 2 || sim >= 0.85) {
          var quality = sim - dist / 100;
          if (!best || quality > best.quality) best = { h: h, dist: dist, sim: sim, quality: quality };
        }
      }
      if (best) {
        claim(spec, best.h, 'fuzzy');
        warnings.push({
          kind: 'fuzzy',
          column: spec.key,
          text: 'Fuzzy match: found "' + best.h.raw + '" (' + best.h.letter +
                ') and used it as "' + spec.label + '" (edit distance ' + best.dist + '). Verify this is correct.'
        });
      }
    });

    // Anything left over is reported by name so staff know a new column isn't being picked up.
    var unmatched = [];
    headerCells.forEach(function (h) {
      if (!h.norm || used[h.col]) return;
      unmatched.push({ name: h.raw, letter: h.letter });
    });

    var missingRequired = expected
      .filter(function (s) { return s.required && !map[s.key]; })
      .map(function (s) { return s.label; });
    var missingOptional = expected
      .filter(function (s) { return !s.required && !map[s.key]; })
      .map(function (s) { return s.label; });

    return {
      map: map, warnings: warnings, unmatched: unmatched, duplicates: duplicates,
      missingRequired: missingRequired, missingOptional: missingOptional
    };
  }

  // ---------------------------------------------------------------------------
  // Block reading
  //
  // Sheet dimensions are NEVER used to find the end of data — five tabs in this
  // workbook carry hundreds of trailing pre-formatted blank rows. Walk from the
  // first data row, anchored on the tab's key column, until either 10 consecutive
  // blank rows or a Totals/summary row.
  // ---------------------------------------------------------------------------

  var MAX_BLANK_RUN = 10;

  /**
   * Read one contiguous table. Returns rows, the resolved range, itemized exclusions,
   * and the row where reading stopped (so multi-block tabs can resume scanning).
   */
  function readBlock(sheet, opts) {
    var spec = opts.expected;
    var colMap = opts.colMap;
    var keyKey = opts.keyKey;
    var startRow = opts.startRow;
    var hardStop = opts.hardStop;
    var dropKeys = opts.dropKeys || {};

    var rows = [];
    var excluded = {};       // reason -> count
    var rejections = [];     // {column, reason, count}
    var rejIndex = {};
    var blankRun = 0;
    var r = startRow;
    var lastDataRow = startRow - 1;
    var stoppedBy = 'end of sheet';
    var stopRow = hardStop;

    function exclude(reason) { excluded[reason] = (excluded[reason] || 0) + 1; }
    function reject(colLabel, reason) {
      var k = colLabel + '||' + reason;
      if (!rejIndex[k]) { rejIndex[k] = { column: colLabel, reason: reason, count: 0 }; rejections.push(rejIndex[k]); }
      rejIndex[k].count++;
    }

    for (; r <= hardStop; r++) {
      // A Totals/summary row terminates the block. Without this, the single blank row
      // that precedes each Totals row would not trip the blank-run rule and the total
      // would be read as data — roughly doubling every monthly figure.
      var firstCell = sheet.cell(r, 0);
      var firstTxt = (firstCell && firstCell.v != null) ? String(firstCell.v).trim() : '';
      if (TOTALS_RE.test(firstTxt)) {
        exclude('totals/summary row (row ' + (r + 1) + ')');
        stoppedBy = 'totals row';
        stopRow = r;
        break;
      }

      var keyCell = colMap[keyKey] ? sheet.cell(r, colMap[keyKey].col) : null;
      var keyBlank = isBlank(keyCell) && !errorText(keyCell);

      if (keyBlank) {
        // Is the row blank everywhere, or only in the key column?
        var anywhere = false;
        for (var k in colMap) {
          if (!isBlank(sheet.cell(r, colMap[k].col))) { anywhere = true; break; }
        }
        if (anywhere) exclude(opts.keyLabel + ' blank (row populated elsewhere)');
        blankRun++;
        if (blankRun >= MAX_BLANK_RUN) { stoppedBy = 'blank run'; stopRow = r; break; }
        continue;
      }
      blankRun = 0;

      // A row that scores as a header mid-block means a new stacked block starts here.
      if (opts.multiBlock && r > startRow) {
        var s = scoreHeaderRow(sheet, r, spec, opts.maxCol);
        if (s.matches >= headerThreshold(spec)) { stoppedBy = 'next block header'; stopRow = r; break; }
      }

      var row = { _row: r + 1 };
      var fatal = null;
      spec.forEach(function (col) {
        if (dropKeys[col.key]) return;              // PII boundary — never constructed
        var m = colMap[col.key];
        if (!m) { row[col.key] = null; return; }
        var cell = sheet.cell(r, m.col);
        var res = COERCERS[col.type || 'text'](cell);
        if (res.ok) { row[col.key] = res.value; return; }
        row[col.key] = null;
        if (res.reason !== 'blank') reject(col.label, res.reason);
        // Only the KEY column can invalidate a whole row. `required` means the
        // COLUMN must exist (enforced in matchColumns); it does not mean every cell
        // must parse. A single "TBD" in Raised must not discard an otherwise good
        // send along with its recipients, open rate and unsub count.
        if (col.key === keyKey) fatal = col.label + ': ' + res.reason;
      });

      if (fatal) { exclude('unusable ' + opts.keyLabel + ' — ' + fatal); continue; }
      rows.push(row);
      lastDataRow = r;
    }

    var cols = Object.keys(colMap).map(function (k) { return colMap[k].col; });
    var minCol = cols.length ? Math.min.apply(null, cols) : 0;
    var maxColUsed = cols.length ? Math.max.apply(null, cols) : 0;
    var dataRange = rows.length
      ? sheet.letter(minCol) + (startRow + 1) + ':' + sheet.letter(maxColUsed) + (lastDataRow + 1)
      : null;

    var rowsExcluded = Object.keys(excluded).map(function (reason) {
      return { reason: reason, count: excluded[reason] };
    });

    return {
      rows: rows,
      dataRange: dataRange,
      firstDataRow: startRow + 1,
      lastDataRow: lastDataRow + 1,
      rowsExcluded: rowsExcluded,
      rejections: rejections,
      stoppedBy: stoppedBy,
      stopRow: stopRow == null ? hardStop : stopRow
    };
  }

  /** The nearest non-empty column-A text above a header row — a stacked block's title. */
  function blockTitle(sheet, headerRow) {
    for (var r = headerRow - 1; r >= 0 && r >= headerRow - 3; r--) {
      var cell = sheet.cell(r, 0);
      if (isBlank(cell)) continue;
      var txt = String(cell.v).trim();
      if (txt && !TOTALS_RE.test(txt)) return txt;
    }
    return null;
  }

  // ---------------------------------------------------------------------------
  // Raw grid capture
  //
  // So a staffer can read the sheet as it looks in Excel, with real row numbers and
  // column letters, without leaving the dashboard. Formatted values (cell.w) are
  // preferred over raw ones so "$27,819.64" reads the way it does in the workbook.
  // ---------------------------------------------------------------------------

  var RAW_ROW_CAP = 2000;

  function displayValue(cell) {
    if (!cell) return '';
    var err = errorText(cell);
    if (err) return err;
    if (cell.w != null && cell.w !== '') return String(cell.w);
    if (cell.v == null) return '';
    if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10);
    return String(cell.v);
  }

  /** Last row in the sheet holding any non-empty cell — never the nominal dimension. */
  function lastPopulatedRow(sheet, maxCol, hardStop) {
    var last = -1;
    for (var r = 0; r <= hardStop; r++) {
      for (var c = 0; c <= maxCol; c++) {
        if (!isBlank(sheet.cell(r, c))) { last = r; break; }
      }
    }
    return last;
  }

  /**
   * Capture a readable grid for one sheet.
   * `omitLetters` drops columns at the capture boundary — this is how the High-Dollar
   * identity columns stay out of the raw view as well as out of the parsed rows.
   */
  function captureRaw(sheet, opts) {
    opts = opts || {};
    if (!sheet.range) return null;
    var maxCol = Math.min(opts.maxCol != null ? opts.maxCol : sheet.range.e.c, sheet.range.e.c);
    var hardStop = Math.min(sheet.range.e.r, RAW_ROW_CAP + 200);
    var lastRow = lastPopulatedRow(sheet, maxCol, hardStop);
    if (lastRow < 0) return { columns: [], rows: [], headerRow: null, totalRows: 0, truncated: false, omitted: [] };

    var headerRow = opts.headerRow != null ? opts.headerRow - 1 : null;
    var omit = {};
    var omitted = [];

    var columns = [];
    for (var c = 0; c <= maxCol; c++) {
      var letter = sheet.letter(c);
      var name = headerRow != null ? displayValue(sheet.cell(headerRow, c)).replace(/[\r\n]+/g, ' ').trim() : '';
      if (opts.omitNames && name && opts.omitNames.indexOf(normalize(name)) !== -1) {
        omit[c] = true;
        omitted.push(name + ' (' + letter + ')');
        continue;
      }
      columns.push({ letter: letter, name: name, col: c });
    }

    var startRow = headerRow != null ? headerRow + 1 : 0;
    var rows = [];
    var truncated = false;
    for (var r = startRow; r <= lastRow; r++) {
      if (rows.length >= RAW_ROW_CAP) { truncated = true; break; }
      var cells = [], any = false;
      columns.forEach(function (col) {
        var v = displayValue(sheet.cell(r, col.col));
        if (v !== '') any = true;
        cells.push(v);
      });
      if (!any) continue;
      rows.push({ r: r + 1, cells: cells });
    }

    return {
      columns: columns.map(function (c) { return { letter: c.letter, name: c.name }; }),
      rows: rows, headerRow: opts.headerRow || null,
      totalRows: rows.length, truncated: truncated, omitted: omitted
    };
  }

  // ---------------------------------------------------------------------------
  // Tab specifications
  //
  // hintHeaderRow is 1-indexed as it appears in Excel and is only a hint: the
  // detector scores rows and reports any shift rather than trusting the number.
  // ---------------------------------------------------------------------------

  var MONTHLY_COLUMNS = [
    { key: 'month',        label: 'Month',         type: 'date',   required: true },
    { key: 'email',        label: 'Email',         type: 'number' },
    { key: 'recurring',    label: 'Recurring',     type: 'number' },
    { key: 'ads',          label: 'Ads',           type: 'number' },
    { key: 'website',      label: 'Website',       type: 'number' },
    { key: 'tandem',       label: 'Tandem',        type: 'number' },
    { key: 'smsBroadcast', label: 'SMS Broadcast', type: 'number' },
    { key: 'smsP2p',       label: 'SMS P2P',       type: 'number' },
    { key: 'social',       label: 'Social',        type: 'number' },
    { key: 'allOther',     label: 'All Other',     type: 'number' },
    { key: 'grossRaised',  label: 'Gross Raised',  type: 'number', aliases: ['raised', 'revenue', 'total raised'] },
    { key: 'grossSpend',   label: 'Gross Spend',   type: 'number', aliases: ['spend', 'cost', 'gross cost'] },
    { key: 'netRaised',    label: 'Net Raised',    type: 'number', aliases: ['net'] }
  ];

  /** Channel columns, in stack order, for the channel-mix panel. */
  var CHANNEL_KEYS = ['email', 'recurring', 'ads', 'website', 'tandem', 'smsBroadcast', 'smsP2p', 'social', 'allOther'];

  var TAB_SPECS = [
    {
      id: 'emailSends', tab: 'Email Statistics', hintHeaderRow: 5, keyKey: 'date', keyLabel: 'Date',
      role: 'primary', maxCol: 27,
      columns: [
        { key: 'year',       label: 'Year',           type: 'number' },
        { key: 'quarter',    label: 'Quarter',        type: 'text' },
        { key: 'monthName',  label: 'Month',          type: 'text' },
        { key: 'dayOfWeek',  label: 'Day of Week',    type: 'text' },
        { key: 'date',       label: 'Date',           type: 'date',   required: true },
        { key: 'sendTime',   label: 'Send Time (ET)', type: 'text' },
        { key: 'ask',        label: 'Ask',            type: 'text' },
        { key: 'subtype',    label: 'Subtype',        type: 'text' },
        { key: 'topic',      label: 'Mailing Topic',  type: 'text' },
        { key: 'sender',     label: 'Sender',         type: 'text' },
        { key: 'subject',    label: 'Subject Line',   type: 'text' },
        { key: 'audience',   label: 'Audience',       type: 'text' },
        { key: 'recipients', label: 'Recipients',     type: 'number', required: true, aliases: ['sent', 'delivered', 'recips', 'quantity'] },
        { key: 'openRate',   label: 'Open Rate',      type: 'rate',   aliases: ['opens', 'open %', 'openrate', 'unique open rate'] },
        { key: 'clickRate',  label: 'Click Rate',     type: 'rate',   aliases: ['clicks', 'ctr', 'click %', 'click through rate'] },
        { key: 'raised',     label: 'Raised',         type: 'number', required: true, aliases: ['gross raised', 'revenue', 'total raised', 'amount raised'] },
        { key: 'donors',     label: 'Donors',         type: 'number' },
        { key: 'average',    label: 'Average',        type: 'number', aliases: ['average gift', 'avg gift'] },
        { key: 'donateRate', label: 'Donate Rate',    type: 'rate' },
        { key: 'actions',    label: 'Actions',        type: 'number' },
        { key: 'actionRate', label: 'Action Rate',    type: 'rate' },
        { key: 'unsubs',     label: 'Unsubs',         type: 'number', aliases: ['unsubscribes', 'unsub', 'opt outs', 'optouts'] },
        { key: 'unsubRate',  label: 'Unsub Rate',     type: 'rate' },
        { key: 'label',      label: 'Label',          type: 'text' }
      ]
    },
    {
      id: 'p2pSends', tab: 'P2P Statistics', hintHeaderRow: 5, keyKey: 'date', keyLabel: 'Date',
      role: 'primary', maxCol: 28,
      columns: [
        { key: 'client',                 label: 'Client',                   type: 'text' },
        { key: 'year',                   label: 'Year',                     type: 'number' },
        { key: 'quarter',                label: 'Quarter',                  type: 'text' },
        { key: 'monthName',              label: 'Month',                    type: 'text' },
        { key: 'dayOfWeek',              label: 'Day of Week',              type: 'text' },
        { key: 'date',                   label: 'Date',                     type: 'date',   required: true },
        { key: 'goal',                   label: 'Goal',                     type: 'text' },
        { key: 'topic',                  label: 'Topic',                    type: 'text' },
        { key: 'audience',               label: 'Audience',                 type: 'text' },
        { key: 'recipients',             label: 'Recipients',               type: 'number', aliases: ['sent', 'delivered', 'recips', 'quantity'] },
        { key: 'grossSpend',             label: 'Gross Spend',              type: 'number', aliases: ['spend', 'cost', 'gross cost'] },
        { key: 'immediateRaised',        label: 'Immediate Raised',         type: 'number' },
        { key: 'immediateRoas',          label: 'Immediate ROAS',           type: 'number' },
        { key: 'clickRate',              label: 'Click Rate',               type: 'rate',   aliases: ['clicks', 'ctr', 'click %'] },
        { key: 'donors',                 label: 'Donors',                   type: 'number' },
        { key: 'avgGift',                label: 'Average Gift',             type: 'number' },
        { key: 'donateRate',             label: 'Donate Rate',              type: 'rate' },
        { key: 'ntl',                    label: 'NTL',                      type: 'number', aliases: ['new to list', 'new-to-list', 'newtolist'] },
        { key: 'grossCpa',               label: 'Gross CPA',                type: 'number' },
        { key: 'ltvFromNtl',             label: 'LTV from NTL',             type: 'number' },
        { key: 'ltvRoasFromNtl',         label: 'LTV ROAS from NTL',        type: 'number' },
        { key: 'immediateRaisedFromNtl', label: 'Immediate Raised from NTL', type: 'number' },
        { key: 'unsubs',                 label: 'Unsubs',                   type: 'number', aliases: ['unsubscribes', 'unsub', 'opt outs'] },
        { key: 'unsubRate',              label: 'Unsub Rate',               type: 'rate' },
        { key: 'sourceCode',             label: 'Source Code',              type: 'text' }
      ]
    },
    {
      id: 'adsMonthly', tab: 'Ads Report - Finance-Adjusted', hintHeaderRow: 3, keyKey: 'month', keyLabel: 'Month',
      maxCol: 12,
      // Row 3 repeats "Lifetime Raised"/"Lifetime ROAS" under the merged row-1 group
      // labels ("Standard Toplines" H:I, "Finance-Adjusted" J:K). Disambiguate by
      // column position, never by name.
      forcedByLetter: {
        lifetimeRaisedStandard: 'H', lifetimeRoasStandard: 'I',
        lifetimeRaisedFinanceAdj: 'J', lifetimeRoasFinanceAdj: 'K'
      },
      columns: [
        { key: 'month',                    label: 'Month',           type: 'date', required: true },
        { key: 'placement',                label: 'Placement',       type: 'text' },
        { key: 'goal',                     label: 'Goal',            type: 'text' },
        { key: 'grossSpend',               label: 'Gross Spend',     type: 'number', aliases: ['spend', 'cost', 'gross cost'] },
        { key: 'ntl',                      label: 'NTL',             type: 'number', aliases: ['new to list', 'new-to-list'] },
        { key: 'grossCpa',                 label: 'Gross CPA',       type: 'number' },
        { key: 'lifetimeRaisedStandard',   label: 'Lifetime Raised', type: 'number' },
        { key: 'lifetimeRoasStandard',     label: 'Lifetime ROAS',   type: 'number' },
        { key: 'lifetimeRaisedFinanceAdj', label: 'Lifetime Raised', type: 'number' },
        { key: 'lifetimeRoasFinanceAdj',   label: 'Lifetime ROAS',   type: 'number' }
      ]
    },
    {
      id: 'monthlyGoals', tab: 'April 2026 Updated Goals', hintHeaderRow: 4, keyKey: 'month', keyLabel: 'Month',
      maxCol: 15, columns: MONTHLY_COLUMNS
    },
    {
      id: 'projections', tab: 'Digital Projections', hintHeaderRow: 2, keyKey: 'month', keyLabel: 'Month',
      // Three stacked scenario blocks (Low / Medium / High Investment), each with its
      // own header row and Totals row. Parsing this as one table concatenates all three.
      multiBlock: true, maxCol: 13, columns: MONTHLY_COLUMNS
    },
    {
      id: 'highDollar', tab: 'High-Dollar Donations', hintHeaderRow: 4, keyKey: 'date', keyLabel: 'date',
      maxCol: 9, pii: true,
      // Also dropped from the RAW view, not just the parsed rows — the raw table
      // must never become a back door to the identity columns.
      omitNames: ['receipt id', 'fundraising page', 'first', 'last', 'email'],
      // Identity columns are dropped inside the reader — the row object is never
      // constructed with them, so they cannot reach state, the DOM, or a console dump.
      dropKeys: { receiptId: true, fundraisingPage: true, first: true, last: true, email: true },
      columns: [
        { key: 'category',        label: 'Category',         type: 'text' },
        { key: 'receiptId',       label: 'receipt_id',        type: 'text' },
        { key: 'date',            label: 'date',              type: 'date', required: true },
        { key: 'amount',          label: 'amount',            type: 'number', required: true, aliases: ['gift amount', 'donation amount'] },
        { key: 'fundraisingPage', label: 'fundraising_page',  type: 'text' },
        { key: 'first',           label: 'first',             type: 'text' },
        { key: 'last',            label: 'last',              type: 'text' },
        { key: 'email',           label: 'email',             type: 'text' },
        { key: 'financeClaim',    label: 'Finance Claim',     type: 'text' },
        { key: 'authenticAdded',  label: 'Authentic Added',   type: 'text' }
      ]
    },
    {
      id: 'emailCalendar', tab: 'Email Sending CalendarTracker', hintHeaderRow: 3, keyKey: 'date', keyLabel: 'Date',
      maxCol: 9, role: 'ops',
      columns: [
        { key: 'date',      label: 'Date',          type: 'date', required: true },
        { key: 'sender',    label: 'Sender',        type: 'text' },
        { key: 'emailName', label: 'Email Name',    type: 'text' },
        { key: 'ask',       label: 'Ask',           type: 'text' },
        { key: 'status',    label: 'Status',        type: 'text' },
        { key: 'targeting', label: 'Targeting',     type: 'text' },
        { key: 'draftLink', label: 'Link to draft', type: 'text' },
        { key: 'notes',     label: 'Notes',         type: 'text' }
      ]
    },
    {
      id: 'p2pCalendar', tab: 'P2P Calendar', hintHeaderRow: 1, keyKey: 'date', keyLabel: 'Date',
      maxCol: 5, role: 'ops',
      columns: [
        { key: 'date',   label: 'Date',   type: 'date', required: true },
        { key: 'type',   label: 'Type',   type: 'text' },
        // C1 carries an embedded newline: "Topic \n(Link copy here)".
        { key: 'topic',  label: 'Topic (Link copy here)', type: 'text', aliases: ['topic'] },
        { key: 'status', label: 'Status', type: 'text' }
      ]
    },
    {
      id: 'partnerToolkits', tab: 'Partner Toolkits', hintHeaderRow: 1, keyKey: 'name', keyLabel: 'Name',
      maxCol: 5, role: 'reference',
      columns: [
        { key: 'name',        label: 'Name',         type: 'text', required: true },
        { key: 'office',      label: 'Office/Org',   type: 'text' },
        { key: 'link',        label: 'Link',         type: 'text' },
        { key: 'lastUpdated', label: 'Last Updated', type: 'date' }
      ]
    }
  ];

  /** Tabs deliberately not parsed, with the reason shown in the Data Sources panel. */
  var NON_DATA_TABS = {
    'Digital Report':   'Formatted cover page — no tabular data to parse.',
    'Paid Media Report': 'Metric definitions — reference text, no tabular data.'
  };

  // ---------------------------------------------------------------------------
  // High-Dollar aggregation
  //
  // Only aggregates ever leave this function. There is no drill-through to an
  // individual donation, by design.
  // ---------------------------------------------------------------------------

  var AMOUNT_BANDS = [
    { label: '$500–999',      min: 500,   max: 1000 },
    { label: '$1,000–2,499',  min: 1000,  max: 2500 },
    { label: '$2,500–4,999',  min: 2500,  max: 5000 },
    { label: '$5,000+',       min: 5000,  max: Infinity }
  ];

  function bandFor(amount) {
    for (var i = 0; i < AMOUNT_BANDS.length; i++) {
      if (amount >= AMOUNT_BANDS[i].min && amount < AMOUNT_BANDS[i].max) return AMOUNT_BANDS[i].label;
    }
    return 'under $500';
  }

  function monthKey(d) {
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0');
  }

  function aggregateHighDollar(rows) {
    function bucket() { return { count: 0, sum: 0 }; }
    var agg = { byCategory: {}, byMonth: {}, byBand: {}, byFinanceClaim: {}, byAuthenticAdded: {},
                total: bucket(), bands: AMOUNT_BANDS.map(function (b) { return b.label; }) };
    rows.forEach(function (row) {
      var amt = typeof row.amount === 'number' ? row.amount : 0;
      var cat = row.category || '(uncategorized)';
      var mk = row.date ? monthKey(row.date) : '(no date)';
      var band = bandFor(amt);
      var claim = row.financeClaim || '(blank)';
      var auth = row.authenticAdded || '(blank)';
      [[agg.byCategory, cat], [agg.byMonth, mk], [agg.byBand, band],
       [agg.byFinanceClaim, claim], [agg.byAuthenticAdded, auth]].forEach(function (pair) {
        var t = pair[0], k = pair[1];
        if (!t[k]) t[k] = bucket();
        t[k].count++; t[k].sum += amt;
      });
      agg.total.count++; agg.total.sum += amt;
    });
    return agg;
  }

  // ---------------------------------------------------------------------------
  // parseWorkbook
  // ---------------------------------------------------------------------------

  function createParser(XLSX) {
    if (!XLSX) throw new Error('CABParse requires the XLSX object');

    function parseTab(wb, spec, report) {
      var entry = {
        id: spec.id, tab: spec.tab, role: spec.role || 'data', status: 'failed',
        hintHeaderRow: spec.hintHeaderRow, headerRow: null, rowsLoaded: 0,
        dataRange: null, dateRange: null, unmatchedColumns: [], missingRequired: [],
        missingOptional: [], warnings: [], rejections: [], rowsExcluded: [], notes: [], blocks: []
      };
      report.tabs.push(entry);

      var ws = wb.Sheets[spec.tab];
      if (!ws) {
        entry.status = 'missing';
        entry.notes.push('Tab "' + spec.tab + '" was not found in this workbook. It may have been renamed or deleted.');
        return null;
      }

      var sheet = makeSheet(XLSX, ws, spec.tab);
      if (!sheet.range) {
        entry.status = 'failed';
        entry.notes.push('Tab is empty (no cell range).');
        return null;
      }
      var maxCol = Math.min(spec.maxCol != null ? spec.maxCol : sheet.range.e.c, sheet.range.e.c);
      var scanTo = Math.min(sheet.range.e.r, 14);

      var head = findHeaderRow(sheet, spec.columns, maxCol, 0, scanTo);
      if (!head) {
        // Still capture the grid: a tab the parser cannot interpret is exactly the
        // one a staffer most needs to be able to read.
        entry.raw = captureRaw(sheet, { maxCol: maxCol, omitNames: spec.omitNames });
        entry.status = 'failed';
        entry.notes.push('No header row found in the first ' + (scanTo + 1) +
          ' rows. Expected columns such as "' + spec.columns.slice(0, 3).map(function (c) { return c.label; }).join('", "') + '".');
        return null;
      }

      entry.headerRow = head.row + 1;
      if (entry.headerRow !== spec.hintHeaderRow) {
        entry.notes.push('Header row found at row ' + entry.headerRow +
          ', not the expected row ' + spec.hintHeaderRow + '. Parsed from row ' + entry.headerRow + '.');
      }

      var allRows = [];
      var blocks = [];
      var cursor = head;
      var guard = 0;

      while (cursor && guard++ < 20) {
        var match = matchColumns(cursor.cells, spec.columns, spec.forcedByLetter);
        if (match.missingRequired.length) {
          entry.raw = captureRaw(sheet, { maxCol: maxCol, headerRow: cursor.row + 1, omitNames: spec.omitNames });
          entry.status = 'failed';
          entry.missingRequired = match.missingRequired;
          entry.notes.push('Required column(s) missing: ' + match.missingRequired.join(', ') +
            '. Panels that depend on this tab cannot be computed.');
          return null;
        }

        var block = readBlock(sheet, {
          expected: spec.columns, colMap: match.map, keyKey: spec.keyKey, keyLabel: spec.keyLabel,
          startRow: cursor.row + 1, hardStop: sheet.range.e.r, maxCol: maxCol,
          multiBlock: !!spec.multiBlock, dropKeys: spec.dropKeys
        });

        var letters = {}, names = {};
        Object.keys(match.map).forEach(function (k) {
          if (spec.dropKeys && spec.dropKeys[k]) return;
          letters[k] = match.map[k].letter;
          names[k] = match.map[k].foundName;
        });

        var b = {
          title: spec.multiBlock ? blockTitle(sheet, cursor.row) : null,
          headerRow: cursor.row + 1,
          dataRange: block.dataRange,
          rows: block.rows,
          rowsAvailable: block.rows.length,
          rowsExcluded: block.rowsExcluded,
          rejections: block.rejections,
          columnLetters: letters,
          columnNames: names,
          warnings: match.warnings.slice(),
          unmatchedColumns: match.unmatched,
          missingOptional: match.missingOptional
        };
        blocks.push(b);
        allRows = allRows.concat(block.rows);

        entry.warnings = entry.warnings.concat(match.warnings);
        entry.unmatchedColumns = entry.unmatchedColumns.concat(match.unmatched);
        entry.rowsExcluded = entry.rowsExcluded.concat(block.rowsExcluded);
        entry.rejections = entry.rejections.concat(block.rejections);
        entry.missingOptional = match.missingOptional;

        if (!spec.multiBlock) break;
        var next = findHeaderRow(sheet, spec.columns, maxCol, block.stopRow, sheet.range.e.r);
        if (!next || next.row <= cursor.row) break;
        cursor = next;
      }

      entry.status = 'parsed';
      entry.raw = captureRaw(sheet, {
        maxCol: maxCol, headerRow: head.row + 1, omitNames: spec.omitNames
      });
      entry.blocks = blocks.map(function (b) {
        return { title: b.title, headerRow: b.headerRow, dataRange: b.dataRange, rows: b.rows.length };
      });
      entry.rowsLoaded = allRows.length;
      entry.dataRange = blocks.length === 1 ? blocks[0].dataRange
        : blocks.map(function (b) { return b.dataRange; }).filter(Boolean).join(', ');

      var dates = allRows.map(function (r) { return r[spec.keyKey]; })
        .filter(function (d) { return d instanceof Date; });
      if (dates.length) {
        var min = new Date(Math.min.apply(null, dates)), max = new Date(Math.max.apply(null, dates));
        entry.dateRange = min.toISOString().slice(0, 10) + ' … ' + max.toISOString().slice(0, 10);
      }
      if (entry.missingOptional && entry.missingOptional.length) {
        entry.notes.push('Optional column(s) not found: ' + entry.missingOptional.join(', ') +
          '. Only the panels that use them are affected.');
      }

      return { entry: entry, blocks: blocks, rows: allRows };
    }

    function parseWorkbook(input) {
      var wb = XLSX.read(input, { type: typeof input === 'string' ? 'binary' : 'array', cellDates: false });
      var report = { tabs: [], warnings: [], sheetNames: wb.SheetNames.slice(), parsedAt: null };
      var out = { parseReport: report };
      var byId = {};

      TAB_SPECS.forEach(function (spec) {
        var res = parseTab(wb, spec, report);
        byId[spec.id] = res;
      });

      // Tabs present in the workbook that no spec claims.
      var claimed = {};
      TAB_SPECS.forEach(function (s) { claimed[s.tab] = true; });
      wb.SheetNames.forEach(function (name) {
        if (claimed[name]) return;
        var ws = wb.Sheets[name];
        if (NON_DATA_TABS[name]) {
          report.tabs.push({ tab: name, status: 'skipped', role: 'reference', rowsLoaded: 0,
            notes: [NON_DATA_TABS[name]], warnings: [], unmatchedColumns: [], rowsExcluded: [], rejections: [], blocks: [],
            raw: ws ? captureRaw(makeSheet(XLSX, ws, name), {}) : null });
          return;
        }
        var note = 'Tab is present in the workbook but is not one the dashboard knows how to read.';
        var warn = null;
        if (ws) {
          var errCells = [];
          Object.keys(ws).forEach(function (addr) {
            if (addr.charAt(0) === '!') return;
            var e = errorText(ws[addr]);
            if (e) errCells.push(addr + ' = ' + e);
          });
          if (errCells.length) {
            warn = 'Tab "' + name + '" contains Excel error cells (' + errCells.slice(0, 3).join(', ') +
              (errCells.length > 3 ? ', …' : '') + '). Skipped — fix the formula in the workbook to use it.';
            note = warn;
          }
        }
        report.tabs.push({ tab: name, status: 'skipped', role: 'unknown', rowsLoaded: 0,
          notes: [note], warnings: warn ? [{ kind: 'sheet', text: warn }] : [],
          unmatchedColumns: [], rowsExcluded: [], rejections: [], blocks: [],
          raw: ws ? captureRaw(makeSheet(XLSX, ws, name), {}) : null });
        if (warn) report.warnings.push({ kind: 'sheet', tab: name, text: warn });
      });

      // Specs whose tab is missing entirely.
      report.tabs.forEach(function (t) {
        if (t.status === 'missing') {
          report.warnings.push({ kind: 'tab', tab: t.tab, text: 'Expected tab "' + t.tab + '" was not found. It may have been renamed.' });
        }
        (t.warnings || []).forEach(function (w) {
          if (w.kind === 'fuzzy' || w.kind === 'alias') {
            report.warnings.push({ kind: w.kind, tab: t.tab, column: w.column, text: t.tab + ': ' + w.text });
          }
        });
      });

      function tableOf(id) {
        var res = byId[id];
        if (!res) return { rows: [], available: false, entry: (report.tabs.filter(function (t) { return t.id === id; })[0] || null) };
        var b = res.blocks[0] || {};
        return {
          rows: res.rows, available: true, entry: res.entry, tab: res.entry.tab,
          headerRow: res.entry.headerRow, dataRange: res.entry.dataRange,
          rowsAvailable: res.rows.length, rowsExcluded: res.entry.rowsExcluded,
          columnLetters: b.columnLetters || {}, columnNames: b.columnNames || {},
          warnings: res.entry.warnings, rejections: res.entry.rejections
        };
      }

      out.emailSends     = tableOf('emailSends');
      out.p2pSends       = tableOf('p2pSends');
      out.adsMonthly     = tableOf('adsMonthly');
      out.monthlyGoals   = tableOf('monthlyGoals');
      out.emailCalendar  = tableOf('emailCalendar');
      out.p2pCalendar    = tableOf('p2pCalendar');
      out.partnerToolkits = tableOf('partnerToolkits');

      // Projections: one entry per stacked scenario block.
      var proj = byId['projections'];
      out.projectionScenarios = [];
      if (proj) {
        proj.blocks.forEach(function (b) {
          var title = b.title || 'Block at row ' + b.headerRow;
          var nonZero = b.rows.filter(function (r) {
            return CHANNEL_KEYS.some(function (k) { return typeof r[k] === 'number' && r[k] !== 0; });
          }).length;
          var months = b.rows.map(function (r) { return r.month; }).filter(function (d) { return d instanceof Date; });
          // Month columns that advance by a day at a time are a broken fill, not months.
          var dayStepped = 0;
          for (var i = 1; i < months.length; i++) {
            if ((months[i] - months[i - 1]) / 86400000 <= 2) dayStepped++;
          }
          var usable = nonZero > 0 && dayStepped < 3;
          var reason = null;
          if (nonZero === 0) reason = 'every channel value in this block is $0.00 or an Excel error — nothing to plot';
          else if (dayStepped >= 3) reason = 'the Month column advances by single days after the first ' +
            (months.length - dayStepped) + ' rows, so these are not monthly figures';
          out.projectionScenarios.push({
            scenario: title, headerRow: b.headerRow, dataRange: b.dataRange, rows: b.rows,
            rowsAvailable: b.rows.length, rowsExcluded: b.rowsExcluded, columnLetters: b.columnLetters,
            columnNames: b.columnNames, warnings: b.warnings, tab: 'Digital Projections',
            usable: usable, unusableReason: reason
          });
          if (!usable) {
            report.warnings.push({ kind: 'data', tab: 'Digital Projections',
              text: 'Scenario "' + title + '" is not usable: ' + reason + '.' });
          }
        });
      }

      // High-Dollar: aggregates only. Raw rows are discarded here and never exposed.
      var hd = byId['highDollar'];
      out.highDollarAgg = {
        available: !!hd,
        agg: hd ? aggregateHighDollar(hd.rows) : null,
        tab: 'High-Dollar Donations',
        headerRow: hd ? hd.entry.headerRow : null,
        dataRange: hd ? hd.entry.dataRange : null,
        rowsAvailable: hd ? hd.rows.length : 0,
        rowsExcluded: hd ? hd.entry.rowsExcluded : [],
        columnLetters: hd && hd.blocks[0] ? hd.blocks[0].columnLetters : {},
        warnings: hd ? hd.entry.warnings : [],
        droppedColumns: ['receipt_id (B)', 'fundraising_page (E)', 'first (F)', 'last (G)', 'email (H)']
      };

      out.channelKeys = CHANNEL_KEYS;
      out.amountBands = AMOUNT_BANDS.map(function (b) { return b.label; });
      return out;
    }

    return { parseWorkbook: parseWorkbook };
  }

  return {
    createParser: createParser,
    // exported for tests
    _internals: { normalize: normalize, levenshtein: levenshtein, tokenSetSim: tokenSetSim,
                  coerceNumber: coerceNumber, coerceDate: coerceDate, coerceRate: coerceRate,
                  CHANNEL_KEYS: CHANNEL_KEYS, AMOUNT_BANDS: AMOUNT_BANDS, TAB_SPECS: TAB_SPECS }
  };
});
