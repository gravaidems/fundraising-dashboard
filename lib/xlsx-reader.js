/* =====================================================================
   xlsx-reader.js — minimal .xlsx reader, shared by the dashboard page and
   the node build/verification scripts.

   An .xlsx file is a ZIP of XML parts. We need only the cached cell values
   (and formulas, for provenance), so this reads:
     xl/workbook.xml            sheet names, in order, with rIds
     xl/_rels/workbook.xml.rels rId -> worksheet part path
     xl/sharedStrings.xml       the string table
     xl/worksheets/sheetN.xml   the cells

   Deliberately not supported (and reported as errors rather than guessed):
   encrypted workbooks, ZIP64 archives, and compression methods other than
   stored (0) and deflate (8). Dates are returned as their raw serial numbers,
   which is fine here because no figure on the dashboard is a date.

   Works in the browser (DecompressionStream) and in node (zlib).
   ===================================================================== */
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.XlsxReader = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  /* ---------- raw inflate, whichever runtime we are in ----------
     An override can be supplied via setInflate(fn) for environments that have
     neither DecompressionStream nor node's zlib (notably the jsdom-based
     tests, which exercise the same code path the browser takes). */
  let inflateOverride = null;
  function setInflate(fn) { inflateOverride = fn || null; }

  async function inflateRaw(bytes) {
    if (inflateOverride) return await inflateOverride(bytes);
    if (typeof DecompressionStream === "function") {
      const ds = new DecompressionStream("deflate-raw");
      const stream = new Blob([bytes]).stream().pipeThrough(ds);
      const buf = await new Response(stream).arrayBuffer();
      return new Uint8Array(buf);
    }
    // node
    const zlib = require("zlib");
    return new Uint8Array(zlib.inflateRawSync(Buffer.from(bytes)));
  }

  function hasInflate() {
    if (inflateOverride) return true;
    if (typeof DecompressionStream === "function") return true;
    try { require("zlib"); return true; } catch (e) { return false; }
  }

  /* ---------- ZIP parsing ---------- */
  const SIG_EOCD = 0x06054b50, SIG_CD = 0x02014b50, SIG_LFH = 0x04034b50;

  function findEOCD(dv, len) {
    // EOCD is at most 22 bytes plus a comment up to 65535
    const start = Math.max(0, len - 22 - 65535);
    for (let i = len - 22; i >= start; i--) {
      if (dv.getUint32(i, true) === SIG_EOCD) return i;
    }
    return -1;
  }

  const utf8 = new TextDecoder("utf-8");

  function readEntries(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const eocd = findEOCD(dv, u8.byteLength);
    if (eocd < 0) throw new Error("Not a .xlsx file: no ZIP end-of-directory record found.");
    let count = dv.getUint16(eocd + 10, true);
    let cdOffset = dv.getUint32(eocd + 16, true);
    const cdSize = dv.getUint32(eocd + 12, true);
    if (count === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) {
      throw new Error("This workbook uses the ZIP64 format, which this reader does not support.");
    }
    const entries = {};
    let p = cdOffset;
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== SIG_CD) break;
      const method = dv.getUint16(p + 10, true);
      const compSize = dv.getUint32(p + 20, true);
      const uncompSize = dv.getUint32(p + 24, true);
      const nameLen = dv.getUint16(p + 28, true);
      const extraLen = dv.getUint16(p + 30, true);
      const commentLen = dv.getUint16(p + 32, true);
      const localOff = dv.getUint32(p + 42, true);
      const name = utf8.decode(u8.subarray(p + 46, p + 46 + nameLen));
      entries[name] = {method: method, compSize: compSize, uncompSize: uncompSize, localOff: localOff};
      p += 46 + nameLen + extraLen + commentLen;
    }
    return {u8: u8, dv: dv, entries: entries};
  }

  async function readPart(zip, name) {
    const e = zip.entries[name];
    if (!e) return null;
    const dv = zip.dv;
    if (dv.getUint32(e.localOff, true) !== SIG_LFH) {
      throw new Error("Corrupt archive: bad local header for " + name);
    }
    const nameLen = dv.getUint16(e.localOff + 26, true);
    const extraLen = dv.getUint16(e.localOff + 28, true);
    const dataStart = e.localOff + 30 + nameLen + extraLen;
    const raw = zip.u8.subarray(dataStart, dataStart + e.compSize);
    let out;
    if (e.method === 0) out = raw;
    else if (e.method === 8) out = await inflateRaw(raw);
    else throw new Error("Unsupported compression (method " + e.method + ") for " + name);
    return utf8.decode(out);
  }

  /* ---------- XML helpers (regex based: these parts are machine generated) ---------- */
  const ENT = {amp: "&", lt: "<", gt: ">", quot: '"', apos: "'"};
  function unesc(s) {
    if (s.indexOf("&") < 0) return s;
    return s.replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, function (m, g) {
      if (g.charAt(0) === "#") {
        const cp = g.charAt(1) === "x" ? parseInt(g.slice(2), 16) : parseInt(g.slice(1), 10);
        return isNaN(cp) ? m : String.fromCodePoint(cp);
      }
      return ENT[g] != null ? ENT[g] : m;
    });
  }
  const attr = (tag, name) => {
    const m = tag.match(new RegExp("\\b" + name + '\\s*=\\s*"([^"]*)"'));
    return m ? unesc(m[1]) : null;
  };

  function parseSharedStrings(xml) {
    if (!xml) return [];
    const out = [];
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g;
    let m;
    while ((m = siRe.exec(xml))) {
      const inner = m[1] || "";
      // concatenate every <t> so rich-text runs join into one string
      let s = "", tRe = /<t\b[^>]*>([\s\S]*?)<\/t>|<t\b[^>]*\/>/g, t;
      while ((t = tRe.exec(inner))) s += unesc(t[1] || "");
      out.push(s);
    }
    return out;
  }

  /* ---------- shared formulas ----------
     XLSX stores a "shared" formula once, on a master cell, and dependent cells
     reference it by index with no text of their own. To report a dependent's
     formula honestly we must translate the master's relative references by the
     row/column offset — reusing the master text verbatim would attribute the
     wrong cells to the wrong row. */
  function parseRefAddr(ref) {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    return m ? {col: colToNum(m[1]), row: parseInt(m[2], 10)} : null;
  }
  // A1-style reference, not preceded/followed by characters that would make it
  // part of a function name (LOG10) or a longer token.
  const REF_RE = /(?<![A-Z0-9_$.])(\$?)([A-Z]{1,3})(\$?)([1-9]\d{0,6})(?![\d(A-Z_])/g;

  function translateFormula(text, fromRef, toRef) {
    const f = parseRefAddr(fromRef), t = parseRefAddr(toRef);
    if (!f || !t) return text;
    const dR = t.row - f.row, dC = t.col - f.col;
    if (!dR && !dC) return text;
    // never rewrite inside double-quoted string literals
    return text.split(/("(?:[^"]|"")*")/).map(function (seg, i) {
      if (i % 2) return seg;
      return seg.replace(REF_RE, function (m0, aC, col, aR, row) {
        const c = aC ? col : numToCol(Math.max(1, colToNum(col) + dC));
        const r = aR ? row : String(Math.max(1, parseInt(row, 10) + dR));
        return aC + c + aR + r;
      });
    }).join("");
  }

  // Cell values are returned as: number, string, boolean, or null.
  function parseSheet(xml, shared) {
    const cells = {};
    if (!xml) return cells;
    const sharedF = {};          // si -> {text, master}
    const pendingShared = [];    // cells awaiting their master
    const cRe = /<c\b([^>]*?)\/>|<c\b([^>]*?)>([\s\S]*?)<\/c>/g;
    let m;
    while ((m = cRe.exec(xml))) {
      const tag = m[1] != null ? m[1] : m[2];
      const inner = m[3] || "";
      const ref = attr(tag, "r");
      if (!ref) continue;
      const t = attr(tag, "t");
      let fm = null;
      const fOpen = inner.match(/<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/);
      if (fOpen) {
        const fAttrs = fOpen[1] || "";
        const fText = fOpen[2] != null ? unesc(fOpen[2]) : "";
        const fType = attr(fAttrs, "t");
        const si = attr(fAttrs, "si");
        if (fType === "shared" && si != null) {
          if (fText) { sharedF[si] = {text: fText, master: ref}; fm = "=" + fText; }
          else { pendingShared.push({ref: ref, si: si}); fm = null; }
        } else if (fText) {
          fm = "=" + fText;
        } else {
          fm = null;   // an empty <f/> tells us nothing; do not claim "="
        }
      }
      let v = null;
      if (t === "inlineStr") {
        let s = "", tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g, x;
        while ((x = tRe.exec(inner))) s += unesc(x[1]);
        v = s;
      } else {
        const vMatch = inner.match(/<v\b[^>]*>([\s\S]*?)<\/v>/);
        const raw = vMatch ? unesc(vMatch[1]) : null;
        if (raw == null) v = null;
        else if (t === "s") { const i = parseInt(raw, 10); v = shared[i] != null ? shared[i] : null; }
        else if (t === "str" || t === "e") v = raw;
        else if (t === "b") v = raw === "1";
        else { const n = parseFloat(raw); v = isNaN(n) ? raw : n; }
      }
      if (v !== null || fm !== null) cells[ref] = {v: v, f: fm};
    }
    // resolve shared-formula dependents now that every master has been seen
    pendingShared.forEach(function (p) {
      const src = sharedF[p.si];
      if (!src) return;
      const translated = "=" + translateFormula(src.text, src.master, p.ref);
      if (cells[p.ref]) cells[p.ref].f = translated;
      else cells[p.ref] = {v: null, f: translated};
    });
    return cells;
  }

  /* ---------- public API ---------- */

  /* Returns {sheetNames:[...], sheets:{name: Sheet}} where Sheet wraps a cell map. */
  async function read(arrayBuffer) {
    if (!hasInflate()) {
      throw new Error("This browser cannot decompress .xlsx files. " +
        "Chrome 103+, Safari 16.4+, or Firefox 113+ is required.");
    }
    const u8 = arrayBuffer instanceof Uint8Array ? arrayBuffer : new Uint8Array(arrayBuffer);
    if (u8.byteLength < 4 || u8[0] !== 0x50 || u8[1] !== 0x4b) {
      throw new Error("That does not look like an .xlsx file (no ZIP signature). " +
        "If it is an .xls or .csv, re-save it as .xlsx first.");
    }
    const zip = readEntries(u8);
    if (zip.entries["EncryptedPackage"]) {
      throw new Error("This workbook is password protected. Remove the protection and try again.");
    }
    const wbXml = await readPart(zip, "xl/workbook.xml");
    if (!wbXml) throw new Error("Not a valid .xlsx workbook: xl/workbook.xml is missing.");
    const relsXml = await readPart(zip, "xl/_rels/workbook.xml.rels");

    // rId -> part path
    const rels = {};
    if (relsXml) {
      const re = /<Relationship\b([^>]*)\/?>/g;
      let m;
      while ((m = re.exec(relsXml))) {
        const id = attr(m[1], "Id"), target = attr(m[1], "Target");
        if (id && target) rels[id] = target.replace(/^\/?xl\//, "").replace(/^\//, "");
      }
    }

    const shared = parseSharedStrings(await readPart(zip, "xl/sharedStrings.xml"));

    const sheetNames = [], sheetPaths = {};
    const sRe = /<sheet\b([^>]*)\/?>/g;
    let sm, idx = 0;
    while ((sm = sRe.exec(wbXml))) {
      const name = attr(sm[1], "name");
      if (!name) continue;
      idx++;
      const rid = attr(sm[1], "r:id") || attr(sm[1], "id");
      let path = rid && rels[rid] ? rels[rid] : "worksheets/sheet" + idx + ".xml";
      sheetNames.push(name);
      sheetPaths[name] = "xl/" + path.replace(/^xl\//, "");
    }

    const sheets = {};
    for (const name of sheetNames) {
      let xml = await readPart(zip, sheetPaths[name]);
      if (xml == null) xml = await readPart(zip, sheetPaths[name].replace("xl/", ""));
      sheets[name] = makeSheet(parseSheet(xml, shared), name);
    }
    return {sheetNames: sheetNames, sheets: sheets};
  }

  /* ---------- column letter helpers ---------- */
  function colToNum(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n;
  }
  function numToCol(n) {
    let s = "";
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /* A thin accessor over the cell map, mirroring the openpyxl calls we make. */
  function makeSheet(cells, name) {
    let maxRow = 0, maxCol = 0;
    for (const ref in cells) {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      const r = parseInt(m[2], 10), c = colToNum(m[1]);
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    }
    return {
      name: name,
      cells: cells,
      maxRow: maxRow,
      maxCol: maxCol,
      raw: ref => cells[ref] || null,
      /* cached value at a reference, or null */
      val: function (ref) { const c = cells[ref]; return c ? c.v : null; },
      /* formula string at a reference, or null */
      formula: function (ref) { const c = cells[ref]; return c && c.f ? c.f : null; },
      at: function (col, row) { return this.val(col + row); },
      /* trimmed string, or null for blanks */
      txt: function (ref) {
        const v = this.val(ref);
        if (v === null || v === undefined) return null;
        const s = String(v).trim();
        return s === "" ? null : s;
      },
      /* number, or null for blanks, '-' and error sentinels */
      num: function (ref) {
        const v = this.val(ref);
        if (v === null || v === undefined || typeof v === "boolean") return null;
        if (typeof v === "number") return v;
        const s = String(v).trim().replace(/[,$%]/g, "");
        if (s === "" || s === "-" || s.charAt(0) === "#") return null;
        const n = parseFloat(s);
        return isNaN(n) ? null : n;
      },
      bool: function (ref) {
        const v = this.val(ref);
        if (typeof v === "boolean") return v;
        if (typeof v === "number") return v !== 0;
        if (typeof v === "string") return v.trim().toUpperCase() === "TRUE";
        return false;
      }
    };
  }

  return {
    read: read,
    setInflate: setInflate,
    colToNum: colToNum,
    numToCol: numToCol,
    hasInflate: hasInflate,
    _internals: {parseSheet: parseSheet, parseSharedStrings: parseSharedStrings,
                 unesc: unesc, translateFormula: translateFormula}
  };
});
