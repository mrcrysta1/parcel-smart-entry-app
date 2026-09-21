/* Parcel Smart Entry - workbook engine core
 *
 * ONE implementation, three callers:
 *   - docs/engine.js        browser, offline single-user build (generated)
 *   - netlify/functions/*   server, shared multi-user build
 *   - tests
 *
 * Everything that describes the proforma comes from shared/schema.json, which
 * app/routes.py also reads, so the Python and JavaScript sides cannot drift.
 *
 * Written as a UMD factory so it loads unchanged in Node (require/import) and
 * in a plain browser <script>. ExcelJS is injected rather than imported: the
 * browser gets it from a CDN global, Node from node_modules.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ParcelEngineFactory = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  return function createEngine(opts) {
    var S = opts.schema;
    var ExcelJS = opts.ExcelJS;

    var FIELDS = S.fields;
    var SEARCH_FIELDS = S.search_fields;
    var REQUIRED = S.required_fields;
    var LABELS = S.labels;
    var FLOAT_FIELDS = S.float_fields;
    var MAX_RESULTS = S.max_results || 50;

    /* ---------------------------------------------------------- helpers */
    function norm(v) {
      if (v === null || v === undefined) return '';
      return String(v).trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
    }

    // Header aliases are normalised the same way lookups are, so 'P/S' and
    // 'ps' and 'p/s' all resolve to the same key.
    var HEADER_MAP = {};
    Object.keys(S.header_map).forEach(function (k) { HEADER_MAP[norm(k)] = S.header_map[k]; });

    var PLACEHOLDERS = (S.placeholder_values || []).map(norm);

    function pad(n) { return (n < 10 ? '0' : '') + n; }

    function numText(f) {
      if (Number.isInteger(f) && Math.abs(f) < 1e15) return String(f);
      return String(parseFloat(f.toPrecision(10)));   // mirrors Python format(f, '.10g')
    }

    function cellText(v) {
      if (v === null || v === undefined) return '';
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (typeof v === 'number') return numText(v);
      if (v instanceof Date) {
        return v.getUTCFullYear() + '-' + pad(v.getUTCMonth() + 1) + '-' + pad(v.getUTCDate()) +
          ' ' + pad(v.getUTCHours()) + ':' + pad(v.getUTCMinutes()) + ':' + pad(v.getUTCSeconds());
      }
      if (typeof v === 'object') {
        if (v.richText) return v.richText.map(function (t) { return t.text; }).join('').trim();
        if (v.formula !== undefined || v.sharedFormula !== undefined) {
          return v.result === undefined || v.result === null ? '' : cellText(v.result);
        }
        if (v.text !== undefined) return String(v.text).trim();   // hyperlink
        if (v.error) return '';
      }
      return String(v).trim();
    }

    function colLetter(n) {
      var s = '';
      while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
      return s;
    }

    function err(message, status, data) {
      var e = new Error(message);
      e.status = status || 400;
      e.data = data || { error: message };
      return e;
    }

    /* ------------------------------------------------- header detection */
    function sheetOf(wb) {
      var ws = wb.worksheets && wb.worksheets[0];
      if (!ws) throw err('That workbook has no sheets.');
      return ws;
    }

    function mapRow(ws, r) {
      var cols = {}, row = ws.getRow(r);
      var width = Math.max(ws.columnCount || 0, row.cellCount || 0);
      for (var c = 1; c <= width; c++) {
        var field = HEADER_MAP[norm(cellText(row.getCell(c).value))];
        if (field && !(field in cols)) cols[field] = c;
      }
      return cols;
    }

    function findHeader(ws) {
      var bestScore = 0, bestRow = null, bestCols = {};
      var scan = Math.min(ws.rowCount || 1, 40);
      for (var r = 1; r <= scan; r++) {
        var cols = mapRow(ws, r);
        if ('parcel' in cols && ('name' in cols || 'cnic' in cols || 'house_code' in cols)) {
          return { row: r, cols: cols };
        }
        var score = Object.keys(cols).length + ('parcel' in cols ? 5 : 0);
        if (score > bestScore) { bestScore = score; bestRow = r; bestCols = cols; }
      }
      if (bestRow) return { row: bestRow, cols: bestCols };
      return { row: 4, cols: mapRow(ws, 4) };
    }

    function lastDataRow(ws, hr, cols) {
      var check = SEARCH_FIELDS.filter(function (f) { return f in cols; })
        .map(function (f) { return cols[f]; });
      if (!check.length) check = Object.keys(cols).map(function (f) { return cols[f]; });
      if (!check.length) return hr;
      for (var r = (ws.rowCount || hr); r > hr; r--) {
        var row = ws.getRow(r);
        for (var i = 0; i < check.length; i++) {
          if (cellText(row.getCell(check[i]).value) !== '') return r;
        }
      }
      return hr;
    }

    /* ------------------------------------------------------- record I/O */
    function rowToRecord(ws, r, cols) {
      var rec = {}, row = ws.getRow(r);
      FIELDS.forEach(function (f) {
        rec[f] = (f in cols) ? cellText(row.getCell(cols[f]).value) : '';
      });
      return rec;
    }

    function readRecords(ws, hr, cols) {
      var out = [], end = lastDataRow(ws, hr, cols);
      for (var r = hr + 1; r <= end; r++) {
        var rec = rowToRecord(ws, r, cols);
        if (!SEARCH_FIELDS.some(function (f) { return rec[f]; })) continue;
        rec.row = r;
        out.push(rec);
      }
      return out;
    }

    function applyDefaults(raw) {
      var rec = {};
      FIELDS.forEach(function (f) {
        var v = raw ? raw[f] : '';
        rec[f] = (v === null || v === undefined) ? '' : String(v).trim();
      });
      if (!rec.parcel || !rec.name) return rec;
      FIELDS.forEach(function (f) {
        if (f !== 'parcel' && f !== 'name' && f !== 'property_type' && !rec[f]) rec[f] = 'Nill';
      });
      if (!rec.property_type) rec.property_type = 'Other';
      return rec;
    }

    function coerce(field, text) {
      var s = (text === null || text === undefined) ? '' : String(text).trim();
      if (s === '') return null;
      if (FLOAT_FIELDS.indexOf(field) >= 0) {
        var f = Number(s);
        return isNaN(f) ? s : f;
      }
      if ((field === 'parcel' || field === 'sr') && /^[1-9]\d{0,8}$/.test(s)) return parseInt(s, 10);
      return s;
    }

    function missingRequired(rec) {
      return REQUIRED.filter(function (f) { return !rec[f]; }).map(function (f) { return LABELS[f]; });
    }

    function searchRecords(records, q, limit) {
      var qn = norm(q);
      if (!qn) return [];
      var scored = [];
      records.forEach(function (rec) {
        var best = null;
        SEARCH_FIELDS.forEach(function (f) {
          var v = norm(rec[f]);
          if (!v) return;
          var s = v === qn ? 0 : (v.indexOf(qn) === 0 ? 1 : (v.indexOf(qn) > 0 ? 2 : -1));
          if (s < 0) return;
          if (best === null || s < best) best = s;
        });
        if (best !== null) scored.push([best, rec.row, rec]);
      });
      scored.sort(function (a, b) { return a[0] - b[0] || a[1] - b[1]; });
      return scored.slice(0, limit || MAX_RESULTS).map(function (t) { return t[2]; });
    }

    /* Distinct values already used in a column, most common first. Filler
     * such as 'Nill' is never offered. Mirrors field_options() in routes.py. */
    function fieldOptions(records, limit) {
      limit = limit || 60;
      var out = {};
      (S.option_fields || []).forEach(function (f) {
        var counts = {}, order = [];
        records.forEach(function (r) {
          var v = String(r[f] === undefined || r[f] === null ? '' : r[f]).trim();
          if (!v || PLACEHOLDERS.indexOf(norm(v)) >= 0) return;
          if (!(v in counts)) { counts[v] = 0; order.push(v); }
          counts[v]++;
        });
        order.sort(function (a, b) {
          return counts[b] - counts[a] || a.toLowerCase().localeCompare(b.toLowerCase());
        });
        out[f] = order.slice(0, limit);
      });
      return out;
    }

    /* Longest run of leading digits shared by every all-digit value. Mirrors
     * common_prefix() in routes.py. */
    function commonPrefix(values) {
      var seen = {}, vals = [];
      values.forEach(function (v) {
        var s = String(v === undefined || v === null ? '' : v).trim();
        if (!/^\d+$/.test(s) || seen[s]) return;
        seen[s] = 1; vals.push(s);
      });
      if (vals.length < 2) return '';
      vals.sort();
      var first = vals[0], last = vals[vals.length - 1], i = 0;
      while (i < Math.min(first.length, last.length) && first[i] === last[i]) i++;
      return first.slice(0, i);
    }

    function fieldPrefixes(records) {
      var out = {};
      (S.prefix_fields || []).forEach(function (f) {
        out[f] = commonPrefix(records.map(function (r) { return r[f]; }));
      });
      return out;
    }

    /* --------------------------------------------------------- workbooks */
    function ensureTemplate() {
      var wb = new ExcelJS.Workbook();
      var ws = wb.addWorksheet('Sheet1');
      ws.getCell('A1').value = 'Proforma for parcel mapping';
      ws.getCell('A1').font = { bold: true, size: 13 };
      ws.getCell('A2').value = 'District name:';
      ws.getCell('C2').value = 'District ID:';
      ws.getCell('E2').value = 'Tehsil name:';
      ws.getCell('G2').value = 'Tehsil ID:';
      ws.getCell('I2').value = 'UC Name:';
      ws.getCell('L2').value = 'UC ID:';
      S.template_headers.forEach(function (h, i) {
        var c = ws.getRow(4).getCell(i + 1);
        c.value = h;
        c.font = { bold: true };
        c.alignment = { wrapText: true, vertical: 'middle' };
        ws.getColumn(i + 1).width = (i + 1) === 11 ? 32 : 16;
      });
      ws.views = [{ state: 'frozen', ySplit: 4 }];
      return wb;
    }

    function b64ToBytes(b64) {
      if (typeof Buffer !== 'undefined') return new Uint8Array(Buffer.from(b64, 'base64'));
      var bin = atob(b64), out = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }

    function bytesToB64(buf) {
      var bytes = new Uint8Array(buf);
      if (typeof Buffer !== 'undefined') return Buffer.from(bytes).toString('base64');
      var bin = '';
      for (var i = 0; i < bytes.length; i += 0x8000) {
        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
      }
      return btoa(bin);
    }

    // ExcelJS cannot round-trip VBA, so .xlsm input is flagged and the user
    // warned that macros will be dropped.
    function hasVba(bytes) {
      var needle = 'vbaProject.bin', limit = bytes.length - needle.length;
      for (var i = 0; i < limit; i++) {
        if (bytes[i] === 118 /* v */) {
          var hit = true;
          for (var j = 1; j < needle.length; j++) {
            if (bytes[i + j] !== needle.charCodeAt(j)) { hit = false; break; }
          }
          if (hit) return true;
        }
      }
      return false;
    }

    async function openBytes(bytes) {
      var wb = new ExcelJS.Workbook();
      try {
        await wb.xlsx.load(bytes.buffer ? bytes.buffer : bytes);
      } catch (e) {
        throw err('Could not read that file: ' + (e && e.message ? e.message : e));
      }
      return wb;
    }

    /* Parse an uploaded workbook into the shared model. */
    async function parseWorkbook(bytes) {
      var wb = await openBytes(bytes);
      var ws = sheetOf(wb);
      var h = findHeader(ws);
      if (!('parcel' in h.cols)) {
        throw err('No Parcel / P-S column found in this sheet. ' +
                  'Check that the proforma header row is present.');
      }
      var records = readRecords(ws, h.row, h.cols);
      return {
        sheetName: ws.name,
        headerRow: h.row,
        cols: h.cols,
        records: records,
        nextRow: lastDataRow(ws, h.row, h.cols) + 1,
        macro: hasVba(bytes),
        missing: SEARCH_FIELDS.filter(function (f) { return !(f in h.cols); })
          .map(function (f) { return LABELS[f]; })
      };
    }

    /* Write records into a workbook and highlight the rows given. */
    function writeRows(ws, cols, records, highlightRows) {
      var highlight = {};
      (highlightRows || []).forEach(function (r) { highlight[r] = 1; });
      var maxCol = Math.max.apply(null, Object.keys(cols).map(function (f) { return cols[f]; }));
      records.forEach(function (rec) {
        var target = ws.getRow(rec.row);
        FIELDS.forEach(function (f) {
          if (f in cols) target.getCell(cols[f]).value = coerce(f, rec[f]);
        });
        if ('sr' in cols && !cellText(target.getCell(cols.sr).value)) {
          target.getCell(cols.sr).value = rec.row - (rec._headerRow || 0);
        }
        if (highlight[rec.row]) {
          for (var c = 1; c <= maxCol; c++) {
            target.getCell(c).fill = {
              type: 'pattern', pattern: 'solid', fgColor: { argb: S.highlight_argb }
            };
          }
        }
        target.commit();
      });
    }

    /* Rebuild a downloadable workbook: the stored original (so formatting,
     * headers and column widths survive) with every current record written
     * back into its row. */
    async function buildWorkbook(originalBytes, sheet) {
      var wb = originalBytes ? await openBytes(originalBytes) : ensureTemplate();
      var ws = sheetOf(wb);
      var h = originalBytes ? findHeader(ws) : { row: 4, cols: mapRow(ws, 4) };
      var cols = h.cols;
      if (!('parcel' in cols)) throw err('No Parcel / P-S column found in the stored file.');
      var recs = sheet.records.map(function (r) {
        var c = Object.assign({}, r);
        c._headerRow = h.row;
        return c;
      });
      writeRows(ws, cols, recs, sheet.records
        .filter(function (r) { return r.touched; })
        .map(function (r) { return r.row; }));
      return await wb.xlsx.writeBuffer();
    }

    /* ------------------------------------------- single-user browser API */
    /* Used by the offline GitHub Pages build, where the whole workbook lives
     * in the tab and travels as base64. */
    async function load(b64) {
      if (!b64) throw err('No workbook supplied.');
      var bytes = b64ToBytes(b64);
      var parsed = await parseWorkbook(bytes);
      var columns = {};
      Object.keys(parsed.cols).forEach(function (f) { columns[f] = colLetter(parsed.cols[f]); });
      return {
        records: parsed.records,
        options: fieldOptions(parsed.records),
        prefixes: fieldPrefixes(parsed.records),
        header_row: parsed.headerRow,
        sheet: parsed.sheetName,
        columns: columns,
        missing: parsed.missing,
        macro: false,   // ExcelJS always writes .xlsx; see warning below
        warning: parsed.macro
          ? 'This is a macro-enabled file. The online version cannot keep macros - '
            + 'it will download as .xlsx. Run the Flask app locally to preserve them.'
          : '',
        downloadName: 'Parcel_Mapping_Updated.xlsx',
        message: parsed.records.length + ' record(s) loaded'
      };
    }

    async function save(b64, rawRecord, row) {
      var rec = applyDefaults(rawRecord);
      var updating = row !== null && row !== undefined;

      var missing = missingRequired(rec);
      if (missing.length) {
        throw err(missing.join(' and ') + ' ' + (missing.length > 1 ? 'are' : 'is') + ' required.');
      }

      var bytes = b64 ? b64ToBytes(b64) : null;
      var wb = bytes ? await openBytes(bytes) : ensureTemplate();
      var ws = sheetOf(wb);
      var h = findHeader(ws);
      var cols = h.cols, hr = h.row;
      if (!('parcel' in cols)) throw err('No Parcel / P-S column found in this sheet.');

      if (updating) {
        row = parseInt(row, 10);
        if (isNaN(row) || row <= hr) throw err('Invalid row reference.');
      } else {
        var all = readRecords(ws, hr, cols), existing = null;
        for (var i = 0; i < all.length; i++) {
          if (norm(all[i].parcel) === norm(rec.parcel)) { existing = all[i]; break; }
        }
        if (existing) {
          var only = {};
          FIELDS.forEach(function (f) { only[f] = existing[f]; });
          var msg = 'P/S ' + rec.parcel + ' already exists at row ' + existing.row +
                    '. Open it to update instead of adding a duplicate.';
          throw err(msg, 409, { error: msg, duplicate: true, row: existing.row, record: only });
        }
        row = lastDataRow(ws, hr, cols) + 1;
      }

      var one = Object.assign({}, rec, { row: row, _headerRow: hr });
      writeRows(ws, cols, [one], [row]);

      var saved = {};
      FIELDS.forEach(function (f) { saved[f] = rec[f]; });
      saved.row = row;

      return {
        workbook: bytesToB64(await wb.xlsx.writeBuffer()),
        record: saved,
        row: row,
        downloadName: 'Parcel_Mapping_Updated.xlsx',
        macro: false,
        action: updating ? 'updated' : 'added'
      };
    }

    async function download(b64) {
      if (!b64) throw err('Nothing to download yet - load a file or save a record first.');
      return new Blob([b64ToBytes(b64)], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
    }

    return {
      // pure helpers
      norm: norm, cellText: cellText, colLetter: colLetter, err: err,
      applyDefaults: applyDefaults, coerce: coerce, missingRequired: missingRequired,
      searchRecords: searchRecords, fieldOptions: fieldOptions,
      commonPrefix: commonPrefix, fieldPrefixes: fieldPrefixes,
      // workbook level
      parseWorkbook: parseWorkbook, buildWorkbook: buildWorkbook,
      ensureTemplate: ensureTemplate, findHeader: findHeader, readRecords: readRecords,
      b64ToBytes: b64ToBytes, bytesToB64: bytesToB64, hasVba: hasVba,
      // single-user browser API
      load: load, save: save, download: download,
      FIELDS: FIELDS, SEARCH_FIELDS: SEARCH_FIELDS
    };
  };
}));
