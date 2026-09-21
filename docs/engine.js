/* Parcel Smart Entry - browser workbook engine
 *
 * A port of app/routes.py to the browser, backed by ExcelJS, so the app can be
 * hosted as a static site (GitHub Pages) with no Python backend. It exposes the
 * same call shapes as the Flask API, so app.js is identical in both modes:
 *
 *   ParcelEngine.load(b64)               -> {records, header_row, ...}
 *   ParcelEngine.save(b64, record, row)  -> {workbook, record, row, action}
 *   ParcelEngine.download(b64, macro)    -> Blob
 *
 * Field names, header aliases and the proforma layout all come from
 * schema.js, which build_static.py compiles from shared/schema.json - the
 * same file app/routes.py reads. The two implementations cannot drift.
 */
'use strict';

window.ParcelEngine = (function () {
  var S = window.PARCEL_SCHEMA;
  var FIELDS = S.fields;
  var SEARCH_FIELDS = S.search_fields;
  var REQUIRED = S.required_fields;
  var LABELS = S.labels;
  var FLOAT_FIELDS = S.float_fields;

  /* ------------------------------------------------------------ helpers */
  function norm(v) {
    if (v === null || v === undefined) return '';
    return String(v).trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
  }

  // Header aliases are normalised the same way lookups are (see routes.py).
  var HEADER_MAP = {};
  Object.keys(S.header_map).forEach(function (k) { HEADER_MAP[norm(k)] = S.header_map[k]; });

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

  function b64ToBytes(b64) {
    var bin = atob(b64), out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function bytesToB64(buf) {
    var bytes = new Uint8Array(buf), bin = '';
    for (var i = 0; i < bytes.length; i += 0x8000) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(bin);
  }

  // ExcelJS cannot round-trip VBA, so .xlsm files are detected up front and
  // the user is warned that macros will be dropped.
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

  /* -------------------------------------------------- header detection */
  function sheetOf(wb) {
    var ws = wb.worksheets && wb.worksheets[0];
    if (!ws) throw err('That workbook has no sheets.');
    return ws;
  }

  function mapRow(ws, r) {
    var cols = {}, row = ws.getRow(r), width = Math.max(ws.columnCount || 0, row.cellCount || 0);
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

  /* ---------------------------------------------------------- records */
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
      var any = SEARCH_FIELDS.some(function (f) { return rec[f]; });
      if (!any) continue;
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

  /* -------------------------------------------------------- workbooks */
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

  async function open(b64) {
    if (!b64) return { wb: ensureTemplate(), macro: false };
    var bytes = b64ToBytes(b64);
    var wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(bytes.buffer);
    } catch (e) {
      throw err('Could not read that file: ' + (e && e.message ? e.message : e));
    }
    return { wb: wb, macro: hasVba(bytes) };
  }

  async function toB64(wb) {
    return bytesToB64(await wb.xlsx.writeBuffer());
  }

  /* --------------------------------------------------------- the API */
  async function load(b64) {
    if (!b64) throw err('No workbook supplied.');
    var opened = await open(b64);
    var ws = sheetOf(opened.wb);
    var h = findHeader(ws);
    if (!('parcel' in h.cols)) {
      throw err('No Parcel / P-S column found in this sheet. ' +
                'Check that the proforma header row is present.');
    }
    var records = readRecords(ws, h.row, h.cols);
    var columns = {};
    Object.keys(h.cols).forEach(function (f) { columns[f] = colLetter(h.cols[f]); });
    return {
      records: records,
      header_row: h.row,
      sheet: ws.name,
      columns: columns,
      missing: SEARCH_FIELDS.filter(function (f) { return !(f in h.cols); })
        .map(function (f) { return LABELS[f]; }),
      // Always false: ExcelJS rewrites the file as .xlsx, so the download must
      // not claim to be .xlsm. The warning below tells the user why.
      macro: false,
      warning: opened.macro
        ? 'This is a macro-enabled file. The online version cannot keep macros - '
          + 'it will download as .xlsx. Run the Flask app locally to preserve them.'
        : '',
      downloadName: 'Parcel_Mapping_Updated.xlsx',
      message: records.length + ' record(s) loaded'
    };
  }

  async function save(b64, rawRecord, row) {
    var rec = applyDefaults(rawRecord);
    var updating = row !== null && row !== undefined;

    var missing = REQUIRED.filter(function (f) { return !rec[f]; }).map(function (f) { return LABELS[f]; });
    if (missing.length) {
      throw err(missing.join(' and ') + ' ' + (missing.length > 1 ? 'are' : 'is') + ' required.');
    }

    var opened = await open(b64);
    var wb = opened.wb;
    var ws = sheetOf(wb);
    var h = findHeader(ws);
    var cols = h.cols, hr = h.row;
    if (!('parcel' in cols)) throw err('No Parcel / P-S column found in this sheet.');

    if (updating) {
      row = parseInt(row, 10);
      if (isNaN(row) || row <= hr) throw err('Invalid row reference.');
    } else {
      var existing = null, all = readRecords(ws, hr, cols);
      for (var i = 0; i < all.length; i++) {
        if (norm(all[i].parcel) === norm(rec.parcel)) { existing = all[i]; break; }
      }
      if (existing) {
        var rowsOnly = {};
        FIELDS.forEach(function (f) { rowsOnly[f] = existing[f]; });
        var msg = 'P/S ' + rec.parcel + ' already exists at row ' + existing.row +
                  '. Open it to update instead of adding a duplicate.';
        throw err(msg, 409, { error: msg, duplicate: true, row: existing.row, record: rowsOnly });
      }
      row = lastDataRow(ws, hr, cols) + 1;
      if ('sr' in cols && !cellText(ws.getRow(row).getCell(cols.sr).value)) {
        ws.getRow(row).getCell(cols.sr).value = row - hr;
      }
    }

    var target = ws.getRow(row);
    FIELDS.forEach(function (f) {
      if (f in cols) target.getCell(cols[f]).value = coerce(f, rec[f]);
    });

    var maxCol = Math.max.apply(null, Object.keys(cols).map(function (f) { return cols[f]; }));
    for (var c = 1; c <= maxCol; c++) {
      target.getCell(c).fill = {
        type: 'pattern', pattern: 'solid', fgColor: { argb: S.highlight_argb }
      };
    }
    target.commit();

    var saved = {};
    FIELDS.forEach(function (f) { saved[f] = rec[f]; });
    saved.row = row;

    return {
      workbook: await toB64(wb),
      record: saved,
      row: row,
      downloadName: 'Parcel_Mapping_Updated.xlsx',
      macro: false,   // ExcelJS always writes .xlsx (see load())
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
    load: load, save: save, download: download,
    // exposed for the test harness
    _internal: { norm: norm, cellText: cellText, findHeader: findHeader, coerce: coerce }
  };
})();
