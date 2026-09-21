/* Parcel Smart Entry - UI
 *
 * Runs in two modes from this one file:
 *   - server mode : talks to the Flask API (/api/load, /api/save, /api/download)
 *   - browser mode: talks to window.ParcelEngine (docs/ build, GitHub Pages)
 * Both expose the same shapes, so everything below this `api` object is shared.
 *
 * The workbook is parsed once on load and the records are cached here, so
 * searching is instant and local in both modes.
 */
'use strict';

var state = {
  workbook: '',      // base64 of the current .xlsx/.xlsm
  records: [],       // [{row, parcel, name, ...}] cache of the sheet
  row: null,         // sheet row being edited, or null for a new record
  fileName: '',
  macro: false,
  dirty: false,
  timer: null,
  results: [],
  options: {},       // {first_name: [...]} values already used in the file
  prefixes: {},      // {house_code: '318468'} shared leading digits
  // shared mode only
  user: '',          // who is entering data
  sheetId: '',       // which server-held file is open
  rev: 0,            // sheet revision the cache is current to
  baseRev: null,     // revision of the record in the form, for conflict checks
  pending: null,     // record awaiting a conflict decision
  poll: null
};

var UI = window.PARCEL_UI || {};
var MODE = UI.mode || 'local';          // 'local' = offline build, 'shared' = Netlify server
var SHARED = MODE === 'shared';
var SEARCH_FIELDS = UI.search_fields || ['parcel', 'cnic', 'house_code', 'name'];
var MAX_RESULTS = UI.max_results || 50;
var STICKY_FIELDS = UI.sticky_fields || [];
var OPTION_FIELDS = UI.option_fields || [];
var CHOICE_FIELDS = UI.choice_fields || {};
var PREFIX_FIELDS = UI.prefix_fields || [];
var OTHER = '\u0000other';   // sentinel option value, cannot collide with real data

var $ = function (id) { return document.getElementById(id); };

/* Hooks the shared (Netlify) build fills in. shared.js loads after this file
 * and replaces these declarations; in the offline build they stay no-ops so
 * the single-user path never has to know about servers or conflicts. */
function adoptSheet() {}
function refreshSheetList() {}
function showConflict(data) { toast(data && data.error ? data.error : 'Edit conflict.', 'error'); }
function hideConflict() {}
function ago() { return ''; }
function startPolling() {}
var inputs = function () { return Array.prototype.slice.call(document.querySelectorAll('[data-k]')); };

/* ------------------------------------------------------------------- api */
async function postJSON(url, body) {
  var res;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (e) {
    throw new Error('Network error - is the server still running?');
  }
  var text = await res.text();
  var data = {};
  try { data = text ? JSON.parse(text) : {}; } catch (e) {
    throw new Error('Unexpected server response (' + res.status + ').');
  }
  if (!res.ok) {
    var err = new Error(data.error || 'Request failed (' + res.status + ').');
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

var engine = function () { return window.ParcelEngine; };

/* REST helper for shared mode. Every call carries the user's name so the
 * server can record who changed what. */
async function rest(method, path, body) {
  var init = { method: method, headers: { 'x-parcel-user': state.user || 'Someone' } };
  // PARCEL_AUTH only exists in the shared build, and only matters when the
  // server has a team password set.
  var tok = (typeof PARCEL_AUTH !== 'undefined') ? PARCEL_AUTH.token() : '';
  if (tok) init.headers['x-parcel-token'] = tok;
  if (body !== undefined) {
    init.headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(Object.assign({ by: state.user }, body));
  }
  var res;
  try { res = await fetch(path, init); }
  catch (e) { throw new Error('Cannot reach the server. Check your connection.'); }
  var text = await res.text();
  var data = {};
  try { data = text ? JSON.parse(text) : {}; }
  catch (e) { throw new Error('Unexpected server response (' + res.status + ').'); }
  if (!res.ok) {
    var err = new Error(data.error || 'Request failed (' + res.status + ').');
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

var api = {
  /* local mode: parse a workbook the browser holds.
     shared mode: upload it, creating a file everyone can open. */
  load: function (b64, name) {
    if (SHARED) return rest('POST', '/api/sheets', { workbook: b64, name: name });
    return engine() ? engine().load(b64) : postJSON('/api/load', { workbook: b64 });
  },
  save: function (b64, record, row) {
    if (SHARED) {
      return row === null || row === undefined
        ? rest('POST', '/api/sheets/' + state.sheetId + '/records', { record: record })
        : rest('PUT', '/api/sheets/' + state.sheetId + '/records/' + row,
               { record: record, baseRev: state.baseRev });
    }
    return engine() ? engine().save(b64, record, row)
                    : postJSON('/api/save', { workbook: b64, record: record, row: row });
  },
  download: async function (b64, macro) {
    if (SHARED) {
      var r = await fetch('/api/sheets/' + state.sheetId + '/export');
      if (!r.ok) {
        var e = {};
        try { e = await r.json(); } catch (x) {}
        throw new Error(e.error || 'Download failed.');
      }
      return r.blob();
    }
    if (engine()) return engine().download(b64, macro);
    var res = await fetch('/api/download', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ workbook: b64, macro: macro })
    });
    if (!res.ok) {
      var d = {};
      try { d = await res.json(); } catch (e) {}
      throw new Error(d.error || 'Download failed.');
    }
    return res.blob();
  },
  listSheets: function () { return rest('GET', '/api/sheets'); },
  openSheet: function (id) { return rest('GET', '/api/sheets/' + id); },
  blankSheet: function (name) { return rest('POST', '/api/sheets', { name: name }); },
  changes: function (id, since) {
    return rest('GET', '/api/sheets/' + id + '/changes?since=' + (since || 0));
  }
};

/* ---------------------------------------------------------------- chrome */
function setStatus(t) { $('status').textContent = t; }

var toastTimer = null;
function toast(msg, kind) {
  var el = $('toast');
  el.textContent = msg;
  el.className = 'toast ' + (kind || 'info');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.className = 'toast hidden'; }, 5200);
}

function setDirty(on) {
  state.dirty = !!on;
  $('dirty').classList.toggle('hidden', !on);
}

function setBanner(text, kind) {
  var el = $('banner');
  if (!text) { el.className = 'banner hidden'; el.textContent = ''; return; }
  el.className = 'banner ' + (kind || 'info');
  el.textContent = text;
}

/* --------------------------------------------------------------- controls */
/* Some fields are plain inputs, some are dropdowns built from the values the
 * loaded file already uses. Both carry data-k, so collect()/fillForm() treat
 * them identically. */

function control(field) { return document.getElementById('f-' + field); }

function optionEl(value, text) {
  var o = document.createElement('option');
  o.value = value;
  o.textContent = text === undefined ? value : text;
  return o;
}

// A select can only hold values it lists, so a value coming from the sheet
// that is not among the choices is added rather than silently dropped.
function ensureOption(sel, value) {
  if (!value) return;
  var has = Array.prototype.some.call(sel.options, function (o) { return o.value === value; });
  if (has) return;
  var other = Array.prototype.filter.call(sel.options, function (o) { return o.value === OTHER; })[0];
  sel.insertBefore(optionEl(value), other || null);
}

function setFieldValue(el, v) {
  v = (v === undefined || v === null) ? '' : String(v);
  if (el.tagName === 'SELECT') ensureOption(el, v);
  el.value = v;
}

function attrsFor(field) {
  if (field === 'property_type') return { list: 'propertyTypes' };
  if (field === 'latitude' || field === 'longitude') return { inputmode: 'decimal' };
  if (field === 'cnic' || field === 'parcel' || field === 'house_code') return { inputmode: 'numeric' };
  return {};
}

function replaceControl(field, el) {
  var old = control(field);
  if (!old) return null;
  el.id = old.id;
  el.name = field;
  el.dataset.k = field;
  old.parentNode.replaceChild(el, old);
  return el;
}

function makeInput(field, value) {
  var el = document.createElement('input');
  el.type = 'text';
  var a = attrsFor(field);
  Object.keys(a).forEach(function (k) { el.setAttribute(k, a[k]); });
  replaceControl(field, el);
  el.value = value || '';
  return el;
}

function makeSelect(field, options, value) {
  var el = document.createElement('select');
  el.appendChild(optionEl('', '— choose —'));
  options.forEach(function (v) { el.appendChild(optionEl(v)); });
  el.appendChild(optionEl(OTHER, 'Other…'));
  replaceControl(field, el);
  setFieldValue(el, value || '');
  return el;
}

// "Other..." turns the dropdown back into a free-text box, so a name that is
// not in the file yet can still be entered.
document.addEventListener('change', function (e) {
  var el = e.target;
  if (!el || !el.dataset || !el.dataset.k || el.value !== OTHER) return;
  var field = el.dataset.k;
  makeInput(field, '').focus();
  setDirty(true);
});

/* Rebuild the enumerator dropdowns for a newly loaded file. One distinct value
 * in the column is filled in automatically; several become a picker. */
function applyFieldOptions(options) {
  state.options = options || {};
  OPTION_FIELDS.forEach(function (f) {
    var opts = state.options[f] || [];
    if (opts.length >= 2) makeSelect(f, opts, '');
    else makeInput(f, opts.length === 1 ? opts[0] : '');
  });
  // Reset fixed dropdowns so values carried over from a previous file go away.
  Object.keys(CHOICE_FIELDS).forEach(function (f) {
    var el = control(f);
    if (el && el.tagName === 'SELECT') {
      el.innerHTML = '';
      el.appendChild(optionEl('', '— choose —'));
      CHOICE_FIELDS[f].forEach(function (v) { el.appendChild(optionEl(v)); });
    }
  });
}

/* Values that carry over to the next record: the enumerator's own details,
 * plus the shared leading digits of the house code. */
function carryOver() {
  var out = {};
  STICKY_FIELDS.forEach(function (f) {
    var el = control(f);
    var v = el ? el.value : '';
    if (v && v !== OTHER) out[f] = v;
    else if ((state.options[f] || []).length === 1) out[f] = state.options[f][0];
  });
  PREFIX_FIELDS.forEach(function (f) {
    if (state.prefixes[f]) out[f] = state.prefixes[f];
  });
  return out;
}

/* ------------------------------------------------------------------ form */
function fillForm(rec) {
  rec = rec || {};
  inputs().forEach(function (i) { setFieldValue(i, rec[i.dataset.k]); });
  setDirty(false);
}

function collect() {
  var r = {};
  inputs().forEach(function (i) {
    var v = i.value === OTHER ? '' : i.value;   // sentinel is never real data
    r[i.dataset.k] = v.trim();
  });
  return r;
}

function editRecord(rec) {
  state.row = rec.row;
  state.baseRev = rec.rev === undefined ? null : rec.rev;
  hideConflict();
  fillForm(rec);
  $('mode').textContent = 'Editing row ' + rec.row + ' · P/S ' + (rec.parcel || '—');
  $('save').textContent = 'Update record';
  setBanner('Editing row ' + rec.row + '. Saving overwrites that row' +
    (SHARED && rec.updatedBy ? ' · last saved by ' + rec.updatedBy + ' ' + ago(rec.updatedAt) : '') +
    '.', 'edit');
  markActive(rec.row);
  setStatus('Record loaded');
  if (window.matchMedia('(max-width: 850px)').matches) {
    $('form').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  var first = $('f-parcel');
  if (first) first.focus({ preventScroll: true });
}

function newRecord(keep) {
  state.row = null;
  state.baseRev = null;
  hideConflict();
  fillForm(keep === undefined ? carryOver() : keep);
  $('mode').textContent = 'New record';
  $('save').textContent = 'Save record';
  setBanner(state.workbook || state.records.length
    ? 'New record — Save appends it to the bottom of the sheet.'
    : 'New file — enter P/S and Name, then Save to create the workbook.', 'new');
  markActive(null);
}

/* --------------------------------------------------------------- results */
function esc(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;').replace(/"/g, '&quot;')
    .replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function markActive(row) {
  Array.prototype.forEach.call(document.querySelectorAll('.result'), function (b) {
    b.classList.toggle('active', row !== null && +b.dataset.row === row);
  });
}

/* In shared mode a poll re-runs the search every couple of seconds. Rewriting
 * innerHTML each time would destroy and rebuild the buttons under the user's
 * finger, which can swallow a tap, so the DOM is only touched when the markup
 * actually differs. */
var lastResultsHtml = null;
function paintResults(html) {
  var el = $('results');
  if (html !== lastResultsHtml) {
    el.innerHTML = html;
    lastResultsHtml = html;
  }
}

function renderResults(matches, query) {
  state.results = matches;
  if (!query) {
    paintResults('<div class="empty">' + (state.records.length
      ? 'Type a P/S, CNIC, H/S Code or Name to search ' + state.records.length + ' record(s).'
      : 'Load an Excel file to begin searching.') + '</div>');
    return;
  }
  if (!matches.length) {
    paintResults('<div class="empty">No matching record. Fill the form and Save to add it as a new record.</div>');
    return;
  }
  paintResults(matches.map(function (m, i) {
    return '<button type="button" class="result" data-row="' + m.row + '" data-i="' + i + '">' +
      '<b>' + esc(m.parcel || '—') + ' · ' + esc(m.name || 'Unnamed') + '</b>' +
      '<span>CNIC: ' + esc(m.cnic || '—') + ' &nbsp;|&nbsp; H/S: ' + esc(m.house_code || '—') +
      ' &nbsp;|&nbsp; Row ' + m.row + '</span>' +
      '<em>Tap to edit</em></button>';
  }).join(''));
  markActive(state.row);
}

// Delegated once, on the container: renderResults() replaces the list wholesale
// on every keystroke, so per-button listeners would be discarded mid-click.
$('results').addEventListener('click', function (e) {
  var btn = e.target.closest ? e.target.closest('.result') : null;
  if (!btn) return;
  var rec = state.results[+btn.dataset.i];
  // Fall back to the row number in case the list re-rendered under the click.
  if (!rec || rec.row !== +btn.dataset.row) {
    rec = state.results.filter(function (r) { return r.row === +btn.dataset.row; })[0]
       || state.records.filter(function (r) { return r.row === +btn.dataset.row; })[0];
  }
  openResult(rec);
});

function openResult(rec) {
  if (!rec) return;
  if (state.dirty && !window.confirm('You have unsaved changes. Discard them and open this record?')) return;
  editRecord(rec);
}

/* ---------------------------------------------------------------- search */
function normalize(v) {
  return String(v === undefined || v === null ? '' : v).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function runSearch() {
  var q = $('search').value.trim();
  var qn = normalize(q);
  if (!qn) {
    renderResults([], '');
    setStatus(state.records.length ? state.records.length + ' record(s) loaded' : 'Ready');
    return;
  }
  var scored = [];
  for (var i = 0; i < state.records.length; i++) {
    var rec = state.records[i], best = null;
    for (var j = 0; j < SEARCH_FIELDS.length; j++) {
      var v = normalize(rec[SEARCH_FIELDS[j]]);
      if (!v) continue;
      var s = v === qn ? 0 : (v.indexOf(qn) === 0 ? 1 : (v.indexOf(qn) > 0 ? 2 : -1));
      if (s < 0) continue;
      if (best === null || s < best) best = s;
    }
    if (best !== null) scored.push([best, rec.row, rec]);
  }
  scored.sort(function (a, b) { return a[0] - b[0] || a[1] - b[1]; });
  var matches = scored.slice(0, MAX_RESULTS).map(function (t) { return t[2]; });
  renderResults(matches, q);
  setStatus(matches.length + ' match(es)');
}

/* ----------------------------------------------------------------- utils */
function toBase64(buf) {
  var bytes = new Uint8Array(buf), bin = '';
  for (var i = 0; i < bytes.length; i += 0x8000) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  }
  return btoa(bin);
}

function upsertCache(rec) {
  for (var i = 0; i < state.records.length; i++) {
    if (state.records[i].row === rec.row) { state.records[i] = rec; return; }
  }
  state.records.push(rec);
  state.records.sort(function (a, b) { return a.row - b.row; });
}

/* -------------------------------------------------------------- handlers */
/* One path for "a workbook arrived", whether the user picked it from disk or
 * the page fetched the copy that ships with the site. */
async function useWorkbook(b64, fileName, note) {
  setStatus('Reading file…');
  try {
    var d = await api.load(b64, fileName.replace(/\.(xlsx|xlsm)$/i, ''));
    state.workbook = SHARED ? '' : b64;
    if (SHARED) { adoptSheet(d); }
    state.records = d.records || [];
    state.macro = !!d.macro;
    state.fileName = fileName;
    state.prefixes = d.prefixes || {};
    applyFieldOptions(d.options);
    $('fileInfo').textContent = fileName + ' · ' + state.records.length +
      ' record(s) · header row ' + d.header_row;
    if (SHARED) { refreshSheetList(); }
    var noteEl = $('demoNote');
    if (noteEl) {
      noteEl.textContent = note || '';
      noteEl.classList.toggle('hidden', !note);
    }
    $('search').disabled = false;
    $('search').value = '';
    newRecord();
    renderResults([], '');
    setStatus(d.message || 'Excel loaded');
    if (d.missing && d.missing.length) {
      toast('Columns not found in this sheet: ' + d.missing.join(', ') + '. Those fields will not be saved.', 'warn');
    } else if (d.warning) {
      toast(d.warning, 'warn');
    } else {
      toast(d.message || 'Excel loaded', 'ok');
    }
  } catch (err) {
    state.workbook = ''; state.records = [];
    $('search').disabled = true;
    $('fileInfo').textContent = 'Could not read that file.';
    setStatus('Load failed');
    toast(err.message, 'error');
  }
}

$('file').addEventListener('change', async function (e) {
  var f = e.target.files[0];
  if (!f) return;
  try {
    await useWorkbook(toBase64(await f.arrayBuffer()), f.name, '');
  } finally {
    e.target.value = '';
  }
});

/* The copy that ships with the site, so nobody has to hunt for a file just to
 * start work. Only rendered when a sample was bundled at build time. */
var sampleBtn = $('sampleBtn');
if (sampleBtn) {
  sampleBtn.addEventListener('click', async function () {
    if (state.dirty && !window.confirm('Discard unsaved changes and open the team file?')) return;
    sampleBtn.disabled = true;
    setStatus('Opening the team file…');
    try {
      var res = await fetch(UI.sample_file, { cache: 'no-cache' });
      if (!res.ok) throw new Error('Could not download the team file (' + res.status + ').');
      await useWorkbook(toBase64(await res.arrayBuffer()), UI.sample_file, UI.sample_note || '');
    } catch (e) {
      setStatus('Could not open');
      toast(e.message, 'error');
    } finally {
      sampleBtn.disabled = false;
    }
  });
}

$('search').addEventListener('input', function () {
  clearTimeout(state.timer);
  state.timer = setTimeout(runSearch, 90);
});

$('search').addEventListener('keydown', function (e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    clearTimeout(state.timer);
    runSearch();
    if (state.results.length) openResult(state.results[0]);
  }
});

$('save').addEventListener('click', async function () {
  var rec = collect();
  if (!rec.parcel || !rec.name) {
    toast('P/S / Parcel Number and Name are required.', 'error');
    (rec.parcel ? $('f-name') : $('f-parcel')).focus();
    return;
  }
  var btn = $('save'), label = btn.textContent;
  btn.disabled = true; btn.textContent = 'Saving…';
  setStatus('Saving…');
  try {
    var d = await api.save(state.workbook, rec, state.row);
    if (!SHARED) { state.workbook = d.workbook; state.macro = !!d.macro; }
    if (SHARED && d.rev) { state.rev = d.rev; }
    upsertCache(d.record);
    editRecord(d.record);
    $('fileInfo').textContent = (state.fileName || 'New workbook') + ' · ' +
      state.records.length + ' record(s)';
    $('search').disabled = false;
    if ($('search').value.trim()) runSearch();
    setStatus(d.action === 'added' ? 'Record added' : 'Record updated');
    var atRow = d.record ? d.record.row : d.row;
    toast(d.action === 'added'
      ? 'New record added at row ' + atRow + '.'
      : 'Row ' + atRow + ' updated.', 'ok');
  } catch (err) {
    if (err.status === 409 && err.data && err.data.conflict) {
      setStatus('Someone else edited this');
      showConflict(err.data, rec);
    } else if (err.status === 409 && err.data && err.data.duplicate) {
      setStatus('Duplicate P/S');
      if (window.confirm(err.data.error + '\n\nOpen the existing record now?')) {
        var existing = err.data.record;
        existing.row = err.data.row;
        upsertCache(existing);
        $('search').value = existing.parcel || '';
        runSearch();
        editRecord(existing);
      }
    } else {
      setStatus('Save failed');
      toast(err.message, 'error');
    }
  } finally {
    btn.disabled = false; btn.textContent = label;
  }
});

$('clear').addEventListener('click', function () {
  if (state.dirty && !window.confirm('Discard unsaved changes?')) return;
  newRecord();
  setStatus('Ready');
});

$('newBtn').addEventListener('click', async function () {
  if (state.dirty && !window.confirm('Discard unsaved changes and start a new file?')) return;
  if (SHARED) {
    var suggested = window.prompt('Name for the new shared file:', 'Parcel mapping');
    if (suggested === null) return;
    setStatus('Creating…');
    try {
      adoptSheet(await api.blankSheet(suggested.trim() || 'Parcel mapping'));
      refreshSheetList();
      toast('Blank file created. Everyone on the team can now open it.', 'ok');
    } catch (e) { setStatus('Could not create'); toast(e.message, 'error'); }
    return;
  }
  state.workbook = ''; state.records = []; state.macro = false; state.fileName = '';
  state.prefixes = {};
  applyFieldOptions({});
  $('fileInfo').textContent = 'New file — a fresh proforma workbook is created on first save.';
  $('search').value = '';
  $('search').disabled = false;
  renderResults([], '');
  newRecord({});
  setStatus('New file mode');
});

$('download').addEventListener('click', async function () {
  if (!state.workbook) { toast('Load a file or save a record first.', 'error'); return; }
  setStatus('Preparing download…');
  try {
    var blob = await api.download(state.workbook, state.macro);
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'Parcel_Mapping_Updated.' + (state.macro ? 'xlsm' : 'xlsx');
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    setStatus('Excel downloaded');
    toast('Workbook downloaded.', 'ok');
  } catch (err) {
    setStatus('Download failed');
    toast(err.message, 'error');
  }
});

document.addEventListener('input', function (e) {
  if (e.target && e.target.dataset && e.target.dataset.k) setDirty(true);
});

window.addEventListener('beforeunload', function (e) {
  // In shared mode everything saved is already on the server; only warn about
  // a half-typed record. In local mode the whole workbook is only in this tab.
  if (state.dirty || (!SHARED && state.workbook)) { e.preventDefault(); e.returnValue = ''; }
});

newRecord({});
renderResults([], '');
