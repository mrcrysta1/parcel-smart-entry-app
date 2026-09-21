/* Parcel Smart Entry - full-page grid
 *
 * Every record in the shared file as one live table. Edits are made in place:
 * click a cell, type, leave the cell, and it is saved. Other people's changes
 * arrive within a couple of seconds and the affected row flashes.
 *
 * Each cell save carries the revision the row was read at, so if someone else
 * changed that row in the meantime the server rejects the write and we ask
 * which version to keep, rather than quietly clobbering their work.
 */
'use strict';

var UI = window.PARCEL_UI || {};
var FIELDS = UI.form_order || UI.fields || [];
var LABELS = UI.labels || {};
var CHOICE_FIELDS = UI.choice_fields || {};
var SEARCH_FIELDS = UI.search_fields || [];
var POLL_MS = 2500;
var USER_KEY = 'parcel.user';

var state = { id: '', name: '', rev: 0, records: [], user: '', filter: '', editing: null, poll: null };

var $ = function (id) { return document.getElementById(id); };

function esc(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;').replace(/"/g, '&quot;')
    .replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function setStatus(t) { $('status').textContent = t; }

var toastTimer = null;
function toast(msg, kind) {
  var el = $('toast');
  el.textContent = msg;
  el.className = 'toast ' + (kind || 'info');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.className = 'toast hidden'; }, 5200);
}

function ago(iso) {
  if (!iso) return '';
  var secs = Math.max(0, (Date.now() - new Date(iso).getTime()) / 1000);
  if (secs < 10) return 'just now';
  if (secs < 60) return Math.floor(secs) + 's ago';
  if (secs < 3600) return Math.floor(secs / 60) + 'm ago';
  if (secs < 86400) return Math.floor(secs / 3600) + 'h ago';
  return new Date(iso).toLocaleDateString();
}

/* -------------------------------------------------------------------- api */
async function rest(method, path, body) {
  var init = { method: method, headers: { 'x-parcel-user': state.user || 'Someone' } };
  if (body !== undefined) {
    init.headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(Object.assign({ by: state.user }, body));
  }
  var res;
  try { res = await fetch(path, init); }
  catch (e) { throw new Error('Cannot reach the server.'); }
  var text = await res.text();
  var data = {};
  try { data = text ? JSON.parse(text) : {}; }
  catch (e) { throw new Error('Unexpected server response (' + res.status + ').'); }
  if (!res.ok) {
    var err = new Error(data.error || 'Request failed (' + res.status + ').');
    err.status = res.status; err.data = data;
    throw err;
  }
  return data;
}

/* ------------------------------------------------------------------- user */
function setUser(name) {
  state.user = String(name || '').trim().slice(0, 40);
  try { localStorage.setItem(USER_KEY, state.user); } catch (e) {}
  $('whoami').textContent = state.user || 'Set name';
  $('gate').classList.toggle('hidden', !!state.user);
}

/* ------------------------------------------------------------------- grid */
function visibleRecords() {
  var q = state.filter.toLowerCase().replace(/[^a-z0-9]+/g, '');
  var list = state.records;
  if (q) {
    list = list.filter(function (r) {
      return SEARCH_FIELDS.some(function (f) {
        return String(r[f] || '').toLowerCase().replace(/[^a-z0-9]+/g, '').indexOf(q) >= 0;
      });
    });
  }
  return list.slice().sort(function (a, b) { return a.row - b.row; });
}

function renderHead() {
  $('gridHead').innerHTML = '<tr><th class="rownum">Row</th>' +
    FIELDS.map(function (f) { return '<th>' + esc(LABELS[f] || f) + '</th>'; }).join('') +
    '<th class="who">Last change</th></tr>';
}

function cellHtml(rec, f) {
  return '<td data-row="' + rec.row + '" data-k="' + esc(f) + '" tabindex="0">' +
    '<span>' + esc(rec[f] || '') + '</span></td>';
}

function rowHtml(rec) {
  return '<tr data-row="' + rec.row + '" data-rev="' + (rec.rev || 0) + '">' +
    '<th class="rownum">' + rec.row + '</th>' +
    FIELDS.map(function (f) { return cellHtml(rec, f); }).join('') +
    '<td class="who">' + esc(rec.updatedBy || '') +
    (rec.updatedAt ? '<span>' + esc(ago(rec.updatedAt)) + '</span>' : '') + '</td></tr>';
}

function renderGrid() {
  var list = visibleRecords();
  $('gridBody').innerHTML = list.map(rowHtml).join('');
  $('gridEmpty').textContent = list.length ? '' :
    (state.records.length ? 'No record matches that filter.' : 'This file has no records yet.');
  $('gridEmpty').style.display = list.length ? 'none' : 'block';
  $('sheetMeta').textContent = state.records.length + ' record(s)' +
    (state.filter ? ' · ' + list.length + ' shown' : '') + ' · updates live';
}

/* Repaint one row in place, so editing elsewhere in the table is undisturbed. */
function repaintRow(rec, flash) {
  var tr = $('gridBody').querySelector('tr[data-row="' + rec.row + '"]');
  if (!tr) { renderGrid(); return; }
  var wrap = document.createElement('tbody');
  wrap.innerHTML = rowHtml(rec);
  var fresh = wrap.firstChild;
  tr.parentNode.replaceChild(fresh, tr);
  if (flash) {
    fresh.classList.add('changed');
    setTimeout(function () { fresh.classList.remove('changed'); }, 2600);
  }
}

function upsert(rec) {
  for (var i = 0; i < state.records.length; i++) {
    if (state.records[i].row === rec.row) { state.records[i] = rec; return false; }
  }
  state.records.push(rec);
  return true;
}

/* ---------------------------------------------------------- cell editing */
function closeEditor(commit) {
  var ed = state.editing;
  if (!ed) return;
  state.editing = null;
  var input = ed.input;
  var value = input.value;
  var td = ed.td;
  td.innerHTML = '<span>' + esc(commit ? value : ed.original) + '</span>';
  if (commit && value !== ed.original) saveCell(ed.row, ed.field, value, td);
}

function openEditor(td) {
  if (state.editing) closeEditor(true);
  var row = +td.dataset.row, field = td.dataset.k;
  var rec = state.records.filter(function (r) { return r.row === row; })[0];
  if (!rec) return;
  var original = String(rec[field] || '');

  var input;
  if (CHOICE_FIELDS[field]) {
    input = document.createElement('select');
    var blank = document.createElement('option');
    blank.value = ''; blank.textContent = '—';
    input.appendChild(blank);
    CHOICE_FIELDS[field].forEach(function (v) {
      var o = document.createElement('option'); o.value = v; o.textContent = v; input.appendChild(o);
    });
    if (original && CHOICE_FIELDS[field].indexOf(original) < 0) {
      var keep = document.createElement('option');
      keep.value = original; keep.textContent = original;
      input.appendChild(keep);
    }
  } else {
    input = document.createElement('input');
    input.type = 'text';
  }
  input.value = original;
  td.innerHTML = '';
  td.appendChild(input);
  input.focus();
  if (input.select) input.select();

  state.editing = { td: td, input: input, row: row, field: field, original: original };

  input.addEventListener('blur', function () { closeEditor(true); });
  input.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
    else if (e.key === 'Escape') { e.preventDefault(); closeEditor(false); }
    else if (e.key === 'Tab') {
      e.preventDefault();
      var cells = Array.prototype.slice.call($('gridBody').querySelectorAll('td[data-k]'));
      var i = cells.indexOf(td);
      closeEditor(true);
      var next = cells[i + (e.shiftKey ? -1 : 1)];
      if (next) openEditor(next);
    }
  });
}

async function saveCell(row, field, value, td) {
  var rec = state.records.filter(function (r) { return r.row === row; })[0];
  if (!rec) return;
  var payload = {};
  (UI.fields || FIELDS).forEach(function (f) { payload[f] = rec[f] || ''; });
  payload[field] = value;

  td.classList.add('saving');
  try {
    var d = await rest('PUT', '/api/sheets/' + state.id + '/records/' + row,
                       { record: payload, baseRev: rec.rev });
    state.rev = d.rev || state.rev;
    upsert(d.record);
    repaintRow(d.record, false);
    setStatus('Saved row ' + row);
  } catch (e) {
    if (e.status === 409 && e.data && e.data.conflict) {
      var theirs = e.data.theirs;
      upsert(theirs);
      var keepMine = window.confirm(
        e.data.error + '\n\n' +
        'Their ' + (LABELS[field] || field) + ': ' + (theirs[field] || '(empty)') + '\n' +
        'Your ' + (LABELS[field] || field) + ': ' + (value || '(empty)') + '\n\n' +
        'OK = keep YOUR value (overwrites theirs)\nCancel = keep THEIR value');
      if (keepMine) {
        var again = {};
        (UI.fields || FIELDS).forEach(function (f) { again[f] = theirs[f] || ''; });
        again[field] = value;
        try {
          var d2 = await rest('PUT', '/api/sheets/' + state.id + '/records/' + row,
                              { record: again, baseRev: theirs.rev });
          state.rev = d2.rev || state.rev;
          upsert(d2.record);
          repaintRow(d2.record, false);
          toast('Your value was saved over theirs.', 'ok');
        } catch (e2) { toast(e2.message, 'error'); repaintRow(theirs, true); }
      } else {
        repaintRow(theirs, true);
        toast('Kept ' + (theirs.updatedBy || 'their') + ' version.', 'warn');
      }
    } else if (e.status === 409 && e.data && e.data.duplicate) {
      toast(e.data.error, 'error');
      repaintRow(rec, true);
    } else {
      toast(e.message, 'error');
      repaintRow(rec, true);
    }
  } finally {
    td.classList.remove('saving');
  }
}

/* ------------------------------------------------------------ live sync */
var pollBusy = false;
async function pollOnce() {
  if (pollBusy || document.hidden || !state.id) return;
  pollBusy = true;
  try {
    var d = await rest('GET', '/api/sheets/' + state.id + '/changes?since=' + state.rev);
    if (d.rev === state.rev) return;
    state.rev = d.rev;
    var added = false;
    (d.records || []).forEach(function (rec) {
      // never yank a cell out from under someone mid-type
      if (state.editing && state.editing.row === rec.row) { upsert(rec); return; }
      if (upsert(rec)) added = true;
      else repaintRow(rec, rec.updatedBy !== state.user);
    });
    if (added) renderGrid();
    else $('sheetMeta').textContent = state.records.length + ' record(s) · updates live';
    if (d.updatedBy && d.updatedBy !== state.user) setStatus(d.updatedBy + ' is editing');
  } catch (e) {
    /* a dropped poll should not interrupt typing */
  } finally {
    pollBusy = false;
  }
}

/* ---------------------------------------------------------------- startup */
async function openSheet(id) {
  setStatus('Loading…');
  try {
    var d = await rest('GET', '/api/sheets/' + id);
    state.id = d.id;
    state.name = d.name;
    state.rev = d.rev;
    state.records = d.records || [];
    $('sheetName').textContent = d.name;
    $('backLink').href = 'index.html?id=' + encodeURIComponent(d.id);
    renderHead();
    renderGrid();
    setStatus('Live');
    if (state.poll) clearInterval(state.poll);
    state.poll = setInterval(pollOnce, POLL_MS);
  } catch (e) {
    setStatus('Could not load');
    $('sheetName').textContent = 'Could not load this file';
    $('gridEmpty').textContent = e.message;
    $('gridEmpty').style.display = 'block';
  }
}

async function addRecord() {
  var parcel = window.prompt('P/S / Parcel Number for the new record:');
  if (parcel === null) return;
  parcel = parcel.trim();
  if (!parcel) return;
  var name = window.prompt('Name for P/S ' + parcel + ':');
  if (name === null) return;
  name = name.trim();
  if (!name) { toast('Name is required.', 'error'); return; }
  try {
    var d = await rest('POST', '/api/sheets/' + state.id + '/records',
                       { record: { parcel: parcel, name: name } });
    state.rev = d.rev || state.rev;
    upsert(d.record);
    renderGrid();
    repaintRow(d.record, true);
    toast('Added at row ' + d.record.row + '.', 'ok');
  } catch (e) {
    toast(e.message, 'error');
  }
}

function init() {
  var saved = '';
  try { saved = localStorage.getItem(USER_KEY) || ''; } catch (e) {}
  if (saved) setUser(saved);

  $('gateForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var v = $('gateName').value.trim();
    if (!v) return;
    setUser(v);
    start();
  });
  $('whoami').addEventListener('click', function () {
    $('gate').classList.remove('hidden');
    $('gateName').value = state.user;
    $('gateName').focus();
  });

  $('gridBody').addEventListener('click', function (e) {
    var td = e.target.closest ? e.target.closest('td[data-k]') : null;
    if (td && (!state.editing || state.editing.td !== td)) openEditor(td);
  });
  $('gridBody').addEventListener('keydown', function (e) {
    if (e.key !== 'Enter') return;
    var td = e.target.closest ? e.target.closest('td[data-k]') : null;
    if (td && !state.editing) { e.preventDefault(); openEditor(td); }
  });

  var t = null;
  $('filter').addEventListener('input', function () {
    clearTimeout(t);
    t = setTimeout(function () { state.filter = $('filter').value.trim(); renderGrid(); }, 100);
  });

  $('addRow').addEventListener('click', addRecord);
  $('download').addEventListener('click', async function () {
    try {
      var r = await fetch('/api/sheets/' + state.id + '/export');
      if (!r.ok) throw new Error('Download failed.');
      var blob = await r.blob();
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = (state.name || 'Parcel_Mapping').replace(/[^A-Za-z0-9_-]+/g, '_') + '.xlsx';
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
      toast('Workbook downloaded.', 'ok');
    } catch (e) { toast(e.message, 'error'); }
  });

  document.addEventListener('visibilitychange', function () { if (!document.hidden) pollOnce(); });

  if (saved) start(); else { $('gate').classList.remove('hidden'); $('gateName').focus(); }
}

function start() {
  var id = new URLSearchParams(location.search).get('id');
  if (!id) { try { id = localStorage.getItem('parcel.lastSheet') || ''; } catch (e) {} }
  if (!id) {
    setStatus('No file chosen');
    $('sheetName').textContent = 'No file chosen';
    $('gridEmpty').innerHTML = 'Open a file from the <a href="index.html">data entry page</a> first.';
    $('gridEmpty').style.display = 'block';
    return;
  }
  openSheet(id);
}

init();
