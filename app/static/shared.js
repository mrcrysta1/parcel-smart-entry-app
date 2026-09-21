/* Parcel Smart Entry - shared-mode additions
 *
 * Loaded only by the Netlify build, after app.js. Adds the parts that only
 * make sense when the workbook lives on the server and several people are in
 * it at once:
 *   - asking who you are, so edits can be attributed
 *   - listing and opening the files held on the server
 *   - polling for other people's changes and folding them into the cache
 *   - the conflict panel shown when two people save the same parcel
 *
 * app.js defines `state`, `SHARED`, `api`, `toast`, `editRecord` and friends,
 * and calls the hooks below (adoptSheet, refreshSheetList, showConflict,
 * hideConflict, ago) which are defined here as no-ops when not in shared mode.
 */
'use strict';

var POLL_MS = 2500;
var USER_KEY = 'parcel.user';

/* ------------------------------------------------------------------ time */
function ago(iso) {
  if (!iso) return '';
  var secs = Math.max(0, (Date.now() - new Date(iso).getTime()) / 1000);
  if (secs < 10) return 'just now';
  if (secs < 60) return Math.floor(secs) + 's ago';
  if (secs < 3600) return Math.floor(secs / 60) + 'm ago';
  if (secs < 86400) return Math.floor(secs / 3600) + 'h ago';
  return new Date(iso).toLocaleDateString();
}

/* ------------------------------------------------------------- who am I */
function setUser(name) {
  state.user = String(name || '').trim().slice(0, 40);
  try { localStorage.setItem(USER_KEY, state.user); } catch (e) {}
  var b = $('whoami');
  if (b) b.textContent = state.user || 'Set name';
  var gate = $('gate');
  if (gate) gate.classList.toggle('hidden', !!state.user);
}

function askWho() {
  var gate = $('gate');
  if (!gate) return;
  gate.classList.remove('hidden');
  var f = $('gateName');
  f.value = state.user || '';
  setTimeout(function () { f.focus(); }, 50);
}

/* --------------------------------------------------------- sheet picker */
function renderSheetList(sheets) {
  var el = $('sheetList');
  if (!el) return;
  if (!sheets.length) {
    el.innerHTML = '<div class="empty">No shared files yet. Upload one to get started — ' +
      'everyone else will then see it here.</div>';
    return;
  }
  el.innerHTML = sheets.map(function (s) {
    var active = s.id === state.sheetId;
    return '<button type="button" class="sheet' + (active ? ' active' : '') + '" data-id="' + esc(s.id) + '">' +
      '<b>' + esc(s.name) + '</b>' +
      '<span>' + s.rows + ' record(s) · last change ' + esc(ago(s.updatedAt)) +
      (s.updatedBy ? ' by ' + esc(s.updatedBy) : '') + '</span>' +
      (active ? '<em>Open</em>' : '') + '</button>';
  }).join('');
}

async function refreshSheetList() {
  if (!SHARED) return;
  try {
    var d = await api.listSheets();
    renderSheetList(d.sheets || []);
  } catch (e) {
    var el = $('sheetList');
    if (el) el.innerHTML = '<div class="empty">Could not reach the server: ' + esc(e.message) + '</div>';
  }
}

/* Take a server snapshot and make it the sheet this tab is working on. */
function adoptSheet(d) {
  state.sheetId = d.id;
  state.rev = d.rev || 0;
  state.records = d.records || [];
  state.prefixes = d.prefixes || {};
  state.fileName = d.name || '';
  state.macro = false;
  applyFieldOptions(d.options);
  $('search').disabled = false;
  $('search').value = '';
  renderResults([], '');
  newRecord();
  describeSheet();
  var link = $('gridLink');
  if (link) {
    link.href = 'sheet.html?id=' + encodeURIComponent(d.id);
    link.classList.remove('hidden');
  }
  try {
    history.replaceState(null, '', '?id=' + encodeURIComponent(d.id));
    localStorage.setItem('parcel.lastSheet', d.id);
  } catch (e) {}
  startPolling();
}

function describeSheet() {
  var info = $('fileInfo');
  if (!info) return;
  info.textContent = state.fileName + ' · ' + state.records.length +
    ' record(s) · shared with your team';
}

async function openSheet(id) {
  if (state.dirty && !window.confirm('Discard unsaved changes and open that file?')) return;
  setStatus('Opening…');
  try {
    adoptSheet(await api.openSheet(id));
    refreshSheetList();
    setStatus(state.records.length + ' record(s) loaded');
  } catch (e) {
    setStatus('Could not open');
    toast(e.message, 'error');
  }
}

/* ------------------------------------------------------- live updates */
/* Poll for records changed since the revision this tab last saw. Only the
 * changed rows come back, so this stays cheap with a few people editing. */
function startPolling() {
  stopPolling();
  if (!SHARED || !state.sheetId) return;
  state.poll = setInterval(pollOnce, POLL_MS);
}

function stopPolling() {
  if (state.poll) { clearInterval(state.poll); state.poll = null; }
}

var pollBusy = false;
async function pollOnce() {
  if (pollBusy || document.hidden || !state.sheetId) return;
  pollBusy = true;
  try {
    var d = await api.changes(state.sheetId, state.rev);
    if (d.rev === state.rev) return;
    state.rev = d.rev;
    var openRowChanged = null;
    (d.records || []).forEach(function (rec) {
      upsertCache(rec);
      if (state.row !== null && rec.row === state.row && rec.updatedBy !== state.user) {
        openRowChanged = rec;
      }
    });
    describeSheet();
    if ($('search').value.trim()) runSearch();
    markLiveActivity(d);
    // Warn before a clash rather than only at save time: if someone else just
    // changed the record open in the form, say so while it can still be merged.
    if (openRowChanged) warnRecordChanged(openRowChanged);
  } catch (e) {
    // a dropped poll is not worth interrupting data entry over
  } finally {
    pollBusy = false;
  }
}

function markLiveActivity(d) {
  if (!d.updatedBy || d.updatedBy === state.user) return;
  setStatus(d.updatedBy + ' edited · ' + (d.total || state.records.length) + ' record(s)');
}

function warnRecordChanged(rec) {
  if (state.dirty) {
    setBanner(rec.updatedBy + ' just saved changes to row ' + rec.row +
      '. Your unsaved edits are still here — saving will be checked against their version.', 'edit');
    toast(rec.updatedBy + ' changed the record you are editing.', 'warn');
  } else {
    editRecord(rec);
    toast(rec.updatedBy + ' changed this record — reloaded it for you.', 'warn');
  }
}

/* ------------------------------------------------------------ conflicts */
function fieldRows(dl, rec, other) {
  dl.innerHTML = (UI.form_order || Object.keys(rec)).filter(function (f) {
    return rec[f] !== undefined && UI.labels && UI.labels[f];
  }).map(function (f) {
    var differs = other && String(rec[f] || '') !== String(other[f] || '');
    return '<div' + (differs ? ' class="diff"' : '') + '><dt>' + esc(UI.labels[f]) +
      '</dt><dd>' + esc(rec[f] || '—') + '</dd></div>';
  }).join('');
}

function showConflict(data, mine) {
  var panel = $('conflict');
  if (!panel) { toast(data.error, 'error'); return; }
  state.pending = { theirs: data.theirs, mine: mine };
  $('conflictHead').textContent = data.error + ' Compare the two and choose which to keep.';
  fieldRows($('conflictTheirs'), data.theirs, mine);
  fieldRows($('conflictMine'), mine, data.theirs);
  panel.classList.remove('hidden');
  panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function hideConflict() {
  var panel = $('conflict');
  if (panel) panel.classList.add('hidden');
  if (typeof state !== 'undefined') state.pending = null;
}

function wireConflict() {
  var theirs = $('takeTheirs'), mine = $('keepMine');
  if (!theirs || !mine) return;

  theirs.addEventListener('click', function () {
    if (!state.pending) return hideConflict();
    var rec = state.pending.theirs;
    upsertCache(rec);
    hideConflict();
    editRecord(rec);
    toast('Kept their version. Your unsaved changes were discarded.', 'ok');
  });

  // Overwrite deliberately: re-save against their revision so the write is
  // still guarded, it is just no longer based on a stale copy.
  mine.addEventListener('click', async function () {
    if (!state.pending) return hideConflict();
    // Read both versions out before hiding the panel - hideConflict() clears
    // state.pending, and the save below still needs them.
    var theirs = state.pending.theirs;
    var myVersion = state.pending.mine;
    state.baseRev = theirs.rev;
    hideConflict();
    setStatus('Saving…');
    try {
      var d = await api.save('', myVersion, state.row);
      state.rev = d.rev || state.rev;
      upsertCache(d.record);
      editRecord(d.record);
      if ($('search').value.trim()) runSearch();
      setStatus('Record updated');
      toast('Your version was saved over theirs.', 'ok');
    } catch (e) {
      setStatus('Save failed');
      toast(e.message, 'error');
    }
  });
}

/* ------------------------------------------------------------------ init */
function initShared() {
  if (!SHARED) return;

  var saved = '';
  try { saved = localStorage.getItem(USER_KEY) || ''; } catch (e) {}
  if (saved) setUser(saved); else askWho();

  var form = $('gateForm');
  if (form) {
    form.addEventListener('submit', function (e) {
      e.preventDefault();
      var v = $('gateName').value.trim();
      if (!v) return;
      setUser(v);
      boot();
    });
  }
  var b = $('whoami');
  if (b) b.addEventListener('click', askWho);

  var list = $('sheetList');
  if (list) {
    list.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('.sheet') : null;
      if (btn) openSheet(btn.dataset.id);
    });
  }

  wireConflict();
  document.addEventListener('visibilitychange', function () {
    if (!document.hidden) pollOnce();
  });

  if (saved) boot();
}

async function boot() {
  await refreshSheetList();
  var want = new URLSearchParams(location.search).get('id');
  if (!want) { try { want = localStorage.getItem('parcel.lastSheet') || ''; } catch (e) {} }
  if (want) {
    try { adoptSheet(await api.openSheet(want)); refreshSheetList(); return; } catch (e) {}
  }
  setStatus('Pick a file to start');
}

initShared();
