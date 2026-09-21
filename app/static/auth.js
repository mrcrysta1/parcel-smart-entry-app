/* Parcel Smart Entry - optional team password (shared build only)
 *
 * The server decides whether a password is needed: GET /api/auth answers
 * {required:true|false}, driven by the TEAM_PASSWORD environment variable. If
 * it is off, nothing here changes how the app behaves.
 *
 * On sign-in the password is exchanged once for a signed token, which is what
 * gets stored and sent from then on. The password itself is never written to
 * localStorage and never sent again.
 */
'use strict';

var PARCEL_AUTH = (function () {
  var TOKEN_KEY = 'parcel.token';
  var MODE_KEY = 'parcel.authRequired';   // last answer, so a revisit is instant
  var required = false;

  function remember(val) {
    try { localStorage.setItem(MODE_KEY, val ? '1' : '0'); } catch (e) {}
  }

  /* What we were told last time: true, false, or null if never asked. Lets the
   * page start rendering without waiting for another round trip; the real
   * answer is confirmed in the background, and a wrong guess still ends up at
   * the 401 handler. */
  function knownRequired() {
    try {
      var v = localStorage.getItem(MODE_KEY);
      return v === null ? null : v === '1';
    } catch (e) { return null; }
  }

  function token() {
    try { return localStorage.getItem(TOKEN_KEY) || ''; } catch (e) { return ''; }
  }

  function store(t) {
    try {
      if (t) localStorage.setItem(TOKEN_KEY, t);
      else localStorage.removeItem(TOKEN_KEY);
    } catch (e) { /* private mode - the session just will not be remembered */ }
  }

  /* Asked once and remembered. Callers can await it freely; the request is
   * only made the first time. It also gives up after a few seconds rather
   * than leaving the sign-in card hanging on a slow network - if a password
   * really is required, the first API call comes back 401 and the page asks
   * then. */
  var checking = null;
  function check() {
    if (checking) return checking;
    checking = (async function () {
      var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
      // Generous: a serverless function that has gone cold can take a while to
      // answer the very first request.
      var timer = setTimeout(function () { if (ctrl) ctrl.abort(); }, 20000);
      try {
        var res = await fetch('/api/auth', {
          headers: { 'cache-control': 'no-cache' },
          signal: ctrl ? ctrl.signal : undefined
        });
        var d = await res.json();
        required = !!d.required;
        remember(required);
      } catch (e) {
        /* Never cache a failure. Remembering "no password needed" because one
         * request timed out would leave the app permanently unable to ask for
         * it. Fall back to whatever we were told last time, let the next call
         * try again, and rely on the 401 handler in the meantime. */
        var last = knownRequired();
        required = last === null ? false : last;
        checking = null;
      } finally {
        clearTimeout(timer);
      }
      return required;
    })();
    return checking;
  }

  async function login(password) {
    var res = await fetch('/api/auth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password: password })
    });
    var d = {};
    try { d = await res.json(); } catch (e) {}
    if (!res.ok) throw new Error(d.error || 'Could not sign in.');
    store(d.token);
    return d;
  }

  return {
    check: check,
    login: login,
    token: token,
    isRequired: function () { return required; },
    knownRequired: knownRequired,
    // A token can expire or be invalidated by changing the password; when the
    // server says so, drop it and ask again.
    clear: function () { store(''); },
    hasToken: function () { return !!token(); }
  };
})();
