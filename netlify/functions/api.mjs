/* Parcel Smart Entry - shared multi-user API (Netlify Function)
 *
 * Holds the workbook on the server so several people edit one file at once.
 *
 * Storage is Netlify Blobs, which needs no account, connection string or
 * environment variable - credentials are injected at runtime. Two keys per
 * sheet:
 *    sheet:<id>   JSON  { meta, records[], rev }   the live data
 *    file:<id>    bytes the original .xlsx         kept for export formatting
 *  The file list is built from each sheet blob's own metadata, so there is no
 *  central index to lock on every save.
 *
 * Concurrency. Every write is a read-modify-write guarded by the blob's ETag
 * (`onlyIfMatch`), so two people saving at the same instant cannot silently
 * overwrite one another - the loser retries against fresh data. On top of
 * that, each record carries its own `rev`; a save quotes the rev it was based
 * on, and if someone else has changed that record in the meantime the save is
 * rejected with 409 and both versions so the user can decide. Reads use strong
 * consistency, because the default eventual mode can lag up to a minute and
 * this needs to feel live.
 */
import { getStore } from '@netlify/blobs';
import { createHmac, createHash, timingSafeEqual } from 'node:crypto';
import ExcelJS from 'exceljs';
import createEngine from '../../shared/engine-core.cjs';
import schema from '../../shared/schema.json';

const engine = createEngine({ schema, ExcelJS });
const FIELDS = schema.fields;

const STORE = 'parcel-sheets';
const MAX_UPLOAD = 25 * 1024 * 1024;
const MAX_RETRY = 8;

const store = () => getStore({ name: STORE, consistency: 'strong' });
// No ':' in keys: list() percent-encodes it in some runtimes, which silently
// breaks prefix filtering. '_' round-trips everywhere.
const SHEET_PREFIX = 'sheet_';
const sheetKey = (id) => `${SHEET_PREFIX}${id}`;
const fileKey = (id) => `file_${id}`;
const now = () => new Date().toISOString();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function fail(status, message, extra) {
  const e = new Error(message);
  e.status = status;
  e.payload = { error: message, ...(extra || {}) };
  return e;
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json', 'cache-control': 'no-store' }
  });
}

function newId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function who(body, req) {
  const name = String((body && body.by) || req.headers.get('x-parcel-user') || '').trim();
  return name.slice(0, 40) || 'Someone';
}

/* ---------------------------------------------------------------- access
 *
 * Optional shared team password, switched on purely by setting the
 * TEAM_PASSWORD environment variable in Netlify. With it unset the app is
 * open exactly as before, so turning this on is a deliberate act and turning
 * it off again is a one-click revert.
 *
 * The password is exchanged once for a signed token that the browser keeps;
 * the password itself is never stored client-side and never travels again
 * after sign-in. The token is an expiry plus an HMAC of that expiry keyed by
 * the password, so no server-side session store is needed - which matters on
 * serverless, where there is nowhere to keep one.
 */
const TEAM_PASSWORD = String(process.env.TEAM_PASSWORD || process.env.PARCEL_PASSWORD || '').trim();
const TOKEN_DAYS = 30;
const authRequired = () => TEAM_PASSWORD.length > 0;

function sign(expiry) {
  return createHmac('sha256', TEAM_PASSWORD).update(String(expiry)).digest('hex');
}

function issueToken() {
  const exp = Date.now() + TOKEN_DAYS * 86400000;
  return { token: exp + '.' + sign(exp), expiresAt: new Date(exp).toISOString() };
}

function validToken(raw) {
  if (typeof raw !== 'string') return false;
  const dot = raw.indexOf('.');
  if (dot < 1) return false;
  const exp = Number(raw.slice(0, dot));
  if (!Number.isFinite(exp) || exp < Date.now()) return false;
  let given;
  try { given = Buffer.from(raw.slice(dot + 1), 'hex'); } catch (e) { return false; }
  const want = Buffer.from(sign(exp), 'hex');
  return given.length === want.length && timingSafeEqual(given, want);
}

// Compare fixed-length digests so the check cannot be timed character by character.
function passwordMatches(given) {
  const a = createHash('sha256').update(String(given == null ? '' : given)).digest();
  const b = createHash('sha256').update(TEAM_PASSWORD).digest();
  return timingSafeEqual(a, b);
}

async function handleAuth(req, method) {
  if (method === 'GET') return json({ required: authRequired() });
  if (method !== 'POST') throw fail(405, 'Method not allowed.');
  if (!authRequired()) return json(issueToken());
  const body = await req.json().catch(() => ({}));
  if (!passwordMatches(body.password)) {
    // Slow guessing down a little; there is no shared store to rate-limit with.
    await sleep(400 + Math.random() * 400);
    throw fail(401, 'That team password is not right.', { authRequired: true });
  }
  return json(issueToken());
}

/* ------------------------------------------------------------- storage */

/* Why a lease and not ETags alone.
 *
 * Netlify Blobs offers compare-and-swap through `onlyIfMatch`, and it works:
 * a write carrying a stale ETag is refused. But the ETag has to come back from
 * a *read*, and that is not guaranteed - the local dev sandbox, for one,
 * returns an empty `etag` on every read. `onlyIfMatch: undefined` silently
 * degrades to an unconditional write, which is last-write-wins, which for a
 * survey means one enumerator's record quietly disappearing.
 *
 * `onlyIfNew` is atomic everywhere, so writes to a sheet are serialised with a
 * short lease built on it, and the ETag is still passed when a read supplies
 * one. Correct where CAS is available, correct where it is not.
 */
const LOCK_TTL_MS = 12000;      // a lease older than this is treated as abandoned
const LOCK_WAIT_MS = 15000;     // how long to queue before giving up

async function withLock(name, fn) {
  const s = store();
  const key = `lock_${name}`;
  const token = newId();
  const deadline = Date.now() + LOCK_WAIT_MS;

  while (Date.now() < deadline) {
    const got = await s.setJSON(key, { token, at: Date.now() }, { onlyIfNew: true });
    if (got && got.modified) {
      try {
        return await fn();
      } finally {
        // Release only our own lease - never one that was stolen from us.
        try {
          const cur = await s.get(key, { type: 'json', consistency: 'strong' });
          if (!cur || cur.token === token) await s.delete(key);
        } catch (e) { /* the lease expires on its own */ }
      }
    }
    // Held by someone else. If the holder died mid-write, take it over.
    const cur = await s.get(key, { type: 'json', consistency: 'strong' });
    if (!cur || Date.now() - (cur.at || 0) > LOCK_TTL_MS) {
      await s.delete(key);
      continue;
    }
    await sleep(40 + Math.random() * 110);
  }
  throw fail(503, 'The sheet is busy with other edits right now. Please try again in a moment.');
}

/** Read-modify-write a JSON blob, serialised per key.
 *  `fn` receives the current value and returns { value, result } to commit,
 *  or { abort: <Response> } to bail out without writing. */
async function mutate(key, fn) {
  const s = store();
  return withLock(key, async () => {
    for (let attempt = 0; attempt < MAX_RETRY; attempt++) {
      const got = await s.getWithMetadata(key, { type: 'json', consistency: 'strong' });
      if (!got || got.data === null || got.data === undefined) {
        throw fail(404, 'That sheet no longer exists.');
      }
      const out = await fn(structuredClone(got.data));
      if (out.abort) return out.abort;
      // Belt and braces: hold the lease AND refuse the write if the blob moved.
      const opts = { metadata: listingMeta(out.value) };
      if (got.etag) opts.onlyIfMatch = got.etag;
      const res = await s.setJSON(key, out.value, opts);
      if (!got.etag || (res && res.modified)) return out.result;
      await sleep(30 + Math.random() * 120);
    }
    throw fail(503, 'Could not save because of competing edits. Please try again.');
  });
}

/* The file list comes from each sheet blob's own metadata, not a central index
 * blob. A shared index would have to be locked on every single record save,
 * which serialises writes across unrelated files for no benefit. Metadata
 * rides along with the write that already happens, and `getMetadata` reads it
 * without downloading the records. */
function listingMeta(sheet) {
  return {
    name: sheet.name,
    rows: String(sheet.records.length),
    rev: String(sheet.rev),
    createdAt: sheet.createdAt,
    createdBy: sheet.createdBy,
    updatedAt: sheet.updatedAt,
    updatedBy: sheet.updatedBy
  };
}

/* -------------------------------------------------------------- shaping */
function publicRecord(r) {
  const out = { row: r.row, rev: r.rev, updatedBy: r.updatedBy, updatedAt: r.updatedAt };
  FIELDS.forEach((f) => { out[f] = r[f] === undefined ? '' : r[f]; });
  return out;
}

function snapshot(sheet) {
  const records = sheet.records.map(publicRecord);
  return {
    id: sheet.id,
    name: sheet.name,
    rev: sheet.rev,
    sheet: sheet.sheetName,
    header_row: sheet.headerRow,
    records,
    options: engine.fieldOptions(records),
    prefixes: engine.fieldPrefixes(records),
    missing: sheet.missing || [],
    macro: !!sheet.macro,
    updatedAt: sheet.updatedAt,
    updatedBy: sheet.updatedBy
  };
}

/* -------------------------------------------------------------- handlers */

async function listSheets() {
  const s = store();
  const { blobs } = await s.list({ prefix: SHEET_PREFIX });
  // Some runtimes hand back percent-encoded keys; normalise before matching.
  const keys = blobs
    .map((b) => { try { return decodeURIComponent(b.key); } catch (e) { return b.key; } })
    .filter((k) => k.startsWith(SHEET_PREFIX));
  const sheets = (await Promise.all(keys.map(async (key) => {
    const id = key.slice(SHEET_PREFIX.length);
    try {
      const { metadata } = await s.getMetadata(key, { consistency: 'strong' });
      if (!metadata || !metadata.name) return null;
      return {
        id,
        name: metadata.name,
        rows: Number(metadata.rows) || 0,
        rev: Number(metadata.rev) || 0,
        createdAt: metadata.createdAt,
        createdBy: metadata.createdBy,
        updatedAt: metadata.updatedAt,
        updatedBy: metadata.updatedBy
      };
    } catch (e) {
      return null;   // a sheet deleted mid-listing should not break the list
    }
  }))).filter(Boolean);
  sheets.sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
  return json({ sheets });
}

async function createSheet(req) {
  const body = await req.json().catch(() => ({}));
  const by = who(body, req);
  const name = String(body.name || '').trim().slice(0, 120) || 'Parcel mapping';

  let parsed, bytes = null;
  if (body.workbook) {
    bytes = engine.b64ToBytes(body.workbook);
    if (bytes.length > MAX_UPLOAD) {
      throw fail(413, 'That file is larger than 25 MB. Please split it before uploading.');
    }
    parsed = await engine.parseWorkbook(bytes);
  } else {
    // start an empty proforma
    const wb = engine.ensureTemplate();
    const buf = await wb.xlsx.writeBuffer();
    bytes = new Uint8Array(buf);
    parsed = await engine.parseWorkbook(bytes);
  }

  const id = newId();
  const stamp = now();
  const sheet = {
    id,
    name,
    sheetName: parsed.sheetName,
    headerRow: parsed.headerRow,
    cols: parsed.cols,
    missing: parsed.missing,
    macro: parsed.macro,
    nextRow: parsed.nextRow,
    rev: 1,
    createdAt: stamp,
    createdBy: by,
    updatedAt: stamp,
    updatedBy: by,
    records: parsed.records.map((r) => ({
      ...r, rev: 1, updatedBy: by, updatedAt: stamp, touched: false
    }))
  };

  const s = store();
  await s.set(fileKey(id), bytes.buffer ? bytes.buffer : bytes);
  const { modified } = await s.setJSON(sheetKey(id), sheet,
    { onlyIfNew: true, metadata: listingMeta(sheet) });
  if (!modified) throw fail(500, 'Could not create the sheet. Please try again.');

  return json(snapshot(sheet), 201);
}

async function getSheet(id) {
  const data = await store().get(sheetKey(id), { type: 'json', consistency: 'strong' });
  if (!data) throw fail(404, 'That sheet no longer exists.');
  return json(snapshot(data));
}

async function getChanges(id, url) {
  const since = Number(url.searchParams.get('since') || 0);
  const data = await store().get(sheetKey(id), { type: 'json', consistency: 'strong' });
  if (!data) throw fail(404, 'That sheet no longer exists.');
  const changed = data.records.filter((r) => (r.rev || 0) > since).map(publicRecord);
  return json({
    rev: data.rev,
    records: changed,
    total: data.records.length,
    updatedAt: data.updatedAt,
    updatedBy: data.updatedBy
  });
}

async function addRecord(id, req) {
  const body = await req.json().catch(() => ({}));
  const by = who(body, req);
  const rec = engine.applyDefaults(body.record || {});
  const missing = engine.missingRequired(rec);
  if (missing.length) {
    throw fail(400, `${missing.join(' and ')} ${missing.length > 1 ? 'are' : 'is'} required.`);
  }

  return await mutate(sheetKey(id), (sheet) => {
    const clash = sheet.records.find((r) => engine.norm(r.parcel) === engine.norm(rec.parcel));
    if (clash) {
      return {
        abort: json({
          error: `P/S ${rec.parcel} already exists at row ${clash.row}` +
                 (clash.updatedBy ? ` (last saved by ${clash.updatedBy})` : '') +
                 '. Open it to update instead of adding a duplicate.',
          duplicate: true, row: clash.row, record: publicRecord(clash)
        }, 409)
      };
    }
    const rev = sheet.rev + 1;
    const row = Math.max(sheet.nextRow, ...sheet.records.map((r) => r.row + 1), sheet.headerRow + 1);
    const saved = { ...rec, row, rev, updatedBy: by, updatedAt: now(), touched: true };
    sheet.records.push(saved);
    sheet.records.sort((a, b) => a.row - b.row);
    sheet.rev = rev;
    sheet.nextRow = row + 1;
    sheet.updatedAt = saved.updatedAt;
    sheet.updatedBy = by;
    return {
      value: sheet,
      result: json({ record: publicRecord(saved), rev, action: 'added' }, 201)
    };
  });
}

async function updateRecord(id, row, req) {
  const body = await req.json().catch(() => ({}));
  const by = who(body, req);
  const baseRev = Number(body.baseRev);
  const rec = engine.applyDefaults(body.record || {});
  const missing = engine.missingRequired(rec);
  if (missing.length) {
    throw fail(400, `${missing.join(' and ')} ${missing.length > 1 ? 'are' : 'is'} required.`);
  }
  const target = parseInt(row, 10);
  if (!Number.isFinite(target)) throw fail(400, 'Invalid row reference.');

  return await mutate(sheetKey(id), (sheet) => {
    const i = sheet.records.findIndex((r) => r.row === target);
    if (i < 0) throw fail(404, `Row ${target} is no longer in this sheet.`);
    const mine = sheet.records[i];

    // Someone else saved this same parcel since it was opened - do not
    // overwrite their work, hand both versions back and let the user choose.
    if (Number.isFinite(baseRev) && (mine.rev || 0) !== baseRev) {
      return {
        abort: json({
          error: `${mine.updatedBy || 'Someone else'} changed this record while you were editing it.`,
          conflict: true,
          theirs: publicRecord(mine),
          by: mine.updatedBy,
          at: mine.updatedAt
        }, 409)
      };
    }

    // Changing a P/S to one that already exists elsewhere is still a duplicate.
    const clash = sheet.records.find(
      (r) => r.row !== target && engine.norm(r.parcel) === engine.norm(rec.parcel));
    if (clash) {
      return {
        abort: json({
          error: `P/S ${rec.parcel} already exists at row ${clash.row}.`,
          duplicate: true, row: clash.row, record: publicRecord(clash)
        }, 409)
      };
    }

    const rev = sheet.rev + 1;
    const saved = { ...mine, ...rec, row: target, rev, updatedBy: by, updatedAt: now(), touched: true };
    sheet.records[i] = saved;
    sheet.rev = rev;
    sheet.updatedAt = saved.updatedAt;
    sheet.updatedBy = by;
    return {
      value: sheet,
      result: json({ record: publicRecord(saved), rev, action: 'updated' })
    };
  });
}

async function exportSheet(id) {
  const s = store();
  const sheet = await s.get(sheetKey(id), { type: 'json', consistency: 'strong' });
  if (!sheet) throw fail(404, 'That sheet no longer exists.');
  const original = await s.get(fileKey(id), { type: 'arrayBuffer' });
  const buf = await engine.buildWorkbook(original ? new Uint8Array(original) : null, sheet);
  const safe = (sheet.name || 'Parcel_Mapping').replace(/[^A-Za-z0-9_-]+/g, '_').slice(0, 60);
  return new Response(buf, {
    status: 200,
    headers: {
      'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'content-disposition': `attachment; filename="${safe}.xlsx"`,
      'cache-control': 'no-store'
    }
  });
}

async function deleteSheet(id) {
  const s = store();
  await s.delete(sheetKey(id));
  await s.delete(fileKey(id));
  return json({ deleted: true });
}

/* ---------------------------------------------------------------- router */
export default async (req) => {
  const url = new URL(req.url);
  const path = url.pathname.replace(/^\/(\.netlify\/functions\/api|api)/, '').replace(/\/+$/, '');
  const parts = path.split('/').filter(Boolean);   // e.g. ['sheets','<id>','records','7']
  const method = req.method.toUpperCase();

  try {
    if (parts[0] === 'auth' && parts.length === 1) return await handleAuth(req, method);

    // Everything below needs a valid token, but only when a password is set.
    if (authRequired() && !validToken(req.headers.get('x-parcel-token'))) {
      throw fail(401, 'Please enter the team password.', { authRequired: true });
    }

    if (parts[0] !== 'sheets') throw fail(404, 'Unknown endpoint.');

    if (parts.length === 1) {
      if (method === 'GET') return await listSheets();
      if (method === 'POST') return await createSheet(req);
      throw fail(405, 'Method not allowed.');
    }

    const id = parts[1];
    if (parts.length === 2) {
      if (method === 'GET') return await getSheet(id);
      if (method === 'DELETE') return await deleteSheet(id);
      throw fail(405, 'Method not allowed.');
    }

    if (parts.length === 3 && parts[2] === 'changes' && method === 'GET') {
      return await getChanges(id, url);
    }
    if (parts.length === 3 && parts[2] === 'export' && method === 'GET') {
      return await exportSheet(id);
    }
    if (parts.length === 3 && parts[2] === 'records' && method === 'POST') {
      return await addRecord(id, req);
    }
    if (parts.length === 4 && parts[2] === 'records' && method === 'PUT') {
      return await updateRecord(id, parts[3], req);
    }
    throw fail(404, 'Unknown endpoint.');
  } catch (e) {
    if (e && e.status) return json(e.payload || { error: e.message }, e.status);
    console.error('api error', e);
    return json({ error: (e && e.message) || 'Unexpected server error.' }, 500);
  }
};

export const config = { path: '/api/*' };
