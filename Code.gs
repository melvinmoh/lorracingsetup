/**
 * LOR Racing Setup Database — Google Sheets sync endpoint.
 *
 * One row per record (championship, driver, round, session), matched by unique id.
 *  - New id            → appended as a new row (rev 1)
 *  - Known id          → that one row is updated, rev + 1 (only if the device edited the latest version)
 *  - Same data resent  → no change (repeat syncs never create duplicates)
 *  - Edited elsewhere  → "conflict": nothing is written; the device asks the user which version to keep
 *  - Deletes           → row kept, "deleted" set to TRUE. Rows are never removed and the sheet is never rewritten.
 *
 * Protected by a secret token stored in Script Properties (SYNC_TOKEN). Run setup() once to create it.
 */

const SHEET_NAME = 'Records';
const HEADERS = ['id', 'type', 'parentId', 'label', 'rev', 'deleted', 'updatedAt', 'updatedBy', 'data'];
const COL = { id: 0, type: 1, parentId: 2, label: 3, rev: 4, deleted: 5, updatedAt: 6, updatedBy: 7, data: 8 };
const TYPES = ['championship', 'driver', 'race', 'session'];

/** Run once from the editor: creates the Records sheet and the sync token (shown in the Execution log). */
function setup() {
  sheet_();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SYNC_TOKEN')) {
    props.setProperty('SYNC_TOKEN', (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
  }
  Logger.log('Sync token (paste into the app ⚙ Sync Settings): ' + props.getProperty('SYNC_TOKEN'));
}

/** Run this only if the token leaks. Every device then needs the new token. */
function rotateToken() {
  PropertiesService.getScriptProperties().deleteProperty('SYNC_TOKEN');
  setup();
}

function doGet() {
  return json_({ ok: true, service: 'LOR setup sync' }); // health check only — no data without the token
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!tokenOk_(req.token)) return json_({ ok: false, error: 'unauthorized' });
    const lock = LockService.getScriptLock();
    lock.waitLock(30000); // one sync at a time
    try { out = sync_(req); } finally { lock.releaseLock(); }
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return json_(out);
}

function sync_(req) {
  const sh = sheet_();
  const width = HEADERS.length;
  const existing = Math.max(sh.getLastRow() - 1, 0);
  const rows = existing ? sh.getRange(2, 1, existing, width).getValues() : [];
  const index = {};
  rows.forEach((r, i) => { if (r[COL.id] !== '') index[String(r[COL.id])] = i; });

  const device = String(req.device || 'unknown').slice(0, 60);
  const now = new Date();
  const touched = {};
  const results = [];

  (Array.isArray(req.changes) ? req.changes : []).forEach(c => {
    const id = String((c && c.id) || '');
    const data = String((c && c.data) || '');
    if (!/^[a-z0-9]{6,40}$/i.test(id) || TYPES.indexOf(c.type) < 0 || !data || data.length > 45000) {
      results.push({ id, status: 'error', error: 'invalid record' });
      return;
    }
    const deleted = c.deleted === true;
    const row = [id, c.type, String(c.parentId || ''), String(c.label || '').slice(0, 300), 1, deleted, now, device, data];
    const i = index[id];

    if (i === undefined) {                                    // new record → new row
      rows.push(row); index[id] = rows.length - 1; touched[rows.length - 1] = true;
      results.push({ id, status: 'ok', rev: 1 });
      return;
    }
    const cur = rows[i];
    const rev = Number(cur[COL.rev]) || 0;
    if (String(cur[COL.data]) === data && (cur[COL.deleted] === true) === deleted) {
      results.push({ id, status: 'ok', rev });                // already stored → no duplicate, no change
      return;
    }
    if (rev !== Number(c.baseRev)) {                          // changed on another device since this one last synced
      results.push({ id, status: 'conflict', server: rec_(cur) });
      return;
    }
    row[COL.rev] = rev + 1; rows[i] = row; touched[i] = true;  // update this one row
    results.push({ id, status: 'ok', rev: rev + 1 });
  });

  // Write only the rows that changed; append new ones below.
  Object.keys(touched).map(Number).filter(i => i < existing)
    .forEach(i => sh.getRange(i + 2, 1, 1, width).setValues([rows[i]]));
  if (rows.length > existing) {
    sh.getRange(existing + 2, 1, rows.length - existing, width).setValues(rows.slice(existing));
  }
  return { ok: true, results, records: rows.filter(r => r[COL.id] !== '').map(rec_) };
}

function rec_(r) {
  const t = r[COL.updatedAt];
  return {
    id: String(r[COL.id]), type: String(r[COL.type]), parentId: String(r[COL.parentId]), label: String(r[COL.label]),
    rev: Number(r[COL.rev]) || 0, deleted: r[COL.deleted] === true,
    updatedAt: t instanceof Date ? t.toISOString() : String(t || ''), updatedBy: String(r[COL.updatedBy]),
    data: String(r[COL.data]),
  };
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:D').setNumberFormat('@'); // ids, labels, data stored as plain text
    sh.getRange('H:I').setNumberFormat('@');
  }
  return sh;
}

function tokenOk_(given) {
  const token = PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN');
  if (!token || typeof given !== 'string' || given.length !== token.length) return false;
  let diff = 0;
  for (let i = 0; i < token.length; i++) diff |= token.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
