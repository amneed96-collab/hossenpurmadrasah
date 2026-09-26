/**
 * SCHOOL MANAGEMENT SYSTEM — Google Apps Script Backend
 * -------------------------------------------------------
 * SETUP:
 * 1. Go to https://script.google.com -> New Project
 * 2. Delete any starter code, paste this entire file in
 * 3. Click "Deploy" -> "New deployment"
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 4. Copy the Web App URL it gives you (ends in /exec)
 * 5. Paste that URL into the School App's Settings screen
 *
 * This script auto-creates all needed sheets/tabs on first run.
 * Every field is stored in its OWN column (auto-expanding headers) —
 * not squeezed into a single JSON column — so the Google Sheet itself
 * is directly readable/editable like a normal spreadsheet.
 */

function doGet(e) {
  return handle(e);
}
function doPost(e) {
  return handle(e);
}

function handle(e) {
  try {
    let params = {};
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      params = e.parameter;
    }
    const action = params.action;
    let result;

    switch (action) {
      case "list":
        result = listRows(params.sheet);
        break;
      case "upsert":
        result = upsertRow(params.sheet, params.row);
        break;
      case "delete":
        result = deleteRow(params.sheet, params.id);
        break;
      case "bulkGet":
        result = bulkGet(params.sheets);
        break;
      case "uploadImage":
        result = uploadImage(params.folder, params.filename, params.base64, params.mimeType);
        break;
      default:
        result = { error: "Unknown action: " + action };
    }
    return respond(result);
  } catch (err) {
    return respond({ error: err.toString() });
  }
}

// Saves an uploaded image to Google Drive (in a folder inside "School App Uploads")
// and returns a direct, publicly-viewable URL. Sheets cells cannot hold large
// base64 image data (50,000 char/cell limit), so images must live in Drive —
// only the short URL is ever stored in a sheet cell.
function uploadImage(folder, filename, base64, mimeType) {
  const rootName = "School App Uploads";
  const folders = DriveApp.getFoldersByName(rootName);
  const root = folders.hasNext() ? folders.next() : DriveApp.createFolder(rootName);

  const subName = folder || "misc";
  const subFolders = root.getFoldersByName(subName);
  const sub = subFolders.hasNext() ? subFolders.next() : root.createFolder(subName);

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", filename || ("upload_" + Date.now()));
  const file = sub.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const id = file.getId();
  // Direct-embeddable image URL (works in <img src="...">)
  const url = "https://lh3.googleusercontent.com/d/" + id;
  return { success: true, url: url, fileId: id };
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(["id", "createdAt", "updatedAt"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

// Ensures every key in `row` has a matching column; adds new columns as needed.
function ensureColumns(sheet, row) {
  let headers = getHeaders(sheet);
  if (headers.length === 0) {
    headers = ["id", "createdAt", "updatedAt"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const existing = new Set(headers);
  const newKeys = Object.keys(row).filter(k => !existing.has(k));
  if (newKeys.length > 0) {
    const startCol = headers.length + 1;
    sheet.getRange(1, startCol, 1, newKeys.length).setValues([newKeys]);
    headers = headers.concat(newKeys);
  }
  return headers;
}

// Nested objects/arrays (e.g. payments: [...]) are JSON-stringified into their single cell
// since a spreadsheet cell can't hold a nested structure — every top-level field
// (name, class, roll, phone, etc.) still gets its own column.
function cellValue(val) {
  if (val === undefined || val === null) return "";
  let out = typeof val === "object" ? JSON.stringify(val) : String(val);
  // Google Sheets hard limit is 50,000 characters per cell — guard against
  // accidentally pasting raw base64 data here (images must go through uploadImage instead).
  if (out.length > 45000) out = out.slice(0, 45000) + "…[TRUNCATED:TOO_LARGE]";
  return out;
}

// Sheets auto-detects date-looking text (e.g. "2026-01-10") typed via setValues() and silently
// stores it as a real Date cell, exactly like typing it into the UI would. When that happens,
// getValues() hands back a native Date object instead of the original string, and by the time
// it's JSON-stringified for the client it becomes a full ISO datetime (e.g.
// "2026-01-10T00:00:00.000Z") instead of the plain "yyyy-MM-dd" the app's date inputs expect —
// which is why an edited record could appear to be missing its Admission Date. This restores
// any such Date cell back into plain text before it ever leaves the server.
function normalizeCellForOutput(val) {
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    const tz = Session.getScriptTimeZone() || "UTC";
    const hasTime = val.getHours() !== 0 || val.getMinutes() !== 0 || val.getSeconds() !== 0;
    return hasTime
      ? Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss")
      : Utilities.formatDate(val, tz, "yyyy-MM-dd");
  }
  return val;
}

function parseCell(val) {
  val = normalizeCellForOutput(val);
  if (typeof val === "string" && (val.startsWith("[") || val.startsWith("{"))) {
    try { return JSON.parse(val); } catch (e) { return val; }
  }
  return val;
}

// Results are cached (Apps Script's built-in CacheService) so that repeat loads — which is
// what "Sync now" and every screen that reads data does — come back almost instantly instead
// of re-reading the spreadsheet every time. The cache is invalidated for a sheet the moment a
// row in it is saved or deleted (see upsertRow/deleteRow), so nobody ever sees stale data —
// they just don't pay the full Sheets-read cost on every single load.
const LIST_CACHE_SECONDS = 300; // 5 minutes
function listCacheKey(sheetName) { return "list_" + sheetName; }

function listRows(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = listCacheKey(sheetName);
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (e) {
    // corrupt/unreadable cache entry — fall through and read fresh from the sheet
  }

  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return { rows: [] };

  // One single getValues() call for header + all data rows together, instead of two
  // separate range reads (header, then data) — halves the number of slow Sheets API calls.
  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0];
  const rows = [];
  for (let i = 1; i < all.length; i++) {
    const rowArr = all[i];
    if (!rowArr[0]) continue; // skip rows without id
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = parseCell(rowArr[idx]);
    });
    rows.push(obj);
  }
  const result = { rows };

  try {
    const json = JSON.stringify(result);
    // CacheService caps each value at 100KB; only cache if it comfortably fits.
    if (json.length < 95000) cache.put(cacheKey, json, LIST_CACHE_SECONDS);
  } catch (e) {
    // if a row contains something unstringify-able, just skip caching it — not fatal
  }
  return result;
}

function invalidateListCache(sheetName) {
  try { CacheService.getScriptCache().remove(listCacheKey(sheetName)); } catch (e) {}
}

function bulkGet(sheetsCsv) {
  const names = (sheetsCsv || "").split(",").map(s => s.trim()).filter(Boolean);
  const out = {};
  names.forEach(n => {
    out[n] = listRows(n).rows;
  });
  return { data: out };
}

// Finds which column actually holds "id" by its header text, instead of assuming it's always
// column 1. Hardcoding column 1 silently reads/deletes the WRONG column's values if a column was
// ever manually inserted before it in the real Google Sheet (a helper/sort column, a reordered
// layout, importing from another source, etc.) — which would make id-matching never find the
// real row (or, worse, coincidentally match something else). Falls back to column 1 only when no
// "id" header exists at all (shouldn't normally happen — getSheet/ensureColumns always put it
// first on a sheet this app created — but a sheet edited outside the app could lack it).
function idColumnIndex(headers) {
  const idx = headers.indexOf("id");
  return idx > -1 ? idx : 0;
}

// upsertRow/deleteRow both do a "read every id in the sheet, then write" sequence. Without a
// lock, two requests arriving close together (a slow mobile connection retrying a request, two
// staff saving at nearly the same moment, an auto-sync overlapping a manual Save) can both read
// the sheet BEFORE either one's write lands — each one then thinks "no existing row with this
// id" and both append a brand-new row, which is exactly what shows up as a record being
// duplicated on Save, or a Delete silently not removing anything because the row index shifted
// under it mid-operation. Wrapping the whole read-check-write sequence in a script lock makes it
// atomic: only one upsert/delete against this spreadsheet runs at a time, so this can't happen.
function upsertRow(sheetName, rowJson) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet(sheetName);
    const row = typeof rowJson === "string" ? JSON.parse(rowJson) : rowJson;
    if (!row.id) {
      row.id = Utilities.getUuid();
    }
    const now = new Date().toISOString();
    if (!row.createdAt) row.createdAt = now;
    row.updatedAt = now;

    const headers = ensureColumns(sheet, row);

    const lastRow = sheet.getLastRow();
    let foundRowIndex = -1;
    if (lastRow >= 2) {
      // Compare as strings — a value typed as a plain number (e.g. an old id that happens to be
      // all digits) can come back from getValues() as a JS Number while row.id sent from the
      // client is always a String; a strict === would miss that match and append a duplicate.
      const idCol = idColumnIndex(headers) + 1; // getRange columns are 1-based
      const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
      const targetId = String(row.id);
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === targetId) {
          foundRowIndex = i + 2;
          break;
        }
      }
    }

    const rowValues = headers.map(h => cellValue(row[h]));

    const targetRow = foundRowIndex > -1 ? foundRowIndex : sheet.getLastRow() + 1;
    const range = sheet.getRange(targetRow, 1, 1, headers.length);
    // Plain-text format BEFORE writing — otherwise Sheets auto-detects date/number-looking
    // strings (like Admission Date "2026-01-10") and silently converts the cell to a real
    // Date/Number type, which breaks the app's <input type="date"> fields on the next edit.
    range.setNumberFormat("@");
    range.setValues([rowValues]);
    invalidateListCache(sheetName);
    return { success: true, id: row.id };
  } finally {
    lock.releaseLock();
  }
}

function deleteRow(sheetName, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    // NOTE: deliberately no "error" key on the not-found outcome below. The frontend's
    // sheetsCall() treats ANY response with a truthy "error" field as a hard failure and throws
    // (showing a scary "Delete failed on server" toast) — but "there was no row with this id to
    // delete" is a perfectly normal, benign outcome (e.g. a student who never had a Promotion
    // row yet, or a row already removed by another device/tab), not a real server error. Using
    // "notFound" instead of "error" keeps that distinction so callers can tell the difference.
    if (lastRow < 2) return { success: false, notFound: true };
    const headers = getHeaders(sheet);
    const idCol = idColumnIndex(headers) + 1;
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    const targetId = String(id);
    // Delete EVERY row matching this id, not just the first one found. If a duplicate row with
    // the same id was ever left behind — e.g. from before the lock above existed, when a slow or
    // retried Save could append a second row instead of updating the first — deleting only "the"
    // first match would leave the duplicate behind, and the student would look un-deleted the
    // next time the app re-reads the Sheet even though this delete genuinely ran and reported
    // success. Deleting bottom-to-top keeps the earlier row numbers valid as rows shift up.
    const matchingRows = [];
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === targetId) matchingRows.push(i + 2);
    }
    if (matchingRows.length === 0) return { success: false, notFound: true };
    matchingRows.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));
    invalidateListCache(sheetName);
    return { success: true, deletedCount: matchingRows.length };
  } finally {
    lock.releaseLock();
  }
}
