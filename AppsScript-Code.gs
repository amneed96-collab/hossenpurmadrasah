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

function parseCell(val) {
  if (typeof val === "string" && (val.startsWith("[") || val.startsWith("{"))) {
    try { return JSON.parse(val); } catch (e) { return val; }
  }
  return val;
}

function listRows(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return { rows: [] };
  const headers = getHeaders(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const rowArr = values[i];
    if (!rowArr[0]) continue; // skip rows without id
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = parseCell(rowArr[idx]);
    });
    rows.push(obj);
  }
  return { rows };
}

function bulkGet(sheetsCsv) {
  const names = (sheetsCsv || "").split(",").map(s => s.trim()).filter(Boolean);
  const out = {};
  names.forEach(n => {
    out[n] = listRows(n).rows;
  });
  return { data: out };
}

function upsertRow(sheetName, rowJson) {
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
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === row.id) {
        foundRowIndex = i + 2;
        break;
      }
    }
  }

  const rowValues = headers.map(h => cellValue(row[h]));

  if (foundRowIndex > -1) {
    sheet.getRange(foundRowIndex, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([rowValues]);
  }
  return { success: true, id: row.id };
}

function deleteRow(sheetName, id) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: "Row not found" };
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, error: "Row not found" };
}
