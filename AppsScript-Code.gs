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
 * No manual sheet setup required.
 */

const SHEET_NAMES = {
  students: "Students",
  teachers: "Teachers",
  finance: "Finance",
  fees: "MasterFees",
  committee: "Committee",
  seatplans: "SeatPlans",
  meta: "Meta"
};

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
      default:
        result = { error: "Unknown action: " + action };
    }
    return respond(result);
  } catch (err) {
    return respond({ error: err.toString() });
  }
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
    sheet.appendRow(["id", "data", "createdAt", "updatedAt"]);
  }
  return sheet;
}

function listRows(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const [id, data, createdAt, updatedAt] = values[i];
    if (!id) continue;
    try {
      rows.push({ id, ...JSON.parse(data), createdAt, updatedAt });
    } catch (err) {
      // skip malformed row
    }
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
  const values = sheet.getDataRange().getValues();
  let foundRowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === row.id) {
      foundRowIndex = i + 1; // 1-indexed sheet row
      break;
    }
  }
  const now = new Date().toISOString();
  const { id, ...rest } = row;
  const dataStr = JSON.stringify(rest);

  if (foundRowIndex > -1) {
    sheet.getRange(foundRowIndex, 2).setValue(dataStr);
    sheet.getRange(foundRowIndex, 4).setValue(now);
  } else {
    const createdAt = now;
    sheet.appendRow([id, dataStr, createdAt, now]);
  }
  return { success: true, id: row.id };
}

function deleteRow(sheetName, id) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Row not found" };
}
