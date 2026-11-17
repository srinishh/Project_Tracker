// Project Tracker backend — paste into Extensions > Apps Script of your Google Sheet.
// Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.
// If you already deployed, use Deploy > Manage deployments > Edit > New version.

const SHEET_NAME = "Projects";
const TOKEN = "change-me"; // must match API_TOKEN in index.html
const HEADERS = ["id", "name", "owner", "due", "status", "progress", "invoiced", "updatedAt"];

function doGet(e) {
  const sheet = getSheet_();
  const cols = headerMap_(sheet);
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // drop header row
  const projects = rows
    .filter(r => r[cols.id - 1] !== "")
    .map(r => toProject_(r, cols));
  return json_({ projects: projects });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return json_({ error: "unauthorized" });

    const sheet = getSheet_();
    if (body.action === "upsert") upsert_(sheet, body.project);
    else if (body.action === "delete") remove_(sheet, body.id);
    else return json_({ error: "unknown action" });
    return json_({ ok: true });
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function toProject_(r, cols) {
  const get = name => (cols[name] ? r[cols[name] - 1] : "");
  return {
    id: String(get("id")),
    name: String(get("name")),
    owner: String(get("owner")),
    due: formatDue_(get("due")),
    status: String(get("status")) || "Not started",
    progress: Number(get("progress")) || 0,
    invoiced: get("invoiced") === true || String(get("invoiced")).toUpperCase() === "TRUE"
  };
}

function upsert_(sheet, p) {
  const cols = headerMap_(sheet);
  const width = sheet.getLastColumn();
  const values = {
    id: p.id,
    name: p.name,
    owner: p.owner || "",
    due: p.due || "",
    status: p.status,
    progress: p.progress,
    invoiced: p.invoiced === true,
    updatedAt: new Date()
  };
  const row = new Array(width).fill("");
  Object.keys(cols).forEach(name => { if (name in values) row[cols[name] - 1] = values[name]; });

  const data = sheet.getDataRange().getValues();
  const idCol = cols.id - 1;
  const rowIndex = data.findIndex((r, i) => i > 0 && String(r[idCol]) === String(p.id));
  if (rowIndex > 0) {
    sheet.getRange(rowIndex + 1, 1, 1, width).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function remove_(sheet, id) {
  const cols = headerMap_(sheet);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][cols.id - 1]) === String(id)) sheet.deleteRow(i + 1);
  }
}

// Maps header name -> 1-based column number, so column order in the sheet doesn't matter.
function headerMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h)] = i + 1; });
  return map;
}

function formatDue_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v ? String(v) : "";
}

// Creates the sheet if missing, and appends any header columns the sheet doesn't have yet.
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    return sheet;
  }
  if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);
  const existing = headerMap_(sheet);
  HEADERS.forEach(h => {
    if (!existing[h]) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(h);
    }
  });
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
