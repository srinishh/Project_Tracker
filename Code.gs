// Project Tracker backend — paste into Extensions > Apps Script of your Google Sheet.
// Deploy: Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone.

const SHEET_NAME = "Projects";
const TOKEN = "change-me"; // must match API_TOKEN in index.html
const HEADERS = ["id", "name", "owner", "due", "status", "progress", "updatedAt"];

function doGet(e) {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // drop header
  const projects = rows
    .filter(r => r[0] !== "")
    .map(r => ({
      id: String(r[0]),
      name: String(r[1]),
      owner: String(r[2]),
      due: formatDue_(r[3]),
      status: String(r[4]),
      progress: Number(r[5]) || 0
    }));
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

function upsert_(sheet, p) {
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((r, i) => i > 0 && String(r[0]) === String(p.id));
  const row = [p.id, p.name, p.owner || "", p.due || "", p.status, p.progress, new Date()];
  if (rowIndex > 0) {
    sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

function remove_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][0]) === String(id)) sheet.deleteRow(i + 1);
  }
}

function formatDue_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v ? String(v) : "";
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
