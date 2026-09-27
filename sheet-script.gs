// Pre-transplant time: shared log.
// Paste this into Extensions > Apps Script of a new Google Sheet, then Deploy > New deployment > Web app,
// Execute as: Me, Who has access: Anyone. Copy the web app URL into SHEET_URL in index.html.

const HEAD = ["id", "date", "time", "physician", "activity", "minutes", "received"];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Log");
  if (!sh) { sh = ss.insertSheet("Log"); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const tz = Session.getScriptTimeZone();
  const fmt = (v, p) => v instanceof Date ? Utilities.formatDate(v, tz, p) : String(v);
  return out_({ ok: true, entries: rows.filter(r => r[0] !== "").map(r => ({
    id: String(r[0]), date: fmt(r[1], "yyyy-MM-dd"), time: fmt(r[2], "HH:mm"),
    doctor: String(r[3]), activity: String(r[4]), minutes: Number(r[5]) || 0
  })) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const sh = sheet_();
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => String(r[0]));
    const adds = body.add || [];
    const dels = body.remove || [];
    const newRows = [];
    adds.forEach(x => {
      const id = String(x.id);
      if (!id || ids.indexOf(id) >= 0) return;
      ids.push(id);
      newRows.push(["'" + id, "'" + x.date, "'" + x.time, x.doctor, x.activity, Number(x.minutes) || 0, new Date()]);
    });
    if (newRows.length) sh.getRange(sh.getLastRow() + 1, 1, newRows.length, HEAD.length).setValues(newRows);
    dels.map(String).forEach(id => {
      const i = ids.indexOf(id);
      if (i > 0) { sh.deleteRow(i + 1); ids.splice(i, 1); }
    });
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
