// Google Apps Script backend for the Grade 7 Motion lesson.
// Setup: create a Google Sheet > Extensions > Apps Script > paste this file.
// Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
// Copy the web app URL into SCRIPT_URL in index.html.

const SHEET_NAME = "Submissions";
const ITEM_KEYS = [];
["dd:6", "sv:7", "gr:6", "vr:5", "quiz:6"].forEach(function (s) {
  var p = s.split(":");
  for (var i = 1; i <= +p[1]; i++) ITEM_KEYS.push(p[0] + i);
});

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      var head = ["Timestamp", "Name", "Section", "Date", "Score", "Total", "Answered"]
        .concat(ITEM_KEYS)
        .concat(["Distance (m)", "Time east (s)", "Time west (s)", "Independent var", "Controlled vars", "Reflection"]);
      sh.appendRow(head);
      sh.setFrozenRows(1);
    }
    var a = d.answers || {};
    var items = ITEM_KEYS.map(function (k) {
      return a[k] ? a[k].answer + (a[k].correct ? " ✓" : " ✗") : "";
    });
    var act = d.activity || {};
    sh.appendRow(
      [d.timestamp, d.name, d.section, d.date, d.score, d.total, d.answered]
        .concat(items)
        .concat([act.distance_m, act.time_east_s, act.time_west_s, act.independent, act.controlled, d.reflection])
    );
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput("Motion lesson backend is running.");
}
