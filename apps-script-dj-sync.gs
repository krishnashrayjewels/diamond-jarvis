// Diamond Jarvis -> KV HISAB (Google Drive) auto-sync
// Ek vaar deploy karnu he, pachi app na entries aapo-aapo KV HISAB ma save thase.
// Deploy steps: Deploy > New deployment > gear icon > Web app
//   Description: Diamond Jarvis sync
//   Execute as: Me
//   Who has access: Anyone
// Authorize: Review permissions > account > Advanced > Go to ... > Allow
// Pachi "Web app URL" copy kari ne Diamond Jarvis app na Settings ma Drive Sync URL ma paste karo.
var SHEET_ID = "1RWcEmDj4qt2QGmGWFL_b0AQh3YcG5DH83AwIbJ50DII"; // KV HISAB
var TAB = "Diamond Jarvis";
var COLS = ["date","type","party","billable_cts","bill","dalali","batav","angadia","net_cash","pieces","vf_payable","emi_fund","profit_buffer","landed70","breakeven70","due_days","saved_at"];

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName(TAB);
    if (!sh) sh = ss.insertSheet(TAB);
    if (sh.getLastRow() === 0) sh.appendRow(COLS);
    var row = [];
    for (var i = 0; i < COLS.length - 1; i++) {
      var v = d[COLS[i]];
      row.push(v === undefined || v === null ? "" : v);
    }
    row.push(new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" }));
    sh.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, err: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Diamond Jarvis sync app is running").setMimeType(ContentService.MimeType.TEXT);
}
