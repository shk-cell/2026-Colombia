function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('미사일 피하기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, user-scalable=no');
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('기록');
  if (!sheet) {
    sheet = ss.insertSheet('기록');
    sheet.appendRow(['이름', '점수', '일시']);
  }
  return sheet;
}

function saveScore(name, score) {
  var sheet = getSheet_();
  var safeName = String(name || '익명').slice(0, 20);
  var safeScore = Math.max(0, Math.floor(Number(score) || 0));
  sheet.appendRow([safeName, safeScore, new Date()]);
  return true;
}

function getRanking() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  data.sort(function (a, b) { return b[1] - a[1]; });
  return data.slice(0, 10);
}
