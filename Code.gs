/**
 * Serve the Web App
 */
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Google Apps Script Calculator")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Include HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Save every calculation to Google Sheets
 */
function saveCalculation(expression, result) {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Logs");

  const now = new Date();

  const date = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "dd-MMM-yyyy"
  );

  const time = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "HH:mm:ss"
  );

  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 4).setValues([
    [date, time, expression, result]
  ]);
  sheet.getRange(row, 1, 1, 4)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  return true;
}
