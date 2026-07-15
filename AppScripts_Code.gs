/**
 * Read-only API for the Hire Ledger dashboard.
 *
 * This script ONLY reads the sheet and returns it as JSON — there is no
 * write endpoint, on purpose. The dashboard is view-only; the sheet is
 * the single place edits happen.
 *
 * SETUP
 * 1. Row 1 of the active sheet must be a header row. Recommended headers,
 *    in any order (matching is case-insensitive and ignores spacing):
 *      Hire Name | Employer | Phone Number | Preferred Payment Method |
 *      Production Cycle | Basic Salary | Commission
 * 2. Deploy -> New deployment -> Web app -> Execute as: Me,
 *    Who has access: Anyone. Copy the /exec URL into the dashboard's
 *    "API URL" box.
 */

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();

  if (values.length < 1) {
    return jsonResponse({ rows: [], updatedAt: new Date().toISOString() });
  }

  var headers = values[0].map(function (h) {
    return String(h).trim();
  });

  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var isBlank = row.every(function (cell) {
      return cell === "" || cell === null;
    });
    if (isBlank) continue;

    var record = {};
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      record[headers[c]] = row[c];
    }
    rows.push(record);
  }

  return jsonResponse({ rows: rows, updatedAt: new Date().toISOString() });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
