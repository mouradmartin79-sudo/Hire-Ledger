/**
 * Read-only API for the Hire Ledger dashboard + Payments tracker.
 *
 * This script ONLY reads the sheet and returns it as JSON — there is no
 * write endpoint, on purpose. Both pages are view-only; the sheet is
 * the single place edits happen.
 *
 * SETUP
 * 1. Tab 1 (any name, defaults to the first sheet) is the Hire Ledger.
 *    Recommended headers, in any order (matching is case-insensitive and
 *    ignores spacing):
 *      Hire Name | Employer | Phone Number | Preferred Payment Method |
 *      Production Cycle | Basic Salary | Commission | Status
 *    ("Status" should contain Active / Inactive.)
 *
 * 2. Add a second tab named exactly "Payments" with headers:
 *      Setter Name | Month | Amount | Payment Status
 *    ("Payment Status" should contain Paid / Pending / Overdue.)
 *
 * 3. Deploy -> New deployment -> Web app -> Execute as: Me,
 *    Who has access: Anyone. Copy the /exec URL into each page's
 *    "API URL" box (Settings). The Payments page automatically requests
 *    ?sheet=Payments — you don't need to do anything extra.
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requestedSheet = e && e.parameter && e.parameter.sheet;

  var sheet = requestedSheet
    ? ss.getSheetByName(requestedSheet)
    : ss.getSheets()[0];

  if (!sheet) {
    return jsonResponse({
      rows: [],
      updatedAt: new Date().toISOString(),
      error: "Sheet not found: " + requestedSheet
    });
  }

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
