/**
 * Read-only API for the Scale Setters dashboard.
 *
 * This script ONLY reads sheets and returns them as JSON — there is no
 * write endpoint, on purpose. The dashboard is view-only; the spreadsheet
 * is the single place edits happen.
 *
 * SETUP
 * Your spreadsheet should have these tabs, named exactly:
 *
 *   "Hires" — one row per setter. Header row (row 1), any order:
 *     Hire Name | Employer | Phone Number | Preferred Payment Method |
 *     Production Cycle | Basic Salary | Commission | Active
 *   ("Active" accepts Yes/No, TRUE/FALSE, or Active/Inactive.)
 *
 *   One tab per month — named exactly "January", "February", ... "December".
 *   Each setter appears ONCE per month tab (no repeated rows). Header row:
 *     Setter Name | Amount | Status
 *   (Setter Name should match the name used in the Hires tab. Status is
 *   "Paid" or "Pending".)
 *
 *   Only create the month tabs you're actually using — any missing ones
 *   are simply skipped, nothing breaks.
 *
 * DEPLOY
 * Deploy -> New deployment -> Web app -> Execute as: Me,
 * Who has access: Anyone. Copy the /exec URL into the dashboard's
 * Settings panel (password-protected on the site).
 *
 * If you edit this script later, use Deploy -> Manage deployments ->
 * Edit -> New version (not "New deployment") so the URL stays the same.
 */

var MONTH_TABS = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hiresSheet = ss.getSheetByName("Hires") || ss.getSheets()[0];
  var hires = sheetToRecords(hiresSheet);

  var payments = [];
  MONTH_TABS.forEach(function (monthName) {
    var sheet = ss.getSheetByName(monthName);
    if (!sheet) return;
    var records = sheetToRecords(sheet);
    records.forEach(function (r) {
      r.Month = monthName;
      payments.push(r);
    });
  });

  // Backward compatibility: still read a "Payments" tab if someone kept one
  var legacySheet = ss.getSheetByName("Payments");
  if (legacySheet) {
    payments = payments.concat(sheetToRecords(legacySheet));
  }

  return jsonResponse({
    hires: hires,
    payments: payments,
    updatedAt: new Date().toISOString()
  });
}

function sheetToRecords(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 1) return [];

  var headers = values[0].map(function (h) {
    return String(h).trim();
  });

  var records = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var isBlank = row.every(function (cell) {
      return cell === "" || cell === null;
    });
    if (isBlank) continue;

    var record = {};
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      var val = row[c];
      if (Object.prototype.toString.call(val) === "[object Date]") {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      record[headers[c]] = val;
    }
    records.push(record);
  }
  return records;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
