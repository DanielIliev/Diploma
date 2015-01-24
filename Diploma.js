/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function scan_err() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};

function fix_err() {
  Logger.log("Fix Errors");
}

function exporter() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var email_addresses = [];
  var number_of_emails = Browser.inputBox("Number of emails", Browser.Buttons.OK_CANCEL);
  var number_of_emails = parseInt(number_of_emails);
  for (var i = 0; i <= numRows - 1; i++) {
    var message = values;       // Data from spreadsheet
    var subject = "Sending emails from a Spreadsheet";
    MailApp.sendEmail(emailAddress, subject, message);
  }
}
/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function application_start() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Scan for errors",
    functionName : "scan_err",
  },null,
  {
    name : "Fix errors",
    functionName : "fix_err",
  },null,
  {
    name : "Send schedule to teachers",
    functionName : "exporter",
  }];
  spreadsheet.addMenu("Schedule_fixer", entries);
};