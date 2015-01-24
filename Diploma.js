// Scans sheet for errors

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


// Fix errors in the schedule

function fix_err() {

  Logger.log("Fix Errors");

}


// Send emails about the schedule to teachers

function exporter() {

  state_check_email = false;
  while (state_check_email !== true ) {
    if (state_check_email === false) {
      var email = Browser.inputBox("Input a valid email of teacher", Browser.Buttons.OK_CANCEL);
      var email_validity_pattern = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
      var validity_result = email_validity_pattern.test(email);
      if (!validity_result) {
        Browser.msgBox("Invalid email, try again");
      } else {
        var url_fetch = SpreadsheetApp.getActiveSpreadsheet();
        var url = url_fetch.getUrl();
        var add_viewer_address = SpreadsheetApp.getActiveSpreadsheet();
        add_viewer_address.addViewer(email);
        var subject = "Welcome";
        MailApp.sendEmail(email, subject, "Welcome --> This is the url to the schedule " + url);
        Browser.msgBox("Success");
        var state_check_email = true;
      }
    }
  }

}


// Menu in spreadsheet implementation

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
