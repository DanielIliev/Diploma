// Create addon menu for user

function start() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Fix Duplicates', 'duplicates')
    .addSeparator()
    .addItem('Cabinet Occupation Check', 'cabinet_occupation')
    .addSeparator()
    .addItem('Teachers Schedule Check', 'teachers_schedule')
    .addSeparator()
    .addItem('Send schedule to email', 'send_to_email')
    .addToUi();
}


// Fix duplicates in schedule - Task 1 - Not Finished

function duplicates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var range = sheet.getRange("A1:AJ52");
  get_data(sheet);
};

function get_data(current_sheet) {
  var data_range = current_sheet.getDataRange();
  var values = data_range.getValues();
  var rows = data_range.getNumRows();
  var cols = data_range.getNumColumns();
  var timestamp_pattern = /[0-9]+[:]+[0-9]+/;
  var digits = /[0-9]+/;
  var cyrillic = /[А-Я]+/;
  
  var clean_arr = []
  for (var row = 0; row < rows; row++) {
    clean_arr.push([]);
  }
  
  for (var row = 0; row < rows; row++) {
    for (var col = clean_arr[row].length; col < cols; col++) {
      clean_arr[row].push("");
    }
  }
  
  for (var row = 2; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      var bad_data_check = values[row][col].toString().search(timestamp_pattern) + 
        values[row][col].toString().search(digits);
      var cyrillic_check = values[row][col].toString().search(cyrillic);
      if (bad_data_check < 0 && 
        cyrillic_check >= 0 &&
        values[row][col].length > 3) {
        var temp_string = values[row][col].toString();
        clean_arr[row].push(temp_string);
      }
    }
  }
  Logger.log(clean_arr);
};

// Check for occupied cabinets - Task 2 - Not Finished

function cabinet_occupation() {
  Logger.log("Cabinet Occupation func invoked");
}


// Check teachers schedule for holes - Task 3 - Not Finished

function teachers_schedule() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data_range = sheet.getDataRange();
  var rows = data_range.getNumRows();
  var cols = data_range.getNumColumns();
  var values = data_range.getValues();
  var range = sheet.getRange("A1:AJ52");
  var initials = /[А-Я]+/;
  
  for (var i = 0; i < rows; i++) {
    for (var j = 0; j < cols; j++) {
      if (values[i][j].length <= 3) {
        var check_initials = values[i][j].toString().search(initials);
        if (check_initials != -1) {
          var cell = range.getCell(i+1,j+1);
          cell.setBackground("red");
        }
      }
    }
  }
}

// Check email validity

function send_to_email() {  
  state_check_email = false;
  while (state_check_email !== true ) {
    if (state_check_email === false) {
      var email = Browser.inputBox("Input a valid email of teacher", Browser.Buttons.OK);
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

};

// Custom functions used for fixing errors in tasks

function cleanArray(current){
  var newArray = new Array();
  for(var i = 0; i < current.length; i++){
    if (current[i]) {
      newArray.push(current[i]);
    }
  }
  
  return newArray;
}
