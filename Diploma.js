function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('График учители', 'teachers')
  .addToUi()
}

function teachers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var num_of_sheets = sheet.length;
  var rows = 0;
  var cols = 0;
  var result = [];
  var schedule_sheet_height = 0;
  var initials = Browser.inputBox('Въведи инициали на учител');
  for (var i = 0; i < num_of_sheets; i++) {
    rows = sheet[i].getDataRange().getNumRows();
    cols = sheet[i].getDataRange().getNumColumns();
    var values = sheet[i].getDataRange().getValues();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        if (values[row][col] == initials) {
          var temp_arr = [];
          temp_arr.push(values[row][col].toString());
          result.push(temp_arr);
          schedule_sheet_height++;
        }
      }
    }
  }
  var schedule = SpreadsheetApp.getActiveSpreadsheet()
    .insertSheet('График Учител', 5)
    .getRange(1,1,result.length,result[0].length)
    .setValues(result);
}
