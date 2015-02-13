function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('График учители', 'teachers')
  .addToUi()
}

function teachers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  if (sheet.length < 6) {
    var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('График Учител', 5)
    var schedule = get_schedule(sheet);
    sheet[5].getRange(1,1,schedule.length,schedule[0].length).setValues(schedule);
    var rows = sheet[5].getDataRange().getNumRows();
    var cols = sheet[5].getDataRange().getNumColumns();
    get_new(schedule,rows,cols);
    sheet[5].clear();
  } else if (sheet.length >= 6) {
    var schedule = get_schedule(sheet,rows,cols);
    sheet[5].getRange(1,1,schedule.length,schedule[0].length).setValues(schedule);
    var rows = sheet[5].getDataRange().getNumRows();
    var cols = sheet[5].getDataRange().getNumColumns();
    get_new(schedule,rows,cols);
    sheet[5].clear();
  }
}

function get_schedule(sheets) {
  var num_of_sheets = sheets.length - 1;
  var result = [];
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  var initials = Browser.inputBox('Въведи инициали на учител');
  var cases = 0;
  var temp_arr = [];
  for (var i = 0; i < num_of_sheets; i++) {
    var rows = sheets[i].getDataRange().getNumRows();
    var cols = sheets[i].getDataRange().getNumColumns();
    var data = sheets[i].getDataRange().getValues();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        if (data[row][col] == initials) {
          if (data[row][col-2].toString().search(timestamp) != -1) {
            cases = 0;
          } else if (data[row][col-3].toString().search(timestamp) != -1) {
            cases = 1;
          } else if (data[row][col-4].toString().search(timestamp) != -1) {
            cases = 2;
          } else if (data[row][col-5].toString().search(timestamp) != -1) {
            cases = 3;
          } else if (data[row][col-6].toString().search(timestamp) != -1) {
            cases = 4;
          }
          switch(cases) {
            case 0:
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              break;
            case 1:
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              break;
            case 2:
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              temp_arr.push(data[row][col-4]);
              break;
            case 3:
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              temp_arr.push(data[row][col-4]);
              temp_arr.push(data[row][col-5]);
              break;
            case 4:
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              temp_arr.push(data[row][col-4]);
              temp_arr.push(data[row][col-5]);
              temp_arr.push(data[row][col-6]);
              break;
          }
        }
      }
    }
  }
    result.push(temp_arr);
    return result;
  }

function day(curr_sheet_data,curr_row,curr_col,sheet_rows,sheet_cols) {
  var curr_day = "";
  for (var i = curr_row; i < sheet_rows; i++) {
    switch(curr_sheet_data[i][0]) {
      case "Понеделник":
        curr_day = "Понеделник";
        break
      case "Вторник":
        curr_day = "Вторник";
        break;
      case "Сряда":
        curr_day = "Сряда";
        break;
      case "Четвъртък":
        curr_day = "Четвъртък";
        break;
      case "Петък":
        curr_day = "Петък";
        break;
    }
  }
  Logger.log(curr_day);
}

function get_new(data,rows,cols) {
  var new_arr = [];
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (!data[row][col] == "") {
        if (isNaN(data[row][col])) {
          new_arr.push(data[row][col]);
        }
      }
    }
  }
  Logger.log(new_arr);
}

/*
  Tasks left:
    - Append days to schedule
    - Check for holes ...
*/
