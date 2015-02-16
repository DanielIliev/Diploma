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
    var rows = sheet[5].getDataRange().getNumRows();
    var cols = sheet[5].getDataRange().getNumColumns();
    sheet[5].clear();
    sheet[5].getRange(1,1,schedule.length,schedule[0].length).setValues(schedule);
  } else if (sheet.length >= 6) {
    var schedule = get_schedule(sheet,rows,cols);
    var rows = sheet[5].getDataRange().getNumRows();
    var cols = sheet[5].getDataRange().getNumColumns();
    sheet[5].clear();
    sheet[5].getRange(1,1,schedule.length,schedule[0].length).setValues(schedule);
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
              var day = get_day(row,data);
              temp_arr.push(day);
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              if (data[row+1][col-1] == data[row][col-1]) {
                temp_arr.push(data[row+1][col]);
                temp_arr.push(data[row+1][col-1]);
                temp_arr.push(data[row+1][col-2]);
              }
              if (data[row+2][col-1] == data[row][col-1]) {
                temp_arr.push(data[row+2][col]);
                temp_arr.push(data[row+2][col-1]);
                temp_arr.push(data[row+2][col-2]);
              }
              break;
            case 1:
              var day = get_day(row,data);
              temp_arr.push(day);
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              break;
            case 2:
              var day = get_day(row,data);
              temp_arr.push(day);
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              temp_arr.push(data[row][col-4]);
              break;
            case 3:
              var day = get_day(row,data);
              temp_arr.push(day);
              temp_arr.push(data[row][col]);
              temp_arr.push(data[row][col-1]);
              temp_arr.push(data[row][col-2]);
              temp_arr.push(data[row][col-3]);
              temp_arr.push(data[row][col-4]);
              temp_arr.push(data[row][col-5]);
              break;
            case 4:
              var day = get_day(row,data);
              temp_arr.push(day);
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

function get_day(row,data) {
  while (row >= 0) {
    var cur_cell = data[row][0];
    switch(cur_cell) {
      case "Понеделник":
        var curr_day = data[row][0];
        row = 0;
        break;
      case "Вторник":
        var curr_day = data[row][0];
        row = 0;
        break;
      case "Сряда":
        var curr_day = data[row][0];
        row = 0;
        break;
      case "Четвъртък":
        var curr_day = data[row][0];
        row = 0;
        break;
      case "Петък":
        var curr_day = data[row][0];
        row = 0;
        break;
    }
    row--;
  }
  Logger.log(curr_day);
  return curr_day;
}

/*
  Tasks left:
    - Append days to schedule
    - Check for holes ...
*/
