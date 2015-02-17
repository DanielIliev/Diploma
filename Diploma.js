function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('График учител', 'teachers')
  .addSeparator()
  .addItem('Заети кабинети', 'cabinets')
  .addToUi()
}

// Task 3
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
  var cab = /[0-9]+/;
  var initials = Browser.inputBox('Въведи инициали на учител');
  var cases = 0;
  result.push([initials]);
  for (var i = 0; i < num_of_sheets; i++) {
    var rows = sheets[i].getDataRange().getNumRows();
    var cols = sheets[i].getDataRange().getNumColumns();
    var data = sheets[i].getDataRange().getValues();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        if (data[row][col] == initials) {
          var cases = 0;
          if (data[row][col-2].toString().search(timestamp) != -1) {
            cases = 0;
          } else if (data[row][col-3].toString().search(timestamp) != -1) {
            cases = 1;
          } else if (data[row][col-7].toString().search(timestamp) != -1) {
            cases = 2;
          }
          switch(cases) {
            case 0:
              var day = get_day(row,data);
              result.push([day]);
              var classes = get_class(rows,cols,data);
              result.push([classes]);
              if (data[row][col-2] != "") {
                result.push([data[row][col-2]]);
              }
              if (data[row][col-1] == data[row+1][col-1]) {
                if (data[row+1][col-2] != "") {
                  result.push([data[row+1][col-2]]);
                }
              }
              if (data[row][col-1] == data[row+2][col-1]) {
                if (data[row+2][col-2] != "") {
                  result.push([data[row+2][col-2]]);
                }
              }
              if (data[row][col-1] == data[row+3][col-1]) {
                if (data[row+3][col-2] != "") {
                  result.push([data[row+3][col-2]]);
                }
              }
              break;
            case 1:
              var day = get_day(row,data);
              result.push([day]);
              var classes = get_class(rows,cols,data);
              result.push([classes]);
              result.push([data[row][col-3]]);
              if (data[row][col-2] == data[row+1][col-2]) {
                if (data[row+1][col-3] != "") {
                  result.push([data[row+1][col-3]]);
                }
              }
              if (data[row][col-2] == data[row+2][col-2]) {
                if (data[row+2][col-3] != "") {
                  result.push([data[row+2][col-3]]);
                }
              }
              break;
            case 2:
              var day = get_day(row,data);
              result.push([day]);
              var classes = get_class(rows,cols,data);
              result.push([classes]);
              if (data[row][col-7] != "") {
                result.push([data[row][col-7]]);
              }
              if (data[row][col-2] == data[row+1][col-2]) {
                if (data[row+1][col-7] != "") {
                  result.push([data[row+1][col-7]]);
                }
              }
              if (data[row][col-2] == data[row+2][col-2]) {
                if (data[row+2][col-7] != "") {
                  result.push([data[row+2][col-7]]);
                }
              }
              break;
          }
        }
      }
    }
  }
    return result;
}

// End of Task 3

// Task 2
function cabinets() {
  
}

//Custom functions
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

function get_class(rows,cols,data) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      switch(data[row][col]) {
        case "XII клас":
          var curr_class = data[row][col];
          break;
        case "XI клас":
          var curr_class = data[row][col];
          break;
        case "X клас":
          var curr_class = data[row][col];
          break;
        case "IX клас":
          var curr_class = data[row][col];
          break;
        case "VIII клас":
          var curr_class = data[row][col];
          break;
      }
    }
  }
  return curr_class;
}
