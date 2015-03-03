function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Разписание учител', 'teachers')
  .addSeparator()
  .addItem('Заетост на кабинети', 'cab_occupation')
  .addSeparator()
  .addItem('Маркирай повтарящи се предмети', 'mark_duplicates')
  .addToUi();
}

function mark_duplicates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var num_sheets = 5;
for (var i = 0; i < num_sheets; i++) {
    var rows = sheet[i].getDataRange().getNumRows();
    var cols = sheet[i].getDataRange().getNumColumns();
    var data = sheet[i].getDataRange().getValues();
    var range = sheet[i].getRange(1, 1, data.length, data[0].length);
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        switch(data[row][col]) {
          case "VIII клас":
            class_VII(rows,cols,data,range);
            break;
          case "IX клас":
            class_IX(rows,cols,data,range);
            break;
          case "X клас":
            class_X_XI_XII(rows,cols,data,range);
            break;
          case "XI клас":
            class_X_XI_XII(rows,cols,data,range);
            break;
          case "XII клас":
            class_X_XI_XII(rows,cols,data,range);
            break;
        }
      }
    }
  }
}

function class_X_XI_XII(rows,cols,data,range) {
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col] != "") {
        if (data[row][col].toString().search(timestamp) != -1) {
          if (data[row][col+1] == data[row][col+12]) {
            if (data[row][col+3] == data[row][col+14]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+13);
              cell2.setBackground("red");
            }
          }
          if (data[row][col+1] == data[row][col+23]) {
            if (data[row][col+3] == data[row][col+25]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+24);
              cell2.setBackground("red");
            }
          }
          if (data[row][col+1] == data[row][col+34]) {
            if (data[row][col+3] == data[row][col+36]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+35);
              cell2.setBackground("red");
            }
          }
        }
      }
    }
  }
}

function class_IX(rows,cols,data,range) {
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col] != "") {
        if (data[row][col].toString().search(timestamp) != -1) {
          if (data[row][col+1] == data[row][col+7]) {
            if (data[row][col+3] == data[row][col+9]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+8);
              cell2.setBackground("red");
            }
          }
          if (data[row][col+1] == data[row][col+13]) {
            if (data[row][col+3] == data[row][col+15]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+14);
              cell2.setBackground("red");
            }
          }
          if (data[row][col+1] == data[row][col+19]) {
            if (data[row][col+3] == data[row][col+21]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+20);
              cell2.setBackground("red");
            }
          }
        }
      }
    }
  }
}

function class_VII(rows,cols,data,range) {
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col] != "") {
        if (data[row][col].toString().search(timestamp) != -1) {
          if (data[row][col+1] == data[row][col+6]) {
            if (data[row][col+2] == data[row][col+7]) {
              Logger.log("Duplicate found");
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+7);
              cell2.setBackground("red");
            }
          }
          if (data[row][col+1] == data[row][col+11]) {
            if (data[row][col+2] == data[row][col+12]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+12);
              cell2.setBackground("red");
            }
          }
          if (data[row][col+1] == data[row][col+16]) {
            if (data[row][col+2] == data[row][col+17]) {
              var cell1 = range.getCell(row+1, col+2);
              cell1.setBackground("red");
              var cell2 = range.getCell(row+1, col+17);
              cell2.setBackground("red");
            }
          }
        }
      }
    }
  }
}

function teachers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var initials = Browser.inputBox("Въведи инициали на преподавател");
  try {
    var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Разписание за: " + initials, sheet.length + 1); 
    var schedule = get_schedule(sheet,initials);
    new_sheet.getRange(1,1,schedule.length,schedule[0].length).setValues(schedule);
  }
  catch(e) {
    Browser.msgBox("Разписанието за избраният учител е изисквано вече, моля изтрийте таблицата на дадения преподавател и опитайте отново");
  }
}

function get_schedule(sheet,inits) {
  var num_of_sheets = 5; // Sheets with school schedules
  var result = [];
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  var cases = 0;
  for (var i = 0; i < num_of_sheets; i++) {
    var rows = sheet[i].getDataRange().getNumRows();
    var cols = sheet[i].getDataRange().getNumColumns();
    var data = sheet[i].getDataRange().getValues();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        if (data[row][col] == inits) {
          if (data[row][col-2].toString().search(timestamp) != -1) {
            cases = 0;
          } else if (data[row][col-3].toString().search(timestamp) != -1) {
            cases = 1;
          }
          switch(cases) {
            case 0:
              result.push([get_day(row,data), data[row][col-2], data[row][col-1], data[row][col], data[row][col+1]]);
              break;
            case 1:
              result.push([get_day(row,data), data[row][col-3], data[row][col-2], data[row][col-1], data[row][col], data[row][col+1]]);
              break;
          }
        }
      }
    }
  }
  return result;
}

function cab_occupation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  try {
    var time = Browser.inputBox("Въведи час");
    var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Разписание за: " + time, sheet.length + 1);
    var schedule = get_cab_oc(sheet,time);
    new_sheet.getRange(1,1,schedule.length,schedule[0].length).setValues(schedule);
  }
  catch(e) {
    Browser.msgBox("Разписанието за избраният учител е изисквано вече, моля изтрийте таблицата на дадения преподавател и опитайте отново");
  }
}

function get_cab_oc(sheet, time_interval) {
  var result = [];
  var num_sheets = 5;
  var cases = 0;
  var cab_pattern = /[0-9]{2}/;
  for (var i = 0; i < num_sheets; i++) {
    var rows = sheet[i].getDataRange().getNumRows();
    var cols = sheet[i].getDataRange().getNumColumns();
    var data = sheet[i].getDataRange().getValues();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        if (data[row][col] == time_interval) {
          if (data[row][col+3].toString().search(cab_pattern) != - 1) {
            cases = 0;
          } else if (data[row][col+4].toString().search(cab_pattern) != -1) {
            cases = 1;
          }
          switch(cases) {
            case 0:
              result.push(["В",get_day(row,data),"Кабинет",data[row][col+3],"е зает"]);
              break;
            case 1:
              result.push(["В",get_day(row,data),"Кабинет",data[row][col+4],"е зает"]);
              break;
          }
        }
      }
    }
  }
  return result;
}

function get_day(curr_row, data) {
  var curr_day = "";
  while(curr_row != 0) {
    switch(data[curr_row][0]) {
      case "Понеделник":
        curr_day = "Понеделник";
        return curr_day;
        curr_row = 0;
        break;
      case "Вторник":
        curr_day = "Вторник";
        return curr_day;
        curr_row = 0;
        break;
      case "Сряда":
        curr_day = "Сряда";
        return curr_day;
        curr_row = 0;
        break;
      case "Четвъртък":
        curr_day = "Четвъртък";
        return curr_day;
        curr_row = 0;
        break;
      case "Петък":
        curr_day = "Петък";
        return curr_day;
        curr_row = 0;
        break;
    }
    curr_row--;
  }
}
/*
function duplicates() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  for (var i = 0; i < 5; i++) {
    var range = sheets[i].getActiveRange();
    var rows = sheets[i].getDataRange().getNumRows();
    var cols = sheets[i].getDataRange().getNumColumns();
    var data = sheets[i].getDataRange().getValues();
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        check_row (rows,cols,data[row],data[row][col]);
      }
    }
  }
}

function check_row(curr_row_pos,cols,curr_row_data,curr_cell) {
  for (var 
}*/
