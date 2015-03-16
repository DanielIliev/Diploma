function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Разписание учител', 'teachers')
  .addSeparator()
  .addItem('Маркирай дубликати', 'mark_duplicates')
  .addToUi();
}

function mark_duplicates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheet.length; i++) {
    var rows = sheet[i].getDataRange().getNumRows();
    var cols = sheet[i].getDataRange().getNumColumns();
    var data = sheet[i].getDataRange().getValues();
    var range = sheet[i].getRange(1, 1, data.length, data[0].length);
    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        switch(data[row][col]) {
          case "VIII клас":
            class_VIII(rows,cols,data,range);
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

function class_VIII(rows,cols,data,range) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col].length > 0 && data[row][col].length <= 3) {
        switch(data[row][col]) {
          case data[row][col+5]:
            var cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+5);
            cell2.setBackground("red");
            break;
          case data[row][col+10]:
            var cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+10);
            cell2.setBackground("red");
            break;
          case data[row][col+15]:
            var cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+15);
            cell2.setBackground("red");
        }
      }
    }
  }
}

function class_IX(rows,cols,data,range) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col].length > 0 && data[row][col].length <= 3) {
        switch(data[row][col]) {
          case data[row][col+6]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+4);
            cell2.setBackground("red");
            break;
          case data[row][col+12]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+11);
            cell2.setBackground("red");
            break;
          case data[row][col+18]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+17);
            cell2.setBackground("red");
            break;
        }
      }
    }
  }
}

function class_X_XI_XII(rows,cols,data,range) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col].length > 0 && data[row][col].length <= 3) {
        switch(data[row][col]) {
          case data[row][col+6]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+6);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+5);
            cell2.setBackground("red");
          case data[row][col+11]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+11);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+10);
            cell2.setBackground("red");
            break;
          case data[row][col+16]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+16);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+15);
            cell2.setBackground("red");
            break;
          case data[row][col+22]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+22);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+21);
            cell2.setBackground("red");
            break;
          case data[row][col+27]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+27);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+26);
            cell2.setBackground("red");
            break;
          case data[row][col+33]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+33);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+32);
            cell2.setBackground("red");
            break;
          case data[row][col+38]:
            var cell1 = range.getCell(row+1,col-1);
            cell1.setBackground("red");
            cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+38);
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+37);
            cell2.setBackground("red");
            break;
        }
      }
    }
  }
}

function teachers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var initials = Browser.inputBox("Въведи инициали на преподавател");
  if (initials != "cancel") {
    for (var i = 0; i < sheet.length; i++) {
      var rows = sheet[i].getDataRange().getNumRows();
      var cols = sheet[i].getDataRange().getNumColumns();
      var data = sheet[i].getDataRange().getValues();
      for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {
          var inits_check = search_init_sheet(rows,cols,data,initials);
          if (inits_check) {
            switch(data[row][col]) {
              case "VIII клас":
                var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("График на " + initials + " за VIII клас", sheet.length + 1)
                var range = new_sheet.getRange(1,1,data.length,data[0].length);
                range.setValues(data);
                get(range, data, initials);
                break;
              case "IX клас":
                var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("График на " + initials + " за IX клас", sheet.length + 1)
                var range = new_sheet.getRange(1,1,data.length,data[0].length);
                range.setValues(data);
                get(range, data, initials);
                break;
              case "X клас":
                var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("График на " + initials + " X клас", sheet.length + 1)
                var range = new_sheet.getRange(1,1,data.length,data[0].length);
                range.setValues(data);
                get(range, data, initials);
                break;
              case "XI клас":
                var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("График на " + initials + " за XI клас", sheet.length + 1)
                var range = new_sheet.getRange(1,1,data.length,data[0].length);
                range.setValues(data);
                get(range, data, initials);
                break;
              case "XII клас":
                var new_sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("График на " + initials + " за XII клас", sheet.length + 1)
                var range = new_sheet.getRange(1,1,data.length,data[0].length);
                range.setValues(data);
                get(range, data, initials);
                break;
            }
          }
        }
      }
    }
  }
}

function get(range, data, inits) {
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  for (var row = 0; row < rows; row++) {
    var time = search_time(cols,data[row]);
    if (time) {
      var init_check = search_inits(cols, data[row], inits);
      if (!init_check) {
        for (var col = 0; col < cols; col++) {
          var cell = range.getCell(row+1,col+1);
          cell.setBackground("red");
        }
      }
    }
  }
}

function search_time(cols,data_row) {
  var timestamp = /[0-9]+[:]{1}[0-9]+/;
  var found = false;
  for (var col = 0; col < cols; col++) {
    if (data_row[col] != "") {
      if (data_row[col].toString().search(timestamp) != -1) {
        found = true;
      }
    }
  }
  return found;
}

function search_init_sheet(rows,cols,data,inits) {
  var found = false;
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col] == inits) {
        found = true;
      }
    }
  }
  return found;
}

function search_inits(cols,data_row,inits) {
  var found = false;
  for (var col = 0; col < cols; col++) {
    if (data_row[col] == inits) {
      found = true;
    }
  }
  return found;
}
