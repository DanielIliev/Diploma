function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Разписание учител', 'teachers')
  .addSeparator()
  .addItem('Маркирай дублиращи', 'mark_duplicates')
  .addSeparator()
  .addItem('График на учебни кабинети', 'cabinets')
  .addToUi()
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
            class_VIII(rows,cols,data,range)
            break;
          case "IX клас":
            class_IX(rows,cols,data,range)
            break;
          case "X клас":
            class_X(rows,cols,data,range)
            break;
          case 'XI клас':
            class_XI(rows,cols,data,range)
            break;
          case "XII клас":
            class_XII(rows,cols,data,range)
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
          case data[row][col+4]:
            var cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+4);
            cell2.setBackground("red");
            break;
          case data[row][col+8]:
            var cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+8);
            cell2.setBackground("red");
            break;
          case data[row][col+12]:
            var cell1 = range.getCell(row+1,col);
            cell1.setBackground("red");
            var cell2 = range.getCell(row+1,col+12);
            cell2.setBackground("red");
        }
      }
    }
  }
}

function class_IX(rows,cols,data,range) {
  for (var row = 1; row < rows; row++) {
    for (var col = 1; col < cols; col++) {
      if (data[row][col].length > 0 && data[row][col].length <= 3) {
        switch(data[row][col]) {
          case data[row][col+5]:
            var cell1 = range.getCell(row+1,col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1,col+5)
            cell2.setBackground("red");
            cell2 = range.getCell(row+1,col+4)
            cell2.setBackground("red")
            break;
          case data[row][col+10]:
            var cell1 = range.getCell(row+1,col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1,col+10)
            cell2.setBackground("red")
            break;
          case data[row][col+15]:
            var cell1 = range.getCell(row+1,col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1,col+15)
            cell2.setBackground("red")
            break;
        }
      }
    }
  }
}

function class_X(rows,cols,data,range) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (data[row][col].length > 0 && data[row][col].length <= 3) {
        switch(data[row][col]) {
          case data[row][col+5]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+5)
            cell2.setBackground("red")
            break;
          case data[row][col+10]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+10)
            cell2.setBackground("red")
            break;
          case data[row][col+15]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+15)
            cell2.setBackground("red")
            break;
          case data[row][col+20]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+20)
            cell2.setBackground("red")
            break;
          case data[row][col+25]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+25)
            cell2.setBackground("red")
            break;
          case data[row][col+30]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+30)
            cell2.setBackground("red")
            break;
          case data[row][col+35]:
            var cell1 = range.getCell(row+1, col)
            cell1.setBackground("red")
            var cell2 = range.getCell(row+1, col+35)
            cell2.setBackground("red")
            break;
        }
      }
    }
  }
}

function class_XI(rows,cols,data,range) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      Logger.log(data[row][col])
    }
  }
}

function class_XII(rows,cols,data,range) {
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      Logger.log(data[row][col])
    }
  }
}

function teachers() {
  Browser.msgBox('Тази функционалност е в процес на разработка')
}

function cabinets() {
  Browser.msgBox('Тази функционалност е в процес на разработка')
}

function search_inits(cols,data_row,inits) {
  var found = false
  for (var col = 0; col < cols; col++) {
    if (data_row[col] == inits) {
      found = true
    }
  }
  return found
}
