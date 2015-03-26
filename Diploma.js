function uiMenu() {
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Разписание учител', 'teachersSchedule')
  .addSeparator()
  .addItem('Маркирай дублиращи', 'markDuplicates')
  .addSeparator()
  .addItem('График на учебни кабинети', 'cabinets')
  .addToUi()
}


function teachersSchedule() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  var rows = sheet[0].getDataRange().getNumRows()
  var cols = sheet[0].getDataRange().getNumColumns()
  var data = sheet[0].getDataRange().getValues()
  var range = sheet[0].getRange(1,1,data.length,data[0].length) 
  var inits = Browser.inputBox('Въведи инициали')
  var status = false
  
  for (var row = 0; row < rows; row++) {
    if (data[row][1] != '') {
      switch(data[row][1]) {
        case 1:
          if (search_inits(cols, data[row], inits)) {
            status = true
          } else if (search_inits(cols, data[row], inits)) {
            status = true
          }
        break;
        case 2:
          if(search_inits(cols, data[row], inits)) {
            status = true
          } else if (search_inits(cols, data[row], inits)) {
            status = true
          }
        break;
        case 3:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              if (!search_inits(cols,data[row-1],inits)) {
                for (var p = 1; p < cols; p++) {
                  var cell = range.getCell(row,p+1)
                  cell.setBackground('red')
                }
              }
            } else {
              status = true
            }
          }
        break;
        case 4:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              
            } else {
              status = true
            }
          }
        break;
        case 5:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              
            } else {
              status = true
            }
          }
        break;
        case 6:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              
            } else {
              status = true
            }
          }
        break;
        case 7:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              
            } else {
              status = true
            }
          }
        break;
        case 8:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              
            }
          }
          status = false
        break;
      }
    }
  }
}

function cabinets() {
  Browser.msgBox('Тази функционалност е в процес на разработка')
}

function markDuplicates() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (var i = 0; i < sheet.length; i++) {
    var rows = sheet[i].getDataRange().getNumRows()
    var cols = sheet[i].getDataRange().getNumColumns()
    var data = sheet[i].getDataRange().getValues()
    var range = sheet[i].getRange(1,1,data.length,data[0].length)
    for (var row = 0; row < rows; row++) {
      for (var col = 2; col < cols; col++) {
        if (data[row][col].length >= 2 && data[row][col].length <= 3) {
          switch(data[row][col]) {
            case data[row][col+5]:
              if (data[2][col-3] != data[2][col+2]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+5)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+4)
                cell.setBackground('red')
              }
            break;
            case data[row][col+10]:
              if (data[2][col-3] != data[2][col+7]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+10)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+9)
                cell.setBackground('red')
              }
            break;
            case data[row][col+15]:
              if (data[2][col-3] != data[2][col+12]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+15)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+14)
                cell.setBackground('red')
              }
            break;
            case data[row][col+20]:
              if (data[2][col-3] != data[2][col+17]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+20)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+19)
                cell.setBackground('red')
              }
            break;
            case data[row][col+25]:
              if (data[2][col-3] != data[2][col+22]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+25)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+24)
                cell.setBackground('red')
              }
            break;
            case data[row][col+30]:
              if (data[2][col-3] != data[2][col+27]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+30)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+29)
                cell.setBackground('red')
              }
            break;
            case data[row][col+35]:
              if (data[2][col-3] != data[2][col+32]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+35)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+34)
                cell.setBackground('red')
              }
            break;
            case data[row][col+40]:
              if (data[2][col-3] != data[2][col+37]) {
                var cell = range.getCell(row+1,col)
                cell.setBackground('red')
                cell = range.getCell(row+1,col-1)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+40)
                cell.setBackground('red')
                cell = range.getCell(row+1,col+39)
                cell.setBackground('red')
              }
            break;
          }
        }
      }
    }
  }
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
