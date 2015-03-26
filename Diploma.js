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
  var inits = Browser.inputBox('Въведи инициали')
  if (inits != '') {
    for (var i = 0; i < sheet.length; i++) {
      var rows = sheet[i].getDataRange().getNumRows()
      var cols = sheet[i].getDataRange().getNumColumns()
      var data = sheet[i].getDataRange().getValues()
      var range = sheet[i].getRange(1,1,data.length,data[0].length)
      var found = false
      var class = ''
      for (var row = 0; row < rows; row++) {
        for (var col = 0; col < cols; col++) {
          if (data[row][col] == inits) {
            found = true
          }
        }
      }
      if (found) {
        var newsheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Разписание на ' + inits + ' за ' + data[0][2], sheet.length + 1)
        newsheet.getRange(1,1,data.length,data[0].length).setValues(data)
        markholes(newsheet, inits)
      }
      found = false
    }
  }
}

function markholes(newsheet, inits) {
  var rows = newsheet.getDataRange().getNumRows()
  var cols = newsheet.getDataRange().getNumColumns()
  var data = newsheet.getDataRange().getValues()
  var range = newsheet.getRange(1,1,data.length,data[0].length)
  var status = false
  var row_position = 0
  for (var row = 0; row < rows; row++) {
    if (data[row][1] != '') {
      switch(data[row][1]) {
        case 1:
          if (search_inits(cols, data[row], inits)) {
            status = true
            row_position = 1
          }
          break;
        case 2:
          if (search_inits(cols,data[row],inits)) {
            status = true
            row_position = 2
          }
          break;
        case 3:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              switch(row_position) {
                case 1:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
              }
            } else {
              status = true
              row_position = 3
            }
          }
          break;
        case 4:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              switch(row_position) {
                case 1:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 2:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
              }
            } else {
              status = true
              row_position = 4
            }
          }
          break;
        case 5:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              switch(row_position) {
                case 1:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 2:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 3:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
              }
            } else {
              status = true
              row_position = 5
            }
          }
          break;
        case 6:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              switch(row_position) {
                case 1:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-4],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-3,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 2:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 3:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 4:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
              }
            } else {
              status = true
              row_position = 6
            }
          }
          break;
        case 7:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              switch(row_position) {
                case 1:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-4],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-3,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-5],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-4,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 2:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-4],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-3,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 3:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 4:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 5:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
              }
            } else {
              status = true
              row_position = 7
            }
          }
          break;
        case 8:
          if (search_inits(cols, data[row], inits)) {
            if (status) {
              switch(row_position) {
                case 1:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-4],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-3,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-5],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-4,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-6],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-5,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 2:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-4],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-3,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-5],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-4,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 3:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-4],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-3,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 4:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-3],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-2,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 5:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  if (!search_inits(cols,data[row-2],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row-1,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
                case 6:
                  if (!search_inits(cols,data[row-1],inits)) {
                    for (var p = 1; p < cols; p++) {
                      var cell = range.getCell(row,p)
                      cell.setBackground('red')
                    }
                  }
                  break;
              }
            }
          }
          status = false
          break;
      }
    }
  } // The for loop for row ends here
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
