function start() {
  SpreadsheetApp.getUi().createAddonMenu().addItem('Search', 'test').addToUi()
}

function test() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var cols = sheet.getDataRange().getNumColumns()
  var rows = sheet.getDataRange().getNumRows()
  var data = sheet.getDataRange().getValues()
  var range = sheet.getRange(1,1,data.length,data[0].length)
  var inits = Browser.inputBox("Въведи инициали")
  var checked_row = false
  var inits_check = false
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (checked_row === false) {
        switch(data[row][col]) {
          case 1:
            switch(search_inits(cols,data[row],inits)) {
              case true:
                checked_row = true
                inits_check = true
                Logger.log('Found')
              break;
            }
          break;
          case 2:
            switch(search_inits(cols,data[row],inits)) {
              case true:
                checked_row = true
                inits_check = true
                Logger.log('Found')
              break;
            }
          break;
          case 3:
            switch(search_inits(cols,data[row],inits)) {
              case true:
                switch(inits_check) {
                  case true:
                    switch(search_inits(cols,data[row-1],inits)) {
                      case false:
                        for (var p = 1; p < data[0].length; p++) {
                          var cell = range.getCell(row,p)
                          cell.setBackground('red')
                        }
                      break;
                    }
                  break;
                  case false:
                    switch(search_inits(cols,data[row],inits)) {
                      case true:
                        inits_check = true
                      break;
                    }
                  break;
                }
                break;
            }
          break;
          case 4:
            switch(search_inits(cols,data[row],inits)) {
              case true:
                switch(inits_check) {
                case true:
                  switch(search_inits(cols,data[row-1],inits)) {
                    case false:
                      for (var p = 1; p < data[0].length; p++) {
                        var cell = range.getCell(row,p)
                        cell.setBackground('red')
                      }
                    break;
                  }
                  switch(search_inits(cols,data[row-2],inits)) {
                    case false:
                      for (var p = 1; p < data[0].length; p++) {
                        var cell = range.getCell(row-1,p)
                        cell.setBackground('red')
                      }
                    break;
                  }
                break;
              }
              break;
            }
          break;
          case 5:
            switch(search_inits(cols,data[row],inits)) {
              case true:
                switch(inits_check) {
                case true:
                  switch(search_inits(cols,data[row-1],inits)) {
                    case false:
                      for (var p = 1; p < data[0].length; p++) {
                        var cell = range.getCell(row,p)
                        cell.setBackground('red')
                      }
                    break;
                  }
                  switch(search_inits(cols,data[row-2],inits)) {
                    case false:
                      for (var p = 1; p < data[0].length; p++) {
                        var cell = range.getCell(row-1,p)
                        cell.setBackground('red')
                      }
                    break;
                  }
                  switch(search_inits(cols,data[row-3],inits)) {
                    case false:
                      for (var p = 1; p < data[0].length; p++) {
                        var cell = range.getCell(row-2,p)
                        cell.setBackground('red')
                      }
                    break;
                  }
                break;
              }
              break;
            }
          break;
          case 6:
          break;
          case 7:
          break;
          case 8:
          break;
        }
      }
    }
    checked_row = false
  }
}

function search_inits(cols,data_row,inits) {
  var found = false
  for (var col = 0; col < cols; col++) {
    if (data_row[col].toString() == inits) {
      found = true
    }
  }
  return found
}


/*
switch(search_inits(cols,data[row],inits)) {
          case true:
            checked_row = true
            Logger.log("Found")
            break;
        }
*/
