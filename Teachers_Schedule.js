function test() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var cols = sheet.getDataRange().getNumColumns()
  var rows = sheet.getDataRange().getNumRows()
  var data = sheet.getDataRange().getValues()
  var range = sheet.getRange(1,1,data.length,data[0].length)
  var inits = Browser.inputBox("Въведи инициали")
  var check = 0
  var row_check = false
  
  for (var row = 0; row < rows; row++) {
    for (var col = 0; col < cols; col++) {
      if (check == 0) {
        switch(data[row][col]) {
          case 1:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 2:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 3:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 4:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 5:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 6:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 7:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
            } else {
              Logger.log("Not Found")
              check++
            }
            break;
          case 8:
            if (search_inits(cols,data[row],inits)) {
              Logger.log("Found")
              check++
              row_check = false
            } else {
              row_check = false
              Logger.log("Not Found")
              check++
            }
            break;
        }
      }
    }
    check = 0
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
