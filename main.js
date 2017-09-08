var excel = require('exceljs');


var file = './xls/十堰市.xlsx'


var workbook = new excel.Workbook();
workbook.xlsx.readFile(file)
    .then(function () {
      // var workSheetCount = workbook._worksheets.length - 1;
      workbook.eachSheet(function (worksheet, sheetId) {

        console.log("*********************************************************************")
        worksheet.eachRow(function (row, rowNumber) {
          // console.log(rowNumber)
          row.eachCell(function (cell) {
            console.log(cell.value)
          })
        })
      })
    })

