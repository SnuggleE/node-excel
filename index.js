var pinyin=require('pinyin')
var mysql=require('mysql');
var excel=require('exceljs');
var fs=require('fs')

var rootDir='./xls/'

var conn=mysql.createConnection({

});
//插入数据库
function insertUser(schoolName) {
  var shortName=getShortName(schoolName);
}
//获取学校名称的拼音首字母
function getShortName(name) {
  var resStr='';
  pinyin(name,{
    style:pinyin.STYLE_NORMAL
  }).forEach(function (item) {
    resStr+=item[0][0]
  })
  return resStr
}

//主函数
function main() {
  var count=0
  fs.readdir(rootDir,function (error,files) {
    files.forEach(function (file) {
      var workbook=new excel.Workbook();
      workbook.xlsx.readFile(rootDir+file)
          .then(function () {
            // console.log(workbook.getWorksheet(3)._rows[3]._cells[0]._value.model.value);
            var workSheetCount=workbook._worksheets.length-1;
            for(var sheetIndex=1;sheetIndex<workSheetCount;sheetIndex++){
              var sheet=workbook.getWorksheet(sheetIndex);
              console.log(sheet._rows[2]._cells[0]._value.model.value);
              var schoolName=sheet._rows[2]._cells[0]._value.model.value
              if(schoolName){
                var shortName=getShortName(schoolName);
                console.log(shortName)
              }

              count++;
            }
          })
    })
    setTimeout(function () {
      console.log(count+'  sheets')
    },10000)
  });

}

main();