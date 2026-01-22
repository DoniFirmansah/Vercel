const idDb = "1eaqIthOSYAQNiORUp6diyXLdHHWRjy74lDBiAJmCkGA"



// function doGet() {
//   return HtmlService.createHtmlOutputFromFile('index');
// }

function doGet(e){

  var page = e.parameter.page
  if(page == null){
    page = "home"
  }
  var output = HtmlService.createTemplateFromFile(page)
  return output
          .evaluate()
          .setTitle("SMKN 9 Semarang")
}

//ini akan menghasilkan link url web kita
function myUrl(){
  var link =  ScriptApp.getService().getUrl()
  var linkFile = link.replace('https://script.google.com','https://script.google.com/a/~')
  //memberikan linknya ke user
  return linkFile
}

//ini untuk ambil potongan source code
function include(nameFile){
   return HtmlService.createTemplateFromFile(nameFile).evaluate().getContent();
}

function coba(){
  // ScriptApp.newTrigger("baca")
  // .timeBased()
  // .at()

  var date = new Date()
  var day = date.getDay()
  var namaHari = getNamaHari(day)
  Logger.log(namaHari)
}

function getData() {
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  //gunakan nama dinamis harian untuk mendapatkan nama sheet
  var date = new Date()
  var day = date.getDay()
  var namaHari = getNamaHari(day)
  var sheet = SpreadsheetApp.openById(idDb).getSheetByName(namaHari)
  var range = sheet.getDataRange();
  var values = range.getValues();
  return values;
}

function getDataColor() {
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  //gunakan nama dinamis harian untuk mendapatkan nama sheet
  var date = new Date()
  var day = date.getDay()
  var namaHari = getNamaHari(day)
  var sheet = SpreadsheetApp.openById(idDb).getSheetByName(namaHari)
  var range = sheet.getDataRange();
  var colors = range.getBackgrounds();
  return colors;
}

function getDataWithColors() {
  var data = getData();
  var colors = getDataColor();
  return [data, colors];
}

function trigerClearWarnaCell(){
  for(var i = 0;i<10;i++){
    resetWarnaCell(2+i)
  }
}

function resetWarnaCell(noKolom){
  var date = new Date()
  var day = date.getDay()
  var namaHari = getNamaHari(day)
  var sheet = SpreadsheetApp.openById(idDb).getSheetByName(namaHari)
  var data = sheet.getDataRange().getValues()
  var filterData = data.filter(function(guru,index){
    if(index>0){
      if(guru[noKolom-1] !=""){
        sheetGuru.getRange((index+1),noKolom).setBackground(null)
        return guru
      }
    }
  }) 
}

function getNamaHari(noHari){
  var hari = ""
  switch(noHari){
    case 1 :
      hari = "Senin"
      break
    case 2 :
      hari = "Selasa"
      break
    case 3 :
      hari = "Rabu"
      break
    case 4 :
      hari = "Kamis"
      break
    case 5 :
      hari = "Jumat"
      break        
  }
  return hari
}



