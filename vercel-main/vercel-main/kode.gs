const idDb = "1eaqIthOSYAQNiORUp6diyXLdHHWRjy74lDBiAJmCkGA"
var date = new Date()
var day = date.getDay()
var namaHari = getNamaHari(day)
var sheet = SpreadsheetApp.openById(idDb).getSheetByName(namaHari)
const ID_JADWAL = '1eaqIthOSYAQNiORUp6diyXLdHHWRjy74lDBiAJmCkGA';


// function doGet() {
//   return HtmlService.createHtmlOutputFromFile('index');
// }

function doGet(e){
  if (e.parameter.api === 'guruizin') {
    return handleApiRequest(e, 'GET');
  }
  
  var page = e.parameter.page
  if(page == null){
    page = "home" 
  }
  var output = HtmlService.createTemplateFromFile(page)
  return output
          .evaluate()
          .setTitle("SMKN 9 Semarang")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)  
}

function doPost(e){
  // API endpoint untuk POST requests
  return handleApiRequest(e, 'POST');
}
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handler untuk API requests
function handleApiRequest(e, method) {
  try {
    let action = null;
    let requestData = null;
    
    // Log untuk debugging
    Logger.log('=== API Request ===');
    Logger.log('Method: ' + method);
    Logger.log('Parameters: ' + JSON.stringify(e.parameter));
    
    if (method === 'GET') {
      action = e.parameter.action;
      requestData = e.parameter;
      Logger.log('Action (GET): ' + action);
    } else if (method === 'POST') {
      // Parse POST data - support both JSON and form-urlencoded
      if (e.postData && e.postData.contents) {
        const contentType = e.postData.type || '';
        try {
          if (contentType.includes('application/json')) {
            requestData = JSON.parse(e.postData.contents);
          } else if (contentType.includes('application/x-www-form-urlencoded')) {
            // Parse URL-encoded form data
            requestData = {};
            const params = e.postData.contents.split('&');
            params.forEach(param => {
              const [key, value] = param.split('=');
              if (key && value) {
                requestData[decodeURIComponent(key)] = decodeURIComponent(value.replace(/\+/g, ' '));
              }
            });
          } else {
            // Try JSON first, then fallback to parameter
            try {
              requestData = JSON.parse(e.postData.contents);
            } catch (e) {
              requestData = e.parameter || {};
            }
          }
          action = requestData.action;
        } catch (parseError) {
          // Jika bukan JSON, coba ambil dari parameter
          requestData = e.parameter || {};
          action = requestData.action;
        }
      } else {
        requestData = e.parameter || {};
        action = requestData.action;
      }
    }
    
    if (!action) {
      Logger.log('Error: Action parameter is required');
      return createJsonResponse({
        success: false,
        error: 'Action parameter is required'
      });
    }
    
    Logger.log('Processing action: ' + action);
    
    if (method === 'GET') {
      switch(action) {
        case 'getDataGuru':
          return createJsonResponse({
            success: true,
            data: getDataGuru()
          });
          
        case 'getDataKelas':
          return createJsonResponse({
            success: true,
            data: getDataKelas()
          });
          
        case 'getDeskripsiTugas':
          const idKelas = requestData.idKelas;
          return createJsonResponse({
            success: true,
            data: getDeskripsiTugasByIdKelas(idKelas)
          });
          
        case 'getAllDeskripsiTugas':
          return createJsonResponse({
            success: true,
            data: getAllDeskripsiTugas()
          });
          
        default:
          return createJsonResponse({
            success: false,
            error: 'Action not found: ' + action
          });
      }
    } else if (method === 'POST') {
      switch(action) {
        case 'saveIjinKeluar':
          // Remove action from data before passing to saveIjinKeluar
          const formData = Object.assign({}, requestData);
          delete formData.action;
          const result = saveIjinKeluar(formData);
          return createJsonResponse({
            success: result.includes('berhasil'),
            message: result
          });
          
        case 'uploadFile':
          const fileContent = requestData.fileContent;
          const fileName = requestData.fileName;
          const mimeType = requestData.mimeType;
          try {
            const fileUrl = uploadFile(fileContent, fileName, mimeType);
            return createJsonResponse({
              success: true,
              data: { url: fileUrl }
            });
          } catch (error) {
            return createJsonResponse({
              success: false,
              error: error.toString()
            });
          }
          
        default:
          return createJsonResponse({
            success: false,
            error: 'Action not found: ' + action
          });
      }
    }
  } catch (error) {
    Logger.log("API Error: " + error.toString());
    return createJsonResponse({
      success: false,
      error: error.toString()
    });
  }
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

  // var date = new Date()
  // var day = date.getDay()
  // var namaHari = getNamaHari(day)
  var data = getData()

  data.slice(1).forEach(function(row, rowIndex) {
      if(row[0] == 'Dra. Handayani Sri Lestari'){
      
        row.forEach(function(cell, colIndex) {
          var kelas = (cell.split("/")[0]).trim()
          // Logger.log(kelas)
          if(kelas == 'X PM 3'){
            sheet.getRange(rowIndex+2,colIndex+1).setBackground('#FFFF00')
            Logger.log("baris : "+(rowIndex+2)+ ' kolom : '+(colIndex+1))
            // #FFFF00
          } 
                  
        });
        
      }
      // Logger.log('masih diteruskan ini')        
      
              
  });

  // Logger.log(data.slice(1))
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
        sheet.getRange((index+1),noKolom).setBackground(null)
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


//===================FORM IJIN KELUAR GURU KONFIGURASI ================================
const sheetDb = SpreadsheetApp.openById(idDb)
const FOLDER_ID = '1LJnFCTlPrZzeI5YB4PS00gbj7noNSvyx'

function getDataGuru(){
  try {
    const sheet = sheetDb.getSheetByName("Data Guru");
    if (!sheet) {
      Logger.log("Sheet Data Guru tidak ditemukan");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log("Total rows in Data Guru: " + data.length);
    
    let output = []
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nama = row[1] ? row[1].toString().trim() : '';
      const noWa = row[2] ? row[2].toString().trim() : '';
      
      if (nama && nama !== '') {
        output.push({
          nama: nama,
          no_wa: noWa
        });
        Logger.log("Found guru: " + nama);
      }
    }
    
    Logger.log("Total guru found: " + output.length);
    return output;
    
  } catch (error) {
    Logger.log("Error in getDataGuru: " + error.toString());
    return [];
  }
}

function getDataKelas() {
  const sheet = sheetDb.getSheetByName("dataKelas");
  if (!sheet) return []

  const data = sheet.getDataRange().getValues();
  let output = []

  data.slice(1).forEach(row => {
    if (row[1] && row[1].toString().trim() !== '') {
      output.push({
        id: row[0],
        kelas: row[1].toString().trim(),
        wali: row[2] ? row[2].toString() : '',
        jurusan: row[3] ? row[3].toString() : ''
      })
    }
  })

  return output
}

function getDeskripsiTugasByIdKelas(idKelas) {
  try {
    const sheet = sheetDb.getSheetByName("deskripsiTugas");
    if (!sheet) return null
    const kelasData = getDataKelas();
    const kelasMap = {};
    kelasData.forEach(kelas => {
      kelasMap[kelas.id] = kelas.kelas;
    })
    const namaKelas = kelasMap[idKelas];
    if (!namaKelas) return null;
    
    const data = sheet.getDataRange().getValues()
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == namaKelas) {
        return {
          id_kelas: idKelas,
          nama_kelas: data[i][0],
          deskripsi: data[i][1] || '',
          urlFoto: data[i][2] || ''
        };
      }
    }
    return null;
  } catch (error) {
    Logger.log("Error in getDeskripsiTugasByIdKelas: " + error.toString());
    return null;
  }
}
function saveDeskripsiTugas(deskripsiData) {
  try {
    const sheet = sheetDb.getSheetByName("deskripsiTugas");
    if (!sheet) return "Sheet deskripsiTugas tidak ditemukan";
    
    const kelasData = getDataKelas();
    const kelasMap = {};
    kelasData.forEach(kelas => {
      kelasMap[kelas.id] = kelas.kelas;
    });
    
    // Jika input ID numerik, ubah ke Nama Kelas
    const namaKelas = kelasMap[deskripsiData.id_kelas] || deskripsiData.id_kelas.toString();
    
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Cari apakah kelas sudah ada di list
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == namaKelas) {
        rowIndex = i + 1; // Baris di sheet (1-based index)
        break;
      }
    }
    
    if (rowIndex > 0) {
      // Update baris yang sudah ada (Update Kolom A, B, C)
      sheet.getRange(rowIndex, 1, 1, 3).setValues([[
        namaKelas,                   // Kolom A
        deskripsiData.deskripsi,    // Kolom B
        deskripsiData.urlFoto       // Kolom C
      ]]);
      Logger.log("Updated deskripsi for kelas: " + namaKelas);
    } else {
      // Tambah baris baru (Append A, B, C)
      sheet.appendRow([
        namaKelas,                   // Kolom A
        deskripsiData.deskripsi,    // Kolom B
        deskripsiData.urlFoto       // Kolom C
      ]);
      Logger.log("Added new deskripsi for kelas: " + namaKelas);
    }
    return "Deskripsi tugas berhasil disimpan";
  } catch (error) {
    Logger.log("Error in saveDeskripsiTugas: " + error.toString());
    return "Error: " + error.toString();
  }
}

function uploadFile(fileContent, fileName, mimeType) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(fileContent), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (error) {
    Logger.log("uploadFile error: " + error);
    throw new Error("Gagal mengupload file: " + error.toString());
  }
}

function testUpload() {
  const base64 = Utilities.base64Encode("Hello world!");
  const url = uploadFile(base64, "coba.txt", "text/plain");
  Logger.log(url);
}

function requestDriveCreateAuth() {
  try {
    var blob = Utilities.newBlob("auth test", "text/plain", "apps_script_auth_test.txt");
    var file = DriveApp.createFile(blob);
    Logger.log("Created test file: " + file.getId());
    
    var fileId = file.getId();
    DriveApp.getFileById(fileId).setTrashed(true);
    Logger.log("Trashed test file: " + fileId);
  } catch (e) {
    Logger.log("requestDriveCreateAuth error: " + e.toString());
    throw e;
  }
}

function requestDriveAuth() {
  DriveApp.getRootFolder(); 
  Logger.log("Drive authorized.");
}

function saveIjinKeluar(formData) {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) return "Sheet FormIjinKeluar tidak ditemukan";
    
    const kelasData = getDataKelas();
    const kelasMap = {};
    kelasData.forEach(kelas => {
      kelasMap[kelas.id] = kelas.kelas;
    });

    const rowData = [
      new Date(),
      formData.namaGuru,
      formData.keperluan,
      formData.meninggalkanPukul,
      formData.estimasiSelesaiPukul,
      formData.meninggalkanBerapaKelas
    ];

    const jumlahKelas = parseInt(formData.meninggalkanBerapaKelas);
    
    for (let i = 1; i <= jumlahKelas; i++) {
      const idKelas = formData[`kelas_id_diampuKe${i}`];
      const namaKelas = idKelas ? kelasMap[idKelas] : '';
      
      // Simpan ke sheet FormIjinKeluar
      rowData.push(namaKelas || '');
      rowData.push(formData[`linkTugasKe${i}`] || '');
      
      // Simpan ke sheet deskripsiTugas
      if (idKelas && namaKelas) {
        const deskripsiData = {
          id_kelas: idKelas, 
          deskripsi: formData[`deskripsiTugasKe${i}`] || `Tugas untuk ${namaKelas} - ${formData.keperluan}`,
          
          // --- PERBAIKAN: urlFoto HANYA mengambil nilai upload. Jangan fallback ke linkTugas ---
          urlFoto: formData[`urlFotoKe${i}`] || '' 
        };
        saveDeskripsiTugas(deskripsiData);
      }
    }

    // Isi sisa kolom jika kurang dari 10 kelas
    const maxKelas = 10;
    for (let i = jumlahKelas + 1; i <= maxKelas; i++) {
      rowData.push('');
      rowData.push('');
    }

    rowData.push('');

    sheet.appendRow(rowData);
    
    // Update warna jadwal
    resetWarnaJadwal();
    warnaiJadwalIzinKeluar();
    
    return "Data izin keluar dan deskripsi tugas berhasil disimpan";

  } catch (error) {
    Logger.log("Error in saveIjinKeluar: " + error.toString());
    return "Error: " + error.toString();
  }
}

function getAllDeskripsiTugas() {
  try {
    const sheet = sheetDb.getSheetByName("deskripsiTugas");
    if (!sheet) {
      console.log('Sheet deskripsiTugas tidak ditemukan');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    let output = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const namaKelas = row[0] ? row[0].toString().trim() : '';
      
      if (namaKelas !== '') {
        output.push({
          id_kelas: namaKelas, // Ini adalah nama kelas, bukan ID numerik
          deskripsi: row[1] ? row[1].toString() : '',
          urlFoto: row[2] ? row[2].toString() : ''
        });
        
        console.log('Deskripsi found:', namaKelas);
      }
    }
    
    console.log('Total deskripsi:', output.length);
    return output;
    
  } catch (error) {
    console.log('Error in getAllDeskripsiTugas:', error.toString());
    return [];
  }
}
function debugFormIjinKeluar() {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) return "Sheet tidak ditemukan";
    
    const data = sheet.getDataRange().getValues();
    const result = {
      headers: data[0],
      totalRows: data.length,
      rows: []
    };
    
    Logger.log("=== DEBUG FORM IJIN KELUAR ===");
    Logger.log("Headers: " + data[0].join(" | "));
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowData = {
        rowNumber: i + 1,
        tanggal: row[0],
        tanggalType: typeof row[0],
        namaGuru: row[1],
        keperluan: row[2],
        jamMulai: row[3],
        jamSelesai: row[4],
        jumlahKelas: row[5],
        kelas: []
      }
      for (let j = 1; j <= 9; j++) {
        const kelasIndex = 6 + ((j - 1) * 2);
        const linkIndex = kelasIndex + 1;
        
        if (kelasIndex < row.length) {
          const kelas = row[kelasIndex] ? row[kelasIndex].toString().trim() : '';
          const link = row[linkIndex] ? row[linkIndex].toString().trim() : '';
          
          if (kelas !== '') {
            rowData.kelas.push({
              no: j,
              nama: kelas,
              link: link
            });
          }
        }
      }
      
      result.rows.push(rowData)
      Logger.log("Row " + (i+1) + ": " + 
                 "Tanggal: " + row[0] + " (" + typeof row[0] + "), " +
                 "Guru: " + row[1] + ", " +
                 "Kelas: " + rowData.kelas.length);
    }
    
    return result;
    
  } catch (error) {
    Logger.log("Error in debugFormIjinKeluar: " + error.toString());
    return "Error: " + error.toString();
  }
}

function debugGetAllIzinData() {
  return debugFormIjinKeluar()
}

function testGetPenugasan() {
  var result = getPenugasanByTanggal("2025-11-16");
  Logger.log("Test result: " + JSON.stringify(result));
  return result;
}
function getDataGuruIjinKeluar() {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) {
      Logger.log("Sheet FormIjinKeluar tidak ditemukan");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const today = new Date()
    const tYear = today.getFullYear();
    const tMonth = today.getMonth();
    const tDate = today.getDate();

    let result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tanggal = row[0];

      if (!(tanggal instanceof Date)) continue
      if (
        tanggal.getFullYear() === tYear &&
        tanggal.getMonth() === tMonth &&
        tanggal.getDate() === tDate
      ) {
        const namaGuru = row[2] || "";
        const keperluan = row[3] || "";
        const mulai = row[4] || "";
        const selesai = row[5] || "";
        const jumlahKelas = row[6] || 0
        let kelasTugas = [];
        const maxKelas = 10;
        let colStart = 7;

        for (let k = 1; k <= maxKelas; k++) {
          const kelas = row[colStart] || "";
          const link = row[colStart + 1] || "";
          if (kelas !== "") {
            kelasTugas.push({ kelas, link });
          }
          colStart += 2;
        }

        result.push({
          tanggal,
          namaGuru,
          keperluan,
          mulai,
          selesai,
          jumlahKelas,
          kelasTugas,
        });
      }
    }

    Logger.log("Total guru ijin hari ini: " + result.length);
    return result;

  } catch (error) {
    Logger.log("Error getDataGuruIjinKeluar: " + error.toString());
    return [];
  }
}
function testIzin() {
  var d = getDataGuruIjinKeluar();
  Logger.log(JSON.stringify(d, null, 2));
}
function getDataGuruIzinHariIni() {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) {
      Logger.log("Sheet FormIjinKeluar tidak ditemukan");
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const today = new Date();
    const todayString = Utilities.formatDate(today, "Asia/Jakarta", "yyyy-MM-dd");
    
    let result = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const tanggal = row[0];
      
      if (!tanggal) continue;
      
      // Format tanggal untuk perbandingan
      const tanggalString = Utilities.formatDate(tanggal, "Asia/Jakarta", "yyyy-MM-dd");
      
      if (tanggalString === todayString) {
        const namaGuru = row[1] || "";
        const keperluan = row[2] || "";
        const jamMulai = formatJam(row[3]);
        const jamSelesai = formatJam(row[4]);
        const jumlahKelas = parseInt(row[5]) || 0;
        
        let kelasTugas = [];
        let colStart = 6; // Kolom pertama untuk kelas dimulai dari index 6

        for (let k = 1; k <= jumlahKelas; k++) {
          const kelas = row[colStart] ? row[colStart].toString().trim() : "";
          const linkTugas = row[colStart + 1] ? row[colStart + 1].toString().trim() : "";
          
          if (kelas !== "") {
            // Ambil deskripsi tugas untuk kelas ini
            const deskripsiData = getDeskripsiTugasByNamaKelas(kelas);
            
            kelasTugas.push({
              namaKelas: kelas,
              linkTugas: linkTugas,
              deskripsi: deskripsiData ? deskripsiData.deskripsi : "",
              urlFoto: deskripsiData ? deskripsiData.urlFoto : ""
            });
          }
          colStart += 2; // Pindah ke pasangan kelas-link berikutnya
        }

        result.push({
          id: i, // ID baris untuk referensi
          namaGuru: namaGuru,
          keperluan: keperluan,
          jamMulai: jamMulai,
          jamSelesai: jamSelesai,
          jumlahKelas: jumlahKelas,
          kelasTugas: kelasTugas
        });
      }
    }

    Logger.log("Total guru izin hari ini: " + result.length);
    return result;

  } catch (error) {
    Logger.log("Error in getDataGuruIzinHariIni: " + error.toString());
    return [];
  }
}
//=======================TUGAS SISWA======================

function getDetailTugasModal(idIzin) {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    
    if (idIzin < 1 || idIzin >= data.length) return null;
    
    const row = data[idIzin];
    const jumlahKelas = parseInt(row[5]) || 0;
    
    let kelasTugas = [];
    let colStart = 6;

    for (let k = 1; k <= jumlahKelas; k++) {
      const kelas = row[colStart] ? row[colStart].toString().trim() : "";
      const linkTugas = row[colStart + 1] ? row[colStart + 1].toString().trim() : "";
      
      if (kelas !== "") {
        const deskripsiData = getDeskripsiTugasByNamaKelas(kelas);
        
        kelasTugas.push({
          namaKelas: kelas,
          linkTugas: linkTugas,
          deskripsi: deskripsiData ? deskripsiData.deskripsi : "",
          urlFoto: deskripsiData ? deskripsiData.urlFoto : ""
        });
      }
      colStart += 2;
    }

    return {
      namaGuru: row[1] || "",
      keperluan: row[2] || "",
      jamMulai: formatJam(row[3]),
      jamSelesai: formatJam(row[4]),
      tanggal: Utilities.formatDate(row[0], "Asia/Jakarta", "dd/MM/yyyy"),
      jumlahKelas: jumlahKelas,
      kelasTugas: kelasTugas
    };
    
  } catch (error) {
    Logger.log("Error in getDetailTugasModal: " + error.toString());
    return null;
  }
}
function formatJam(jam) {
  if (jam instanceof Date) {
    return Utilities.formatDate(jam, "Asia/Jakarta", "HH:mm");
  }
  return jam ? jam.toString() : "-";
}
function getDetailTugasById(idIzin, indexKelas) {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    
    if (idIzin < 1 || idIzin >= data.length) return null;
    
    const row = data[idIzin];
    const jumlahKelas = parseInt(row[5]) || 0;
    
    if (indexKelas < 1 || indexKelas > jumlahKelas) return null;
    
    const colStart = 6 + ((indexKelas - 1) * 2);
    const namaKelas = row[colStart] ? row[colStart].toString().trim() : "";
    const linkTugas = row[colStart + 1] ? row[colStart + 1].toString().trim() : "";
    
    if (!namaKelas) return null;
    
    // Ambil deskripsi lengkap
    const deskripsiData = getDeskripsiTugasByNamaKelas(namaKelas);
    
    return {
      namaGuru: row[1] || "",
      keperluan: row[2] || "",
      jamMulai: formatJam(row[3]),
      jamSelesai: formatJam(row[4]),
      namaKelas: namaKelas,
      linkTugas: linkTugas,
      deskripsi: deskripsiData ? deskripsiData.deskripsi : "",
      urlFoto: deskripsiData ? deskripsiData.urlFoto : "",
      tanggal: Utilities.formatDate(row[0], "Asia/Jakarta", "dd/MM/yyyy")
    };
    
  } catch (error) {
    Logger.log("Error in getDetailTugasById: " + error.toString());
    return null;
  }
}
function getAllTugasByKelas(namaKelas) {
  try {
    const semuaIzin = getDataGuruIzinHariIni();
    let tugasKelas = [];
    
    semuaIzin.forEach(izin => {
      izin.kelasTugas.forEach(tugas => {
        if (tugas.namaKelas === namaKelas) {
          tugasKelas.push({
            guru: izin.namaGuru,
            keperluan: izin.keperluan,
            jam: `${izin.jamMulai} - ${izin.jamSelesai}`,
            deskripsi: tugas.deskripsi,
            linkTugas: tugas.linkTugas,
            urlFoto: tugas.urlFoto
          });
        }
      });
    });
    
    return tugasKelas;
    
  } catch (error) {
    Logger.log("Error in getAllTugasByKelas: " + error.toString());
    return [];
  }
}
function formatDataUntukTampilan() {
  try {
    const dataIzin = getDataGuruIzinHariIni();
    
    const formattedData = dataIzin.map(izin => {
      const buttonDetails = izin.kelasTugas.map((tugas, index) => {
        return {
          kelas: tugas.namaKelas,
          buttonText: `Detail Tugas ${tugas.namaKelas}`,
          dataAttributes: {
            idIzin: izin.id,
            indexKelas: index + 1
          }
        };
      });
      
      return {
        guru: izin.namaGuru,
        keperluan: izin.keperluan,
        jam: `${izin.jamMulai} - ${izin.jamSelesai}`,
        jumlahKelas: izin.jumlahKelas,
        buttonDetails: buttonDetails
      };
    });
    
    return formattedData;
    
  } catch (error) {
    Logger.log("Error in formatDataUntukTampilan: " + error.toString());
    return [];
  }
}
function testDataGuruIzin() {
  Logger.log("=== TEST DATA GURU IZIN ===");
  
  const dataIzin = getDataGuruIzinHariIni();
  
  dataIzin.forEach((izin, index) => {
    Logger.log(`\n--- Guru Izin #${index + 1} ---`);
    Logger.log(`Nama: ${izin.namaGuru}`);
    Logger.log(`Keperluan: ${izin.keperluan}`);
    Logger.log(`Jam: ${izin.jamMulai} - ${izin.jamSelesai}`);
    Logger.log(`Jumlah Kelas: ${izin.jumlahKelas}`);
    
    izin.kelasTugas.forEach((tugas, tugasIndex) => {
      Logger.log(`  Kelas ${tugasIndex + 1}: ${tugas.namaKelas}`);
      Logger.log(`  Link: ${tugas.linkTugas}`);
      Logger.log(`  Deskripsi: ${tugas.deskripsi}`);
    });
  });
  
  return dataIzin;
}

function testDetailTugas() {
  Logger.log("=== TEST DETAIL TUGAS ===");
  
  // Test mendapatkan detail tugas untuk guru pertama, kelas pertama
  const dataIzin = getDataGuruIzinHariIni();
  if (dataIzin.length > 0) {
    const detail = getDetailTugasById(dataIzin[0].id, 1);
    Logger.log("Detail tugas:");
    Logger.log(JSON.stringify(detail, null, 2));
    return detail;
  }
  
  return null;
}
function getTugasHariIni() {
  try {
    const sheet = sheetDb.getSheetByName("FormIjinKeluar");
    if (!sheet) {
      Logger.log("Sheet FormIjinKeluar tidak ditemukan");
      return [];
    }
    
    const today = new Date();
    const dataIzin = sheet.getDataRange().getValues();
    const tugasHariIni = []
    for (let i = 1; i < dataIzin.length; i++) {
      const row = dataIzin[i];
      const tglIzin = row[0]
      if (!isSameDate(tglIzin, today)) {
        continue;
      }
      
      const namaGuru = row[1];
      const keperluan = row[2];
      const jamMulai = row[3];
      const jamSelesai = row[4];
      const jumlahKelas = parseInt(row[5]) || 0
      for (let j = 1; j <= jumlahKelas; j++) {
        const kelasIndex = 6 + ((j - 1) * 2);
        const linkIndex = kelasIndex + 1;
        
        const namaKelas = row[kelasIndex] ? row[kelasIndex].toString().trim() : '';
        const linkTugas = row[linkIndex] ? row[linkIndex].toString().trim() : '';
        
        if (namaKelas !== '') {
          const deskripsiData = getDeskripsiTugasByNamaKelas(namaKelas)
          
          tugasHariIni.push({
            namaGuru: namaGuru,
            keperluan: keperluan,
            jamMulai: formatJam(jamMulai),
            jamSelesai: formatJam(jamSelesai),
            namaKelas: namaKelas,
            linkTugas: linkTugas,
            deskripsi: deskripsiData ? deskripsiData.deskripsi : '',
            urlFoto: deskripsiData ? deskripsiData.urlFoto : ''
          });
        }
      }
    }
    
    Logger.log("Total tugas hari ini: " + tugasHariIni.length);
    return tugasHariIni;
    
  } catch (error) {
    Logger.log("Error in getTugasHariIni: " + error.toString());
    return [];
  }
}
function formatJam(jam) {
  if (jam instanceof Date) {
    return Utilities.formatDate(jam, "Asia/Jakarta", "HH:mm");
  }
  return jam;
}

function getDeskripsiTugasByNamaKelas(namaKelas) {
  try {
    const sheet = sheetDb.getSheetByName("deskripsiTugas");
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    
    // Loop mulai dari baris ke-2 (index 1) karena baris ke-1 adalah header
    for (let i = 1; i < data.length; i++) {
      // Cek Kolom A (index 0) apakah sama dengan namaKelas yang dicari
      if (data[i][0] == namaKelas) {
        return {
          nama_kelas: data[i][0],       // Ambil dari Kolom A
          deskripsi: data[i][1] || '',    // Ambil dari Kolom B
          urlFoto: data[i][2] || ''      // Ambil dari Kolom C (URL Foto)
        };
      }
    }
    return null; // Jika kelas tidak ditemukan
  } catch (error) {
    Logger.log("Error in getDeskripsiTugasByNamaKelas: " + error.toString());
    return null;
  }
}
//===============testing tugas

function testGetTugasHariIni() {
  const tugas = getTugasHariIni();
  Logger.log("=== TEST GET TUGAS HARI INI ===");
  Logger.log("Total tugas: " + tugas.length);
  
  tugas.forEach(function(t, index) {
    Logger.log("\n--- Tugas " + (index + 1) + " ---");
    Logger.log("Guru: " + t.namaGuru);
    Logger.log("Kelas: " + t.namaKelas);
    Logger.log("Jam: " + t.jamMulai + " - " + t.jamSelesai);
    Logger.log("Keperluan: " + t.keperluan);
    Logger.log("Deskripsi: " + t.deskripsi);
    Logger.log("Link: " + t.linkTugas);
    Logger.log("Foto: " + t.urlFoto);
  });
  
  return tugas;
}
//======================WARNA IZIN================
function submitIjin(formData) {
  const result = saveIjinKeluar(formData);
  
  if (result.includes("berhasil")) {
    warnaiJadwalIzinKeluar();
  }

  return result;
}


function warnaiJadwalIzinKeluar() { 
  const ss = SpreadsheetApp.openById(idDb);
  const sheetIzin = ss.getSheetByName("FormIjinKeluar");
  if (!sheetIzin) return;

  const today = new Date();
  const namaHari = getNamaHari(today.getDay());
  if (!namaHari) return;

  const sheetJadwal = ss.getSheetByName(namaHari);
  if (!sheetJadwal) return;

  const dataIzin = sheetIzin.getDataRange().getValues();
  const dataJadwal = sheetJadwal.getDataRange().getValues();
  const headerJam = dataJadwal[0].slice(1);

  for (let i = 1; i < dataIzin.length; i++) {
    const row = dataIzin[i];
    const tglIzin = row[0];
    if (!isSameDate(tglIzin, today)) continue;

    const namaGuru = row[1];
    const jamMulai = parseWaktu(row[3]);
    const jamSelesai = parseWaktu(row[4]);

    const rowIndexGuru = dataJadwal.findIndex(r => r[0] === namaGuru);
    if (rowIndexGuru === -1) continue;

    headerJam.forEach((jamHeader, j) => {
      if (!jamHeader) return;
      const jamKe = parseInt(jamHeader);
      if (isNaN(jamKe)) return;

      // Hitung waktu pelajaran sesuai jam ke
      const waktuMulaiPel = new Date(today);
      waktuMulaiPel.setHours(7, (jamKe - 1) * 45, 0, 0);

      const waktuSelesaiPel = new Date(waktuMulaiPel);
      waktuSelesaiPel.setMinutes(waktuMulaiPel.getMinutes() + 45);

      if (isWaktuOverlap(jamMulai, jamSelesai, waktuMulaiPel, waktuSelesaiPel)) {
        sheetJadwal.getRange(rowIndexGuru + 1, j + 2).setBackground("#ff4444");
      }
    });
  }
}

// ============= FUNGSI HELPER =============

function parseWaktu(waktuStr) {
  if (!waktuStr) return null;
  
  try {
    const today = new Date();
    
    // Handle jika waktuStr adalah Date object dari Google Sheets
    if (waktuStr instanceof Date) {
      // Ambil jam dan menit dari Date object
      const jam = waktuStr.getHours();
      const menit = waktuStr.getMinutes();
      
      // Set ke hari ini dengan jam/menit yang sama
      today.setHours(jam, menit, 0, 0);
      return today;
    }
    
    // Handle jika waktuStr adalah string "HH:MM"
    const waktuStrClean = waktuStr.toString().trim();
    const parts = waktuStrClean.split(":");
    if (parts.length !== 2) return null;
    
    const jam = parseInt(parts[0]);
    const menit = parseInt(parts[1]);
    
    if (isNaN(jam) || isNaN(menit)) return null;
    
    today.setHours(jam, menit, 0, 0);
    return today;
  } catch (e) {
    Logger.log("Error parseWaktu: " + e.toString());
    return null;
  }
}

function isSameDate(date1, date2) {
  if (!date1 || !date2) return false;
  
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  
  return d1.getDate() === d2.getDate() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getFullYear() === d2.getFullYear();
}

function isWaktuOverlap(mulai1, selesai1, mulai2, selesai2) {
  return mulai1 < selesai2 && selesai1 > mulai2;
}
function getNamaHari(day) {
  const map = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  return map[day] || "";
}

// ============= FUNGSI UNTUK DIPANGGIL SETELAH SAVE =============

function updateWarnaJadwalIzin() {
  Logger.log("=== Memulai update warna jadwal ===");
  warnaiJadwalIzinKeluar();
}
// ============= FUNGSI TESTING =============

function testWarnaiJadwal() {
  Logger.log("=== TESTING WARNAI JADWAL ===");
  Logger.log("Hari ini: " + getNamaHari(new Date().getDay()));
  warnaiJadwalIzinKeluar();
}

function testDataIzinHariIni() {
  const ss = SpreadsheetApp.openById(idDb);
  const sheetIzin = ss.getSheetByName("FormIjinKeluar");
  const dataIzin = sheetIzin.getDataRange().getValues();
  const today = new Date();
  
  Logger.log("=== DATA IZIN HARI INI ===");
  Logger.log("Tanggal: " + Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy"));
  
  let count = 0;
  dataIzin.slice(1).forEach((row, index) => {
    if (isSameDate(row[0], today)) {
      count++;
      Logger.log("\nIzin #" + count);
      Logger.log("Guru: " + row[1]);
      Logger.log("Jam: " + row[3] + " - " + row[4]);
      Logger.log("Jumlah Kelas: " + row[5]);
      for (let i = 1; i <= parseInt(row[5]); i++) {
        const kelasIndex = 6 + ((i - 1) * 2);
        Logger.log("  Kelas " + i + ": " + row[kelasIndex]);
      }
    }
  });
  
  Logger.log("\n=== Total izin hari ini: " + count + " ===");
}

function isSameDate(date1, date2) {
  if (!date1 || !date2) return false;
  
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  
  return d1.getDate() === d2.getDate() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getFullYear() === d2.getFullYear();
}

function isWaktuOverlap(mulai1, selesai1, mulai2, selesai2) {
  // Cek apakah ada overlap antara dua rentang waktu
  return mulai1 < selesai2 && selesai1 > mulai2;
}
// ============= FUNGSI TESTING =============
function testDataIzinHariIni() {
  const ss = SpreadsheetApp.openById(idDb);
  const sheetIzin = ss.getSheetByName("FormIjinKeluar");
  const dataIzin = sheetIzin.getDataRange().getValues();
  const today = new Date();
  
  Logger.log("=== DATA IZIN HARI INI ===");
  Logger.log("Tanggal: " + Utilities.formatDate(today, "Asia/Jakarta", "dd/MM/yyyy"));
  
  let count = 0;
  dataIzin.slice(1).forEach((row, index) => {
    if (isSameDate(row[0], today)) {
      count++;
      Logger.log("\nIzin #" + count);
      Logger.log("Guru: " + row[1]);
      Logger.log("Jam: " + row[3] + " - " + row[4]);
      Logger.log("Jumlah Kelas: " + row[5]);
      for (let i = 1; i <= parseInt(row[5]); i++) {
        const kelasIndex = 6 + ((i - 1) * 2);
        Logger.log("  Kelas " + i + ": " + row[kelasIndex]);
      }
    }
  });
  
  Logger.log("\n=== Total izin hari ini: " + count + " ===");
}

function isSameDate(date1, date2) {
  if (!date1 || !date2) return false;
  
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  
  return d1.getDate() === d2.getDate() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getFullYear() === d2.getFullYear();
}

function isWaktuOverlap(mulai1, selesai1, mulai2, selesai2) {
  return mulai1 < selesai2 && selesai1 > mulai2;
}

function resetWarnaJadwal() {
  try {
    const ss = SpreadsheetApp.openById(idDb);
    const today = new Date();
    const namaHari = getNamaHari(today.getDay());
    
    if (!namaHari || namaHari === "") return;
    
    const sheetJadwal = ss.getSheetByName(namaHari);
    if (!sheetJadwal) return;
    
    const lastRow = sheetJadwal.getLastRow();
    const lastCol = sheetJadwal.getLastColumn();
    
    if (lastRow > 1 && lastCol > 1) {
      const backgrounds = sheetJadwal.getRange(2, 2, lastRow - 1, lastCol - 1).getBackgrounds();
      
      for (let i = 0; i < backgrounds.length; i++) {
        for (let j = 0; j < backgrounds[i].length; j++) {
          const warna = backgrounds[i][j].toLowerCase();
          if (warna === "#ff4444" || warna === "#f44" || warna === "#ff4444ff") {
            sheetJadwal.getRange(i + 2, j + 2).setBackground(null);
          }
        }
      }
    }
    
    Logger.log("Warna merah jadwal direset, warna lain tetap");
  } catch (error) {
    Logger.log("Error resetWarnaJadwal: " + error.toString());
  }
}
function updateWarnaJadwalIzin() {
  resetWarnaJadwal();
  warnaiJadwalIzinKeluar();
}
 //======================= RESET WARNA IZIN SETIAP HARI========================
function setupTriggerResetHarian() {
  const triggers = ScriptApp.getProjectTriggers()
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'resetDanWarnaOtomatis') {
      ScriptApp.deleteTrigger(trigger);
    }
  })
  ScriptApp.newTrigger('resetDanWarnaOtomatis')
    .timeBased()
    .atHour(0) 
    .everyDays(1) 
    .create();
  
  Logger.log('Trigger otomatis berhasil dibuat!');
  Logger.log('Reset warna akan jalan otomatis setiap hari jam 00:05');
}

function resetDanWarnaOtomatis() {
  try {
    Logger.log('=== RESET OTOMATIS DIMULAI ===');
    Logger.log('Waktu: ' + new Date());
    resetWarnaJadwal()
    warnaiJadwalIzinKeluar();
    
    Logger.log('Reset otomatis selesai');
  } catch (error) {
    Logger.log('Error reset otomatis: ' + error.toString());
  }
}

function cekStatusTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  const triggerAktif = triggers.filter(t => 
    t.getHandlerFunction() === 'resetDanWarnaOtomatis'
  );
  
  if (triggerAktif.length > 0) {
    Logger.log('✓ Trigger otomatis AKTIF');
    triggerAktif.forEach(trigger => {
      Logger.log('  - Handler: ' + trigger.getHandlerFunction());
      Logger.log('  - Tipe: Time-based (Daily)');
    });
  } else {
    Logger.log('Trigger otomatis TIDAK AKTIF');
    Logger.log('Jalankan setupTriggerResetHarian() untuk mengaktifkan');
  }
  
  return triggerAktif.length > 0;
}
