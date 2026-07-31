const SPREADSHEET_ID = "1UC3ib4HfNRGW3SqQtg8k_tPAQzym7EohO8mdKKWE5XE";

// Mengambil objek Spreadsheet utuh (dipakai modul yang butuh akses ke banyak sheet sekaligus)
function getDatabase() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// Fungsi Koneksi ke satu Sheet tertentu
function getSheet(sheetName) {
  return getDatabase().getSheetByName(sheetName);
}

// Utilitas Mengubah Data Sheet Menjadi Format JSON
function getSheetDataAsJson(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = data.slice(1);
  
  return rows.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header.toString().trim()] = row[index];
    });
    return obj;
  });
}

// Keamanan: Enkripsi Password Menggunakan SHA-256
function hashPassword(text) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, 
    text, 
    Utilities.Charset.UTF_8
  );
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    var byteStr = byteVal.toString(16);
    if (byteStr.length == 1) byteStr = '0' + byteStr;
    txtHash += byteStr;
  }
  return txtHash;
}