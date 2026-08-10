// ============================================================
// CONFIG.GS
// ============================================================

// DATABASE UTAMA
const SPREADSHEET_ID = "1UC3ib4HfNRGW3SqQtg8k_tPAQzym7EohO8mdKKWE5XE";

// DATABASE FORMULIR/PENGADUAN
const SPREADSHEET_ID_FORM = "1LLYKAyaCLUWnzzdYADLeW3i8uspBnMj0bJXU3QBJCAk";

// ============================================================
// FUNGSI AKSES DATABASE
// ============================================================
function getDatabase() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(sheetName) {
  return getDatabase().getSheetByName(sheetName);
}

function getFormDatabase() {
  return SpreadsheetApp.openById(SPREADSHEET_ID_FORM);
}

function getFormSheet(sheetName) {
  return getFormDatabase().getSheetByName(sheetName);
}

// ============================================================
// UTILITAS KONVERSI DATA
// ============================================================
function getSheetDataAsJson(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header.toString().trim()] = row[index];
    });
    return obj;
  });
}

// ============================================================
// ENKRIPSI PASSWORD
// ============================================================
function hashPassword(text) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, 
    text, 
    Utilities.Charset.UTF_8
  );
  let txtHash = "";
  for (let i = 0; i < rawHash.length; i++) {
    let byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    let byteStr = byteVal.toString(16);
    if (byteStr.length === 1) byteStr = "0" + byteStr;
    txtHash += byteStr;
  }
  return txtHash;
}