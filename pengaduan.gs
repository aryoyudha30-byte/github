// ============================================================
// PENGADUAN.GS — LOGIKA BACKEND
// ============================================================
const PENGADUAN_SHEET_NAME = "Pengaduan";
const PENGADUAN_TRACKER_SHEETS = ["Diterima", "Diproses", "Selesai"];
const PENGADUAN_STATUS_COL = 9; // Kolom I = Status
const TRACKER_NIK_COL = 3;      // Kolom C = NIK di tracker

// ============================================================
// AMBIL SEMUA DATA PENGADUAN
// ============================================================
function getPengaduan(token) {
  requireRole(token, ["administrator", "petugas pelayanan"]);

  const sheet = getFormSheet(PENGADUAN_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Pengaduan" tidak ditemukan.');
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  return values.map(row => {
    if (row[0] && Object.prototype.toString.call(row[0]) === "[object Date]") {
      row[0] = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    }
    return row;
  });
}

// ============================================================
// MUAT HALAMAN HTML
// ============================================================
function getPengaduanPage() {
  return HtmlService.createHtmlOutputFromFile("Pengaduan").getContent();
}

// ============================================================
// CARI BARIS DATA BERDASARKAN TIMESTAMP + NIK
// ============================================================
function findPengaduanRow(sheet, timestamp, nik) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowTime = normalizeTimestamp(row[0]);
    const targetTime = normalizeTimestamp(timestamp);
    const rowNik = String(row[2] || "").trim();
    const targetNik = String(nik || "").trim();
    if (rowTime === targetTime && rowNik === targetNik) return i + 2;
  }
  return -1;
}

function normalizeTimestamp(value) {
  if (!value) return "";

  // Kasus 1: value adalah objek Date asli dari sheet
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }

  const str = String(value).trim();

  // Kasus 2: value berformat "dd/MM/yyyy HH:mm" (hasil format dari getPengaduan())
  const match = str.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
  if (match) {
    const [, dd, mm, yyyy, hh, min] = match;
    return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
  }

  // Kasus 3: value berupa string ISO ("2026-08-09T20:59:33.000Z" dsb) atau format lain
  const parsed = new Date(str.replace("T", " "));
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  }

  return str;
}

// ============================================================
// UBAH STATUS PENGADUAN + SINCRON TRACKER
// ============================================================
function ubahStatusPengaduan(token, timestamp, nik, statusBaru) {
  requireRole(token, ["administrator", "petugas pelayanan"]);

  const statusValid = ["Diterima", "Diproses", "Selesai"];
  if (statusValid.indexOf(statusBaru) === -1) throw new Error("Status tidak valid.");
  
  const sheet = getFormSheet(PENGADUAN_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Pengaduan" tidak ditemukan.');
  
  const rowNumber = findPengaduanRow(sheet, timestamp, nik);
  if (rowNumber === -1) throw new Error("Data pengaduan tidak ditemukan.");
  
  // Update status di sheet Pengaduan
  sheet.getRange(rowNumber, PENGADUAN_STATUS_COL).setValue(statusBaru);
  
  // Ambil data untuk tracker
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const trackerRow = [row[0], row[1], row[2], row[3], row[4], row[5] || "Pengaduan", statusBaru];
  
  // Hapus dari tracker lama, tambah ke tracker baru
  removeFromAllTrackers(row[2], row[0]);
  const targetTracker = getFormSheet(statusBaru);
  if (!targetTracker) throw new Error('Sheet tracker "' + statusBaru + '" tidak ditemukan.');
  targetTracker.appendRow(trackerRow);
  
  return true;
}

// ============================================================
// HAPUS PENGADUAN
// ============================================================
function hapusPengaduan(token, timestamp, nik) {
  requireRole(token, ["administrator", "petugas pelayanan"]);

  const sheet = getFormSheet(PENGADUAN_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "Pengaduan" tidak ditemukan.');
  
  const rowNumber = findPengaduanRow(sheet, timestamp, nik);
  if (rowNumber === -1) throw new Error("Data pengaduan tidak ditemukan.");
  
  sheet.deleteRow(rowNumber);
  removeFromAllTrackers(nik, timestamp);
  return true;
}

// ============================================================
// HAPUS DATA DARI SEMUA TRACKER
// ============================================================
function removeFromAllTrackers(nik, timestamp) {
  PENGADUAN_TRACKER_SHEETS.forEach(sheetName => {
    const sheet = getFormSheet(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return;
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const rowNik = String(row[TRACKER_NIK_COL - 1] || "").trim();
      const rowTime = normalizeTimestamp(row[0]);
      if (rowNik === String(nik).trim() && rowTime === normalizeTimestamp(timestamp)) {
        sheet.deleteRow(i + 2);
      }
    }
  });
}

// ============================================================
// TEST KONEKSI (JALANKAN MANUAL UNTUK CEK)
// ============================================================
function testKoneksiPengaduan() {
  const ss = getFormDatabase();
  Logger.log("Spreadsheet: " + ss.getName());
  ss.getSheets().forEach(s => Logger.log("Sheet: " + s.getName()));
  const p = getFormSheet("Pengaduan");
  if (!p) throw new Error('Sheet "Pengaduan" tidak ditemukan.');
  Logger.log("✅ Koneksi Pengaduan berhasil!");
}

function tesDuaDatabase() {
  // 1. Tes Database Utama
  try {
    const dbUtama = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log("✅ Database Utama BERHASIL diakses: " + dbUtama.getName());
  } catch (err) {
    Logger.log("❌ Database Utama GAGAL: " + err.message);
  }

  // 2. Tes Database Formulir (tempat Pengaduan)
  try {
    const dbForm = SpreadsheetApp.openById(SPREADSHEET_ID_FORM);
    Logger.log("✅ Database Formulir BERHASIL diakses: " + dbForm.getName());
    
    const sheetPengaduan = dbForm.getSheetByName("Pengaduan");
    if (sheetPengaduan) {
      Logger.log("✅ Tab 'Pengaduan' DITEMUKAN, jumlah baris: " + sheetPengaduan.getLastRow());
    } else {
      Logger.log("❌ Tab 'Pengaduan' TIDAK ADA!");
      Logger.log("📋 Daftar tab yang ada: " + dbForm.getSheets().map(s => s.getName()).join(", "));
    }
  } catch (err) {
    Logger.log("❌ Database Formulir GAGAL: " + err.message);
  }
}