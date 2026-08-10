// ======================================
// PENGADUAN.GS
// Membaca & mengelola data Pengaduan yang SUDAH dikategorikan
// oleh Router Script (dari Google Form -> tab "Pengaduan" -> tab status).
//
// Struktur kolom (9 kolom, SAMA di semua tab status):
// A=Timestamp B=Nama C=NIK D=Alamat E=NoHP
// F=KategoriPengaduan G=Deskripsi H=BuktiKejadian I=Status
//
// PENTING: Router Script TIDAK diubah. File ini hanya MEMBACA
// dan MEMINDAHKAN baris antar tab status yang sudah ada
// (Diterima / Diproses / Selesai) -- pola yang sama dengan
// automasi Sheets paling awal di project ini.
// ======================================

const STAGE_SHEETS_PENGADUAN = ["Diterima", "Diproses", "Selesai"];

function getPengaduanPage(){
  return HtmlService
    .createHtmlOutputFromFile("Pengaduan")
    .getContent();
}

// ======================================
// AMBIL SEMUA DATA (gabungan dari 3 tab status)
// ======================================
function getPengaduan(token){
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  let hasil = [];

  STAGE_SHEETS_PENGADUAN.forEach(function(namaTab){
    const sh = getSheet(namaTab);
    if (!sh) return; // tab belum ada, lewati saja (jangan error)

    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i].map(function(cell){
        if (Object.prototype.toString.call(cell) === "[object Date]") {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        }
        return cell;
      });
      hasil.push(row);
    }
  });

  return hasil;
}

// ======================================
// UBAH STATUS -> PINDAHKAN BARIS KE TAB TUJUAN
// Dicari berdasarkan kombinasi Timestamp + NIK (kunci unik,
// karena data dari Form tidak punya kolom ID angka).
// ======================================
function ubahStatusPengaduan(token, timestampAsli, nik, statusBaru){
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const targetNama = STAGE_SHEETS_PENGADUAN.find(function(nama){
    return nama.toLowerCase() === statusBaru.toString().trim().toLowerCase();
  });
  if (!targetNama) throw new Error("Status tujuan tidak dikenali: " + statusBaru);

  // Cari baris di SEMUA tab status
  for (let s = 0; s < STAGE_SHEETS_PENGADUAN.length; s++) {
    const namaTab = STAGE_SHEETS_PENGADUAN[s];
    const sh = getSheet(namaTab);
    if (!sh) continue;

    const values = sh.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const tsCell = values[i][0];
      const tsString = Object.prototype.toString.call(tsCell) === "[object Date]"
        ? Utilities.formatDate(tsCell, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : tsCell.toString();

      const nikCell = values[i][2] ? values[i][2].toString().trim() : "";

      if (tsString === timestampAsli.toString().trim() && nikCell === nik.toString().trim()) {

        // Kalau target sama dengan tab sekarang, cukup update kolom Status saja
        if (namaTab === targetNama) {
          sh.getRange(i + 1, 9).setValue(targetNama);
          return true;
        }

        // Ambil baris, update kolom Status, pindahkan ke tab tujuan
        const rowData = values[i].slice();
        rowData[8] = targetNama; // kolom I = Status

        const targetSheet = getSheet(targetNama);
        if (!targetSheet) throw new Error("Tab tujuan '" + targetNama + "' tidak ditemukan.");

        targetSheet.appendRow(rowData);
        sh.deleteRow(i + 1);
        return true;
      }
    }
  }

  throw new Error("Data pengaduan tidak ditemukan (mungkin sudah dipindahkan sebelumnya, coba muat ulang halaman).");
}