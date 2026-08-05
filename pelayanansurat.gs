// ======================================
// PELAYANANSURAT.GS
// Modul Pelayanan Surat
//
// Struktur Sheet "PelayananSurat":
// A = ID
// B = No Permohonan
// C = NIK
// D = Nama
// E = Jenis Surat
// F = Tanggal Permohonan
// G = Tanggal Selesai
// H = Keperluan
// I = Status
// J = File
// K = Petugas
// L = No HP
// M = Keterangan
// ======================================

// ======================================
// HALAMAN HTML
// ======================================
function getPelayananSuratPage(){
  return HtmlService
    .createHtmlOutputFromFile("PelayananSurat")
    .getContent();
}

// ======================================
// AMBIL DATA
// ======================================
function getPelayananSurat(token){
  const session = verifySession(token);
  if(!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if(!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const data = sh.getDataRange().getValues();
  const safeData = data.map(row => row.map(cell => {
    if(Object.prototype.toString.call(cell) === "[object Date]"){
      return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
    return cell;
  }));
  return safeData;
}

// AMBIL DATA BERDASARKAN ID
function getPelayananSuratById(token, id){
  const session = verifySession(token);
  if(!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if(!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]) === String(id)){
      return {
        id: data[i][0],
        noPermohonan: data[i][1],
        nik: data[i][2],
        nama: data[i][3],
        jenisSurat: data[i][4],
        tanggalPermohonan: Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        tanggalSelesai: data[i][6] ? Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
        keperluan: data[i][7],
        status: data[i][8],
        file: data[i][9],
        petugas: data[i][10],
        nohp: data[i][11],
        keterangan: data[i][12]
      };
    }
  }
  return null;
}

// ======================================
// FOLDER GOOGLE DRIVE
// ======================================
function getFolderPelayananPersyaratan(){
  const nama = "Pelayanan Surat - Persyaratan";
  const folder = DriveApp.getFoldersByName(nama);
  return folder.hasNext() ? folder.next() : DriveApp.createFolder(nama);
}

function getFolderPelayananSuratJadi(){
  const nama = "Pelayanan Surat - Surat Jadi";
  const folder = DriveApp.getFoldersByName(nama);
  return folder.hasNext() ? folder.next() : DriveApp.createFolder(nama);
}

// ======================================
// UPLOAD FILE PERSYARATAN
// ======================================
function uploadPersyaratanPelayanan(token, base64Data, fileName, mimeType){
  const session = verifySession(token);
  if(!session) throw new Error("Sesi login tidak valid.");

  const folder = getFolderPelayananPersyaratan();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ======================================
// UPLOAD FILE SURAT JADI
// ======================================
function uploadSuratJadiPelayanan(token, base64Data, fileName, mimeType){
  const session = verifySession(token);
  if(!session) throw new Error("Sesi login tidak valid.");

  const folder = getFolderPelayananSuratJadi();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ======================================
// UBAH STATUS
// ======================================
function updateStatusPelayanan(token, id, status){
  const session = verifySession(token);
  if(!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if(!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]) === String(id)){
      sh.getRange(i + 1, 9).setValue(status);
      return true;
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// ======================================
// SIMPAN SURAT JADI & TANDAI SELESAI
// ======================================
function simpanSuratJadi(token, id, fileURL){
  const session = verifySession(token);
  if(!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if(!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const data = sh.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(String(data[i][0]) === String(id)){
      sh.getRange(i + 1, 10).setValue(fileURL); // Kolom J = File
      sh.getRange(i + 1, 9).setValue("Selesai"); // Kolom I = Status
      sh.getRange(i + 1, 7).setValue(new Date()); // Kolom G = Tanggal Selesai
      return true;
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// ======================================================
// TAMBAH PELAYANAN SURAT
// ======================================================
function tambahPelayananSurat(token, data) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  // Validasi
  if (!data.nik) throw new Error("NIK wajib diisi.");
  if (!data.nama) throw new Error("Nama wajib diisi.");
  if (!data.jenisSurat) throw new Error("Jenis Surat wajib dipilih.");

  // Hitung ID baru
  const values = sh.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < values.length; i++) {
    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }

  const newId = maxId + 1;
  const nomorPermohonan = "PS-" + Utilities.formatString("%04d", newId);

  // Simpan data sesuai urutan kolom
  sh.appendRow([
    newId,                        // A
    nomorPermohonan,              // B
    data.nik,                     // C
    data.nama,                    // D
    data.jenisSurat,              // E
    data.tanggal || new Date(),   // F
    "",                           // G Tanggal selesai
    data.keperluan || "",         // H
    "Menunggu",                   // I Status
    data.fileURL || "",           // J
    "",                           // K Petugas
    data.nohp || "",              // L
    data.keterangan || ""         // M
  ]);

  return true;
}

//====================================================
// EDIT PELAYANAN SURAT
//====================================================
function editPelayananSurat(token, id, data) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const values = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      row = i + 1;
      break;
    }
  }
  if (row === -1) throw new Error("Data tidak ditemukan.");

  const fileLama = values[row - 1][9];

  // Update sesuai urutan kolom, tidak bergeser
  sh.getRange(row, 2, 1, 12).setValues([[
    values[row-1][1],          // B No Permohonan (tetap)
    data.nik,                  // C
    data.nama,                 // D
    data.jenisSurat,           // E
    data.tanggal,              // F
    data.tanggalSelesai || "", // G
    data.keperluan,            // H
    data.status || "Menunggu", // I
    data.fileURL || fileLama,  // J
    data.petugas || "",        // K
    data.nohp || "",           // L
    data.keterangan || ""      // M
  ]]);

  return true;
}

//====================================================
// HAPUS PELAYANAN SURAT
//====================================================
function hapusPelayananSurat(token, id) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return true;
    }
  }

  throw new Error("Data tidak ditemukan.");
}