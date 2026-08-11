// ======================================
// PELAYANANSURAT.GS
// Struktur Sheet "PelayananSurat":
// A=ID B=NoPermohonan C=NIK D=Nama E=JenisSurat F=TanggalPermohonan
// G=TanggalSelesai H=Keperluan I=Status J=File K=Petugas L=NoHP M=Keterangan
// ======================================

function getPelayananSuratPage(){
  return HtmlService
    .createHtmlOutputFromFile("PelayananSurat")
    .getContent();
}

function getOrCreateFolderPelayanan(){
  const namaFolder = "Lampiran Pelayanan Surat - Sistem Desa";
  const folders = DriveApp.getFoldersByName(namaFolder);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(namaFolder);
}

function uploadFilePelayanan(token, base64Data, fileName, mimeType){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  if (!session) throw new Error("Sesi login tidak valid.");

  const folder = getOrCreateFolderPelayanan();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getPelayananSurat(token){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const data = sh.getDataRange().getValues();
  return data.map(row =>
    row.map(cell => {
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return cell;
    })
  );
}

function tambahPelayananSurat(token, data){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  if (!data.nik) throw new Error("NIK wajib diisi.");
  if (!data.nama) throw new Error("Nama wajib diisi.");
  if (!data.jenisSurat) throw new Error("Jenis Surat wajib dipilih.");

  const values = sh.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < values.length; i++) {
    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }

  const newId = maxId + 1;
  const nomorPermohonan = "PS-" + Utilities.formatString("%04d", newId);

  sh.appendRow([
    newId,                          // A ID
    nomorPermohonan,                // B No Permohonan
    data.nik,                       // C NIK
    data.nama,                      // D Nama
    data.jenisSurat,                // E Jenis Surat
    data.tanggal || new Date(),     // F Tanggal Permohonan
    "",                             // G Tanggal Selesai
    data.keperluan || "",           // H Keperluan
    data.status || "Menunggu",      // I Status
    data.fileURL || "",             // J File
    data.petugas || "",             // K Petugas
    data.nohp || "",                // L No HP
    data.keterangan || ""           // M Keterangan
  ]);

  return true;
}

function editPelayananSurat(token, id, data){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  if (!data.nik) throw new Error("NIK wajib diisi.");
  if (!data.nama) throw new Error("Nama wajib diisi.");
  if (!data.jenisSurat) throw new Error("Jenis Surat wajib dipilih.");

  const values = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id.toString()) {
      row = i + 1;
      break;
    }
  }
  if (row === -1) throw new Error("Data tidak ditemukan.");

  const fileLama = values[row - 1][9]; // kolom J (index 9)

  sh.getRange(row, 2, 1, 12).setValues([[
    values[row-1][1],               // B No Permohonan (tetap)
    data.nik,                       // C
    data.nama,                      // D
    data.jenisSurat,                // E
    data.tanggal || "",             // F
    data.tanggalSelesai || "",      // G
    data.keperluan || "",           // H
    data.status || "Menunggu",      // I
    data.fileURL || fileLama || "", // J
    data.petugas || "",             // K
    data.nohp || "",                // L
    data.keterangan || ""           // M
  ]]);

  return true;
}

function hapusPelayananSurat(token, id){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("PelayananSurat");
  if (!sh) throw new Error("Sheet PelayananSurat tidak ditemukan.");

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id.toString()) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error("Data tidak ditemukan.");
}