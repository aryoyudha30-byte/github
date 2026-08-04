// ======================================
// SURATMASUK.GS
// Struktur kolom sheet "SuratMasuk":
// ID | NoAgenda | NomorSurat | TanggalSurat | TanggalTerima | Pengirim | Perihal | Keterangan | Status | File
// ======================================

// ======================================
// Ambil Data Surat Masuk
// ======================================
function getSuratMasuk(token){
  const session = verifySession(token);
  if(!session){
    throw new Error("Sesi login tidak valid.");
  }
  const sh = getSheet("SuratMasuk");
  if(!sh){
    throw new Error("Sheet SuratMasuk tidak ditemukan.");
  }
  const data = sh.getDataRange().getValues();
  const safeData = data.map(row =>
    row.map(cell => {
      if(Object.prototype.toString.call(cell)==="[object Date]"){
        return Utilities.formatDate(
          cell,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
        );
      }
      return cell;
    })
  );
  return safeData;
}

// ======================================
// Upload File Lampiran Surat
// Menerima file dalam bentuk base64 dari browser, simpan ke folder
// khusus di Drive, kembalikan link filenya untuk disimpan ke sheet.
// ======================================
function uploadFileSurat(token, base64Data, fileName, mimeType){
  const session = verifySession(token);
  if(!session){
    throw new Error("Sesi login tidak valid.");
  }

  const folder = getOrCreateFolderSuratMasuk();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = folder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function getOrCreateFolderSuratMasuk(){
  const namaFolder = "Lampiran Surat Masuk - Sistem Desa";
  const folders = DriveApp.getFoldersByName(namaFolder);
  if(folders.hasNext()){
    return folders.next();
  }
  return DriveApp.createFolder(namaFolder);
}

// ======================================
// Tambah Surat Masuk
// ======================================
function tambahSuratMasuk(token,data){
  const session = verifySession(token);
  if(!session){
    throw new Error("Sesi login tidak valid.");
  }
  const sh = getSheet("SuratMasuk");
  if(!sh){
    throw new Error("Sheet SuratMasuk tidak ditemukan.");
  }
  if(!data.nomorSurat){
    throw new Error("Nomor Surat wajib diisi.");
  }
  if(!data.pengirim){
    throw new Error("Pengirim wajib diisi.");
  }
  if(!data.perihal){
    throw new Error("Perihal wajib diisi.");
  }

  const values = sh.getDataRange().getValues();
  let maxId = 0;

  for(let i=1;i<values.length;i++){
    if(values[i][2] && values[i][2].toString() === data.nomorSurat.toString()){
      throw new Error("Nomor Surat sudah ada.");
    }
    const id = Number(values[i][0]);
    if(!isNaN(id) && id > maxId){
      maxId = id;
    }
  }

  const newId = maxId + 1;
  const noAgenda = "SM-" + Utilities.formatString("%04d", newId);

  sh.appendRow([
    newId,
    noAgenda,
    data.nomorSurat,
    data.tanggalSurat || '',
    data.tanggalTerima || '',
    data.pengirim,
    data.perihal,
    data.keterangan || '',
    data.status || 'Belum Diproses',
    data.fileURL || ''
  ]);
  return true;
}

// ======================================
// Edit Surat Masuk
// ======================================
function editSuratMasuk(token,id,data){
  const session = verifySession(token);
  if(!session){
    throw new Error("Sesi login tidak valid.");
  }
  const sh = getSheet("SuratMasuk");
  if(!sh){
    throw new Error("Sheet SuratMasuk tidak ditemukan.");
  }
  if(!data.nomorSurat){
    throw new Error("Nomor Surat wajib diisi.");
  }
  if(!data.pengirim){
    throw new Error("Pengirim wajib diisi.");
  }
  if(!data.perihal){
    throw new Error("Perihal wajib diisi.");
  }

  const values = sh.getDataRange().getValues();
  let row = -1;

  for(let i=1;i<values.length;i++){
    if(values[i][0].toString() === id.toString()){
      row = i + 1;
      continue;
    }
    if(values[i][2] && values[i][2].toString() === data.nomorSurat.toString()){
      throw new Error("Nomor Surat sudah dipakai surat lain.");
    }
  }

  if(row === -1){
    throw new Error("Data tidak ditemukan.");
  }

  const fileURLLama = values[row-1][9];
  const fileURLBaru = data.fileURL || fileURLLama || '';

  sh.getRange(row,2,1,9).setValues([[
    values[row-1][1],
    data.nomorSurat,
    data.tanggalSurat || '',
    data.tanggalTerima || '',
    data.pengirim,
    data.perihal,
    data.keterangan || '',
    data.status || 'Belum Diproses',
    fileURLBaru
  ]]);
  return true;
}

// ======================================
// Hapus Surat Masuk
// ======================================
function hapusSuratMasuk(token,id){
  const session = verifySession(token);
  if(!session){
    throw new Error("Sesi login tidak valid.");
  }
  const sh = getSheet("SuratMasuk");
  if(!sh){
    throw new Error("Sheet SuratMasuk tidak ditemukan.");
  }
  const values = sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(values[i][0].toString() === id.toString()){
      sh.deleteRow(i+1);
      return true;
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// ======================================
// Halaman
// ======================================
function getSuratMasukPage(){
  return HtmlService
    .createHtmlOutputFromFile("Suratmasuk")
    .getContent();
}