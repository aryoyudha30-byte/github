// ======================================
// SURATKELUAR.GS
// Struktur Sheet "SuratKeluar"
// A=ID B=NoAgenda C=NomorSurat D=TanggalSurat E=Tujuan
// F=Perihal G=Penandatangan H=Keterangan I=Status J=File
// ======================================

function parseDateInput_(dateStr){
  if(!dateStr) return "";
  if(Object.prototype.toString.call(dateStr)==="[object Date]"){
    return dateStr;
  }
  const parts=dateStr.toString().split("-");
  if(parts.length===3){
    return new Date(parts[0],parts[1]-1,parts[2]);
  }
  return "";
}

function validasiSuratKeluar_(data){
  if(!data.nomorSurat || data.nomorSurat.toString().trim()==""){
    throw new Error("Nomor Surat wajib diisi.");
  }
  if(!data.tujuan || data.tujuan.toString().trim()==""){
    throw new Error("Tujuan wajib diisi.");
  }
  if(!data.perihal || data.perihal.toString().trim()==""){
    throw new Error("Perihal wajib diisi.");
  }
  if(!data.penandatangan || data.penandatangan.toString().trim()==""){
    throw new Error("Penandatangan wajib diisi.");
  }
  if(data.nomorSurat.length>100){
    throw new Error("Nomor Surat maksimal 100 karakter.");
  }
  if(data.tujuan.length>150){
    throw new Error("Tujuan maksimal 150 karakter.");
  }
  if(data.perihal.length>250){
    throw new Error("Perihal maksimal 250 karakter.");
  }
  if(data.penandatangan.length>100){
    throw new Error("Nama Penandatangan maksimal 100 karakter.");
  }
  const statusValid=["Draft","Diproses","Dikirim","Diterima","Diarsipkan"];
  if(!statusValid.includes(data.status)){
    data.status="Draft";
  }
}

// ======================================
// Upload File Lampiran Surat Keluar
// ======================================
function uploadFileSuratKeluar(token, base64Data, fileName, mimeType){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  if(!session){
    throw new Error("Sesi login tidak valid.");
  }
  const folder = getOrCreateFolderSuratKeluar();
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateFolderSuratKeluar(){
  const namaFolder = "Lampiran Surat Keluar - Sistem Desa";
  const folders = DriveApp.getFoldersByName(namaFolder);
  if(folders.hasNext()){
    return folders.next();
  }
  return DriveApp.createFolder(namaFolder);
}

// ======================================
// Ambil Data
// ======================================
function getSuratKeluar(token){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  const sh=getSheet("SuratKeluar");
  if(!sh){
    throw new Error("Sheet SuratKeluar tidak ditemukan.");
  }
  const values=sh.getDataRange().getValues();
  return values.map(function(row,index){
    if(index==0) return row;
    return row.map(function(cell){
      if(Object.prototype.toString.call(cell)==="[object Date]"){
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return cell;
    });
  });
}

// ======================================
// Tambah
// ======================================
function tambahSuratKeluar(token,data){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const session = requireRole(token, ["administrator", "petugas pelayanan"]);
    validasiSuratKeluar_(data);

    const sh=getSheet("SuratKeluar");
    const values=sh.getDataRange().getValues();
    let maxId=0;

    for(let i=1;i<values.length;i++){
      if(values[i][2] &&
         values[i][2].toString().trim().toLowerCase()==
         data.nomorSurat.toString().trim().toLowerCase()){
        throw new Error("Nomor Surat sudah terdaftar.");
      }
      const id=Number(values[i][0]);
      if(id>maxId){
        maxId=id;
      }
    }

    const newId=maxId+1;
    const noAgenda="SK-"+Utilities.formatString("%04d",newId);

    sh.getRange(sh.getLastRow()+1, 1, 1, 10).setValues([[
      newId,
      noAgenda,
      data.nomorSurat.trim(),
      parseDateInput_(data.tanggalSurat),
      data.tujuan.trim(),
      data.perihal.trim(),
      data.penandatangan.trim(),
      data.keterangan || "",
      data.status || "Draft",
      data.fileURL || ""
    ]]);
    return true;

  }catch(err){
    Logger.log(err);
    throw err;
  }finally{
    lock.releaseLock();
  }
}

// ======================================
// Edit
// ======================================
function editSuratKeluar(token,id,data){
  const lock=LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    const session = requireRole(token, ["administrator", "petugas pelayanan"]);
    validasiSuratKeluar_(data);

    const sh=getSheet("SuratKeluar");
    const values=sh.getDataRange().getValues();
    let row=-1;

    for(let i=1;i<values.length;i++){
      if(values[i][0].toString()==id.toString()){
        row=i+1;
        continue;
      }
      if(values[i][2] &&
         values[i][2].toString().trim().toLowerCase()==
         data.nomorSurat.toString().trim().toLowerCase()){
        throw new Error("Nomor Surat sudah dipakai.");
      }
    }

    if(row==-1){
      throw new Error("Data tidak ditemukan.");
    }

    const fileURLLama = values[row-1][9];
    const fileURLBaru = data.fileURL || fileURLLama || "";

    sh.getRange(row, 3, 1, 8).setValues([[
      data.nomorSurat.trim(),
      parseDateInput_(data.tanggalSurat),
      data.tujuan.trim(),
      data.perihal.trim(),
      data.penandatangan.trim(),
      data.keterangan || "",
      data.status || "Draft",
      fileURLBaru
    ]]);
    return true;

  }catch(err){
    Logger.log(err);
    throw err;
  }finally{
    lock.releaseLock();
  }
}

// ======================================
// Hapus
// ======================================
function hapusSuratKeluar(token,id){
  const session = requireRole(token, ["administrator", "petugas pelayanan"]);
  const sh=getSheet("SuratKeluar");
  const values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(values[i][0].toString()==id.toString()){
      sh.deleteRow(i+1);
      return true;
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// ======================================
// Halaman HTML
// PENTING: nama file persis "Suratkeluar" (k kecil), sesuai
// nama file asli di Apps Script -- BUKAN "SuratKeluar" dari GitHub.
// ======================================
function getSuratKeluarPage(){
  return HtmlService
    .createHtmlOutputFromFile("SuratKeluar")
    .getContent();
}