// ===========================================================
// SURATKELUAR.GS
// Modul Pengelolaan Surat Keluar
//
// Struktur Sheet "SuratKeluar"
// -----------------------------------------------------------
// A : ID
// B : No Agenda
// C : Nomor Surat
// D : Tanggal Surat
// E : Tujuan
// F : Perihal
// G : Penandatangan
// H : Keterangan
// I : Status
// ===========================================================



// ===========================================================
// HALAMAN HTML
// ===========================================================

function getSuratKeluarPage() {
  return HtmlService
    .createHtmlOutputFromFile("Suratkeluar")
    .getContent();
}



// ===========================================================
// HELPER
// ===========================================================

function parseDateInput_(value){

  if(!value) return "";

  if(Object.prototype.toString.call(value)==="[object Date]"){
    return value;
  }

  const p=value.toString().split("-");

  if(p.length===3){
    return new Date(
      Number(p[0]),
      Number(p[1])-1,
      Number(p[2])
    );
  }

  return value;

}



// ===========================================================
// Generate Nomor Agenda
// Contoh:
// SK-0001
// SK-0002
// ===========================================================

function generateNoAgenda_(id){

  return "SK-" + Utilities.formatString("%04d",id);

}



// ===========================================================
// Validasi Data
// ===========================================================

function validasiSuratKeluar_(data){

  if(!data.nomorSurat || data.nomorSurat.toString().trim()==""){
    throw new Error("Nomor Surat wajib diisi.");
  }

  if(data.nomorSurat.length>100){
    throw new Error("Nomor Surat maksimal 100 karakter.");
  }

  if(!data.tujuan || data.tujuan.toString().trim()==""){
    throw new Error("Tujuan wajib diisi.");
  }

  if(data.tujuan.length>200){
    throw new Error("Tujuan maksimal 200 karakter.");
  }

  if(!data.perihal || data.perihal.toString().trim()==""){
    throw new Error("Perihal wajib diisi.");
  }

  if(data.perihal.length>300){
    throw new Error("Perihal maksimal 300 karakter.");
  }

  if(!data.penandatangan || data.penandatangan.toString().trim()==""){
    throw new Error("Penandatangan wajib diisi.");
  }

  const statusValid=[
    "Draft",
    "Diproses",
    "Dikirim",
    "Diterima",
    "Diarsipkan"
  ];

  if(data.status && !statusValid.includes(data.status)){
    throw new Error("Status tidak valid.");
  }

}



// ===========================================================
// Ambil Semua Surat Keluar
// ===========================================================

function getSuratKeluar(token){

  const session=verifySession(token);

  if(!session){
    throw new Error("Sesi login tidak valid.");
  }

  const sheet=getSheet("SuratKeluar");

  if(!sheet){
    throw new Error("Sheet SuratKeluar tidak ditemukan.");
  }

  const data=sheet.getDataRange().getValues();

  const safeData=data.map(row=>

    row.map(cell=>{

      if(Object.prototype.toString.call(cell)==="[object Date]"){

        return Utilities.formatDate(
          cell,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );

      }

      return cell;

    })

  );

  return safeData;

}



// ===========================================================
// Tambah Surat Keluar
// ===========================================================

function tambahSuratKeluar(token,data){

  const session=verifySession(token);

  if(!session){
    throw new Error("Sesi login tidak valid.");
  }

  const lock=LockService.getScriptLock();

  if(!lock.tryLock(10000)){
    throw new Error("Sistem sedang sibuk. Silakan coba kembali.");
  }

  try{

    const sheet=getSheet("SuratKeluar");

    if(!sheet){
      throw new Error("Sheet SuratKeluar tidak ditemukan.");
    }

    validasiSuratKeluar_(data);

    const values=sheet.getDataRange().getValues();

    let maxId=0;

    const nomorSurat=data.nomorSurat
      .toString()
      .trim()
      .toLowerCase();

    for(let i=1;i<values.length;i++){

      const id=Number(values[i][0]);

      if(!isNaN(id) && id>maxId){
        maxId=id;
      }

      if(
        values[i][2] &&
        values[i][2]
          .toString()
          .trim()
          .toLowerCase()==nomorSurat
      ){

        throw new Error(
          "Nomor Surat sudah terdaftar."
        );

      }

    }

    const newId=maxId+1;

    const noAgenda=generateNoAgenda_(newId);

    sheet.appendRow([

      newId,

      noAgenda,

      data.nomorSurat.trim(),

      parseDateInput_(data.tanggalSurat),

      data.tujuan.trim(),

      data.perihal.trim(),

      data.penandatangan.trim(),

      data.keterangan
        ? data.keterangan.trim()
        : "",

      data.status || "Draft"

    ]);

    return true;

  }

  catch(err){

    Logger.log(err);

    throw err;

  }

  finally{

    lock.releaseLock();

  }

}

// ===========================================================
// Edit Surat Keluar
// ===========================================================

function editSuratKeluar(token, id, data){

  const session = verifySession(token);

  if(!session){
    throw new Error("Sesi login tidak valid.");
  }

  const lock = LockService.getScriptLock();

  if(!lock.tryLock(10000)){
    throw new Error("Sistem sedang sibuk. Silakan coba kembali.");
  }

  try{

    const sheet = getSheet("SuratKeluar");

    if(!sheet){
      throw new Error("Sheet SuratKeluar tidak ditemukan.");
    }

    validasiSuratKeluar_(data);

    const values = sheet.getDataRange().getValues();

    let targetRow = -1;

    const nomorSurat = data.nomorSurat
      .toString()
      .trim()
      .toLowerCase();

    for(let i=1;i<values.length;i++){

      if(values[i][0].toString() === id.toString()){
        targetRow = i + 1;
        continue;
      }

      if(
        values[i][2] &&
        values[i][2]
          .toString()
          .trim()
          .toLowerCase() === nomorSurat
      ){

        throw new Error("Nomor Surat sudah digunakan oleh surat lain.");

      }

    }

    if(targetRow==-1){
      throw new Error("Data Surat Keluar tidak ditemukan.");
    }

    sheet.getRange(targetRow,3,1,7).setValues([[

      data.nomorSurat.trim(),

      parseDateInput_(data.tanggalSurat),

      data.tujuan.trim(),

      data.perihal.trim(),

      data.penandatangan.trim(),

      data.keterangan
        ? data.keterangan.trim()
        : "",

      data.status || "Draft"

    ]]);

    return true;

  }

  catch(err){

    Logger.log(err);

    throw err;

  }

  finally{

    lock.releaseLock();

  }

}



// ===========================================================
// Hapus Surat Keluar
// ===========================================================

function hapusSuratKeluar(token,id){

  const session = verifySession(token);

  if(!session){
    throw new Error("Sesi login tidak valid.");
  }

  const lock = LockService.getScriptLock();

  if(!lock.tryLock(10000)){
    throw new Error("Sistem sedang sibuk. Silakan coba kembali.");
  }

  try{

    const sheet = getSheet("SuratKeluar");

    if(!sheet){
      throw new Error("Sheet SuratKeluar tidak ditemukan.");
    }

    const values = sheet.getDataRange().getValues();

    for(let i=1;i<values.length;i++){

      if(values[i][0].toString()==id.toString()){

        sheet.deleteRow(i+1);

        return true;

      }

    }

    throw new Error("Data Surat Keluar tidak ditemukan.");

  }

  catch(err){

    Logger.log(err);

    throw err;

  }

  finally{

    lock.releaseLock();

  }

}



// ===========================================================
// Ambil Surat Keluar Berdasarkan ID
// ===========================================================

function getSuratKeluarById(token,id){

  const session = verifySession(token);

  if(!session){
    throw new Error("Sesi login tidak valid.");
  }

  const sheet = getSheet("SuratKeluar");

  if(!sheet){
    throw new Error("Sheet SuratKeluar tidak ditemukan.");
  }

  const values = sheet.getDataRange().getValues();

  for(let i=1;i<values.length;i++){

    if(values[i][0].toString()==id.toString()){

      return{

        id:values[i][0],

        noAgenda:values[i][1],

        nomorSurat:values[i][2],

        tanggalSurat:
          values[i][3]
          ? Utilities.formatDate(
              new Date(values[i][3]),
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : "",

        tujuan:values[i][4],

        perihal:values[i][5],

        penandatangan:values[i][6],

        keterangan:values[i][7],

        status:values[i][8]

      };

    }

  }

  throw new Error("Data Surat Keluar tidak ditemukan.");

}