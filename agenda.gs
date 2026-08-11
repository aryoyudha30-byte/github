// ============================================================
// AGENDA.GS
// MODUL AGENDA SISTEM ADMINISTRASI DESA
// ============================================================
//
// Struktur Sheet "Agenda":
//
// A = ID
// B = Tanggal
// C = Waktu
// D = Nama Agenda
// E = Tempat
// F = Penanggung Jawab
// G = Status
// H = Keterangan
//
// ============================================================


// ============================================================
// HALAMAN AGENDA
// ============================================================
function getAgendaPage() {
  return HtmlService.createHtmlOutputFromFile("Agenda").getContent();
}


// ============================================================
// MENGAMBIL DATA AGENDA
// ============================================================
function getAgenda(token) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Agenda");
  if (!sh) throw new Error("Sheet Agenda tidak ditemukan.");

  const data = sh.getDataRange().getValues();
  if (data.length === 0) return [];

  return data.map(function (row, index) {
    return row.map(function (cell) {
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        // Kolom tanggal (baris data, bukan header)
        if (index > 0) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
      }
      return cell;
    });
  });
}


// ============================================================
// TAMBAH AGENDA
// ============================================================
function tambahAgenda(token, data) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Agenda");
  if (!sh) throw new Error("Sheet Agenda tidak ditemukan.");

  if (!data.tanggal) throw new Error("Tanggal agenda wajib diisi.");
  if (!data.waktu) throw new Error("Waktu agenda wajib diisi.");
  if (!data.namaAgenda) throw new Error("Nama agenda wajib diisi.");

  // Cari ID terbesar
  const values = sh.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < values.length; i++) {
    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  const newId = maxId + 1;

  sh.appendRow([
    newId,
    data.tanggal,
    data.waktu,
    data.namaAgenda,
    data.tempat || "",
    data.penanggungJawab || "",
    data.status || "Terjadwal",
    data.keterangan || ""
  ]);

  return true;
}


// ============================================================
// EDIT AGENDA
// ============================================================
function editAgenda(token, id, data) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Agenda");
  if (!sh) throw new Error("Sheet Agenda tidak ditemukan.");

  if (!data.tanggal) throw new Error("Tanggal agenda wajib diisi.");
  if (!data.waktu) throw new Error("Waktu agenda wajib diisi.");
  if (!data.namaAgenda) throw new Error("Nama agenda wajib diisi.");

  const values = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      row = i + 1;
      break;
    }
  }
  if (row === -1) throw new Error("Data agenda tidak ditemukan.");

  sh.getRange(row, 2, 1, 7).setValues([[
    data.tanggal,
    data.waktu,
    data.namaAgenda,
    data.tempat || "",
    data.penanggungJawab || "",
    data.status || "Terjadwal",
    data.keterangan || ""
  ]]);

  return true;
}


// ============================================================
// HAPUS AGENDA
// ============================================================
function hapusAgenda(token, id) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Agenda");
  if (!sh) throw new Error("Sheet Agenda tidak ditemukan.");

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return true;
    }
  }

  throw new Error("Data agenda tidak ditemukan.");
}