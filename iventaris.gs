// ============================================================
// INVENTARIS.GS
// MODUL INVENTARIS DESA
// ============================================================
//
// Struktur Sheet "Inventaris":
//
// A ID
// B Kode Barang
// C Nama Barang
// D Kategori
// E Jumlah
// F Satuan
// G Kondisi
// H Lokasi
// I Tahun Perolehan
// J Sumber Dana
// K Keterangan
//
// ============================================================


function getInventarisPage() {
  return HtmlService.createHtmlOutputFromFile("Inventaris").getContent();
}


// ============================================================
// GET DATA
// ============================================================
function getInventaris(token) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Inventaris");
  if (!sh) throw new Error("Sheet Inventaris tidak ditemukan.");

  const data = sh.getDataRange().getValues();

  return data.map(function (row) {
    return row.map(function (cell) {
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return cell;
    });
  });
}


// ============================================================
// TAMBAH
// ============================================================
function tambahInventaris(token, data) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Inventaris");
  if (!sh) throw new Error("Sheet Inventaris tidak ditemukan.");

  if (!data.namaBarang) throw new Error("Nama barang wajib diisi.");
  if (data.jumlah === "" || data.jumlah === null || data.jumlah === undefined) {
    throw new Error("Jumlah barang wajib diisi.");
  }

  const values = sh.getDataRange().getValues();
  let maxId = 0;
  for (let i = 1; i < values.length; i++) {
    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) maxId = id;
  }
  const newId = maxId + 1;
  const kodeBarang = "INV-" + Utilities.formatString("%04d", newId);

  sh.appendRow([
    newId,
    kodeBarang,
    data.namaBarang,
    data.kategori || "",
    Number(data.jumlah),
    data.satuan || "Unit",
    data.kondisi || "Baik",
    data.lokasi || "",
    data.tahunPerolehan || "",
    data.sumberDana || "",
    data.keterangan || ""
  ]);

  return true;
}


// ============================================================
// EDIT
// ============================================================
function editInventaris(token, id, data) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Inventaris");
  if (!sh) throw new Error("Sheet Inventaris tidak ditemukan.");

  if (!data.namaBarang) throw new Error("Nama barang wajib diisi.");

  const values = sh.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      row = i + 1;
      break;
    }
  }
  if (row === -1) throw new Error("Barang tidak ditemukan.");

  // Kolom C sampai K (9 kolom, dimulai dari kolom ke-3)
  sh.getRange(row, 3, 1, 9).setValues([[
    data.namaBarang,
    data.kategori || "",
    Number(data.jumlah),
    data.satuan || "Unit",
    data.kondisi || "Baik",
    data.lokasi || "",
    data.tahunPerolehan || "",
    data.sumberDana || "",
    data.keterangan || ""
  ]]);

  return true;
}


// ============================================================
// HAPUS
// ============================================================
function hapusInventaris(token, id) {
  const session = verifySession(token);
  if (!session) throw new Error("Sesi login tidak valid.");

  const sh = getSheet("Inventaris");
  if (!sh) throw new Error("Sheet Inventaris tidak ditemukan.");

  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 1);
      return true;
    }
  }

  throw new Error("Barang tidak ditemukan.");
}