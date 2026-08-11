// ===============================
// KK.GS
// Struktur kolom sheet "KK":
// ID | NoKK | KepalaKeluarga | Alamat | RT | RW | JumlahAnggota
// ===============================

// ===============================
// Ambil Data KK
// ===============================
function getKK(token) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa. Silakan login ulang.");
  }
  const sh = getSheet("KK");
  if (!sh) {
    throw new Error("Sheet KK tidak ditemukan.");
  }

  const data = sh.getDataRange().getValues();

  // Ubah semua objek Date jadi teks biasa (jaga-jaga kalau ada kolom tanggal ke depannya)
  const safeData = data.map(row =>
    row.map(cell => {
      if (Object.prototype.toString.call(cell) === "[object Date]") {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return cell;
    })
  );

  return safeData;
}

// ===============================
// Validasi data KK (dipakai bersama Tambah & Edit)
// ===============================
function validasiDataKK(data) {
  if (!data.noKK || data.noKK.toString().trim() == "") {
    throw new Error("Nomor KK tidak boleh kosong.");
  }
  if (!data.kepalaKeluarga || data.kepalaKeluarga.toString().trim() == "") {
    throw new Error("Nama Kepala Keluarga tidak boleh kosong.");
  }
  const noKKBersih = data.noKK.toString().trim();
  if (!/^\d{16}$/.test(noKKBersih)) {
    throw new Error("Nomor KK harus terdiri dari 16 digit angka.");
  }
  return noKKBersih;
}

// ===============================
// Tambah KK
// ===============================
function tambahKK(token, data) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const sh = getSheet("KK");
  if (!sh) {
    throw new Error("Sheet KK tidak ditemukan.");
  }

  const noKKBersih = validasiDataKK(data);

  const values = sh.getDataRange().getValues();
  let maxId = 0;

  for (let i = 1; i < values.length; i++) {
    if (values[i][1] && values[i][1].toString().trim() === noKKBersih) {
      throw new Error("Nomor KK sudah terdaftar.");
    }
    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) {
      maxId = id;
    }
  }

  const newId = maxId + 1;

  const rowData = [
    newId,
    "'" + noKKBersih,
    data.kepalaKeluarga.toString().trim(),
    data.alamat || '',
    data.rt || '',
    data.rw || '',
    data.jumlahAnggota || 0
  ];

  sh.appendRow(rowData);
  return true;
}

// ===============================
// Edit KK
// ===============================
function editKK(token, id, data) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const sh = getSheet("KK");
  if (!sh) {
    throw new Error("Sheet KK tidak ditemukan.");
  }

  const noKKBersih = validasiDataKK(data);
  const values = sh.getDataRange().getValues();

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id.toString()) {
      targetRow = i + 1;
      continue;
    }
    if (values[i][1] && values[i][1].toString().trim() === noKKBersih) {
      throw new Error("Nomor KK sudah dipakai KK lain.");
    }
  }

  if (targetRow === -1) {
    throw new Error("Data dengan ID tersebut tidak ditemukan.");
  }

  const rowData = [
    id,
    "'" + noKKBersih,
    data.kepalaKeluarga.toString().trim(),
    data.alamat || '',
    data.rt || '',
    data.rw || '',
    data.jumlahAnggota || 0
  ];

  sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  return true;
}

// ===============================
// Hapus KK
// ===============================
function hapusKK(token, id) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const sh = getSheet("KK");
  if (!sh) {
    throw new Error("Sheet KK tidak ditemukan.");
  }
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id.toString()) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error("Data tidak ditemukan.");
}

// ===============================
// Halaman KK (dipanggil dari menu sidebar)
// ===============================
function getKKPage() {
  return HtmlService.createHtmlOutputFromFile("kk").getContent();
}