// ===============================
// PENDUDUK.GS
// ===============================

// ===============================
// Ambil Data Penduduk
// ===============================
function getPenduduk(token) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa. Silakan login ulang.");
  }
  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");
  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
  }

  const data = sh.getDataRange().getValues();

  // Ubah semua objek Date jadi teks biasa sebelum dikirim ke browser,
  // supaya tidak gagal terkirim (jadi null) gara-gara kolom Tanggal Lahir.
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
// Validasi data Penduduk (dipakai bersama oleh Tambah & Edit)
// ===============================
function validasiDataPenduduk(data) {
  if (!data.nik || data.nik.toString().trim() == "") {
    throw new Error("NIK tidak boleh kosong.");
  }
  if (!data.nama || data.nama.toString().trim() == "") {
    throw new Error("Nama tidak boleh kosong.");
  }
  const nikBersih = data.nik.toString().trim();
  if (!/^\d{16}$/.test(nikBersih)) {
    throw new Error("NIK harus terdiri dari 16 digit angka.");
  }
  return nikBersih;
}

// ===============================
// Tambah Penduduk
// ===============================
function tambahPenduduk(token, data) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");
  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
  }

  const nikBersih = validasiDataPenduduk(data);

  const values = sh.getDataRange().getValues();
  let maxId = 0;

  for (let i = 1; i < values.length; i++) {
    if (values[i][1] && values[i][1].toString().trim() === nikBersih) {
      throw new Error("NIK sudah terdaftar.");
    }
    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) {
      maxId = id;
    }
  }

  const newId = maxId + 1;

  const rowData = [
    newId,
    "'" + nikBersih,
    data.nama.toString().trim(),
    data.jk || '',
    data.tempat || '',
    data.tgl || '',
    data.alamat || '',
    data.rt || '',
    data.rw || '',
    "'" + (data.hp ? data.hp.toString().trim() : '')
  ];

  sh.appendRow(rowData);
  return true;
}

// ===============================
// Edit Penduduk (BARU)
// ===============================
function editPenduduk(token, id, data) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");
  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
  }

  const nikBersih = validasiDataPenduduk(data);
  const values = sh.getDataRange().getValues();

  // Cari baris dengan ID yang cocok
  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][0].toString() === id.toString()) {
      targetRow = i + 1; // +1 karena spreadsheet mulai baris 1, index array mulai 0
      continue;
    }
    // Cek NIK tidak bentrok dengan penduduk LAIN (selain baris yang sedang diedit)
    if (values[i][0].toString() !== id.toString() &&
        values[i][1] && values[i][1].toString().trim() === nikBersih) {
      throw new Error("NIK sudah dipakai penduduk lain.");
    }
  }

  if (targetRow === -1) {
    throw new Error("Data dengan ID tersebut tidak ditemukan.");
  }

  const rowData = [
    id,
    "'" + nikBersih,
    data.nama.toString().trim(),
    data.jk || '',
    data.tempat || '',
    data.tgl || '',
    data.alamat || '',
    data.rt || '',
    data.rw || '',
    "'" + (data.hp ? data.hp.toString().trim() : '')
  ];

  sh.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  return true;
}

// ===============================
// Hapus Penduduk
// ===============================
function hapusPenduduk(token, id) {
  const session = requireRole(token, ["administrator", "sekretaris kaur"]);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");
  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
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