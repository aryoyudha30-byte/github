// ===============================
// PENDUDUK.GS
// ===============================

// ===============================
// Ambil Data Penduduk
// ===============================
function getPenduduk(token) {
  const session = verifySession(token);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa. Silakan login ulang.");
  }
  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");
  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
  }
  return sh.getDataRange().getValues();
}

// ===============================
// Tambah Penduduk
// ===============================
function tambahPenduduk(token, data) {
  const session = verifySession(token);
  if (!session) {
    throw new Error("Sesi login tidak valid atau sudah kadaluarsa.");
  }
  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");
  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
  }

  // ==========================
  // VALIDASI DATA
  // ==========================
  if (!data.nik || data.nik.toString().trim() == "") {
    throw new Error("NIK tidak boleh kosong.");
  }
  if (!data.nama || data.nama.toString().trim() == "") {
    throw new Error("Nama tidak boleh kosong.");
  }

  // NIK Indonesia standarnya 16 digit angka
  const nikBersih = data.nik.toString().trim();
  if (!/^\d{16}$/.test(nikBersih)) {
    throw new Error("NIK harus terdiri dari 16 digit angka.");
  }

  // ==========================
  // BACA DATA SEKALI SAJA (untuk cek duplikat NIK & cari ID terbesar)
  // ==========================
  const values = sh.getDataRange().getValues();
  let maxId = 0;

  for (let i = 1; i < values.length; i++) {
    // Cek NIK duplikat
    if (values[i][1] && values[i][1].toString().trim() === nikBersih) {
      throw new Error("NIK sudah terdaftar.");
    }

    const id = Number(values[i][0]);
    if (!isNaN(id) && id > maxId) {
      maxId = id;
    }
  }

  const newId = maxId + 1;

  // ==========================
  // DATA YANG DISIMPAN
  // Urutan HARUS PERSIS sama dengan kolom sheet:
  // ID, NIK, NAME, JENIS KELAMIN, TEMPAT LAHIR, TANGGAL LAHIR, ALAMAT, RT, RW, NoHP
  // ==========================
  const rowData = [
    newId,
    "'" + nikBersih,                       // Ditambah ' agar NIK tersimpan sebagai Text murni
    data.nama.toString().trim(),
    data.jk || '',
    data.tempat || '',
    data.tgl || '',
    data.alamat || '',
    data.rt || '',
    data.rw || '',
    "'" + (data.hp ? data.hp.toString().trim() : '') // Ditambah ' agar No HP yang berawalan 0 tidak hilang
  ];

  sh.appendRow(rowData);
  return true;
}

// ===============================
// Hapus Penduduk
// ===============================
function hapusPenduduk(token, id) {
  const session = verifySession(token);
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