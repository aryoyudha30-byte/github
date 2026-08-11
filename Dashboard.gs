// ======================================
// HALAMAN DASHBOARD (CONTENT ONLY)
// Dipanggil saat klik ulang menu Dashboard di sidebar.
// PENTING: harus "DashboardContent" (cuma kartu+aktivitas),
// BUKAN "Dashboard" (itu layout lengkap dengan sidebar,
// kalau salah ini bikin sidebar dobel).
// ======================================
function getDashboardPage() {
  return HtmlService
    .createHtmlOutputFromFile("DashboardContent")
    .getContent();
}

// ======================================
// STATISTIK DASHBOARD
// ======================================
function getDashboardStats(token) {
  requireRole(token, ["administrator", "petugas pelayanan", "sekretaris kaur"]);

  return {
    penduduk: getJumlahData("Penduduk"),
    kk: getJumlahData("KK"),
    suratMasuk: getJumlahData("SuratMasuk"),
    suratKeluar: getJumlahData("SuratKeluar"),
    pelayananSurat: getJumlahData("PelayananSurat"),
    pengaduan: getJumlahDataForm("Pengaduan"),
    agenda: getJumlahData("Agenda"),
    inventaris: getJumlahData("Inventaris")
  };
}

function getJumlahData(namaSheet) {
  const sheet = getSheet(namaSheet);
  if (sheet == null) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  return lastRow - 1;
}

// Sama seperti getJumlahData(), tapi membaca dari spreadsheet FORM
// (SPREADSHEET_ID_FORM), tempat data Pengaduan sebenarnya tersimpan.
function getJumlahDataForm(namaSheet) {
  const sheet = getFormSheet(namaSheet);
  if (sheet == null) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  return lastRow - 1;
}