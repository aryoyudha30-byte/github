function getDashboardStats() {
  return {
    penduduk: getJumlahData("Penduduk"),
    kk: getJumlahData("KK"),
    suratMasuk: getJumlahData("SuratMasuk"),
    suratKeluar: getJumlahData("SuratKeluar"),
    pengaduan: getJumlahData("Pengaduan"),
    agenda: getJumlahData("Agenda")
  };
}

function getJumlahData(namaSheet) {
  const sheet = getSheet(namaSheet);
  if (sheet == null) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  return lastRow - 1;
}