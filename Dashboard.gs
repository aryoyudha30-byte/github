function getDashboardData() {
  return {
    penduduk: countRows("Penduduk"),
    kk: countRows("KK"),
    suratMasuk: countRows("SuratMasuk"),
    suratKeluar: countRows("SuratKeluar"),
    pengaduan: countRows("Pengaduan"),
    agenda: countRows("Agenda")
  };
}

function countRows(sheetName) {
  const sh = getSheet(sheetName);
  if (!sh) return 0;

  const lastRow = sh.getLastRow();
  return lastRow <= 1 ? 0 : lastRow - 1;
}
function getDashboardData() {

  return {
    penduduk: getJumlahData("Penduduk"),
    kk: getJumlahData("KK"),
    suratMasuk: getJumlahData("SuratMasuk"),
    suratKeluar: getJumlahData("SuratKeluar"),
    pengaduan: getJumlahData("Pengaduan"),
    agenda: getJumlahData("Agenda")
  };

}

function getJumlahData(namaSheet){

  const sheet = getSheet(namaSheet);

  if(sheet == null) return 0;

  const lastRow = sheet.getLastRow();

  if(lastRow <= 1) return 0;

  return lastRow - 1;

}