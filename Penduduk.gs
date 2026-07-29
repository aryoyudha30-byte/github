function getPenduduk() {

  const ss = getDatabase();
  const sh = ss.getSheetByName("Penduduk");

  if (!sh) {
    throw new Error("Sheet Penduduk tidak ditemukan.");
  }

  return sh.getDataRange().getValues();

}
function tambahPenduduk(data) {

  const ss = getDatabase();

  const sh = ss.getSheetByName("Penduduk");

  sh.appendRow(data);

  return true;

}