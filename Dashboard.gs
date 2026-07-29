function getDashboardStats() {
  try {
    const totalPenduduk = getSheetDataAsJson('Penduduk').length;
    const totalKK = getSheetDataAsJson('KK').length;
    const totalSuratMasuk = getSheetDataAsJson('SuratMasuk').length;
    const totalSuratKeluar = getSheetDataAsJson('SuratKeluar').length;
    const totalPengaduan = getSheetDataAsJson('Pengaduan').length;

    return {
      success: true,
      stats: {
        penduduk: totalPenduduk,
        kk: totalKK,
        suratMasuk: totalSuratMasuk,
        suratKeluar: totalSuratKeluar,
        pengaduan: totalPengaduan
      }
    };
  } catch(e) {
    return {
      success: false,
      message: e.message
    };
  }
}