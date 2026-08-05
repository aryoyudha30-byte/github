function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Administrasi Desa')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Memanggil file HTML lain (dipakai oleh index.html dan dashboard.html
// lewat sintaks <?!= include('NamaFile'); ?>). Pakai createTemplateFromFile
// + evaluate() (bukan cuma createHtmlOutputFromFile) supaya file yang
// dipanggil BOLEH berisi <?!= include(...) ?> lagi di dalamnya -- misalnya
// nanti DashboardContent.html mau include ChartDashboard.html, dll.
function include(filename) {
  return HtmlService
    .createTemplateFromFile(filename)
    .evaluate()
    .getContent();
}

// PENTING: pakai createTemplateFromFile().evaluate(), BUKAN
// createHtmlOutputFromFile() -- supaya kode <?!= include(...) ?>
// di dalam dashboard.html benar-benar diproses, bukan tampil sebagai teks mentah.
function getDashboard() {
  return HtmlService
    .createTemplateFromFile("dashboard")
    .evaluate()
    .getContent();
}

function getPendudukPage() {
  return HtmlService.createHtmlOutputFromFile("penduduk").getContent();
}

// Fungsi login lama (hardcode admin/123456) SUDAH DIHAPUS.
// Login sepenuhnya ditangani checkLogin() di auth.gs.