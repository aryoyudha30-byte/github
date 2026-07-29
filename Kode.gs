function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistem Administrasi Desa')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Memanggil file HTML lain
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDashboard() {
  return HtmlService
    .createHtmlOutputFromFile("dashboard")
    .getContent();
}

function getPendudukPage() {
  return HtmlService.createHtmlOutputFromFile("penduduk").getContent();
}

// Login sederhana (sementara)
function login(username, password) {

  if (username == "admin" && password == "123456") {
    return true;
  }

  return false;
}
