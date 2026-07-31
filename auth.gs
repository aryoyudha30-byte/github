function checkLogin(username, password) {
  try {
    const users = getSheetDataAsJson('Users');
    const inputHash = hashPassword(password); // Enkripsi password inputan
 
    // Cari user yang cocok username DAN hash password-nya (WAJIB hash, tidak ada pengecualian polos)
    const foundUser = users.find(u =>
      u.Username.toString().trim() === username.toString().trim() &&
      u.Password.toString().trim() === inputHash
    );
 
    if (foundUser) {
      // Buat token sesi acak, simpan di cache server selama 6 jam
      const token = Utilities.getUuid();
      CacheService.getScriptCache().put(token, JSON.stringify({
        username: foundUser.Username,
        nama: foundUser.Nama,
        role: foundUser.Role
      }), 21600); // 21600 detik = 6 jam
 
      logActivity(foundUser.Username, 'LOGIN', 'User berhasil login ke sistem');
 
      return {
        success: true,
        message: 'Login Berhasil!',
        token: token,
        user: {
          username: foundUser.Username,
          nama: foundUser.Nama,
          role: foundUser.Role
        }
      };
    } else {
      return {
        success: false,
        message: 'Username atau Password salah!'
      };
    }
  } catch (error) {
    return {
      success: false,
      message: 'Gagal menghubungkan database: ' + error.message
    };
  }
}
 
/**
 * Memeriksa apakah token sesi masih valid. Dipanggil oleh SEMUA fungsi
 * yang membaca/mengubah data (Penduduk, KK, Surat, dll) sebelum
 * mengizinkan aksi apapun. Mengembalikan null kalau sesi tidak valid/kadaluarsa.
 */
function verifySession(token) {
  if (!token) return null;
  const data = CacheService.getScriptCache().get(token);
  return data ? JSON.parse(data) : null;
}
 
// Log Aktivitas Akses User untuk Audit Trail Keamanan
function logActivity(username, aksi, detail) {
  try {
    const sheet = getSheet('LogAktivitas');
    if (sheet) {
      sheet.appendRow([new Date(), username, aksi, detail]);
    }
  } catch(e) {
  
  }
  function getAktivitasTerbaru(token) {
  const session = verifySession(token);
  if (!session) {
    throw new Error("Sesi login tidak valid. Silakan login ulang.");
  }

  const sheet = getSheet('LogAktivitas');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // Ambil 5 aktivitas terbaru saja, urutan terbalik (paling baru di atas)
  const rows = data.slice(1).reverse().slice(0, 5);

  return rows.map(r => ({
    waktu: r[0],
    username: r[1],
    aksi: r[2],
    detail: r[3]
  }));
}

}