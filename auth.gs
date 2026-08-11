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
    // Silently fail agar tidak mengganggu alur login
  }
}

// PENTING: fungsi ini SEKARANG BERDIRI SENDIRI (sejajar dengan logActivity),
// bukan lagi bersarang di dalamnya -- supaya bisa dipanggil dari google.script.run.
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

// ============================================================
// LOGOUT / INVALIDASI SESSION (SERVER-SIDE)
// ============================================================
function logoutSession(token) {
  if (!token) return true;

  try {
    const sessionData = CacheService.getScriptCache().get(token);

    if (sessionData) {
      const session = JSON.parse(sessionData);
      logActivity(session.username, "LOGOUT", "User keluar dari sistem");
    }

    // Cabut token dari server -- setelah ini token TIDAK BISA dipakai lagi,
    // walaupun browser masih menyimpannya di localStorage.
    CacheService.getScriptCache().remove(token);

    return true;
  } catch (error) {
    // Tetap anggap logout berhasil di sisi client walau ada error server
    return true;
  }
}

// ============================================================
// HAK AKSES BERDASARKAN ROLE (dipakai mulai Tahap 3, opsional dipakai sekarang)
// Contoh pemakaian di modul lain nanti:
//   const session = requireRole(token, ["administrator"]);
// ============================================================
function requireRole(token, allowedRoles) {
  const session = verifySession(token);
  if (!session) {
    throw new Error("Sesi login tidak valid. Silakan login kembali.");
  }

  const userRole = String(session.role || "").trim().toLowerCase();
  const allowed = allowedRoles.some(r => String(r).trim().toLowerCase() === userRole);

  if (!allowed) {
    logActivity(session.username || "-", "ACCESS_DENIED", "Akses ditolak. Role: " + (session.role || "-"));
    throw new Error("Anda tidak memiliki hak akses untuk melakukan tindakan ini.");
  }

  return session;
}