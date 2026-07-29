function checkLogin(username, password) {
  try {
    const users = getSheetDataAsJson('Users');
    const inputHash = hashPassword(password); // Enkripsi password inputan
    
    // Cari user yang cocok username dan hash passwordnya
    const foundUser = users.find(u => 
      u.Username.toString().trim() === username.toString().trim() && 
      (u.Password.toString().trim() === inputHash || u.Password.toString().trim() === password)
    );

    if (foundUser) {
      // Catat log aktivitas jika diperlukan
      logActivity(foundUser.Username, 'LOGIN', 'User berhasil login ke sistem');

      return {
        success: true,
        message: 'Login Berhasil!',
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

// Log Aktivitas Akses User untuk Audit Trail Keamanan
function logActivity(username, aksi, detail) {
  try {
    const sheet = getSheet('LogAktivitas');
    if (sheet) {
      sheet.appendRow([new Date(), username, aksi, detail]);
    }
  } catch(e) {
  
  }
}