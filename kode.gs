// ==============================================================================
// SERVER-SIDE LOGIC: NutriPlan MBG Enterprise (Google Apps Script)
// ==============================================================================

/**
 * Fungsi wajib GAS untuk menyajikan file HTML sebagai Web App.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('NutriPlan MBG Enterprise')
      .setFaviconUrl('https://img.icons8.com/color/48/000000/leaf.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi untuk memverifikasi Login pengguna dari Google Sheets.
 * @param {string} username - Input username dari antarmuka
 * @param {string} password - Input password dari antarmuka
 * @returns {object} Status sukses, data user, dan pesan.
 */
function verifyUserLogin(username, password) {
  try {
    // Membuka spreadsheet aktif (Pastikan script ini terikat/bound dengan Spreadsheet Anda)
    // Atau gunakan SpreadsheetApp.openById('ID_SPREADSHEET_ANDA') jika script berdiri sendiri (standalone)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetUsers = ss.getSheetByName("Users"); 
    
    if (!sheetUsers) {
      return { success: false, message: "Error: Sheet 'Users' tidak ditemukan di database!" };
    }
    
    // Mengambil semua data (Array 2D)
    var data = sheetUsers.getDataRange().getValues();
    
    // Looping mulai dari baris 1 (mengabaikan baris 0 yang merupakan Header)
    for (var i = 1; i < data.length; i++) {
      var dbUser = data[i][0].toString().trim();
      var dbPass = data[i][1].toString().trim();
      var dbRole = data[i][2] ? data[i][2].toString().trim() : "user";
      
      // Jika cocok
      if (dbUser === username.trim() && dbPass === password.trim()) {
        return {
          success: true,
          user: {
            username: dbUser,
            role: dbRole
          },
          message: "Login berhasil diverifikasi."
        };
      }
    }
    
    // Jika loop selesai dan tidak ada yang cocok
    return { success: false, message: "Username atau Password salah!" };
    
  } catch (error) {
    return { success: false, message: "Gagal terhubung ke Database: " + error.message };
  }
}
