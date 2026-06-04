/**
 * NutriPlan MBG Enterprise - GAS Backend (v4.3.0)
 * Bound Spreadsheet Database Handler with Role-Based Authentication
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0')
      .setTitle('NutriPlan MBG Enterprise')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Memastikan lembar kerja sheet "Users" dan "DataInput" siap digunakan.
 * Jika belum ada, maka akan dibuat otomatis beserta data bawaan (fallback).
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Inisialisasi Sheet Users
  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
    usersSheet.appendRow(["username", "password", "role"]);
    // Data default untuk pengujian awal
    usersSheet.appendRow(["Fajri", "123", "administrator"]);
    usersSheet.appendRow(["Staff", "456", "user"]);
    
    // Rapikan header
    usersSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#10b981").setFontColor("#ffffff");
  }
  
  // 2. Inisialisasi Sheet DataInput
  let dataInputSheet = ss.getSheetByName("DataInput");
  if (!dataInputSheet) {
    dataInputSheet = ss.insertSheet("DataInput");
    dataInputSheet.appendRow(["Timestamp", "NamaUser", "AppStateJSON"]);
    
    // Rapikan header
    dataInputSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#0ea5e9").setFontColor("#ffffff");
  }
}

/**
 * Memvalidasi kredensial pengguna berdasarkan data di Google Sheets.
 * @param {string} username - Nama pengguna
 * @param {string} password - Sandi pengguna
 * @return {object} Status login beserta role pengguna jika sukses
 */
function checkLogin(username, password) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    
    // Mulai membaca baris data (lewati baris header)
    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][0]).trim();
      const dbPass = String(data[i][1]).trim();
      const dbRole = String(data[i][2] || 'user').trim().toLowerCase();
      
      if (dbUser.toLowerCase() === username.toLowerCase().trim() && dbPass === String(password).trim()) {
        return {
          success: true,
          username: dbUser,
          role: dbRole
        };
      }
    }
    
    return { success: false, message: "Kombinasi nama pengguna atau kata sandi tidak valid!" };
  } catch (error) {
    return { success: false, message: "Terjadi gangguan sistem: " + error.toString() };
  }
}

/**
 * Menyimpan data menu kerja (AppState) pengguna langsung ke Google Sheets (DataInput)
 * @param {string} username - Pengguna aktif
 * @param {string} stateJSON - Stringified state data
 */
function saveUserDataToSheet(username, stateJSON) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DataInput");
    const data = sheet.getDataRange().getValues();
    const timestamp = new Date();
    
    let userRowIndex = -1;
    // Cari baris data milik user jika sudah ada sebelumnya
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === username.toLowerCase()) {
        userRowIndex = i + 1; // Baris riil di sheet (offset 1-based index)
        break;
      }
    }
    
    if (userRowIndex !== -1) {
      // Overwrite / Update data yang sudah ada
      sheet.getRange(userRowIndex, 1).setValue(timestamp);
      sheet.getRange(userRowIndex, 3).setValue(stateJSON);
    } else {
      // Buat baris baru untuk user baru
      sheet.appendRow([timestamp, username, stateJSON]);
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, message: "Gagal mencadangkan data ke Cloud Sheets: " + error.toString() };
  }
}

/**
 * Memuat data draf menu khusus milik pengguna yang sedang masuk dari Google Sheets
 * @param {string} username - Pengguna aktif
 * @return {object} Data draf kerja AppState lama jika ditemukan
 */
function loadUserDataFromSheet(username) {
  try {
    initializeSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("DataInput");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === username.toLowerCase()) {
        return {
          success: true,
          found: true,
          stateJSON: data[i][2]
        };
      }
    }
    
    return { success: true, found: false };
  } catch (error) {
    return { success: false, message: "Gagal memuat data draf dari Cloud: " + error.toString() };
  }
}
