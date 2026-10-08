/**
 * Sirupat - Fasilitasi Penggunaan Ruang Rapat Kantor DPD RI KALTIM
 * Backend Logic (Google Apps Script)
 */

const SPREADSHEET_ID = ""; // Kosongkan jika script terikat pada sheet
const SHEET_NAME = "Profil";
const SHEET_FASILITASI = "Fasilitasi";
const SHEET_OPTION = "Option";
const DRIVE_FOLDER_ID = "15PKP3NXptgolD688GUq0SVkOpINkGvmx"; // Pastikan ID Folder ini benar dan ANDA (pemilik script) memiliki akses edit

/**
 * --- PENTING: JALANKAN FUNGSI INI DULU! ---
 * Fungsi ini ada untuk memancing permintaan izin akses Google Drive.
 * 1. Pilih 'checkPermissions' di dropdown atas.
 * 2. Klik 'Run'.
 * 3. Berikan izin (Allow).
 */
function checkPermissions() {
  console.log("Memulai pemeriksaan izin...");
  try {
    // Memanggil Root Folder untuk memicu scope Drive
    const root = DriveApp.getRootFolder();
    console.log("Izin Drive Dasar: OK. Root folder: " + root.getName());
    
    // Memanggil Folder Tujuan
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    console.log("Akses ke Folder Tujuan (" + DRIVE_FOLDER_ID + "): BERHASIL.");
    console.log("Nama Folder: " + folder.getName());
    console.log("SILAKAN LAKUKAN DEPLOYMENT BARU SEKARANG.");
  } catch (e) {
    console.error("GAGAL: " + e.toString());
    console.log("Pastikan ID Folder benar dan Anda memiliki akses edit ke folder tersebut.");
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sirupat - DPD RI KALTIM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet(name) {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAME) {
      sheet.appendRow(["ID", "Nama", "Email", "HP", "Alamat", "Instansi", "Role", "Password"]);
    } else if (name === SHEET_FASILITASI) {
      // --- PERBAIKAN 1: Tambahkan header "Link PDF" di Kolom 22 (Indeks ke-22 setelah Timestamp) ---
      sheet.appendRow([
        "ID", "Status", "Tanggal", "Mulai", "Selesai", "Peserta", "Tempat", "Agenda",
        "Pengguna", "Instansi", "Rekomendasi", "Penanggungjawab", "Email", "Alamat", "Kontak",
        "Setting Ruangan", "Peralatan", "Link KTP", "Link Surat", "OwnerUID", "Timestamp", "Link PDF"
      ]);
    }
  }
  return sheet;
}

/**
 * UTILS: Upload File
 */
function uploadToDrive(base64Data, fileName, mimeType) {
  try {
    if (!base64Data) return "";
    
    // Pastikan folder ada dengan ID yang benar
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    
    const splitBase = base64Data.split(',');
    // Cek apakah format base64 valid (menangani ada/tidaknya prefix data:image/...)
    const dataString = splitBase.length > 1 ? splitBase[1] : splitBase[0];
    
    const blob = Utilities.newBlob(Utilities.base64Decode(dataString), mimeType, fileName);
    const file = folder.createFile(blob);
    
    // Set permission agar bisa dilihat publik (opsional, agar admin bisa lihat tanpa request access)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    // Log error untuk debug di editor
    console.error("Error Upload: " + e.toString());
    // Melempar error agar ditangkap oleh frontend
    throw new Error("Gagal Upload: " + e.toString());
  }
}

/**
 * UTILS: Get Options untuk Dropdown
 */
function getOptionData() {
  const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_OPTION);
  if (!sheet) return { success: false, message: "Sheet Option tidak ditemukan" };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, tempat: [], pengguna: [], rekomendasi: [], instansi: [], libur: [], setting: [] };

  const dataA = sheet.getRange("A2:A" + lastRow).getValues().flat().filter(String); // Tempat
  const dataC = sheet.getRange("C2:C" + lastRow).getValues().flat().filter(String); // Pengguna
  const dataE = sheet.getRange("E2:E" + lastRow).getValues().flat().filter(String); // Rekomendasi
  const dataG = sheet.getRange("G2:G" + lastRow).getValues().flat().filter(String); // Instansi Khusus
  const dataK = sheet.getRange("K2:K" + lastRow).getValues().flat().filter(String); // Hari Libur
  const dataM = sheet.getRange("M2:M" + lastRow).getValues().flat().filter(String); // Setting Ruangan
  const dataO = sheet.getRange("O2:O" + lastRow).getValues().flat().filter(String); // Tahun

  const liburFormatted = dataK.map(d => {
    let date = new Date(d);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  });

  return {
    success: true,
    tempat: dataA,
    pengguna: dataC,
    rekomendasi: dataE,
    instansi: dataG,
    libur: liburFormatted,
    setting: dataM,
    tahun: dataO
  };
}

/**
 * Fungsi Utama Proses Simpan/Update
 */
function processFasilitasi(form, userUid, userName, isUpdate, updateId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    const sheet = getSheet(SHEET_FASILITASI);
    const data = sheet.getDataRange().getValues();

    let targetRow = -1;
    let linkKtp = "";
    let linkSurat = "";

    // --- 0. SECURITY CHECK & SEARCH ---
    if (isUpdate) {
       // Cari baris dan cek kepemilikan
       for(let i=1; i<data.length; i++) {
          if(String(data[i][0]) === String(updateId)) {
            targetRow = i + 1;
            linkKtp = data[i][17]; 
            linkSurat = data[i][18];
            break;
          }
       }
       if(targetRow === -1) return { success: false, message: "Data tidak ditemukan." };
    }

    // --- 1. VALIDASI HARI LIBUR & TABRAKAN JADWAL ---
    const options = getOptionData();
    const inputDate = new Date(form.tanggal);
    const inputDateStr = Utilities.formatDate(inputDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    if (options.libur.includes(inputDateStr)) {
      return { success: false, message: "Layanan ditiadakan pada tanggal tersebut (Libur/Pemeliharaan)." };
    }

    // Persiapan variabel input
    const inputStart = parseInt(form.mulai.replace(/:/g, ""));
    const inputEnd = parseInt(form.selesai.replace(/:/g, ""));
    const inputTempat = String(form.tempat).trim().toLowerCase();
    
    // Nilai jeda persiapan (1 jam = 100)
    const BUFFER_TIME = 10;

    // Ambil Display Values untuk menghindari offset menit (anti-selisih 25 menit)
    const displayValues = sheet.getDataRange().getDisplayValues();

    for (let i = 1; i < data.length; i++) {
      // Lewati jika sedang update baris yang sama
      if (isUpdate && String(data[i][0]) === String(updateId)) continue;
      
      // Lewati jika status pengajuan sudah ditolak
      let rowStatus = String(data[i][1]);
      if (rowStatus === "Ditolak") continue;

      // Normalisasi Tanggal dari Sheet
      let rowDateStr = displayValues[i][2]; 
      if (rowDateStr.includes("/")) {
         rowDateStr = Utilities.formatDate(new Date(data[i][2]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      }

      // Normalisasi Tempat
      let rowTempat = String(data[i][6]).trim().toLowerCase();

      // Jika Tanggal dan Tempat sama, lakukan pengecekan jam + jeda
      if (rowDateStr === inputDateStr && rowTempat === inputTempat) {
        
        // Ambil jam murni dari tampilan sheet (Anti-Offset)
        let rowStart = parseInt(displayValues[i][3].replace(/[^0-9]/g, ""));
        let rowEnd   = parseInt(displayValues[i][4].replace(/[^0-9]/g, ""));

        if (!isNaN(rowStart) && !isNaN(rowEnd)) {
          // LOGIKA INTERSECTION DENGAN JEDA (BUFFER)
          // Memeriksa apakah waktu input bertabrakan dengan (Waktu Sheet + Jeda 1 Jam)
          if (inputStart < (rowEnd + BUFFER_TIME) && (inputEnd + BUFFER_TIME) > rowStart) {
            const penanggungjawab = data[i][11] || "Pihak Lain";
            
            return { 
              success: false, 
              message: `Mohon maaf, "${form.tempat}" pada hari/tanggal tersebut digunakan hingga jam ${displayValues[i][4]}. Harus ada jeda 1 jam untuk persiapan atau silahkan coba pilih ruangan lainnya.` 
            };
          }
        }
      }
    }

    // --- 2. UPLOAD FILE ---
    try {
      if (form.fileKtp && form.fileKtp.data) linkKtp = uploadToDrive(form.fileKtp.data, "KTP_" + userName + "_" + Date.now(), form.fileKtp.mime);
      if (form.fileSurat && form.fileSurat.data) linkSurat = uploadToDrive(form.fileSurat.data, "Surat_" + userName + "_" + Date.now(), form.fileSurat.mime);
    } catch (e) { return { success: false, message: "Gagal Upload: " + e.toString() }; }

    // --- 3. LOGIKA SIMPAN ---
    const valTanggal = Utilities.formatDate(new Date(form.tanggal), "GMT+7", "yyyy-MM-dd");
    let finalMessage = "";

    if (isUpdate) {
      const statusLama = data[targetRow - 1][1]; 
      let statusTarget = statusLama;

      if (statusLama === "Acc") {
        statusTarget = "Blm";
        notifyAdminOfRevision(userName, form.agenda);
      }

      sheet.getRange(targetRow, 2).setValue(statusTarget);
      sheet.getRange(targetRow, 3, 1, 15).setValues([[
        valTanggal, form.mulai, form.selesai, form.peserta, form.tempat, form.agenda,
        form.pengguna, form.instansi, form.rekomendasi, form.pj, form.email, form.alamat, form.kontak,
        form.setting, form.peralatan
      ]]);
      
      sheet.getRange(targetRow, 18).setValue(linkKtp);
      sheet.getRange(targetRow, 19).setValue(linkSurat);
      sheet.getRange(targetRow, 27).setValue(""); // Kosongkan kolom SENT (AA)

      finalMessage = (statusLama === "Acc") 
        ? "Data diperbarui & status di-reset. Admin telah dinotifikasi untuk meninjau ulang." 
        : "Perubahan data berhasil disimpan.";
    } else {
      // Simpan Baris Baru (Kolom ke-21 diisi Timestamp)
      let lastId = data.length > 1 ? parseInt(data[data.length - 1][0]) || 0 : 0;
      sheet.appendRow([
        lastId + 1, "Blm", valTanggal, form.mulai, form.selesai, form.peserta, form.tempat, form.agenda,
        form.pengguna, form.instansi, form.rekomendasi, form.pj, form.email, form.alamat, form.kontak,
        form.setting, form.peralatan, linkKtp, linkSurat, userUid, new Date()
      ]);
      finalMessage = "Berhasil dikirim. Silakan tunggu verifikasi email dari Admin.";
    }

    return { success: true, message: finalMessage };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  } finally { lock.releaseLock(); }
}

/**
 * FUNGSI PEMBAHARU STATUS (DIPANGGIL TOMBOL ADMIN)
 */
function changeStatusFromButton(id, statusBaru) {
  const sheet = getSheet(SHEET_FASILITASI);
  const data = sheet.getDataRange().getValues();
  let targetRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow !== -1) {
    sheet.getRange(targetRow, 2).setValue(statusBaru);
    
    if (statusBaru === "Acc") {
      generateAndSendEmail(targetRow, false); // Panggil fungsi kirim PDF
    } else if (statusBaru === "Ditolak") {
      sendRejectionEmail(targetRow); // Panggil fungsi kirim email penolakan
    }
    return { success: true, message: "Status berhasil diperbarui." };
  }
  return { success: false, message: "Data tidak ditemukan." };
}


/**
 * FITUR FASILITASI: Ambil Data
 */
function getFasilitasiData(userUid) {
  const sheetP = getSheet(SHEET_NAME);
  const dataP = sheetP.getDataRange().getValues();
  let role = "Konstituen";
  let isAdmin = false;
  
  for(let i=1; i<dataP.length; i++) {
    if(String(dataP[i][0]) === String(userUid)) {
      role = dataP[i][6];
      break;
    }
  }
  if(role === "Admin") isAdmin = true;

  const sheet = getSheet(SHEET_FASILITASI);
  // PERUBAHAN: Ambil data nilai asli DAN teks tampilannya
  const range = sheet.getDataRange();
  const data = range.getValues();
  const displayValues = range.getDisplayValues(); // Mengambil teks yang TERLIHAT di sheet
  const result = [];

  for (let i = 1; i < data.length; i++) {
    // Fungsi ini sekarang mengambil teks langsung dari layar Spreadsheet
    const getCleanTimeFromSheet = (rowIdx, colIdx) => {
      let text = displayValues[rowIdx][colIdx];
      if (!text || text === "-") return "-";
      // Ambil HH:mm saja jika ada detik (misal 08:00:00 -> 08:00)
      return text.split(':').slice(0, 2).join(':');
    };

    result.push({
      id: data[i][0],
      status: data[i][1],
      tanggal: (data[i][2] instanceof Date) 
               ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), "yyyy-MM-dd")
               : data[i][2].toString().replace("'", "").trim(),
      // MENGGUNAKAN DISPLAY VALUES (Kolom D=3, E=4)
      mulai: getCleanTimeFromSheet(i, 3), 
      selesai: getCleanTimeFromSheet(i, 4),
      peserta: data[i][5],
      tempat: data[i][6],
      agenda: data[i][7],
      pengguna: data[i][8],
      instansi: data[i][9],
      rekomendasi: data[i][10],
      pj: data[i][11],
      email: data[i][12],
      alamat: data[i][13],
      kontak: data[i][14],
      setting: data[i][15],
      peralatan: data[i][16],
      linkKtp: data[i][17],
      linkSurat: data[i][18],
      ownerUid: data[i][19],
      linkPdf: data[i][21] // Mengambil link PDF dari kolom 22
    });
  }
  
  result.sort((a, b) => b.id - a.id);
  return { success: true, data: result, isAdmin: isAdmin };
}

function deleteFasilitasi(id, userUid) {
  const sheet = getSheet(SHEET_FASILITASI);
  const data = sheet.getDataRange().getValues();
  
  const sheetP = getSheet(SHEET_NAME);
  const dataP = sheetP.getDataRange().getValues();
  let isAdmin = false;
  for(let i=1; i<dataP.length; i++) {
    if(String(dataP[i][0]) === String(userUid) && dataP[i][6] === "Admin") {
      isAdmin = true;
      break;
    }
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      // Validasi: Jika bukan admin, hanya bisa hapus punya sendiri
      if (!isAdmin && String(data[i][19]) !== String(userUid)) {
        return { success: false, message: "Anda tidak berhak menghapus data ini." };
      }
      sheet.deleteRow(i + 1);
      return { success: true, message: "Data berhasil dihapus." };
    }
  }
  return { success: false, message: "Data tidak ditemukan." };
}

function toggleStatusFasilitasi(id, newStatus, userUid) {
  const sheetP = getSheet(SHEET_NAME);
  const dataP = sheetP.getDataRange().getValues();
  let isAdmin = false;
  for(let i=1; i<dataP.length; i++) {
    if(String(dataP[i][0]) === String(userUid) && dataP[i][6] === "Admin") {
      isAdmin = true;
      break;
    }
  }
  
  if(!isAdmin) return { success: false, message: "Hanya Admin yang dapat menyetujui." };

  const sheet = getSheet(SHEET_FASILITASI);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.getRange(i + 1, 2).setValue(newStatus);
      return { success: true, message: `Status berhasil diubah menjadi ${newStatus}.` };
    }
  }
  return { success: false, message: "Data tidak ditemukan." };
}

    /**
     * FITUR FASILITASI: Ambil Opsi Filter & Dropdown
     * Diambil dari SHEET_OPTION sesuai kolom di Spreadsheet
     */
    function getFilterOptions() {
      try {
        const sheet = getSheet(SHEET_OPTION);
        const lastRow = sheet.getLastRow();
        
        // Jika sheet kosong
        if (lastRow < 2) return { success: true, rekomendasi: [], instansi: [] };

        // Ambil data (Sesuaikan kolom: E=5, G=7)
        const rangeRek = sheet.getRange("E2:E" + lastRow).getValues().flat();
        const rangeIns = sheet.getRange("G2:G" + lastRow).getValues().flat();

        // Cleaning: Hapus kosong, hapus "-", hapus duplikat, lalu sort A-Z
        const cleanRek = [...new Set(rangeRek.filter(item => item && item !== "-"))].sort();
        const cleanIns = [...new Set(rangeIns.filter(item => item && item !== "-"))].sort();

        return { 
          success: true, 
          rekomendasi: cleanRek, 
          instansi: cleanIns 
        };
      } catch (e) {
        console.error(e);
        return { success: false, message: e.toString() };
      }
    }

// --- FUNGSI PROFIL LAMA TETAP ADA ---

function registerUser(data) {
  const sheet = getSheet(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === data.email) return { success: false, message: "Email sudah terdaftar." };
  }
  let lastId = rows.length > 1 ? (parseInt(rows[rows.length - 1][0]) || 0) : 0;
  sheet.appendRow([lastId + 1, data.nama, data.email, data.hp, data.alamat, data.instansi, "Konstituen", data.password]);
  return { success: true, message: "Registrasi berhasil! Silakan login." };
}

function loginUser(email, password) {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === email && data[i][7] === password) {
      return {
        success: true,
        user: { uid: data[i][0], nama: data[i][1], email: data[i][2], hp: data[i][3], alamat: data[i][4], instansi: data[i][5], role: data[i][6] }
      };
    }
  }
  return { success: false, message: "Email atau Password salah." };
}

function getAllProfiles(adminUid) {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const users = [];
  let isAdmin = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(adminUid) && data[i][6] === "Admin") { isAdmin = true; break; }
  }
  if (!isAdmin) return { success: false, message: "Akses ditolak." };
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    headers.forEach((header, index) => {
      let key = header.toLowerCase() === 'id' ? 'uid' : header.toLowerCase();
      obj[key] = data[i][index];
    });
    users.push(obj);
  }
  return { success: true, data: users };
}

function updateProfile(uid, updatedData, requesterUid) {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let requesterRole = "";
  
  // Ambil role peminta
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(requesterUid)) {
      requesterRole = data[i][6];
      break; 
    }
  }

  // Cek Izin
  if (String(uid) !== String(requesterUid) && requesterRole !== "Admin") {
    return { success: false, message: "Tidak memiliki izin." };
  }

  // Proses Update
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(uid)) {
      // Hanya update jika data tidak kosong/hanya spasi
      if(updatedData.nama && updatedData.nama.trim()) sheet.getRange(i + 1, 2).setValue(updatedData.nama.trim());
      if(updatedData.hp) sheet.getRange(i + 1, 4).setValue(updatedData.hp);
      if(updatedData.alamat) sheet.getRange(i + 1, 5).setValue(updatedData.alamat);
      if(updatedData.instansi) sheet.getRange(i + 1, 6).setValue(updatedData.instansi);
      
      // Khusus password: Hanya ganti jika user mengisi kolom password
      if(updatedData.password && updatedData.password.trim() !== "") {
        sheet.getRange(i + 1, 8).setValue(updatedData.password);
      }
      
      // Update role (Hanya Admin)
      if(requesterRole === "Admin" && updatedData.role) {
        sheet.getRange(i + 1, 7).setValue(updatedData.role);
      }
      
      return { success: true, message: "Data berhasil diperbarui." };
    }
  }
  return { success: false, message: "User tidak ditemukan." };
}

function deleteProfile(uid, adminUid) {
  const sheet = getSheet(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let isAdmin = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(adminUid) && data[i][6] === "Admin") { isAdmin = true; break; }
  }
  if (!isAdmin) return { success: false, message: "Akses ditolak." };
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(uid)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "User dihapus." };
    }
  }
  return { success: false, message: "User tidak ditemukan." };
}

function getAllStats(startDate = null, endDate = null, selectedYear = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Stat');
  
  // 1. Logika Penulisan Filter Tanggal (H1 & I1)
  if (startDate && endDate) {
    sheet.getRange("H1").setValue(startDate); 
    sheet.getRange("I1").setValue(endDate);
  } else {
    sheet.getRange("H1").setValue("2010-01-01");
    sheet.getRange("I1").setValue("2050-12-31");
  }

  // 2. Logika Penulisan Tahun (H2) - REVISI DI SINI
  if (selectedYear) {
    sheet.getRange("H2").setValue(selectedYear);
  } else {
    // Ambil tahun berjalan secara otomatis jika selectedYear kosong (null)
    const currentYear = new Date().getFullYear();
    sheet.getRange("H2").setValue(currentYear);
  }
  
  SpreadsheetApp.flush(); 

  return {
    scorecards: {
      acc: sheet.getRange("B4").getValue() || 0,
      blm: sheet.getRange("B5").getValue() || 0,
      ditolak: sheet.getRange("B6").getValue() || 0
    },
    pengguna: sheet.getRange("A11:B14").getValues().filter(r => r[0] !== ""),
    anggota: sheet.getRange("A19:B25").getValues().filter(r => r[0] !== ""),
    rekomendasi: sheet.getRange("A30:B33").getValues().filter(r => r[0] !== ""),
    tempat: sheet.getRange("A54:B57").getValues().filter(r => r[0] !== ""),
    bulanan: sheet.getRange("A38:B49").getValues().filter(r => r[0] !== "")
  };
}

/**
 * FUNGSI KIRIM PDF (Hanya saat Acc)
 */
function generateAndSendEmail(targetRow, isUpdate) {
  const sheet = getSheet(SHEET_FASILITASI);
  const row = sheet.getRange(targetRow, 1, 1, 27).getValues()[0];
  const displayRow = sheet.getRange(targetRow, 1, 1, 27).getDisplayValues()[0];

  const emailTujuan = row[12];
  if (!emailTujuan) return;

  try {
    const templateId = "1iOMFFr8ZQUxjVxBrhMqvnWVU9-xiP-un3epGI0t1ht0"; 
    const folderId   = "1fTPAzC2i9GKPivhpavVB5i0x8gPk6_FJ"; 
    
    const templateFile = DriveApp.getFileById(templateId);
    const folder       = DriveApp.getFolderById(folderId);
    
    // --- Format Nama File Baru ---
    const dateObj = new Date(row[2]);
    const formattedDate = Utilities.formatDate(dateObj, "GMT+7", "yyyy-MM-dd");
    const namaInstansi = row[9] ? String(row[9]).replace(/[/\\?%*:|"<>\s]/g, "_") : "Tanpa_Instansi"; 
    const fileName = `Bukti_Pengajuan_${formattedDate}_${namaInstansi}`;
    
    const tempCopy = templateFile.makeCopy(fileName, folder);
    const tempDoc  = DocumentApp.openById(tempCopy.getId());
    const body     = tempDoc.getBody();

    const hariArr   = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const bulanArr  = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

    // Mapping Placeholders
    body.replaceText("<<ID>>", row[0]);
    body.replaceText("<<Pengguna>>", row[8]);
    body.replaceText("<<Nama  penanggung jawab>>", row[11]);
    body.replaceText("<<Alamat Penanggung jawab>>", row[13]);
    body.replaceText("<<Nomor HP>>", row[14]);
    body.replaceText("<<Nama Instansi/Organisasi>>", row[9]);
    body.replaceText("<<Nama Kegiatan>>", row[7]);
    body.replaceText("<<tgl>>", Utilities.formatDate(dateObj, "GMT+7", "dd/MM/yyyy"));
    body.replaceText("<<Waktu Mulai>>", displayRow[3]);
    body.replaceText("<<Waktu Selesai>>", displayRow[4]);
    body.replaceText("<<Tempat>>", row[6]);
    body.replaceText("<<Jumlah Peserta>>", row[5]);
    body.replaceText("<<hari>>", hariArr[dateObj.getDay()]);
    body.replaceText("<<tanggals>>", dateObj.getDate());
    body.replaceText("<<bulan>>", bulanArr[dateObj.getMonth()]);
    body.replaceText("<<penanggungjawab>>", row[11]);

    tempDoc.saveAndClose();
    const pdfBlob = tempCopy.getAs('application/pdf');
    const savedPdf = folder.createFile(pdfBlob); 
    savedPdf.setName(fileName + ".pdf");

    // --- PERBAIKAN 2: Ambil URL File PDF dan Simpan ke Lembar Google Sheet (Kolom 22) ---
    const pdfUrl = savedPdf.getUrl();
    sheet.getRange(targetRow, 22).setValue(pdfUrl);
    // ---------------------------------------------------------------------------------

    MailApp.sendEmail({
      to: emailTujuan,
      subject: "✅ DISETUJUI: Bukti Pengajuan Fasilitasi - " + row[11],
      htmlBody: `<p>Halo <b>${row[11]}</b>,</p><p>Permohonan Anda untuk agenda <b>${row[7]}</b> telah disetujui, permohonan fasilitasi Anda telah berhasil kami terima. Silakan simpan dokumen bukti pengajuan yang terlampir di bawah ini.</p>`,
      attachments: [pdfBlob]
    });

    tempCopy.setTrashed(true);
    sheet.getRange(targetRow, 27).setValue("SENT"); 
  } catch (e) { console.error("Gagal Email PDF: " + e.toString()); }
}

/**
 * FUNGSI EMAIL PENOLAKAN
 */
function sendRejectionEmail(targetRow) {
  const sheet = getSheet(SHEET_FASILITASI);
  const row = sheet.getRange(targetRow, 1, 1, 15).getValues()[0];
  MailApp.sendEmail({
    to: row[12],
    subject: "❌ Update Permohonan: " + row[7],
    htmlBody: `<h2 style="color: #d32f2f;">Permohonan Belum Disetujui</h2>
      <p>Halo <b>${row[11]}</b>,</p>
      <p>Terima kasih telah mengajukan permohonan fasilitasi untuk kegiatan: <b>${row[7]}</b>.</p>
      <p>Mohon maaf, saat ini kami belum dapat menyetujui permohonan Anda dikarenakan alasan teknis atau jadwal yang penuh.</p>
      <p>Silakan ajukan kembali di waktu lain atau hubungi Admin untuk informasi lebih lanjut.</p>
      <p>Terima kasih.</p>`
  });
}

/**
 * NOTIFIKASI ADMIN SAAT USER EDIT DATA ACC
 */
function notifyAdminOfRevision(user, agenda) {
  // 1. Ambil email dari 'Brankas' (Script Properties)
  const props = PropertiesService.getScriptProperties();
  const adminEmail = props.getProperty('ADMIN_EMAIL');

  // 2. Proteksi jika properti lupa diisi
  if (!adminEmail) {
    Logger.log("ERROR: Email admin belum dikonfigurasi di Script Properties.");
    return;
  }

  const htmlBody = `
    <div style="font-family: sans-serif; background: #fff4e5; padding: 20px; border-left: 5px solid #ffa000;">
      <h3>⚠️ Revisi Data (Penting)</h3>
      <p>User <b>${user}</b> baru saja mengubah data agenda: <b>${agenda}</b>.</p>
      <p>Karena data ini sebelumnya sudah berstatus <b>Acc</b>, sistem telah mereset statusnya menjadi <b>Blm (Menunggu)</b>.</p>
      <p>Mohon segera cek dashboard Admin untuk meninjau ulang perubahan tersebut.</p>
    </div>
  `;

  MailApp.sendEmail({ 
    to: adminEmail, 
    subject: "⚠️ NOTIFIKASI REVISI: " + agenda, 
    htmlBody: htmlBody 
  });
}