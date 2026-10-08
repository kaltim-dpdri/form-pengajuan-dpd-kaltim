# SIRUPAT — Migrasi Vercel + Google Apps Script

Arsitektur:

Vercel (`index.html`) → `fetch()` POST → Google Apps Script Web App (`kode.js`) → Google Spreadsheet / Google Drive / Gmail

## 1. Apps Script

1. Buka project Google Apps Script yang sekarang digunakan SIRUPAT.
2. Backup `kode.js` lama.
3. Ganti isi file backend dengan `kode.js` hasil migrasi ini.
4. Pastikan ID Spreadsheet, Folder Drive, dan Script Properties yang sudah digunakan sistem tetap benar.
5. Jalankan `checkPermissions()` sekali dari editor Apps Script untuk memastikan akses Drive.
6. Deploy → New deployment → Web app.
7. Execute as: Me.
8. Who has access: Anyone.
9. Salin URL `/exec` hasil deployment.

## 2. Frontend Vercel

1. Upload `index.html` ke repository GitHub.
2. Pastikan konstanta `API_URL` di `index.html` menunjuk ke URL Web App Apps Script `/exec`.
3. Import repository ke Vercel.
4. Framework Preset: Other / static.
5. Build Command: kosong.
6. Output Directory: `.`
7. Deploy.

## 3. Perubahan utama

- Semua `google.script.run` di frontend sudah dihapus.
- Frontend memakai `fetch()` melalui fungsi `api(action, payload)`.
- Request dikirim sebagai `text/plain` agar tidak memicu CORS preflight `application/json`.
- Apps Script menerima request melalui `doPost(e)` dan melakukan routing berdasarkan `action`.
- Upload KTP/surat tetap dikirim sebagai Base64 ke Apps Script lalu disimpan ke Google Drive.
- PDF dan email tetap diproses di Apps Script.
- Validasi Admin untuk perubahan status ditambahkan di backend.
- Validasi kepemilikan data saat edit fasilitasi ditambahkan di backend.
- Statistik memakai Spreadsheet ID melalui `getSheet("Stat")`, bukan `getActiveSpreadsheet()`, agar tetap bekerja sebagai Web App backend.

## 4. Action API yang tersedia

`loginUser`
`registerUser`
`getFasilitasiData`
`getFilterOptions`
`getOptionData`
`processFasilitasi`
`deleteFasilitasi`
`changeStatusFromButton`
`updateProfile`
`getAllProfiles`
`deleteProfile`
`getAllStats`

## 5. Catatan keamanan

Jangan menaruh password, token rahasia, atau credential Google di `index.html` atau repository publik. URL Web App boleh diketahui frontend, tetapi akses terhadap Spreadsheet/Drive/Gmail tetap dilakukan oleh Apps Script.
