/**
 * Setup.gs - "Pemasang" sistem untuk klien baru.
 *
 * TUJUAN: kode .gs & .html sistem ini 100% SAMA untuk semua klien (universal).
 * Yang berbeda per klien HANYA:
 *   1. Satu spreadsheet baru (kosong) per klien.
 *   2. Nilai "Nama Klien" yang disimpan di Script Properties spreadsheet itu.
 *   3. URL deployment Web App-nya sendiri (otomatis beda tiap deploy baru).
 *
 * CARA PAKAI untuk klien baru:
 *   1. Buat Google Spreadsheet baru (kosong).
 *   2. Extensions > Apps Script > tempel SEMUA file .gs ini apa adanya (tanpa diubah).
 *   3. Reload spreadsheet-nya (supaya menu custom muncul).
 *   4. Klik menu "⚙️ Menu Absensi RAY" > "🚀 1. Setup Sistem Baru (Buat Semua Sheet)".
 *      -> Ini akan tanya "Nama Klien" lalu otomatis membuat SEMUA sheet yang
 *         dibutuhkan (Data Master Karyawan, Data Master Lokasi, dst) lengkap
 *         dengan header kolomnya. Aman dijalankan ulang - sheet yang sudah ada
 *         tidak akan ditimpa/dihapus datanya.
 *   5. Lanjut isi menu "⚙️ 2. Pengaturan API Key & Folder Foto" seperti biasa.
 *   6. Deploy > New deployment > Web app, salin URL-nya.
 *   7. Di repo GitHub Pages (index.html & admin.html), tempel URL itu ke
 *      variabel API_URL. Kode HTML-nya sendiri tidak perlu diubah sama sekali -
 *      nama klien akan otomatis muncul di layar karena diambil lewat API.
 */

var PROP_NAMA_KLIEN = 'NAMA_KLIEN';

// ============ WIZARD: SETUP SISTEM BARU (dipanggil dari menu) ============

function pasangSistemBaru() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var namaKlienLama = props.getProperty(PROP_NAMA_KLIEN);

  var promptTeks = namaKlienLama
    ? 'Nama klien saat ini: "' + namaKlienLama + '".\nKosongkan lalu OK untuk membiarkan tetap sama, atau isi nama baru untuk mengganti:'
    : 'Sistem ini akan dipakai untuk klien apa? (contoh: "PT Sinar Jaya Abadi")\nNama ini akan muncul di layar login karyawan & admin.';

  var res = ui.prompt('🚀 Setup Sistem Baru', promptTeks, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var namaBaru = res.getResponseText().trim();
  if (namaBaru) {
    props.setProperty(PROP_NAMA_KLIEN, namaBaru);
  } else if (!namaKlienLama) {
    ui.alert('Dibatalkan', 'Nama klien tidak boleh kosong untuk setup pertama kali.', ui.ButtonSet.OK);
    return;
  }

  var ringkasan = pasangSemuaSheet_();
  var folderKlien = buatFolderKlien_();

  ui.alert(
    '✅ Setup Selesai',
    'Klien: ' + props.getProperty(PROP_NAMA_KLIEN) + '\n\n' +
    'Sheet yang dibuat baru:\n' + (ringkasan.dibuat.length ? '- ' + ringkasan.dibuat.join('\n- ') : '(tidak ada, semua sudah ada)') +
    '\n\nSheet yang sudah ada sebelumnya (tidak diubah):\n' + (ringkasan.sudahAda.length ? '- ' + ringkasan.sudahAda.join('\n- ') : '(tidak ada)') +
    '\n\nFolder foto Drive: ' + folderKlien.getUrl() +
    ' (berisi sub-folder Absensi/Izin/Patroli/Kegiatan, sejajar dengan file spreadsheet ini)' +
    '\n\nLangkah selanjutnya: jalankan menu "⚙️ 2. Pengaturan API Key & Folder Foto" (isi password admin & API key), lalu Deploy > New deployment > Web app.',
    ui.ButtonSet.OK
  );
}

// Idempotent: aman dipanggil berkali-kali, tidak pernah menghapus/menimpa sheet yang sudah ada.
function pasangSemuaSheet_() {
  var dibuat = [];
  var sudahAda = [];

  var daftarSheetInti = [
    { nama: SHEET_KARYAWAN, fungsi: getOrBuatSheetKaryawan_ },
    { nama: SHEET_LOKASI, fungsi: getOrBuatSheetLokasiMaster_ },
    { nama: SHEET_ABSENSI, fungsi: getOrBuatSheetAbsensiMaster_ },
    { nama: SHEET_JAM_KERJA, fungsi: getOrBuatSheetJamKerja_ },
    { nama: SHEET_IJIN, fungsi: getOrBuatSheetIjin_ },
    { nama: SHEET_PENDAFTARAN, fungsi: getOrBuatSheetPendaftaran_ },
    { nama: SHEET_TITIK_PATROLI, fungsi: getOrBuatSheetTitikPatroli_ },
    { nama: SHEET_DATA_PATROLI, fungsi: getOrBuatSheetDataPatroli_ },
    { nama: SHEET_MASTER_TUGAS, fungsi: getOrBuatSheetMasterTugas_ },
    { nama: SHEET_LAPORAN_KEGIATAN, fungsi: getOrBuatSheetLaporanKegiatan_ }
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  daftarSheetInti.forEach(function (item) {
    var sudahAdaSebelumnya = !!ss.getSheetByName(item.nama);
    item.fungsi(); // getOrBuat...: buat kalau belum ada, biarkan kalau sudah ada
    if (sudahAdaSebelumnya) {
      sudahAda.push(item.nama);
    } else {
      dibuat.push(item.nama);
    }
  });

  // Bersihkan sheet default kosong "Sheet1" bawaan Google Spreadsheet (kalau ada & benar-benar kosong)
  var sheetDefault = ss.getSheetByName('Sheet1');
  if (sheetDefault && sheetDefault.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheetDefault);
  }

  return { dibuat: dibuat, sudahAda: sudahAda };
}

// ============ FOLDER DRIVE OTOMATIS (foto absen/izin/patroli/kegiatan) ============
// Folder klien dibuat SEJAJAR dengan file spreadsheet ini (folder induknya sama),
// supaya di mana pun spreadsheet-nya Anda taruh/pindahkan, folder foto ikut di situ.
// Aman dipanggil ulang: kalau folder dengan nama yang sama sudah ada, dipakai lagi
// (tidak membuat folder duplikat).

var PROP_FOLDER_UTAMA = 'FOLDER_ID';
var PROP_FOLDER_ABSENSI = 'FOLDER_ID_ABSENSI';
var PROP_FOLDER_IJIN = 'FOLDER_ID_IJIN';
var PROP_FOLDER_PATROLI = 'FOLDER_ID_PATROLI';
var PROP_FOLDER_KEGIATAN = 'FOLDER_ID_KEGIATAN';

function buatFolderKlien_() {
  var props = PropertiesService.getScriptProperties();
  var namaKlien = props.getProperty(PROP_NAMA_KLIEN) || 'Klien Tanpa Nama';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fileSpreadsheet = DriveApp.getFileById(ss.getId());
  var indukIterator = fileSpreadsheet.getParents();
  var folderInduk = indukIterator.hasNext() ? indukIterator.next() : DriveApp.getRootFolder();

  var folderKlien = _cariAtauBuatFolder_(folderInduk, namaKlien);
  var folderAbsensi = _cariAtauBuatFolder_(folderKlien, 'Absensi');
  var folderIjin = _cariAtauBuatFolder_(folderKlien, 'Izin');
  var folderPatroli = _cariAtauBuatFolder_(folderKlien, 'Patroli');
  var folderKegiatan = _cariAtauBuatFolder_(folderKlien, 'Kegiatan');

  props.setProperty(PROP_FOLDER_UTAMA, folderKlien.getId());
  props.setProperty(PROP_FOLDER_ABSENSI, folderAbsensi.getId());
  props.setProperty(PROP_FOLDER_IJIN, folderIjin.getId());
  props.setProperty(PROP_FOLDER_PATROLI, folderPatroli.getId());
  props.setProperty(PROP_FOLDER_KEGIATAN, folderKegiatan.getId());

  return folderKlien;
}

function _cariAtauBuatFolder_(folderInduk, namaFolder) {
  var it = folderInduk.getFoldersByName(namaFolder);
  if (it.hasNext()) return it.next();
  return folderInduk.createFolder(namaFolder);
}

// ============ DOWNLOAD KODE WEB (untuk di-upload ke GitHub Pages) ============
// Mengambil isi file "Template_Index.html" & "Template_Admin.html" yang sudah Anda
// tempel di project Apps Script ini (SEKALI saja, tidak berubah per klien), mengganti
// placeholder __API_URL__ dengan URL Web App deployment yang aktif SEKARANG, lalu
// menyimpan hasilnya ke folder Drive klien - siap didownload & upload ke GitHub.
//
// CATATAN: file "Template_Index.html" dan "Template_Admin.html" harus Anda tambahkan
// manual 1x lewat Apps Script Editor > File > New > HTML, isinya SAMA PERSIS dengan
// index.html / admin.html final, HANYA baris "const API_URL = '...'" diganti jadi
// "const API_URL = '__API_URL__';".

function unduhKodeWeb() {
  var ui = SpreadsheetApp.getUi();
  var urlWebApp = ScriptApp.getService().getUrl();
  if (!urlWebApp) {
    ui.alert('Belum Di-deploy', 'Web App belum di-deploy sebagai "Anyone can access". Jalankan Deploy > New deployment > Web app dulu, baru jalankan menu ini lagi.', ui.ButtonSet.OK);
    return;
  }

  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(PROP_FOLDER_UTAMA);
  if (!folderId) {
    ui.alert('Folder Belum Ada', 'Jalankan menu "🚀 1. Setup Sistem Baru" dulu supaya folder Drive klien tersedia.', ui.ButtonSet.OK);
    return;
  }

  var htmlIndex, htmlAdmin;
  try {
    htmlIndex = HtmlService.createHtmlOutputFromFile('Template_Index').getContent();
    htmlAdmin = HtmlService.createHtmlOutputFromFile('Template_Admin').getContent();
  } catch (e) {
    ui.alert(
      'File Template Belum Ada',
      'Pastikan file "Template_Index.html" dan "Template_Admin.html" sudah ditambahkan di project Apps Script ini ' +
      '(Apps Script Editor > File > New > HTML, beri nama persis "Template_Index" dan "Template_Admin").',
      ui.ButtonSet.OK
    );
    return;
  }

  htmlIndex = htmlIndex.split('__API_URL__').join(urlWebApp);
  htmlAdmin = htmlAdmin.split('__API_URL__').join(urlWebApp);

  var folder = DriveApp.getFolderById(folderId);
  var fileIndex = _simpanAtauGantiFile_(folder, 'index.html', htmlIndex, MimeType.HTML);
  var fileAdmin = _simpanAtauGantiFile_(folder, 'admin.html', htmlAdmin, MimeType.HTML);

  ui.alert(
    '✅ Kode Web Siap Didownload',
    'File "index.html" dan "admin.html" (URL Web App sudah otomatis terisi) sudah disimpan di folder Drive klien:\n\n' +
    folder.getUrl() +
    '\n\nBuka folder itu, download kedua file tadi, lalu upload ke repo GitHub Pages Anda bersama file "icon-192.png" (logo tidak ikut digenerate di sini, tetap pakai file logo yang sama seperti biasa).',
    ui.ButtonSet.OK
  );
}

function _simpanAtauGantiFile_(folder, namaFile, konten, mimeType) {
  var it = folder.getFilesByName(namaFile);
  while (it.hasNext()) { it.next().setTrashed(true); } // buang versi lama supaya tidak dobel
  return folder.createFile(namaFile, konten, mimeType);
}

// ============ SHEET INTI YANG SEBELUMNYA DIBUAT MANUAL ============
// (Data Master Karyawan, Data Master Lokasi, Data Absensi tadinya harus dibuat
// tangan oleh admin sebelum sistem bisa dipakai - sekarang ikut dibuat otomatis
// oleh wizard di atas, dengan pola getOrBuatSheet..._ yang sama seperti modul lain.)

function getOrBuatSheetKaryawan_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_KARYAWAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_KARYAWAN);
    sheet.appendRow(['Nama', 'No HP (opsional)', 'Alamat (opsional)', 'Bagian Kerja', 'PIN']);
  }
  return sheet;
}

function getOrBuatSheetLokasiMaster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LOKASI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOKASI);
    sheet.appendRow(['Nama Lokasi', 'Latitude', 'Longitude', 'Radius', 'Satuan Radius']);
  }
  return sheet;
}

function getOrBuatSheetAbsensiMaster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_ABSENSI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ABSENSI);
    sheet.appendRow(['Nama', 'Jenis Absen', 'Tanggal', 'Waktu', 'Foto', 'Latitude', 'Longitude', 'Akurasi',
      'Kelurahan', 'Kecamatan', 'Kabupaten', 'Link Maps', 'Bagian Kerja']);
  }
  return sheet;
}

// ============ INFO SISTEM (dipanggil frontend, tanpa login) ============
// Dipakai index.html & admin.html untuk menampilkan nama klien di layar,
// supaya kode HTML-nya universal (tidak perlu di-hardcode per klien).

function getInfoSistem() {
  var namaKlien = PropertiesService.getScriptProperties().getProperty(PROP_NAMA_KLIEN);
  return {
    sukses: true,
    namaKlien: namaKlien || 'Sistem Absensi RAY'
  };
}
