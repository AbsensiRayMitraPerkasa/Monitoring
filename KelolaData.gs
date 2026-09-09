/**
 * KelolaData.gs - CRUD data master (Karyawan, Lokasi Absensi, Titik Patroli,
 * Tugas Kebersihan, Jam Kerja Standar) untuk panel admin web.
 *
 * KHUSUS peran 'internal' (PT Ray Mitra Perkasa) - admin client TIDAK boleh
 * akses fungsi-fungsi ini (ditegakkan lewat _cekAdminInternal_ di setiap fungsi,
 * sama seperti pola di Ijin.gs & Pendaftaran.gs, bukan cuma disembunyikan di UI).
 */

function _cekAdminInternal_(password) {
  var cek = loginAdmin(password);
  if (!cek.sukses) return cek;
  if (cek.peran !== 'internal') return { sukses: false, pesan: 'Fitur ini hanya untuk Admin Internal.' };
  return { sukses: true };
}

function _hapusBarisSheet_(sheet, baris) {
  var baris_ = parseInt(baris, 10);
  if (isNaN(baris_) || baris_ < 2 || baris_ > sheet.getLastRow()) {
    return { sukses: false, pesan: 'Baris tidak valid atau sudah terhapus.' };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }
  try {
    sheet.deleteRow(baris_);
    return { sukses: true, pesan: 'Data berhasil dihapus.' };
  } finally {
    lock.releaseLock();
  }
}

function _cekDuplikatKodeQr_(sheet, kolomKode, kodeQr, kecualiBaris) {
  kodeQr = String(kodeQr || '').trim();
  if (!kodeQr) return null; // kode boleh kosong (belum dipakai QR), tidak perlu dicek unik
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((i + 1) === kecualiBaris) continue;
    if (String(data[i][kolomKode] || '').trim() === kodeQr) {
      return 'Kode Titik "' + kodeQr + '" sudah dipakai titik lain, gunakan kode yang berbeda.';
    }
  }
  return null;
}

function _validasiTitikLokasi_(nama, lat, lng) {
  if (!nama) return 'Nama tidak boleh kosong.';
  var latNum = parseFloat(String(lat).replace(',', '.'));
  var lngNum = parseFloat(String(lng).replace(',', '.'));
  if (isNaN(latNum) || isNaN(lngNum)) return 'Latitude/Longitude harus berupa angka.';
  if (Math.abs(latNum) > 90 || Math.abs(lngNum) > 180) return 'Latitude/Longitude di luar jangkauan yang valid.';
  return null;
}


// ============ KARYAWAN ============

function getDaftarKaryawanAdmin(password) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  var sheet = getOrBuatSheetKaryawan_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    out.push({
      baris: i + 1,
      nama: String(data[i][0] || ''),
      noHp: String(data[i][1] || ''),
      alamat: String(data[i][2] || ''),
      bagianKerja: String(data[i][3] || ''),
      pin: String(data[i][4] || '')
    });
  }
  return { sukses: true, daftar: out };
}

function _validasiKaryawan_(nama, bagianKerja, pin) {
  if (!nama) return 'Nama tidak boleh kosong.';
  if (DAFTAR_BAGIAN_KERJA.indexOf(bagianKerja) === -1) return 'Bagian kerja tidak valid.';
  if (!/^[0-9]{4}$/.test(String(pin || ''))) return 'PIN harus 4 digit angka.';
  return null;
}

function tambahKaryawanAdmin(password, nama, noHp, alamat, bagianKerja, pin) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;

  nama = String(nama || '').trim();
  var err = _validasiKaryawan_(nama, bagianKerja, pin);
  if (err) return { sukses: false, pesan: err };

  var sheet = getOrBuatSheetKaryawan_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === nama) return { sukses: false, pesan: 'Nama "' + nama + '" sudah terdaftar.' };
  }

  sheet.appendRow([nama, String(noHp || '').trim(), String(alamat || '').trim(), bagianKerja, String(pin).trim()]);
  return { sukses: true, pesan: 'Karyawan "' + nama + '" berhasil ditambahkan.' };
}

function editKaryawanAdmin(password, baris, nama, noHp, alamat, bagianKerja, pin) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;

  nama = String(nama || '').trim();
  var err = _validasiKaryawan_(nama, bagianKerja, pin);
  if (err) return { sukses: false, pesan: err };

  var sheet = getOrBuatSheetKaryawan_();
  var baris_ = parseInt(baris, 10);
  if (isNaN(baris_) || baris_ < 2 || baris_ > sheet.getLastRow()) return { sukses: false, pesan: 'Baris tidak valid.' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((i + 1) !== baris_ && String(data[i][0]).trim() === nama) {
      return { sukses: false, pesan: 'Nama "' + nama + '" sudah dipakai karyawan lain.' };
    }
  }

  sheet.getRange(baris_, 1, 1, 5).setValues([[nama, String(noHp || '').trim(), String(alamat || '').trim(), bagianKerja, String(pin).trim()]]);
  return { sukses: true, pesan: 'Data karyawan "' + nama + '" berhasil diperbarui.' };
}

function hapusKaryawanAdmin(password, baris) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  return _hapusBarisSheet_(getOrBuatSheetKaryawan_(), baris);
}


// ============ LOKASI ABSENSI ============

function getDaftarLokasiAdmin(password) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  var sheet = getOrBuatSheetLokasiMaster_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    out.push({
      baris: i + 1, nama: String(data[i][0] || ''),
      lat: data[i][1], lng: data[i][2], radius: data[i][3],
      satuan: String(data[i][4] || 'Meter')
    });
  }
  return { sukses: true, daftar: out };
}

function tambahLokasiAdmin(password, nama, lat, lng, radius, satuan) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  nama = String(nama || '').trim();
  var err = _validasiTitikLokasi_(nama, lat, lng);
  if (err) return { sukses: false, pesan: err };
  getOrBuatSheetLokasiMaster_().appendRow([nama, lat, lng, parseFloat(radius) || 50, satuan || 'Meter']);
  return { sukses: true, pesan: 'Lokasi "' + nama + '" berhasil ditambahkan.' };
}

function editLokasiAdmin(password, baris, nama, lat, lng, radius, satuan) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  nama = String(nama || '').trim();
  var err = _validasiTitikLokasi_(nama, lat, lng);
  if (err) return { sukses: false, pesan: err };
  var sheet = getOrBuatSheetLokasiMaster_();
  var baris_ = parseInt(baris, 10);
  if (isNaN(baris_) || baris_ < 2 || baris_ > sheet.getLastRow()) return { sukses: false, pesan: 'Baris tidak valid.' };
  sheet.getRange(baris_, 1, 1, 5).setValues([[nama, lat, lng, parseFloat(radius) || 50, satuan || 'Meter']]);
  return { sukses: true, pesan: 'Lokasi "' + nama + '" berhasil diperbarui.' };
}

function hapusLokasiAdmin(password, baris) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  return _hapusBarisSheet_(getOrBuatSheetLokasiMaster_(), baris);
}


// ============ TITIK PATROLI ============

function getDaftarTitikPatroliAdmin(password) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  var sheet = getOrBuatSheetTitikPatroli_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    out.push({
      baris: i + 1, namaPos: String(data[i][0] || ''),
      lat: data[i][1], lng: data[i][2], radius: data[i][3],
      satuan: String(data[i][4] || 'Meter'), kodeQr: String(data[i][5] || '')
    });
  }
  return { sukses: true, daftar: out };
}

function tambahTitikPatroliAdmin(password, namaPos, lat, lng, radius, satuan, kodeQr) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  namaPos = String(namaPos || '').trim();
  var err = _validasiTitikLokasi_(namaPos, lat, lng);
  if (err) return { sukses: false, pesan: err };
  var sheet = getOrBuatSheetTitikPatroli_();
  var errDup = _cekDuplikatKodeQr_(sheet, 5, kodeQr, -1);
  if (errDup) return { sukses: false, pesan: errDup };
  sheet.appendRow([namaPos, lat, lng, parseFloat(radius) || 50, satuan || 'Meter', String(kodeQr || '').trim()]);
  return { sukses: true, pesan: 'Titik patroli "' + namaPos + '" berhasil ditambahkan.' };
}

function editTitikPatroliAdmin(password, baris, namaPos, lat, lng, radius, satuan, kodeQr) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  namaPos = String(namaPos || '').trim();
  var err = _validasiTitikLokasi_(namaPos, lat, lng);
  if (err) return { sukses: false, pesan: err };
  var sheet = getOrBuatSheetTitikPatroli_();
  var baris_ = parseInt(baris, 10);
  if (isNaN(baris_) || baris_ < 2 || baris_ > sheet.getLastRow()) return { sukses: false, pesan: 'Baris tidak valid.' };
  var errDup = _cekDuplikatKodeQr_(sheet, 5, kodeQr, baris_);
  if (errDup) return { sukses: false, pesan: errDup };
  sheet.getRange(baris_, 1, 1, 6).setValues([[namaPos, lat, lng, parseFloat(radius) || 50, satuan || 'Meter', String(kodeQr || '').trim()]]);
  return { sukses: true, pesan: 'Titik patroli "' + namaPos + '" berhasil diperbarui.' };
}

function hapusTitikPatroliAdmin(password, baris) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  return _hapusBarisSheet_(getOrBuatSheetTitikPatroli_(), baris);
}


// ============ TUGAS KEBERSIHAN ============

function getDaftarTugasAdmin(password) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  var sheet = getOrBuatSheetMasterTugas_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    out.push({
      baris: i + 1, namaTugas: String(data[i][0] || ''), area: String(data[i][1] || ''),
      lat: data[i][2], lng: data[i][3], radius: data[i][4], satuan: String(data[i][5] || 'Meter'),
      kodeQr: String(data[i][6] || ''), keterangan: String(data[i][7] || '')
    });
  }
  return { sukses: true, daftar: out };
}

// Catatan: lat/lng untuk Tugas Kebersihan bersifat OPSIONAL (lihat Cleaning.gs -
// kalau kosong, laporan kegiatan untuk tugas itu tidak divalidasi geofencing).
// Karena itu validasinya sedikit beda dari Lokasi/Titik Patroli yang wajib isi.
function _validasiTugas_(namaTugas, lat, lng) {
  if (!namaTugas) return 'Nama tugas tidak boleh kosong.';
  var latKosong = (lat === '' || lat === null || lat === undefined);
  var lngKosong = (lng === '' || lng === null || lng === undefined);
  if (latKosong && lngKosong) return null; // boleh keduanya kosong = tanpa validasi lokasi
  var latNum = parseFloat(String(lat).replace(',', '.'));
  var lngNum = parseFloat(String(lng).replace(',', '.'));
  if (isNaN(latNum) || isNaN(lngNum)) return 'Latitude/Longitude harus berupa angka, atau dikosongkan KEDUANYA.';
  return null;
}

function tambahTugasAdmin(password, namaTugas, area, lat, lng, radius, satuan, kodeQr, keterangan) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  namaTugas = String(namaTugas || '').trim();
  var err = _validasiTugas_(namaTugas, lat, lng);
  if (err) return { sukses: false, pesan: err };
  var sheet = getOrBuatSheetMasterTugas_();
  var errDup = _cekDuplikatKodeQr_(sheet, 6, kodeQr, -1);
  if (errDup) return { sukses: false, pesan: errDup };
  sheet.appendRow([namaTugas, String(area || '').trim(), lat || '', lng || '', parseFloat(radius) || 50, satuan || 'Meter', String(kodeQr || '').trim(), String(keterangan || '').trim()]);
  return { sukses: true, pesan: 'Tugas "' + namaTugas + '" berhasil ditambahkan.' };
}

function editTugasAdmin(password, baris, namaTugas, area, lat, lng, radius, satuan, kodeQr, keterangan) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  namaTugas = String(namaTugas || '').trim();
  var err = _validasiTugas_(namaTugas, lat, lng);
  if (err) return { sukses: false, pesan: err };
  var sheet = getOrBuatSheetMasterTugas_();
  var baris_ = parseInt(baris, 10);
  if (isNaN(baris_) || baris_ < 2 || baris_ > sheet.getLastRow()) return { sukses: false, pesan: 'Baris tidak valid.' };
  var errDup = _cekDuplikatKodeQr_(sheet, 6, kodeQr, baris_);
  if (errDup) return { sukses: false, pesan: errDup };
  sheet.getRange(baris_, 1, 1, 8).setValues([[namaTugas, String(area || '').trim(), lat || '', lng || '', parseFloat(radius) || 50, satuan || 'Meter', String(kodeQr || '').trim(), String(keterangan || '').trim()]]);
  return { sukses: true, pesan: 'Tugas "' + namaTugas + '" berhasil diperbarui.' };
}

function hapusTugasAdmin(password, baris) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  return _hapusBarisSheet_(getOrBuatSheetMasterTugas_(), baris);
}


// ============ JAM KERJA STANDAR ============

function getDaftarJamKerjaAdmin(password) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  var sheet = getOrBuatSheetJamKerja_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (!String(data[i][0] || '').trim()) continue;
    out.push({ baris: i + 1, bagianKerja: String(data[i][0] || ''), jamMasuk: formatJamSel_(data[i][1]), toleransi: data[i][2] });
  }
  return { sukses: true, daftar: out };
}

function tambahJamKerjaAdmin(password, bagianKerja, jamMasuk, toleransi) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  if (DAFTAR_BAGIAN_KERJA.indexOf(bagianKerja) === -1) return { sukses: false, pesan: 'Bagian kerja tidak valid.' };
  if (menitDariJam_(jamMasuk) === null) return { sukses: false, pesan: 'Format jam harus HH:mm, contoh 08:00.' };
  var sheet = getOrBuatSheetJamKerja_();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === bagianKerja) return { sukses: false, pesan: 'Bagian kerja "' + bagianKerja + '" sudah punya pengaturan jam kerja.' };
  }
  sheet.appendRow([bagianKerja, jamMasuk, parseInt(toleransi, 10) || 0]);
  return { sukses: true, pesan: 'Jam kerja untuk "' + bagianKerja + '" berhasil ditambahkan.' };
}

function editJamKerjaAdmin(password, baris, bagianKerja, jamMasuk, toleransi) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  if (DAFTAR_BAGIAN_KERJA.indexOf(bagianKerja) === -1) return { sukses: false, pesan: 'Bagian kerja tidak valid.' };
  if (menitDariJam_(jamMasuk) === null) return { sukses: false, pesan: 'Format jam harus HH:mm, contoh 08:00.' };
  var sheet = getOrBuatSheetJamKerja_();
  var baris_ = parseInt(baris, 10);
  if (isNaN(baris_) || baris_ < 2 || baris_ > sheet.getLastRow()) return { sukses: false, pesan: 'Baris tidak valid.' };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((i + 1) !== baris_ && String(data[i][0]).trim() === bagianKerja) {
      return { sukses: false, pesan: 'Bagian kerja "' + bagianKerja + '" sudah punya pengaturan jam kerja di baris lain.' };
    }
  }
  sheet.getRange(baris_, 1, 1, 3).setValues([[bagianKerja, jamMasuk, parseInt(toleransi, 10) || 0]]);
  return { sukses: true, pesan: 'Jam kerja untuk "' + bagianKerja + '" berhasil diperbarui.' };
}

function hapusJamKerjaAdmin(password, baris) {
  var cek = _cekAdminInternal_(password);
  if (!cek.sukses) return cek;
  return _hapusBarisSheet_(getOrBuatSheetJamKerja_(), baris);
}
