/**
 * Code.gs — Backend Aplikasi Input Tabungan Per Kelas
 * Spreadsheet: Tabungan-Per-Kelas-7A-9F (script TERIKAT ke file ini)
 *
 * Struktur tiap sheet kelas (7A–9F):
 *   Baris 1 : NO | NAMA | NIS | NOREK | KELAS | TANGGAL MENABUNG (gabungan F1:…)
 *   Baris 2 : tanggal-tanggal menabung (satu kolom = satu tanggal), mulai kolom F
 *   Baris 3+ : data murid; sel pada kolom tanggal = nominal setoran hari itu
 *
 * Aksi:
 *   GET  ?action=ping
 *   GET  ?action=list&kelas=7A&tanggal=2026-10-23
 *   POST JSON {action:"simpan", kelas:"7A", tanggal:"2026-10-23",
 *              items:[{nis:"3215", jumlah:5000}, ...]}
 */

var KOLOM_TANGGAL_MULAI = 6; // kolom F

function labelTanggal_(iso) {
  var bulan = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var p = String(iso || '').split('-');
  if (p.length !== 3) return String(iso || '');
  return parseInt(p[2], 10) + ' ' + bulan[parseInt(p[1], 10) - 1] + ' ' + p[0];
}

function sheetKelas_(kelas) {
  var nama = String(kelas || '').trim().toUpperCase();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
  if (!sh) throw new Error('Sheet kelas tidak ditemukan: ' + nama);
  return sh;
}

/** Cari kolom tanggal. Mengembalikan {kolom, daftar:[{kolom, label}]} — kolom = -1 bila belum ada. */
function petaTanggal_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), KOLOM_TANGGAL_MULAI);
  var baris2 = sh.getRange(2, KOLOM_TANGGAL_MULAI, 1, lastCol - KOLOM_TANGGAL_MULAI + 1).getValues()[0];
  var daftar = [];
  var kosongPertama = -1;
  for (var i = 0; i < baris2.length; i++) {
    var v = baris2[i];
    var label = '';
    if (v instanceof Date) {
      label = Utilities.formatDate(v, 'Asia/Jakarta', 'd MMM yyyy');
    } else {
      label = String(v == null ? '' : v).trim();
    }
    if (label) daftar.push({ kolom: KOLOM_TANGGAL_MULAI + i, label: label });
    else if (kosongPertama < 0) kosongPertama = KOLOM_TANGGAL_MULAI + i;
  }
  return { daftar: daftar, kosongPertama: kosongPertama, lastCol: lastCol };
}

function cariKolomTanggal_(sh, label) {
  var peta = petaTanggal_(sh);
  for (var i = 0; i < peta.daftar.length; i++) {
    if (peta.daftar[i].label === label) return peta.daftar[i].kolom;
  }
  return -1;
}

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (p.action === 'list') {
      var sh = sheetKelas_(p.kelas);
      var lastRow = sh.getLastRow();
      var data = sh.getRange(3, 1, Math.max(lastRow - 2, 0), 5).getValues();
      var label = p.tanggal ? labelTanggal_(p.tanggal) : '';
      var kolom = label ? cariKolomTanggal_(sh, label) : -1;
      var nilaiAda = {};
      if (kolom > 0 && lastRow >= 3) {
        var kolomNilai = sh.getRange(3, kolom, lastRow - 2, 1).getValues();
        for (var r = 0; r < data.length; r++) {
          var nisK = String(data[r][2] == null ? '' : data[r][2]).trim();
          var v = kolomNilai[r][0];
          if (nisK && v !== '' && v != null) nilaiAda[nisK] = v;
        }
      }
      var siswa = [];
      for (var i = 0; i < data.length; i++) {
        var nama = String(data[i][1] == null ? '' : data[i][1]).trim();
        var nis = String(data[i][2] == null ? '' : data[i][2]).trim();
        if (!nama) continue;
        siswa.push({ baris: i + 3, nis: nis, nama: nama, jumlahAda: nilaiAda[nis] || 0 });
      }
      var peta = petaTanggal_(sh);
      return json_({ ok: true, kelas: String(p.kelas).toUpperCase(), labelTanggal: label,
                     kolomTanggal: kolom, daftarTanggal: peta.daftar.map(function (d) { return d.label; }),
                     siswa: siswa });
    }
    return json_({ ok: true, pesan: 'Backend Input Tabungan Per Kelas aktif.' });
  } catch (err) {
    return json_({ ok: false, pesan: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    if (body.action !== 'simpan') return json_({ ok: false, pesan: 'Aksi tidak dikenal.' });
    var sh = sheetKelas_(body.kelas);
    var label = labelTanggal_(body.tanggal);
    if (!label) return json_({ ok: false, pesan: 'Tanggal tidak valid.' });

    // Petakan NIS -> nomor baris
    var lastRow = sh.getLastRow();
    var dataNIS = sh.getRange(3, 3, Math.max(lastRow - 2, 0), 1).getValues();
    var barisDariNIS = {};
    for (var r = 0; r < dataNIS.length; r++) {
      var nisK = String(dataNIS[r][0] == null ? '' : dataNIS[r][0]).trim();
      if (nisK) barisDariNIS[nisK] = r + 3;
    }

    // Tentukan kolom tanggal: pakai yang ada, slot kosong pertama, atau kolom baru
    var kolom = cariKolomTanggal_(sh, label);
    if (kolom < 0) {
      var peta = petaTanggal_(sh);
      if (peta.kosongPertama > 0) {
        kolom = peta.kosongPertama;
      } else {
        sh.insertColumnAfter(peta.lastCol);
        kolom = peta.lastCol + 1;
      }
      sh.getRange(2, kolom).setValue(label);
    }

    var hasil = [];
    (body.items || []).forEach(function (it) {
      var nis = String(it.nis == null ? '' : it.nis).trim();
      var jumlah = Number(it.jumlah);
      var baris = barisDariNIS[nis];
      if (!baris) { hasil.push({ nis: nis, ok: false, pesan: 'NIS tidak ditemukan di sheet ' + body.kelas }); return; }
      if (!(jumlah > 0)) { hasil.push({ nis: nis, ok: false, pesan: 'Nominal tidak valid' }); return; }
      sh.getRange(baris, kolom).setValue(jumlah);
      hasil.push({ nis: nis, ok: true, pesan: 'Tersimpan' });
    });
    return json_({ ok: true, kelas: String(body.kelas).toUpperCase(), labelTanggal: label, kolom: kolom, hasil: hasil });
  } catch (err) {
    return json_({ ok: false, pesan: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
