# Input Tabungan Per Kelas — SMPN 4 Majenang

Aplikasi web untuk menginput setoran tabungan murid ke spreadsheet
**Tabungan-Per-Kelas-7A-9F** (18 sheet: 7A–9F).

- Pilih kelas dan tanggal → daftar murid tampil → isi nominal (satu-satu
  atau sekaligus untuk nominal yang sama) → **Simpan Setoran**.
- Setoran masuk ke sheet kelas yang benar, baris nama murid (cocok NIS),
  dan kolom tanggalnya; kolom tanggal baru dibuat otomatis dan bertambah
  terus tanpa batas.
- Nilai yang sudah tersimpan tampil terisi dan dapat diedit ulang pada
  tanggal yang sama.

Berkas:
- `index.html` — aplikasi (satu berkas, URL Web App sudah terpasang).
- `Code.gs` — backend Apps Script yang terikat ke spreadsheet.
- `PANDUAN.txt` — panduan pakai dan pemasangan.
