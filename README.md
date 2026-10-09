# bot-blasting-cawi-ASN

Bot Playwright untuk FASIH-SM: mengubah mode assignment **OPEN** menjadi **CAWI** dan mengirim tautan survei ke responden via email/WhatsApp, berdasarkan daftar kontak di Excel.

Alur per assignment: cari Nama Keluarga → cek status OPEN → titik 3 → Ganti Mode → CAWI → isi email & nomor → nyalakan *Kirim Pesan* → *Ubah Mode Pendataan*.

## Persiapan

```bash
pip install -r requirements.txt
copy config.example.json config.json
```

Isi `config.json`:
- `data_url`: URL halaman **Data** survei di FASIH.
- `xlsx_path`: path Excel kontak, relatif terhadap folder ini.
- `sheet`: nama sheet.

Kolom Excel yang dibaca: `No`, `Assignment ID`, `Nama Keluarga`, `Alamat`, `Email`, `No. WhatsApp`, `Keterangan`, `Keberadaan Keluarga di Fasih-SM`.

Aturan pemilihan baris:
- Baris tanpa email/nomor dilewati.
- Baris *Ditemukan/Baru* dilewati.
- Baris dengan Keterangan PENSIUN/TIDAK DITEMUKAN dilewati.
- Kalau satu keluarga punya lebih dari 1 ASN (`NAMA: nilai; ...`), dipakai kontak kepala keluarga, atau ASN pertama kalau kepala keluarga bukan ASN.

## Menjalankan

1. Jalankan `buka_chrome.bat`, lalu login SSO secara manual. Bot akan menumpang ke Chrome ini lewat port 9222.
2. Coba dulu tanpa mengirim apa pun. Mode default adalah **dry-run**: bot berhenti sebelum klik final dan menyimpan screenshot.
   ```bash
   python broadcast_cawi.py --limit 5
   ```
3. Kirim sungguhan:
   ```bash
   python broadcast_cawi.py --execute
   ```

| Opsi | Fungsi |
|---|---|
| `--execute` | Benar-benar klik *Ubah Mode Pendataan* |
| `--limit N` | Proses maksimal N assignment |
| `--prepare-only` | Hanya membuat `daftar_target.csv` untuk dicek |
| `--recon` | Dry-run 1 assignment, simpan HTML & screenshot tiap langkah (untuk debug) |
| `--port P` | Port Chrome (default 9222) |
| `--part K/N` | Kerjakan bagian ke-K dari N (paralel banyak window) |
| `--exclude-part K/N` | Lewati target milik bagian K/N yang sedang berjalan di window lain |

**Paralel 2 window.** Setiap window memakai Chrome sendiri dan login sendiri:
```bash
buka_chrome.bat 9222
```
```bash
python broadcast_cawi.py --execute --part 1/2 --port 9222
```
Di terminal lain:
```bash
buka_chrome.bat 9223
```
```bash
python broadcast_cawi.py --execute --part 2/2 --port 9223
```

**Berhenti dengan aman.** Buat file bernama `STOP` di folder ini. Semua bot berhenti setelah assignment yang sedang dikerjakan selesai. Hapus file itu sebelum menjalankan bot lagi.

**Sesi habis.** Bot berbunyi beep dan menunggu Anda login ulang (maksimal 30 menit).

## Log & rekap

- Setiap assignment dicatat ke `logs/log_execute*.csv` (waktu, status, pesan).
- Saat dijalankan ulang, bot melewati semua yang sudah `SUCCESS` di log mana pun, jadi tidak ada kirim dobel.
- Yang gagal otomatis dicoba lagi.

| Status | Arti |
|---|---|
| `SUCCESS` / `DRYRUN_OK` | Berhasil / dry-run berhasil |
| `SKIP_TIDAK_DITEMUKAN` | Nama tidak ditemukan, walau halaman sudah dimuat ulang |
| `SKIP_BUKAN_OPEN` | Status assignment bukan OPEN (tidak dicoba lagi) |
| `SKIP_AMBIGU` | Lebih dari 1 assignment OPEN dengan nama sama |
| `ERROR*` | Gagal teknis (timeout, sesi, dll.), dicoba lagi saat run berikutnya |

```bash
python rekap.py
```
Membuat `rekap_broadcast.xlsx` dengan sheet Ringkasan, Perlu Manual, Berhasil, dan Belum Diproses.

```bash
python catat_manual.py "NAMA KELUARGA"
```
Mengecek apakah assignment sudah dikirim. Tambahkan `--catat` setelah memproses manual, supaya bot melewatinya.

## Privasi

`.gitignore` memakai pola *allowlist*: hanya file kode dan dokumentasi yang bisa masuk repo. Data responden, log, screenshot, rekap, `config.json`, dan profil Chrome (berisi cookie login) tidak pernah ikut ter-commit.
