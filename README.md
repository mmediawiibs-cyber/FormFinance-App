# Finance STIT App

Aplikasi web untuk mengelola formulir permintaan pembayaran, menyimpan data ke Sheet-A, membaca data dari Sheet-B, menampilkan print preview, dan melanjutkan proses print PDF.

## Fitur utama

- Form input data keuangan/PP
- Simpan data otomatis ke Sheet-A
- Baca dan tampilkan data dari Sheet-B
- Print preview form PP
- Cetak ke PDF via browser
- Siap di-deploy ke Vercel

## Struktur proyek

- index.html = aplikasi frontend utama
- apps-script.gs = backend Google Apps Script
- vercel.json = konfigurasi deploy Vercel

## Cara pakai

### 1. Siapkan spreadsheet Google

Buat spreadsheet baru di Google Sheets, lalu buat 2 sheet dengan nama:

- Sheet-A
- Sheet-B

Sheet-A dipakai untuk data input form. Sheet-B dipakai untuk data yang ditampilkan ke aplikasi.

### 2. Buat backend Apps Script

1. Buka Google Apps Script: https://script.google.com
2. Buat project baru
3. Paste isi file `apps-script.gs`
4. Ganti value `SPREADSHEET_ID` dengan ID spreadsheet Anda
5. Deploy sebagai Web App
6. Salin URL deployment hasil deploy

### 3. Hubungkan aplikasi frontend

Pada aplikasi web, masukkan URL Web App di pengaturan koneksi atau atur variabel `SCRIPT_URL` di file `index.html`.

Contoh:

```js
let SCRIPT_URL = "https://script.google.com/macros/s/XXXX/exec";
```

### 4. Deploy ke Vercel

1. Upload project ini ke GitHub / GitLab
2. Masuk ke Vercel
3. Import repository
4. Pilih root project
5. Deploy tanpa build command
6. Vercel akan otomatis mendeteksi file `index.html`

## Catatan penting

- Sheet-A adalah sumber data input dari form
- Sheet-B adalah data yang dibaca oleh aplikasi untuk preview dan print
- Untuk proses PDF, gunakan endpoint `?action=cetakPDF` dari Apps Script yang sudah di-deploy

## Tips keamanan

Jangan simpan token rahasia di frontend. Jika nanti ingin dipakai untuk kebutuhan multi-user, gunakan backend yang lebih aman untuk autentikasi.

## Lisensi

Untuk kebutuhan internal dan project kerja Anda.
