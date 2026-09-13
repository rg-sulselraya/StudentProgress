# Student Task Tracker

Aplikasi web + backend untuk monitoring pengumpulan tugas siswa berbasis QR/barcode. Sumber data utama dapat menggunakan Google Apps Script Web App sehingga tidak memerlukan Service Account JSON.

## Cara menjalankan

Jalankan perintah dari folder aplikasi (bukan dari `/Users/fa-13744`):

```bash
cd "/Users/fa-13744/Documents/Codex/2026-09-12/files-pasted-by-the-user-buatkan/outputs"
cp .env.example .env
npm install
npm start
```

Deploy `apps-script/Code.gs` sebagai Web App (Execute as: Me, Who has access: Anyone), lalu salin URL yang berakhiran `/exec` ke `.env`:

```env
GOOGLE_APPS_SCRIPT_URL=https://script.google.com/macros/s/DEPLOYMENT_ID/exec
```

Buka `http://127.0.0.1:4173`. Service Account tidak diperlukan. Uji Apps Script langsung dengan menambahkan `?action=test` pada URL deployment.

Untuk diagnosis koneksi tanpa frontend, buka `http://127.0.0.1:4173/api/students/test`. Endpoint ini menampilkan status koneksi, spreadsheet ID, nama tab, range, header, jumlah baris, dan maksimal tiga contoh siswa tanpa credential.

Student code dibuat stabil dari email siswa (atau kombinasi nama + kelas + sekolah bila email kosong), sehingga tidak bergantung pada nomor baris. QR Code hanya menyimpan student code.

## Sumber Google Sheet

Aplikasi menggunakan spreadsheet ID `1FcnkMsR24hU6A-bF5BYil2AX5iq-Ib38-Ta0dN_luBM`. Header wajib: `Nama Siswa`, `Email Siswa`, `Kelas`, `Tanggal Bayar`, `Hari belajar`, dan `Nama Sekolah`.

Jika database tidak dapat diakses, backend memberikan pesan setup yang jelas (URL Apps Script, deployment, ID spreadsheet, nama tab, atau credential legacy). Frontend tidak melakukan fallback ke daftar siswa hard-coded.

Apps Script menyediakan endpoint `?action=test` untuk diagnosis koneksi dan membaca range `siswa!A:F` dengan header dinamis. Backend mempertahankan nilai kolom `Kelas` lengkap setelah `trim()` dan membuat daftar kelas unik secara dinamis; kode rombel seperti `R4.01` tidak dipotong. Diagnostic membedakan `totalRows`/`rowsRead` (baris berisi data), `rangeRowsRead` (baris rentang teknis), `classCounts`, `classNames`, `excludedRows`, dan `studentsLoaded`.

Fitur utama:

- Dashboard target 2× per minggu dengan statistik dan daftar tindak lanjut.
- Scan kamera (jika browser mengizinkan akses kamera) atau input kode manual.
- Validasi jadwal personal, peringatan di luar jadwal, dan pencegahan scan dekat.
- Rekap per hari, riwayat scan, detail siswa, pengelolaan jadwal, export CSV, dan kartu QR.
