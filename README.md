# Sistem Otomatisasi Rekonsiliasi Data Akuntansi dan Bank (BCA)

Sistem ini dirancang untuk memproses dan mencocokkan (rekonsiliasi) data transaksi dari sistem akuntansi internal dengan mutasi rekening bank BCA secara otomatis. Program ini menangani pembersihan data mentah, pemformatan nilai mata uang, hingga logika pencocokan saldo dan transaksi.

## Deskripsi Program
Program bekerja dengan mengambil dua sumber data utama yaitu laporan akuntansi (Acc.xls) dan mutasi bank (Bca.xlsx). Melalui serangkaian skrip pengolahan, sistem akan mengidentifikasi transaksi yang cocok (matched) dan transaksi yang tidak memiliki pasangan (unmatched) untuk memudahkan proses audit atau penyeimbangan kas.

## Struktur Folder dan File
Untuk menjalankan program ini, pastikan struktur direktori Anda adalah sebagai berikut:

- Jalankan Rekonsiliasi.py (Skrip utama untuk menjalankan seluruh alur)
- Acc.xls (File input data akuntansi) --> Kas & Bank --> Buku Bank --> Print --> Centang Fast Export
- Bca.xlsx (File input data mutasi bank) --> Letakkan file mutasi dengan awalan BCA ... (Misalnya BCA 5444 1 MEI 26, dst)
- Dapur/ (Folder sistem pendukung)
    - 1_AccCleaner.py (Pembersih data akuntansi)
    - 2_BcaCleaner.py (Pembersih data bank)
    - 3_ProcessingData.py (Logika rekonsiliasi data)
    - __init__.py

## Persyaratan Sistem
Pastikan Anda telah menginstal Python dan library pendukung berikut:
- pandas
- numpy
- openpyxl

Anda dapat menginstalnya menggunakan perintah:
pip install pandas numpy openpyxl

## Langkah-Langkah Penggunaan

1. Persiapan File: Letakkan file 'Acc.xls' dan 'Bca.xlsx' di direktori yang sama dengan 'Jalankan Rekonsiliasi.py'.
2. Eksekusi: Jalankan skrip utama dengan perintah:
   python "Jalankan Rekonsiliasi.py"
3. Proses: Program akan secara otomatis memindahkan file input ke folder 'Dapur', membersihkan data, dan menjalankan logika pencocokan.
4. Hasil: Setelah selesai, program akan menghasilkan file 'Hasil_Rekonsiliasi.xlsx' di direktori utama.

## Detail Alur Kerja Skrip

### 1. Jalankan Rekonsiliasi.py
Berperan sebagai orkestrator. Skrip ini memastikan semua file syarat tersedia, membersihkan folder kerja dari sisa proses sebelumnya, dan memanggil skrip pemrosesan secara berurutan.

### 2. 1_AccCleaner.py
Membaca file 'Acc.xls', mencari header laporan secara dinamis (Tanggal, No. Sumber, Keterangan, Penambahan, Pengurangan, Saldo), membersihkan data dari baris kosong, dan memformat kolom angka menjadi tipe data numerik yang rapi. Hasilnya disimpan sebagai 'Acc_temp.xlsx'.

### 3. 2_BcaCleaner.py
Membaca file 'Bca.xlsx', memproses kolom mutasi bank (Debet/Kredit), dan menyelaraskan format tanggal serta nilai uang. Skrip ini memastikan data bank siap untuk dibandingkan dengan data akuntansi. Hasilnya disimpan sebagai 'Bca_temp.xlsx'.

### 4. 3_ProcessingData.py
Ini adalah inti dari sistem rekonsiliasi. Skrip ini melakukan:
- Pembersihan mata uang dari simbol atau format string yang tidak konsisten.
- Pencocokan transaksi berdasarkan nilai (amount) dan rentang tanggal.
- Identifikasi transaksi yang menggantung (tidak ditemukan kecocokannya).
- Memberikan pewarnaan (highlight) pada file excel output: Hijau untuk data yang cocok dan Merah untuk data yang tidak cocok.

## Output Final
File 'Hasil_Rekonsiliasi.xlsx' akan berisi laporan lengkap yang memisahkan antara data yang sudah tervalidasi dan data yang memerlukan pemeriksaan lebih lanjut oleh tim finansial.
