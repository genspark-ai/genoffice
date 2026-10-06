# Pengaturan, bahasa, tema, dan integrasi MCP

## Membuka pengaturan

Baris akun di kiri bawah Beranda membuka panel pengaturan (bacaannya Masuk saat Anda belum masuk). Panel ini punya enam bagian: Akun, Model AI, Media & Pencarian AI, Umum, Integrasi, dan Tentang.

![Pengaturan ▸ Umum, tempat bahasa, tema, Autosave, dan sakelar statistik penggunaan berada](img/settings-general.png)

Konfigurasi model punya bab tersendiri; di **Media & Pencarian AI** Anda menyalakan pembuatan gambar, analisis gambar, analisis video, pencarian web, dan pencarian file lokal untuk tiap penyedia.

## Bahasa

- Pengaturan menawarkan **21 bahasa antarmuka**: Inggris, Tionghoa Sederhana, Jepang, Korea, Prancis, Jerman, Spanyol, Thai, Indonesia, Rusia, Arab, Portugis, Italia, Polandia, Ceko, Belanda, Melayu, Ibrani, Hindi, Tionghoa Tradisional, Vietnam.
- Beralih bahasa berlaku seketika dan bertahan; bilah menu sistem dibangun ulang mengikuti bahasa tersebut.

## Tema

Terang / Gelap / Ikuti Sistem. Mode Ikuti Sistem mengikuti tampilan sistem operasi, dan para editor mengganti tema senada tanpa berkedip.

## Pengikatan aplikasi predefinida

Pengaturan dapat mendaftarkan GenOffice sebagai aplikasi penangan .docx / .xlsx / .pptx / .pdf dan sejenisnya (pendaftaran aplikasi predefinida tingkat sistem; konfirmasi saat diminta).

## Pemberitahuan pihak ketiga dan pembaruan

- Bantuan ▸ Pemberitahuan Perangkat Lunak Pihak Ketiga: inventaris lengkap lisensi OSS yang disertakan dengan aplikasi.
- Bantuan ▸ Periksa Pembaruan…: memicu pemeriksaan manual; versi yang lebih baru menawarkan dipasang.

## Masuk dengan Genspark

- Jalur masuk (di pengaturan atau di daftar proyek cloud) memakai alur **device-code**: GenOffice menampilkan kode dan membuka login di browser; prosesnya dilanjutkan otomatis setelah selesai.
- Login hanya dipakai untuk: daftar proyek cloud dan model terkelola Genspark. Tanpa login, semua fitur lokal dan model kustom tetap berfungsi.
- Keluar cukup satu klik di pengaturan.

## Integrasi MCP (untuk pengguna lanjutan / klien AI)

**Integrasi** adalah panel yang menghubungkan GenOffice dengan agen pemrograman, dan ia punya artikel tersendiri: Menghubungkan agen pemrograman. Versi singkat — pilih jalurnya (baris perintah, atau MCP), ikuti bagian itu, lalu mulai obrolan baru dan bertanya.

![Pengaturan ▸ Integrasi: tiga langkah, lalu baris skill dan opsi MCP](img/settings-integrations.png)

Di bawah **Server HTTP lokal**, aplikasi juga bisa menjalankan server itu sendiri — sakelar aktif dan port — dan **Lanjutan** menambahkan URL pemeriksaan kesehatan serta file log, alih-alih menghandalkannya kepada asisten. Ia hanya mendengarkan di localhost.

## Lembar pintasan baris perintah

| Perintah           | Fungsinya                |
| ------------------ | ------------------------ |
| `genoffice <file>` | membuka sebuah file      |
| `genoffice mcp`    | memulai server MCP lokal |
| `genoffice --help` | semua perintah dan opsi  |
