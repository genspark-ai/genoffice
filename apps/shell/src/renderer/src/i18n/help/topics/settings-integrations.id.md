# Pengaturan, bahasa, tema, dan integrasi MCP

## Membuka pengaturan

Baris akun di kiri bawah Beranda membuka panel pengaturan (bacaannya Masuk saat Anda belum masuk); opsi yang berkaitan dengan AI berada di bagian Model AI miliknya.

![A janela de configurações](img/settings-integrations.png) — a configuração dos modelos está no capítulo Model AI dan pengaturan.

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

GenOffice menyematkan **server MCP** lokal agar klien AI eksternal (Claude Desktop, Cursor, ...) dapat langsung membaca dan menulis dokumen Anda:

- Mulai: `genoffice mcp` di baris perintah (port dan token autentikasi dapat dikonfigurasi; hanya loopback secara bawaan).
- Kemampuan: membuat/membuka/mengedit docx, xlsx, dan pptx, membaca isi, mengonversi format, mengekspor PDF, dan lainnya — perangkat alat yang sama dengan yang dipakai aplikasi desktop.
- Keamanan: autentikasi token bersifat opsional tetapi disarankan; pendengar tetap di mesin lokal secara bawaan; lihat `genoffice mcp --help`.

## Lembar pintasan baris perintah

| Perintah           | Fungsinya                |
| ------------------ | ------------------------ |
| `genoffice <file>` | membuka sebuah file      |
| `genoffice mcp`    | memulai server MCP lokal |
| `genoffice --help` | semua perintah dan opsi  |
