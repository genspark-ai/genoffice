# Mulai cepat: antarmuka dan dasar-dasarnya

GenOffice adalah rangkaian perkantoran yang berjalan sepenuhnya di komputer Anda: satu jendela, satu deret tab, dengan enam editor — Docs (pemrosesan kata), Sheets (spreadsheet), Slides (presentasi), PDF, Markdown, dan HTML. File yang dibuka dan disimpan adalah .docx / .xlsx / .pptx / .pdf asli, sepenuhnya serasi dengan Word, Excel, dan PowerPoint. Tidak perlu jaringan.

## Sekilas antarmuka

![Layar Beranda](img/home-screen.png)

Jendela ini terdiri dari tiga bagian:

- **Bilah tab (atas)**: setiap file yang terbuka adalah sebuah tab. Tab Beranda paling kiri selalu ada dan tidak bisa ditutup; tab lainnya adalah dokumen Anda. Klik dua kali sebuah tab untuk mengganti nama file-nya secara langsung.
- **Area konten**: editor (atau Beranda) milik tab yang sedang aktif.
- **Bilah menu**: pada menu bar sistem di macOS, di bagian atas jendela pada Windows/Linux. Menu File/Edit/Tampilan berganti mengikuti editor yang aktif.

## Membuat dokumen

Pilih salah satu:

- Klik salah satu kartu pembuatan cepat di bagian [Mulai cepat](help://getting-started) pada **Beranda** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **File ▸ Baru**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML, atau PDF.
- Seret sebuah file ke jendela, atau klik dua kali di aplikasi file Anda (jika GenOffice adalah aplikasi default).

Dokumen baru terbuka tanpa nama; file di disk baru dibuat saat disimpan pertama kali.

## Membuka file

- Menu **File ▸ Buka…** (⌘O/ctrl+O) membuka pemilih file sistem: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Klik apa pun di daftar **Terbaru** pada Beranda.
- `genoffice <file>` dari terminal juga dapat membuka file.

## Cara kerja penyimpanan

- **Simpan manual**: ⌘S/ctrl+S, atau File ▸ Simpan / Simpan Sebagai. Penyimpanan pertama kali akan meminta lokasi dan nama.
- **Simpan Otomatis** baru aktif setelah Anda menyimpan file secara manual setidaknya sekali — PDF yang hanya Anda baca tidak akan pernah diam-diam ditulis ulang. Simpan Otomatis berjalan sedikit setelah konten berubah.
- Menutup tab yang belum disimpan akan lebih dulu meminta Simpan / Buang / Batal.
- Setiap penulisan bersifat atomik (file sementara + ganti nama), sehingga kalau listrik mati tidak akan tertinggal file yang hanya separuh.

## Pintasan yang sering dipakai

| Tindakan         | macOS | Windows / Linux |
| ---------------- | ----- | --------------- |
| Dokumen baru     | ⌘N    | ctrl+N          |
| Buka             | ⌘O    | ctrl+O          |
| Simpan           | ⌘S    | ctrl+S          |
| Tutup tab        | ⌘W    | ctrl+W          |
| Buka panduan ini | F1    | F1              |
| Ciutkan Pita     | ⌥⌘R   | Ctrl+F1         |

**Ciutkan Pita** berfungsi di setiap editor. Baris tab tetap ada dan pita perintah di bawahnya tersembunyi; tab yang sedang dipilih sekaligus menjadi kendali pencetakan, sehingga selama pita terlipat tidak ada tab yang dipilih dan menekan tab mana pun mengembalikan pita itu. Klik ganda pada tab melakukan hal yang sama. Cara Anda meninggalkannya diingat per editor.

Pintasan di dalam tiap editor (penyalin format, cari dan ganti, operasi tabel, ...) ada di bab masing-masing; Docs juga menyediakan dialog pintasan papan tikik yang bisa dicari (**⌘/**) (lihat babnya).

## Pintasan Option+Command

Option+Command adalah lapisan yang dicadangkan Word untuk lompatan terstruktur, dan GenOffice mengisinya dengan cara yang sama. Semua yang berikut ini milik Docs sendiri:

| Pintasan (macOS) | Fungsinya              | Windows / Linux    |
| ---------------- | ---------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3  | Judul 1 / 2 / 3        | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0              | Normal                 | Ctrl+Alt+0         |
| ⌥⌘M              | Paragraf               | Ctrl+Alt+M         |
| ⌥⌘A              | Komentar Baru          | Ctrl+Alt+A         |
| ⌥⌘F              | Sisipkan Catatan Kaki  | Ctrl+Alt+F         |
| ⌥⌘E              | Sisipkan Catatan Akhir | Ctrl+Alt+D         |
| ⌥⌘G              | Pergi ke               | Ctrl+G             |

Dua di antaranya berpindah di Windows, karena alasan yang sama seperti saat Word memisahkannya. **macOS memiliki ⌥⌘D** — ia menampilkan dan menyembunyikan Dock — sehingga catatan akhir memakai ⌥⌘E di Mac dan Ctrl+Alt+D di tempat lain. Dan **Pergi ke** melepas Alt: Ctrl+G, sedangkan pintasan di Mac tetap membawanya.

Dengan begitu ⌥⌘D tetap bebas untuk dipakai GenOffice di macOS, bila suatu perintah memutuskan membutuhkannya.

## Ke mana selanjutnya

- Letak file Anda: [Layar Beranda](help://home-screen).
- Mengelola banyak file terbuka: [Tab dan pengelolaan jendela](help://tabs-and-windows).
- Membiarkan AI mengerjakan pekerjaan: [Panel asisten AI](help://ai-panel).
- Bahasa, tema, aplikasi default: [Pengaturan, bahasa, tema, dan integrasi MCP](help://settings-integrations).
