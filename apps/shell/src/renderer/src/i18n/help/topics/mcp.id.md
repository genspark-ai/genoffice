# Menghubungkan agen pemrograman

GenOffice berbicara dalam Model Context Protocol, jadi agen pemrograman dapat membaca, menulis, dan merender dokumen Anda lewat mesin yang sama dengan yang dipakai aplikasi. Agen itu tidak sedang menebak-nebak format berkas: ia memperoleh skema op bertipe dari definisi yang sama dengan yang dipakai eksekutor untuk memvalidasi.

## Mendaftarkannya dari aplikasi

Inilah tempatnya: **Pengaturan ▸ Integrasi**. Panel ini punya dua bagian, dan Anda bisa memakai salah satunya atau keduanya.

**Skill-nya.** Satu baris per agen pemrograman yang ditemukan di komputer ini — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, masing-masing dengan **Pasang**, **Perbarui**, dan **Hapus**, ditambah **Pasang ke folder lain…**, **Unduh skill (zip)**, dan **Salin path**. Jika asisten Anda tidak ada di daftar, tunjukkan folder tempat ia membaca `SKILL.md` kepada GenOffice, atau simpan zip-nya dan biarkan asisten yang memasangnya. Skill dan MCP bisa berdampingan: asisten memilih salah satu, dan keduanya melakukan hal yang persis sama.

**MCP.** Ada dua jalur yang ditawarkan: **Dijalankan oleh asisten (disarankan)**, tempat Anda menambahkan konfigurasi yang ditampilkan ke klien dan asisten menjalankan servernya sendiri, serta **Server HTTP lokal**, yang dijalankan aplikasi untuk Anda. Cara mana pun, asisten akhirnya berbicara dengan GenOffice dan Anda tidak pernah mengetik satu perintah pun.

## Mendaftarkannya dari baris perintah

Hal yang sama dari sebuah terminal — inilah jalur mahir, dan yang sebaiknya dipakai ketika agen ada di tempat yang tidak bisa ditemukan panel:

```sh
genoffice mcp install all
```

Perintah ini menemukan agen pemrograman di komputer ini — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — dan menulis entri server stdio ke dalam konfigurasi masing-masing, meninggalkan sisa berkas itu persis seperti ditemukannya.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Agen yang Anda pasang di lokasi tidak biasa menerima `--dir <path>`; `--force` menulis ulang entri yang sudah ada di sana.

Semua yang diterima server ada dalam satu layar — bentuk untuk mendaftar, melepas, dan mencantumkan, menyajikan lewat HTTP, serta dua opsi skema:

![Keluaran nyata dari genoffice mcp --help: bentuk install, uninstall, dan list, beserta opsi --http, --host, --token, --compact-schemas, --dir, dan --force](img/mcp.png)

## Menjalankannya tanpa asisten

Untuk klien di komputer lain, sajikan lewat HTTP sebagai gantinya:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` menentukan ke mana ia mendengarkan. Berikan token yang sama kepada kliennya.

Berkas juga bisa lewat HTTP: `PUT /files/<name>` mengunggah satu berkas, setiap alat menerima URL `http(s)` menggantikan path, dan keluarannya kembali sebagai URL unduhan — serta, bila cukup kecil, sebagai resource yang ditanamkan. Op, spec, dan Markdown diteruskan inline baik ke sini maupun ke sana.

## Yang diterima agen

Setiap perintah adalah sebuah alat. Yang menarik:

- **`docs`, `sheet`, `slides`** — membaca dan menyunting berkas lewat jalur tulis milik aplikasi itu sendiri, satu **op** setiap kali. Deck baru berjalan `deck_start`, `deck_page`, `deck_build`.
- **`render`** — satu PNG per halaman, diletakkan oleh perender aplikasi, jadi agen bisa melihat sebuah slide alih-alih menebaknya.
- **`pdf`** — lapisan teks sebuah PDF, halaman demi halaman, tanpa proses aplikasi.
- **`info`** — metadata dan ringkasan struktur, yang biasanya merupakan panggilan pertama yang paling murah untuk berkas yang belum dikenal.
- **`search`, `image`, `media`** — penyedia yang dikonfigurasi di aplikasi, jadi agen tidak memerlukan kunci sendiri.
- **`merge`** — mengisi templat `{{key}}`.

## Skema, dan anggaran yang lebih kecil

`apply` dan `create` mengumumkan parameter `ops`, `cells`, dan `data` dengan skema bertipe per-op, yang dibangkitkan dari `genoffice guide <domain> --json`. Itu akurat, dan ukurannya tidak kecil. Klien dengan jendela konteks sempit bisa meminta larik biasa sebagai gantinya:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Mengapa perlu skill

Agen yang tidak mengenal kosakata op akan menebak. Skill itu membawa referensi dan panduan desain — materi yang sama dengan yang dicetak `genoffice guide` —, sehingga asisten menulis op yang spesifikasinya memang sudah dibacanya. Pasang dari panel di atas, atau dengan `genoffice skill` dari terminal.

## Apa yang bisa dicapainya di dalam aplikasi

Server itu tidak dibatasi pada berkas di disk. Selama GenOffice berjalan, agen juga bisa bekerja melalui jendela:

- **`open_in_genoffice`** membuka berkas di sebuah tab dan memfokuskannya.
- **`open_documents`** mencantumkan setiap dokumen yang Anda buka — id, jenis, path, dan apakah ada perubahan yang belum disimpan —, lalu membaca isi terkini salah satunya atau menutupnya, dengan menyimpan lebih dulu kecuali Anda menyuruhnya membuang.
- **Alat konten** mengambil id itu (atau path-nya) sebagai argumen `document`, sehingga hasil suntingan mendarat di tab yang memang sudah Anda buka, dan jendela beralih menampilkannya.

Dua hal tetap di luar jangkauan: tidak ada panel AI, dan dialog pembaruan di dalam aplikasi tidak berlaku.

## Panel Server HTTP lokal

Di bawah **Server HTTP lokal**, aplikasi menjalankan servernya sendiri alih-alih menyerahkannya kepada asisten: sebuah sakelar pengaktif, kolom **Port**, indikator **Berjalan / Tidak berjalan**, dan **Generasi latar belakang** (menulis dokumen langsung ke sebuah path tanpa membuka UI) dan **Contoh konfigurasi klien** untuk disalin. Membuka **Lanjutan** menambahkan dua URL koneksi — Streamable HTTP dan URL SSE yang lama —, URL **Pemeriksaan kesehatan**, dan sakelar **Pencatuman log** yang mencatat aktivitas server dan alat ke berkas lokal yang dapat Anda **Buka**, **Segarkan**, atau **Hapus** dari sana. Ia hanya mendengarkan di localhost.
