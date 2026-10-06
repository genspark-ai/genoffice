# Menghubungkan agen pemrograman

GenOffice berbicara dalam Model Context Protocol, jadi agen pemrograman dapat membaca, menulis, dan merender dokumen Anda lewat mesin yang sama dengan yang dipakai aplikasi. Agen itu tidak sedang menebak-nebak format berkas: ia memperoleh skema op bertipe dari definisi yang sama dengan yang dipakai eksekutor untuk memvalidasi.

## Mendaftarkannya

Kasus yang umum hanyalah satu perintah:

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

## Menjalankannya sendiri

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

## Skill-nya

Agen yang tidak mengenal kosakata op akan menebak. `genoffice skill` memasang skill GenOffice ke dalam agen-agen yang ditemukannya, lengkap dengan referensi dan panduan desain — materi yang sama dengan yang dicetak `genoffice guide`.

## Apa yang bukan

Server MCP itu membaca dan menulis berkas. Ia bukan jendela: tidak ada panel AI, dan dialog pembaruan di dalam aplikasi tidak berlaku. Jika sebuah langkah butuh jendela, bukalah berkasnya.
