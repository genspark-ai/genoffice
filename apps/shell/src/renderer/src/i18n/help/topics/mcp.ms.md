# Menyambungkan ejen penulisan kod

GenOffice bercakap dengan Model Context Protocol, jadi ejen penulisan kod boleh membaca, menulis dan memaparkan dokumen anda melalui enjin yang sama seperti yang digunakan oleh aplikasi. Ejen itu tidak meneka format fail: ia mendapat skema op bertip daripada definisi yang sama seperti yang disahkan oleh pelaku.

## Mendaftarkannya dari aplikasi

Tempatnya ialah **Tetapan ▸ Integrasi**. Panel itu mempunyai dua bahagian, dan anda boleh menggunakan satu sahaja atau kedua-duanya.

**Skill.** Satu baris bagi setiap ejen penulisan kod yang ditemui pada mesin ini — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf —, setiap satu dengan **Pasang**, **Kemas kini** dan **Nyahpasang**, tambahan lagi **Pasang ke folder lain…**, **Muat turun skill (zip)** dan **Salin laluan**. Jika pembantu anda tidak tersenarai, tunjukkan GenOffice folder tempat ia membaca `SKILL.md`, atau simpan fail zip itu dan biarkan pembantu memasangnya. Skill dan MCP boleh bersebelahan: pembantu akan memilih salah satu, dan keduanya melakukan perkara yang sama sekali.

**MCP.** Dua laluan ditawarkan: **Dimulakan oleh pembantu (disyorkan)**, di mana anda menambah konfigurasi yang dipaparkan pada klien anda dan pembantu memulakan pelayan itu sendiri, serta **Pelayan HTTP tempatan**, yang dijalankan aplikasi untuk anda. Bagaimanapun, pembantu akhirnya bercakap dengan GenOffice dan anda tidak sekali pun perlu menaip arahan.

## Mendaftarkannya dari baris arahan

Perkara yang sama daripada satu terminal — inilah laluan lanjutan, dan laluan yang perlu digunakan apabila ejen berada di tempat yang tidak dapat dicari oleh panel:

```sh
genoffice mcp install all
```

Ia menemui ejen penulisan kod pada mesin ini — Claude Code, Codex, Cursor, Gemini CLI, Copilot CLI, OpenCode, Windsurf — lalu menulis entri pelayan stdio ke dalam konfigurasi sendiri setiap satu, dan meninggalkan bahagian lain fail itu seperti yang dijumpainya.

```sh
genoffice mcp list             # where each agent stands
genoffice mcp install cursor    # just one
genoffice mcp uninstall cursor  # take it back out
```

Ejen yang anda pasang di tempat yang luar biasa menggunakan `--dir <path>`; `--force` menulis semula entri yang sudah ada di situ.

Segala yang diterima oleh pelayan itu muat pada satu skrin — borang untuk memasang, mengeluarkan dan menyenaraikan, menyajikan melalui HTTP, dan dua pilihan skema:

![Keluaran sebenar genoffice mcp --help: borang install, uninstall dan list, bersama pilihan --http, --host, --token, --compact-schemas, --dir dan --force](img/mcp.png)

## Menjalankannya tanpa pembantu

Bagi klien pada mesin lain, layankannya melalui HTTP:

```sh
genoffice mcp --http 8765 --token <secret>
```

`--host <addr>` menukar tempat ia mendengarkan. Berikan token yang sama kepada klien.

Fail bergerak melalui HTTP juga: `PUT /files/<name>` memuat naik satu fail, setiap alat menerima URL `http(s)` menggantikan laluan, dan output kembali sebagai URL muat turun — dan, apabila cukup kecil, sebagai sumber terbenam. Op, spec dan Markdown dihantar secara langsung dalam kedua-dua kes.

## Apa yang diterima oleh ejen

Setiap perintah ialah sebuah alat. Yang menarik:

- **`docs`, `sheet`, `slides`** — membaca dan menyunting fail melalui laluan tulis aplikasi itu sendiri, satu **op** pada satu masa. Dek baharu bermula dengan `deck_start`, `deck_page`, `deck_build`.
- **`render`** — satu PNG bagi setiap halaman, disusun oleh penjana aplikasi, jadi ejen boleh melihat slaid itu dan bukan meneka.
- **`pdf`** — lapisan teks PDF, halaman demi halaman, tanpa proses aplikasi.
- **`info`** — metadatan dan ringkasan struktur, yang biasanya merupakan panggilan pertama yang termurah bagi fail yang tidak dikenali.
- **`search`, `image`, `media`** — pembekal yang dikonfigurasi dalam aplikasi, jadi ejen tidak perlukan kunci sendiri.
- **`merge`** — mengisi templat `{{key}}`.

## Skema, dan bajet yang lebih kecil

`apply` dan `create` membentangkan parameter `ops`, `cells` dan `data` mereka bersama skema bertip bagi setiap op, yang dijana daripada `genoffice guide <domain> --json`. Itu tepat, tetapi tidak kecil. Klien dengan tetingkap konteks yang sempit boleh meminta tatasusunan biasa sebagai gantinya:

```sh
genoffice mcp --compact-schemas
GENOFFICE_MCP_COMPACT_SCHEMAS=1 genoffice mcp
```

## Mengapa perlu skill

Ejen yang tidak tahu kosa kata op akan meneka. Skill itu membawa rujukan dan panduan reka bentuk — bahan yang sama seperti yang dicetak oleh `genoffice guide` —, jadi pembantu menulis op yang spesifikasinya memang telah dibaca. Pasang ia daripada panel di atas, atau dengan `genoffice skill` daripada terminal.

## Apa yang boleh dicapainya dalam aplikasi

Pelayan itu tidak terhad kepada fail pada cakera. Selagi GenOffice berjalan, ejen juga boleh bekerja melalui tetingkap:

- **`open_in_genoffice`** membuka fail dalam satu tab dan memberikan tumpu kepadanya.
- **`open_documents`** menyenaraikan setiap dokumen yang anda buka — id, jenis, laluan, dan sama ada ia mempunyai perubahan yang belum disimpan —, kemudian membaca kandungan terkini sesuatu dokumen atau menutupnya, dengan menyimpan dahulu melainkan anda menyuruhnya membuang.
- **Alat kandungan** mengambil id itu (atau laluannya) sebagai argumen `document`, jadi satu suntingan mendarat pada tab yang memang sudah anda buka, dan tetingkap bertukar untuk mempaparkannya.

Dua perkara kekal di luar jangkauan: tiada panel AI, dan dialog kemas kini dalam aplikasi tidak terpakai.

## Panel Pelayan HTTP tempatan

Di bawah **Pelayan HTTP tempatan**, aplikasi menjalankan pelayan itu sendiri dan bukannya menyerahkankannya kepada pembantu: satu suis mendayakan, medan **Port**, penunjuk **Sedang berjalan / Tidak berjalan** dan **Penjanaan latar belakang** (menulis dokumen terus ke laluan tanpa membuka UI), serta **Contoh tetapan klien** untuk disalin. Membuka **Lanjutan** menambah dua URL sambungan — Streamable HTTP dan URL SSE yang lama —, URL **Semakan kesihatan** dan suis **Pencatuman log** yang mencatat aktiviti pelayan dan alat ke dalam fail setempat yang boleh anda **Buka**, **Segarkan** atau **Kosongkan** dari situ. Ia hanya mendengarkan pada localhost.
