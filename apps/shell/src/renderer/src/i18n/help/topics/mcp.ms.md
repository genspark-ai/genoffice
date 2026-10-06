# Menyambungkan ejen penulisan kod

GenOffice bercakap dengan Model Context Protocol, jadi ejen penulisan kod boleh membaca, menulis dan memaparkan dokumen anda melalui enjin yang sama seperti yang digunakan oleh aplikasi. Ejen itu tidak meneka format fail: ia mendapat skema op bertip daripada definisi yang sama seperti yang disahkan oleh pelaku.

## Mendaftarkannya

Kes yang biasa ialah satu perintah:

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

## Menjalankannya sendiri

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

## Kemahiran

Ejen yang tidak tahu kosa kata op akan meneka. `genoffice skill` memasang kemahiran GenOffice ke dalam ejen yang ditemuinya, serta membawa rujukan dan panduan reka bentuk — bahan yang sama seperti yang dicetak oleh `genoffice guide`.

## Apa yang ia bukan

Pelayan MCP membaca dan menulis fail. Ia bukan tetingkap: tiada panel AI, dan dialog kemas kini dalam aplikasi tidak terpakai. Jika sesuatu langkah memerlukan tetingkap, bukalah fail itu.
