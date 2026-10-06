# Baris arahan dan ejen

Setiap pemasangan meliputi satu perintah `genoffice` yang memandu enjin yang sama seperti yang digunakan oleh tetingkap — penghurai yang sama, penulis yang sama, penjana yang sama. Fail yang disimpan aplikasi dan fail yang ditulis oleh perintah itu ialah fail yang sama, dan semakan yang melepasi panel AI aplikasi juga melepasi perintah itu.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

## Mendapatkan perintah itu

macOS dan Windows menghantarnya di dalam bungkus aplikasi. Untuk menggunakannya mengikut nama, jalankan `genoffice install-cli` sekali: ia membuat symlink bagi binari yang dibungkus itu ke dalam `/usr/local/bin`, atau ke dalam `PATH` pengguna anda pada Windows.

## Perintah yang perlu diketahui

| Perintah | Apa yang ia lakukan |
| --- | --- |
| `open` | Membuka dokumen dalam aplikasi; memulakan aplikasi jika ia belum berjalan. |
| `convert` | Menukar antara format menggunakan enjin aplikasi itu sendiri. |
| `create` | Mencipta dokumen daripada kandungan berstruktur. |
| `render` | Satu PNG bagi setiap halaman, mengikut susun atur penjana. |
| `pdf` | Membaca lapisan teks PDF halaman demi halaman, tanpa proses aplikasi. |
| `info` | Metadatan dan ringkasan struktur sesebuah dokumen. |
| `search` | Carian web atau imej melalui pembekal yang dikonfigurasi dalam aplikasi. |
| `image` / `media` | Menghasilkan imej, atau menghuraikan dan bertanya soalan tentang fail imej, video atau audio. |
| `merge` | Mengisi ruang letak `{{key}}` dalam templat `.docx`, `.pptx` atau `.xlsx`. |
| `capabilities` | Melaporkan keupayaan awan yang dikonfigurasi pada mesin ini. |
| `guide` | Rujukan op dan panduan reka bentuk, dijana daripada definisi yang sama seperti yang disahkan oleh pelaku — jadi ia tidak mungkin berbeza daripada apa yang diterima oleh `apply`. `--json` mengembalikannya bersama skema setiap op. |
| `install-cli` | Meletakkan `genoffice` pada `PATH`. |
| `skill` | Menyenaraikan ejen penulisan kod yang ditemui pada mesin ini, dan memasang atau mengemas kini kemahiran GenOffice di dalamnya. |
| `mcp` | Menyediakan setiap perintah sebagai alat Model Context Protocol. Lihat **Menyambungkan ejen penulisan kod**. |

## Penyuntingan: docs, sheets, slides

`genoffice docs`, `genoffice sheet` dan `genoffice slides` membaca dan menyunting melalui laluan tulis yang sama seperti yang digunakan oleh aplikasi, dan mereka berkongsi satu kosa kata: **op** ialah satu suntingan, dan **spec** ialah senarai op yang digunakan mengikut urutan.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` melaporkan apa yang akan dilakukan oleh satu kelompok tanpa menulis apa-apa, dan itulah cara termurah untuk menyemak sesuatu spec sebelum ia benar-benar digunakan. Panel AI dalam aplikasi berjalan tepat pada op yang sama ini, jadi apa sahaja yang boleh anda mintanya boleh juga anda tulis sebagai skrip.

## Model Context Protocol

`genoffice mcp` menyediakan setiap perintah sebagai alat MCP, dan `genoffice mcp install <agent|all>` mendaftarkannya dalam konfigurasi sendiri ejen penulisan kod. Lihat **Menyambungkan ejen penulisan kod** untuk butiran lanjut.

## Apa yang tidak dilakukan oleh perintah itu

Ia membaca dan menulis fail. Ia bukan aplikasi: tiada tetingkap, dan dialog kemas kini dalam aplikasi tidak terpakai. Apa-apa yang memerlukan tiningkap — panel AI, langkah kawalan kualiti ke atas slaid yang telah dipaparkan — mesti menunggu sehingga anda membuka fail itu. `genoffice render` memberi anda piksel tanpa itu, dan `genoffice slides` mengaudit susun atur sesebuah dek secara sendiri.
