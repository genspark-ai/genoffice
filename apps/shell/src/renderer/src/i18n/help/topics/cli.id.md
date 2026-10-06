# Perintah baris perintah dan agen

Setiap pemasangan membawa perintah `genoffice` yang menjalankan mesin yang sama dengan jendela — parser yang sama, penulis yang sama, perender yang sama. Berkas yang disimpan aplikasi dan berkas yang ditulis perintah itu adalah berkas yang sama, dan pemeriksaan yang lolos di panel AI aplikasi juga lolos di perintah.

```sh
genoffice --help          # every command
genoffice guide slides    # the op reference, for writing your own
```

Itulah seluruh permukaannya dalam satu layar — setiap perintah dengan satu baris yang menjelaskan apa yang dilalukannya, lalu opsi global dan kode keluar:

![Keluaran nyata dari genoffice --help: spanduk versi, setiap perintah dengan deskripsi satu baris, serta opsi global dan kode keluar](img/cli.png)

## Mendapatkan perintahnya

macOS dan Windows menyertakannya di dalam bundel aplikasi. Untuk memakainya dengan namanya, jalankan `genoffice install-cli` sekali: perintah ini membuat symbolic link dari biner yang disertakan ke `/usr/local/bin`, atau ke `PATH` milik pengguna Anda di Windows.

## Perintah yang perlu diketahui

| Perintah          | Yang dikerjakan                                                                                                                                                                                                                                          |
| ----------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `open`            | Membuka dokumen di aplikasi; memulai aplikasi bila belum berjalan.                                                                                                                                                                                       |
| `selection`       | Apa yang dipilih pengguna di berkas itu sekarang juga, selagi aplikasi berjalan — penunjuk untuk "yang ini" / "di sini". Mengembalikan rentang blok, rentang sel di lembar, elemen slide, atau satu halaman, tergantung editor tempat berkas itu berada. |
| `convert`         | Mengonversi antarformat memakai mesin milik aplikasi itu sendiri.                                                                                                                                                                                        |
| `create`          | Membuat dokumen dari konten terstruktur.                                                                                                                                                                                                                 |
| `render`          | Satu PNG per halaman, sebagaimana diletakkan oleh perender.                                                                                                                                                                                              |
| `pdf`             | Membaca lapisan teks PDF halaman demi halaman, tanpa proses aplikasi.                                                                                                                                                                                    |
| `info`            | Metadata dan ringkasan struktur sebuah dokumen.                                                                                                                                                                                                          |
| `search`          | Pencarian web atau gambar lewat penyedia yang dikonfigurasi di aplikasi.                                                                                                                                                                                 |
| `image` / `media` | Menghasilkan gambar, atau mendeskripsikan dan mengajukan pertanyaan tentang berkas gambar, video, atau audio.                                                                                                                                            |
| `merge`           | Mengisi placeholder `{{key}}` pada templat `.docx`, `.pptx`, atau `.xlsx`.                                                                                                                                                                               |
| `capabilities`    | Melaporkan fitur awan mana yang sudah dikonfigurasi di komputer ini.                                                                                                                                                                                     |
| `guide`           | Referensi op dan panduan desain, dibangkitkan dari definisi yang sama dengan yang dipakai eksekutor untuk memvalidasi — jadi tidak mungkin melenceng dari apa yang diterima `apply`. `--json` mengembalikannya beserta skema tiap op.                    |
| `install-cli`     | Menaruh `genoffice` di `PATH`.                                                                                                                                                                                                                           |
| `skill`           | Mendaftarkan agen pemrograman yang ditemukan di komputer ini, lalu memasang atau memperbarui skill GenOffice di dalamnya.                                                                                                                                |
| `mcp`             | Menyajikan semua perintah sebagai alat Model Context Protocol. Lihat **Menghubungkan agen pemrograman**.                                                                                                                                                 |

## Menyunting: dokumen, lembar, presentasi

`genoffice docs`, `genoffice sheet`, dan `genoffice slides` membaca dan menulis lewat jalur tulis yang sama dengan yang dipakai aplikasi, dan mereka berbagi satu kosakata: sebuah **op** adalah satu suntingan, dan sebuah **spec** adalah daftar op yang diterapkan berurutan.

```sh
genoffice sheet read report.xlsx
genoffice slides apply deck.pptx --ops ops.json --dry-run
genoffice docs apply brief.docx --ops ops.json
```

`--dry-run` melaporkan apa yang akan dilakukan sebuah batch tanpa menulis apa pun, dan itulah cara murah untuk memeriksa sebuah spec sebelum dijalankan. Panel AI di aplikasi berjalan persis pada op-op ini, jadi apa pun yang bisa Anda minta kepadanya bisa Anda skripkan.

## Model Context Protocol

`genoffice mcp` menyajikan setiap perintah sebagai alat MCP, dan `genoffice mcp install <agent|all>` mendaftarkannya di konfigurasi milik agen pemrograman itu sendiri. Lihat **Menghubungkan agen pemrograman** untuk sisi tersebut.

## Yang tidak dilakukan perintah

Perintah itu membaca dan menulis berkas. Ia bukan aplikasinya: tidak ada jendela, dan dialog pembaruan di dalam aplikasi tidak berlaku. Apa pun yang membutuhkan jendela — panel AI, pemeriksaan mutu atas satu slide hasil render — harus menunggu Anda membuka berkasnya. `genoffice render` memberi Anda pikselnya tanpa jendela, dan `genoffice slides` menata letak sebuah deck secara mandiri.
