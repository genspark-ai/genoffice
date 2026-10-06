# Pintasan papan ketik

Setiap editor punya pengikatan tombol sendiri, dan hanya sebagian yang berlaku di mana saja. Halaman ini adalah daftar lengkapnya, dikelompokkan menurut aplikasi, dengan pintasan macOS dan Windows / Linux berdampingan.

## Baca ini dulu

**Hampir tidak ada yang di sini bersifat global.** Tiga tombol akses di bawah ini memang global; sisanya milik aplikasi yang sedang Anda hadapi, dan berubah saat Anda mengganti tab.

| Tindakan                                                         | macOS | Windows / Linux |
| ---------------------------------------------------------------- | ----- | --------------- |
| Buka Panduan Pengguna — bisa dipakai di mana saja dalam aplikasi | F1    | F1              |
| Ciutkan / perluas pita — setiap editor                           | ⌥⌘R   | Ctrl+F1         |
| Buka lembar Pintasan Papan Ketik — hanya Word                    | ⌘/    | Ctrl+/          |

- **F1 membuka manual ini**, bukan lembar pintasan. Tombol itu milik menu aplikasi, jadi ia berlaku dari Word, Sheets, Slides, PDF, maupun tab Beranda.
- **Tombol pita adalah `Ctrl+F1`, atau `⌥⌘R` di macOS.** Menekan `F1` polos tidak melakukan apa pun — F1 adalah manualnya, dan tombol pita sengaja tidak berbagi kunci dengan F1. `Ctrl+F1` juga diterima di macOS, meski `⌥⌘R` adalah tombol yang dipakai Office for Mac.
- **Lembar Pintasan Papan Ketik hanya milik Word** (`⌘/`). Isinya persis tabel Word di bawah ini. Tidak ada editor lain yang memilikinya.
- **`⌘O`, `⌘S`, dan `⌘P` diimplementasikan terpisah di balik tiap editor**, dan perilakunya tidak sepenuhnya sama. `⌘P` milik Word menyusun pratinjau cetak lebih dulu sehingga satu lembar cetak sama dengan satu halaman di layar; milik PDF membuka dialog cetak PDF; Sheets dan Slides punya versinya sendiri lagi.
- **Selebihnya bersifat per-aplikasi.** Saat tab Word aktif, `⌘B` adalah Tebal milik Word. Saat Slides aktif, `⌘B` adalah Tebal milik Slides, dan menu tempatnya tinggal bukan menu yang biasa Anda pakai. Wherever sebuah pengikatan hanya aktif dengan ada sesuatu yang terpilih, atau di panel tertentu, kolom Tindakan mengatakannya.

## Konvensi

- Kolom macOS memakai simbol yang familier: `⌘` Command, `⌥` Option, `⇧` Shift, `⌃` Control.
- Kolom Windows / Linux mengeja nama tombolnya: `Ctrl`, `Alt`, `Shift`.
- Kalau GenOffice memang butuh tombol berbeda di dua platform, tabelnya menunjukkannya. Kasus yang umum adalah `⌃H` untuk Temukan & Ganti (macOS memakai `⌘H` untuk Sembunyikan), `⌥⌘G` untuk Ke, dan tombol tutup milik Slides.
- Pintasan yang hanya berlaku saat ada seleksi ditulis di kolom Tindakan, bukan diberi kolom sendiri, supaya tabelnya tetap lebar tiga kolom.

## Word

Word punya bilah menu penuh, jadi pintasan berkas dan format di bawah ini berlaku kapan pun sebuah tab Word punya fokus, dokumen apa pun yang terbuka. Pintasan penyuntingan membutuhkan dokumen yang sudah termuat dan bisa disunting, dengan caret di dalam teks; pintasan itu tidak aktif saat Anda mengetik di kotak Temukan, kotak perintah AI, atau kolom isian lain apa pun.

### Berkas

| Tindakan        | macOS | Windows / Linux |
| --------------- | ----- | --------------- |
| Baru            | ⌘N    | Ctrl+N          |
| Jendela Baru    | ⇧⌘N   | Ctrl+Shift+N    |
| Buka…           | ⌘O    | Ctrl+O          |
| Simpan          | ⌘S    | Ctrl+S          |
| Simpan Sebagai… | ⇧⌘S   | Ctrl+Shift+S    |
| Cetak           | ⌘P    | Ctrl+P          |
| Tutup (tab)     | ⌘W    | Ctrl+W          |

`⌘P` bukan perintah cetak peramban. Word menyusun pratinjau cetak lebih dulu, sehingga satu lembar cetak sama persis dengan satu halaman di layar, lalu membuka dialog cetiknya sendiri.

### Edit

| Tindakan                  | macOS | Windows / Linux |
| ------------------------- | ----- | --------------- |
| Urungkan                  | ⌘Z    | Ctrl+Z          |
| Ulangi                    | ⇧⌘Z   | Ctrl+Shift+Z    |
| Potong                    | ⌘X    | Ctrl+X          |
| Salin                     | ⌘C    | Ctrl+C          |
| Tempel                    | ⌘V    | Ctrl+V          |
| Tempel sebagai Teks Biasa | ⌥⇧⌘V  | Ctrl+Shift+V    |
| Pilih Semua               | ⌘A    | Ctrl+A          |
| Temukan                   | ⌘F    | Ctrl+F          |
| Ganti                     | ⌃H    | Ctrl+H          |
| Ke                        | ⌥⌘G   | Ctrl+G          |
| Pintasan Papan Ketik      | ⌘/    | Ctrl+/          |

Ganti memakai Ctrl+H di kedua platform karena `⌘H` menyembunyikan jendela di macOS dan tidak pernah sampai ke Word. Ke mengikuti pembagian milik Word sendiri: `⌥⌘G` di Mac supaya `⇧⌘G` tetap kosong untuk Jumlah Kata.

### Format Teks

| Tindakan                             | macOS    | Windows / Linux       |
| ------------------------------------ | -------- | --------------------- |
| Tebal                                | ⌘B       | Ctrl+B                |
| Miring                               | ⌘I       | Ctrl+I                |
| Garis bawah                          | ⌘U       | Ctrl+U                |
| Font… (dialog)                       | ⌘D       | Ctrl+D                |
| Perbesar Ukuran Font                 | ⇧⌘.      | Ctrl+Shift+.          |
| Perkecil Ukuran Font                 | ⇧⌘,      | Ctrl+Shift+,          |
| Perbesar Ukuran Font 1 pt            | ⌘]       | Ctrl+]                |
| Perkecil Ukuran Font 1 pt            | ⌘[       | Ctrl+[                |
| Superskrip                           | ⇧⌘= / ⌘. | Ctrl+Shift+= / Ctrl+. |
| Subskrip                             | ⌘,       | Ctrl+,                |
| Ubah Huruf Besar-Kecil               | ⇧F3      | Shift+F3              |
| Salin Format (dengan teks terpilih)  | ⇧⌘C      | Ctrl+Shift+C          |
| Tempel Format (setelah Salin Format) | ⇧⌘V      | Ctrl+Alt+Shift+V      |
| Hapus Semua Format                   | ⌃␣       | Ctrl+Space            |

Dua pintasan penyalin format adalah satu-satunya tempat tombol Windows bukan sekadar nama ulang: Tempel Format memakai Alt karena di sana Ctrl+Shift+V sudah menempel sekaligus mencocokkan gaya.

### Format Paragraf

| Tindakan                                    | macOS | Windows / Linux |
| ------------------------------------------- | ----- | --------------- |
| Rata Kiri                                   | ⌘L    | Ctrl+L          |
| Rata Tengah                                 | ⌘E    | Ctrl+E          |
| Rata Kanan                                  | ⌘R    | Ctrl+R          |
| Rata Penuh                                  | ⌘J    | Ctrl+J          |
| Jarak Baris 1.0                             | ⌘1    | Ctrl+1          |
| Jarak Baris 1.5                             | ⌘5    | Ctrl+5          |
| Jarak Baris 2.0                             | ⌘2    | Ctrl+2          |
| Tambah Indentasi                            | ⌃M    | Ctrl+M          |
| Kurangi Indentasi                           | ⌃⇧M   | Ctrl+Shift+M    |
| Indentasi Menggantung                       | ⌘T    | Ctrl+T          |
| Hapus Indentasi Menggantung                 | ⇧⌘T   | Ctrl+Shift+T    |
| Hapus Format Paragraf                       | ⌃Q    | Ctrl+Q          |
| Normal (gaya)                               | ⌥⌘0   | Ctrl+Alt+0      |
| Judul 1 (gaya)                              | ⌥⌘1   | Ctrl+Alt+1      |
| Judul 2 (gaya)                              | ⌥⌘2   | Ctrl+Alt+2      |
| Judul 3 (gaya)                              | ⌥⌘3   | Ctrl+Alt+3      |
| Paragraf… (dialog)                          | ⌥⌘M   | Ctrl+Alt+M      |
| Pindahkan Paragraf ke Atas                  | ⌥⇧↑   | Alt+Shift+↑     |
| Pindahkan Paragraf ke Bawah                 | ⌥⇧↓   | Alt+Shift+↓     |
| Turunkan Tingkat Daftar Item (di awal item) | Tab   | Tab             |
| Naikkan Tingkat Daftar Item (di awal item)  | ⇧Tab  | Shift+Tab       |

Indentasi, indentasi menggantung, dan Hapus Format Paragraf hanya memakai Ctrl di kedua platform: `⌘M` memperkecil jendela dan `⌘Q` keluar, jadi pintasan Command akan menjadi suntingan diam-diam di tengah jalan keluar. Tab hanya mengubah tingkat daftar ketika caret berada di awal sebuah item atau membentang di beberapa item; di dalam teks ia menyisipkan karakter tab, dan `⇧Tab` selalu menaikkan tingkat.

### Sisipkan

| Tindakan                       | macOS | Windows / Linux  |
| ------------------------------ | ----- | ---------------- |
| Pemisah Halaman                | ⌘⏎    | Ctrl+Enter       |
| Pemisah Kolom                  | ⇧⌘⏎   | Ctrl+Shift+Enter |
| Pemisah Baris                  | ⇧⏎    | Shift+Enter      |
| Tautan…                        | ⌘K    | Ctrl+K           |
| Komentar Baru                  | ⌥⌘A   | Ctrl+Alt+A       |
| Sisipkan Catatan Kaki          | ⌥⌘F   | Ctrl+Alt+F       |
| Sisipkan Catatan Akhir         | ⌥⌘E   | Ctrl+Alt+D       |
| Tanggal (field)                | ⌥⇧D   | Alt+Shift+D      |
| Waktu (field)                  | ⌥⇧T   | Alt+Shift+T      |
| Spasi Tanpa Pemenggalan        | ⇧⌘␣   | Ctrl+Shift+Space |
| Tanda Hubung Tanpa Pemenggalan | ⇧⌘-   | Ctrl+Shift+-     |

Sisipkan Catatan Akhir adalah pembagian platform yang sungguhan, bukan sekadar nama ulang: Word memakai `⌥⌘E` di Mac karena `⌥⌘D` milik Dock di sana.

### Tinjau & Alat

| Tindakan           | macOS | Windows / Linux |
| ------------------ | ----- | --------------- |
| Lacak Perubahan    | ⇧⌘E   | Ctrl+Shift+E    |
| Jumlah Kata        | ⇧⌘G   | Ctrl+Shift+G    |
| Proofread          | F7    | F7              |
| Perbarui Field     | F9    | F9              |
| Alihkan Kode Field | ⌥F9   | Alt+F9          |

F7 menjalankan proofread AI dan F9 memperbarui field; keduanya butuh dokumen yang sudah termuat. Begitu pembatasan penyuntingan menyalakan Lacak Perubahan, itu tidak bisa dimatikan lagi.

### Tampilan

| Tindakan                           | macOS | Windows / Linux |
| ---------------------------------- | ----- | --------------- |
| Perbesar                           | ⌘=    | Ctrl+=          |
| Perkecil                           | ⌘-    | Ctrl+-          |
| Perbesar ke 100%                   | ⌘0    | Ctrl+0          |
| Tampilkan/sembunyikan tanda format | ⌘8    | Ctrl+Shift+8    |

Pintasan tanda format mengikuti pembagian milik Word sendiri, `⌘8` melawan Ctrl+Shift+8. Word mengambil pembesaran hanya lewat tiga pintasan menu di atas — `⌘+` tidak melakukan apa pun di sini.

F1 membuka **Panduan Pengguna**, bukan daftar ini. Lembar pintasannya adalah `⌘/` atau Ctrl+/, dan itu sebuah item menu Word, jadi hanya muncul selama tab Word aktif.

## Sheets

GenOffice Sheets adalah kisi yang kompatibel dengan Excel. Sebagian besar pintasan di bawah ini adalah pintasan Excel yang sudah Anda kenal.

> **Penyuntingan menang.** Hampir semua pintasan di Sheets diabaikan selama Anda menyunting sel atau mengetik di **Bilah Rumus**, supaya mengetik, memindahkan caret, dan Backspace tetap bermakna penyuntingan teks biasa. Kalau sebuah tombol tidak melakukan apa pun, periksa apakah Anda sedang di tengah menyunting: tekan Esc atau Enter dulu, lalu coba lagi. Perintah yang bekerja pada seleksi juga tidak melakukan apa pun selama dialog terbuka. Di mana sebuah pintasan menjadi pengecualian terhadap aturan ini, tabel menyatakannya.

Dua konvensi berikut bersifat khusus untuk Sheets. `⌘` dan `⌃` adalah slot modifier yang sama — keduanya diterima di setiap platform, jadi `⌘1` dan `Ctrl+1` adalah pintasan yang sama. **Sel Windows / Linux yang kosong berarti pintasan itu hanya ada di macOS.**

### Navigasi

| Tindakan                                               | macOS      | Windows / Linux |
| ------------------------------------------------------ | ---------- | --------------- |
| Lembar kerja berikutnya                                | ⌘PgDn      | Ctrl+PgDn       |
| Lembar kerja sebelumnya                                | ⌘PgUp      | Ctrl+PgUp       |
| Lembar kerja berikutnya (setara Excel untuk Mac)       | ⌥→         |                 |
| Lembar kerja sebelumnya (setara Excel untuk Mac)       | ⌥←         |                 |
| Lompat ke sel pertama yang tidak dibekukan             | ⌘Home      | Ctrl+Home       |
| Lompat ke sel terakhir yang terpakai (isi atau format) | ⌘End       | Ctrl+End        |
| Dialog Ke                                              | ⌘G, F5     | Ctrl+G, F5      |
| Sisipkan lembar kerja baru                             | ⇧F11       | Shift+F11       |
| Pindahkan sel aktif satu layar ke bawah (dan gulir)    | Page Down  | Page Down       |
| Pindahkan sel aktif satu layar ke atas (dan gulir)     | Page Up    | Page Up         |
| Pindahkan sel aktif satu layar ke kanan (dan gulir)    | ⌥Page Down | Alt+Page Down   |
| Pindahkan sel aktif satu layar ke kiri (dan gulir)     | ⌥Page Up   | Alt+Page Up     |

### Seleksi

| Tindakan                                                               | macOS           | Windows / Linux                                     |
| ---------------------------------------------------------------------- | --------------- | --------------------------------------------------- |
| Pilih seluruh kolom sel aktif                                          | ⌃Space          | Ctrl+Space                                          |
| Pilih seluruh baris sel aktif                                          | ⇧Space          | Shift+Space                                         |
| Lompat ke tepi blok data — salah satu dari empat panah                 | ⌘↑ ⌘↓ ⌘← ⌘→     | Ctrl+↑ Ctrl+↓ Ctrl+← Ctrl+→                         |
| Perluas seleksi sampai tepi blok data; tekan lagi untuk menyusutkannya | ⌘⇧↑ ⌘⇧↓ ⌘⇧← ⌘⇧→ | Ctrl+Shift+↑ Ctrl+Shift+↓ Ctrl+Shift+← Ctrl+Shift+→ |
| Lompat ke kolom pertama baris sekarang                                 | Home            | Home                                                |

### Penyuntingan

| Tindakan                                                                                        | macOS                     | Windows / Linux       |
| ----------------------------------------------------------------------------------------------- | ------------------------- | --------------------- |
| Isi seluruh seleksi dengan isian sekarang                                                       | ⌃⏎ atau ⌘⏎                | Ctrl+Enter            |
| Mulai baris baru di dalam sel                                                                   | ⌥⏎, atau ⌃⌥⏎ / ⌘⌥⏎ di Mac | Alt+Enter             |
| Sunting sel aktif di tempat                                                                     | F2                        | F2                    |
| Tebal                                                                                           | ⌘B                        | Ctrl+B                |
| Miring                                                                                          | ⌘I                        | Ctrl+I                |
| Garis bawah                                                                                     | ⌘U                        | Ctrl+U                |
| Temukan                                                                                         | ⌘F                        | Ctrl+F                |
| Isi ke kanan dari kolom terpilih paling kiri                                                    | ⌘R                        | Ctrl+R                |
| Berganti rujukan rumus antara relatif dan absolut                                               | F4                        | F4                    |
| Berganti rujukan rumus antara relatif dan absolut (setara Excel untuk Mac, saat menyunting sel) | ⌘T                        |                       |
| Salin                                                                                           | ⌘C                        | Ctrl+C                |
| Potong                                                                                          | ⌘X                        | Ctrl+X                |
| Tempel                                                                                          | ⌘V                        | Ctrl+V                |
| Tempel nilai saja                                                                               | ⌘⇧V                       | Ctrl+Shift+V          |
| Isi Cepat dari pola di atasnya                                                                  | ⌘E                        | Ctrl+E                |
| Urungkan                                                                                        | ⌘Z                        | Ctrl+Z                |
| Ulangi                                                                                          | ⇧⌘Z                       | Ctrl+Shift+Z          |
| Pilih seluruh lembar kerja                                                                      | ⌘A                        | Ctrl+A                |
| Sisipkan tanggal hari ini                                                                       | ⌘;                        | Ctrl+;                |
| Sisipkan waktu saat ini                                                                         | ⌘⇧;                       | Ctrl+Shift+;          |
| Kosongkan isi setiap sel terpilih                                                               | ⌫ atau ⌦                  | Backspace atau Delete |

JumlahOtomatis hanya ada di macOS. Tekan ⌘⇧T.

### Format

| Tindakan                                                      | macOS | Windows / Linux |
| ------------------------------------------------------------- | ----- | --------------- |
| Dialog Format Sel                                             | ⌘1    | Ctrl+1          |
| Coret (dengan sel terpilih)                                   | ⌘5    | Ctrl+5          |
| Perbesar ukuran font (dengan sel terpilih)                    | ⌘⇧.   | Ctrl+Shift+.    |
| Perkecil ukuran font (dengan sel terpilih)                    | ⌘⇧,   | Ctrl+Shift+,    |
| Format angka: Umum                                            | ⌘⇧`   | Ctrl+Shift+`    |
| Format angka: angka dengan dua desimal                        | ⌘⇧1   | Ctrl+Shift+1    |
| Format angka: waktu `h:mm AM/PM`                              | ⌘⇧2   | Ctrl+Shift+2    |
| Format angka: tanggal `d-mmm-yy`                              | ⌘⇧3   | Ctrl+Shift+3    |
| Format angka: mata uang                                       | ⌘⇧4   | Ctrl+Shift+4    |
| Format angka: persen                                          | ⌘⇧5   | Ctrl+Shift+5    |
| Format angka: ilmiah                                          | ⌘⇧6   | Ctrl+Shift+6    |
| Semua garis batas (dengan sel terpilih)                       | ⌘⇧7   | Ctrl+Shift+7    |
| Tanpa garis batas (dengan sel terpilih)                       | ⌘⇧-   | Ctrl+Shift+-    |
| Batalkan Penyalin Format (hanya saat Penyalin Format menyala) | Esc   | Esc             |

### Sisipkan & hapus

| Tindakan                                                                                    | macOS                                    | Windows / Linux                             |
| ------------------------------------------------------------------------------------------- | ---------------------------------------- | ------------------------------------------- |
| Sisipkan sel — baris atau kolom penuh langsung tersisip, selain itu membuka dialog Sisipkan | ⌘⇧=                                      | Ctrl+Shift+=                                |
| Sisipkan sel — tombol `+` pada keypad numerik, dialog yang sama                             | ⌘ bersama tombol `+` pada keypad numerik | Ctrl bersama tombol `+` pada keypad numerik |
| Hapus sel — baris atau kolom penuh langsung terhapus, selain itu membuka dialog Hapus       | ⌘-                                       | Ctrl+-                                      |
| Hapus sel — tombol `-` pada keypad numerik, dialog yang sama                                | ⌘ bersama tombol `-` pada keypad numerik | Ctrl bersama tombol `-` pada keypad numerik |
| Sembunyikan baris terpilih                                                                  | ⌘9                                       | Ctrl+9                                      |
| Tampilkan lagi baris tersembunyi dalam seleksi                                              | ⌘⇧9                                      | Ctrl+Shift+9                                |
| Sembunyikan kolom terpilih                                                                  | ⌘0                                       | Ctrl+0                                      |
| Tampilkan lagi kolom tersembunyi dalam seleksi                                              | ⌘⇧0                                      | Ctrl+Shift+0                                |

### Data & tinjau

| Tindakan                                   | macOS | Windows / Linux |
| ------------------------------------------ | ----- | --------------- |
| Dialog Sisipkan Fungsi                     | ⇧F3   | Shift+F3        |
| Manajer Nama                               | ⌃F3   | Ctrl+F3         |
| Tambah atau sunting catatan pada sel aktif | ⇧F2   | Shift+F2        |
| Sisipkan hipertautan                       | ⌘K    | Ctrl+K          |
| Lacak presedan (sel harus berisi rumus)    | ⌘[    | Ctrl+[          |
| Lacak dependen                             | ⌘]    | Ctrl+]          |
| Kelompokkan baris terpilih                 | ⌥⇧→   | Alt+Shift+→     |
| Keluarkan dari kelompok baris terpilih     | ⌥⇧←   | Alt+Shift+←     |
| Hitung ulang seluruh buku kerja            | F9    | F9              |
| Hitung ulang hanya lembar kerja aktif      | ⇧F9   | Shift+F9        |
| Alihkan filter pada rentang terpilih       | ⌘⇧F   |                 |

Pintasan filter milik Excel, Ctrl+Shift+L, tidak melakukan apa pun di Windows maupun Linux. Tekan ⌘⇧F di macOS, atau pakai perintah filter pada pita di platform mana pun.

### Jendela

| Tindakan                                                          | macOS            | Windows / Linux     |
| ----------------------------------------------------------------- | ---------------- | ------------------- |
| Simpan buku kerja                                                 | ⌘S               | Ctrl+S              |
| Simpan Sebagai                                                    | ⇧⌘S              | Ctrl+Shift+S        |
| Buka buku kerja                                                   | ⌘O               | Ctrl+O              |
| Cetak                                                             | ⌘P               | Ctrl+P              |
| Tutup buku kerja (macOS) / Keluar dari aplikasi (Windows & Linux) | ⌘W               | Ctrl+Q              |
| Buka Panduan Pengguna                                             | F1               | F1                  |
| Perbesar                                                          | ⌘=               | Ctrl+=              |
| Perbesar dengan roda tetikus                                      | ⌘ + roda tetikus | Ctrl + roda tetikus |
| Tampilkan atau sembunyikan rumus di dalam sel                     | ⌘`               | Ctrl+`              |
| Perluas atau ciutkan Bilah Rumus                                  | ⌘⇧U              | Ctrl+Shift+U        |

## Slides

GenOffice Slides adalah editor presentasi bergaya PowerPoint. Sebagian besar tombol melakukan hal yang sama di mana pun Anda berada, tetapi beberapa berubah makna mengikuti fokus — di dalam kotak teks, atau ketika **rel thumbnail**, **kerangka garis besar**, atau **penyusun slide** punya fokus papan tik. Kasus itu ditandai di kolom Tindakan.

### Berkas dan aplikasi

| Tindakan                     | macOS | Windows / Linux |
| ---------------------------- | ----- | --------------- |
| Simpan                       | ⌘S    | Ctrl+S          |
| Simpan sebagai               | ⇧⌘S   | Ctrl+Shift+S    |
| Buka                         | ⌘O    | Ctrl+O          |
| Tutup tab sekarang           | ⌘W    | Ctrl+Q          |
| Cetak                        | ⌘P    | Ctrl+P          |
| Urungkan                     | ⌘Z    | Ctrl+Z          |
| Ulangi                       | ⇧⌘Z   | Ctrl+Shift+Z    |
| Ulangi (pintasan alternatif) | ⌘Y    | Ctrl+Y          |

Menutup tab adalah satu-satunya tempat kedua platform benar-benar terpisah: ⌘W di macOS, sedangkan di Windows dan Linux Ctrl+Q, yang keluar dari seluruh aplikasi, bukan hanya tabnya.

### Penyuntingan

| Tindakan                                           | macOS  | Windows / Linux |
| -------------------------------------------------- | ------ | --------------- |
| Slide baru                                         | ⇧⌘N    | Ctrl+M          |
| Temukan (bukan saat menyunting teks dalam bentuk)  | ⌘F     | Ctrl+F          |
| Anotasi seleksi dengan AI (dengan bentuk terpilih) | ⌘K     | Ctrl+K          |
| Perbesar                                           | ⌘=     | Ctrl+=          |
| Perkecil                                           | ⌘-     | Ctrl+-          |
| Ukuran asli (100%)                                 | ⌘0     | Ctrl+0          |
| Pilih semua                                        | ⌘A     | Ctrl+A          |
| Berganti seleksi bentuk maju / mundur              | ⇥ / ⇧⇥ | Tab / Shift+Tab |

Slide baru adalah pembagian platform yang lain. Tiap platform punya tombolnya sendiri dan tombol platform lain tidak melakukan apa pun: ⇧⌘N di macOS, Ctrl+M di Windows dan Linux. `⌘M` tidak bisa mengerjakan itu di macOS karena itulah tombol Perkecil milik sistem.

Pilih semua mengikuti fokus: dengan rel thumbnail, kerangka garis besar, atau penyusun slide yang fokus, ia memilih setiap slide, dan di tempat lain ia memilih setiap elemen pada slide sekarang. Tab berganti antarbentuk di slide itu, dan hanya ketika kanvas slide itu sendiri yang fokus.

### Format teks

Ini berlaku pada teks bentuk yang terpilih, atau pada paragraf tempat caret berada ketika Anda di dalam kotak teks.

| Tindakan                                     | macOS | Windows / Linux |
| -------------------------------------------- | ----- | --------------- |
| Tebal                                        | ⌘B    | Ctrl+B          |
| Miring                                       | ⌘I    | Ctrl+I          |
| Garis bawah                                  | ⌘U    | Ctrl+U          |
| Rata kiri                                    | ⌘L    | Ctrl+L          |
| Rata tengah                                  | ⌘E    | Ctrl+E          |
| Rata kanan                                   | ⌘R    | Ctrl+R          |
| Rata penuh                                   | ⌘J    | Ctrl+J          |
| Perbesar ukuran font sebesar 1 pt            | ⌘]    | Ctrl+]          |
| Perkecil ukuran font sebesar 1 pt            | ⌘[    | Ctrl+[          |
| Besarkan font ke ukuran berikutnya di tangga | ⇧⌘>   | Ctrl+Shift+>    |
| Perkecil font ke ukuran sebelumnya di tangga | ⇧⌘<   | Ctrl+Shift+<    |

Tombol tangga font itu juga menerima tombol . dan , dengan Shift ditahan, untuk tata letak papan tik yang melaporkan keduanya di tempat tombol tanda baca.

### Papan klip dan objek

| Tindakan                                                 | macOS       | Windows / Linux    |
| -------------------------------------------------------- | ----------- | ------------------ |
| Salin                                                    | ⌘C          | Ctrl+C             |
| Potong                                                   | ⌘X          | Ctrl+X             |
| Tempel                                                   | ⌘V          | Ctrl+V             |
| Tempel format saja                                       | ⇧⌘V         | Ctrl+Shift+V       |
| Salin format saja (dengan bentuk terpilih)               | ⇧⌘C         | Ctrl+Shift+C       |
| Duplikat di tempat (dengan bentuk terpilih)              | ⌘D          | Ctrl+D             |
| Kelompokkan seleksi (dengan bentuk terpilih)             | ⌘G          | Ctrl+G             |
| Keluarkan seleksi dari kelompok (dengan bentuk terpilih) | ⇧⌘G         | Ctrl+Shift+G       |
| Hapus seleksi (dengan bentuk terpilih)                   | ⌫ / Delete  | Backspace / Delete |
| Hapus titik yang terpilih (di Edit Titik)                | ⌫ / Delete  | Backspace / Delete |
| Geser seleksi (dengan bentuk terpilih)                   | ↑ ↓ ← →     | Arrow keys         |
| Geser satu piksel layar                                  | ⌘↑ ⌘↓ ⌘← ⌘→ | Ctrl+Arrow keys    |
| Ubah ukuran terhadap titik tengah                        | ⇧↑ ⇧↓ ⇧← ⇧→ | Shift+Arrow keys   |

Di dalam kotak teks, salin, potong, dan tempel bekerja pada teks, bukan pada bentuknya. Shift bersama tombol panah mengubah ukuran tiap bentuk terhadap titik tengahnya sendiri, tidak menyentuh konektor, dan menggabungkan seleksi banyak menjadi satu langkah urungkan.

### Slides — rel thumbnail, kerangka garis besar, atau penyusun slide yang fokus

Perintah tingkat slide bergaya PowerPoint hanya aktif ketika salah satu dari tiga kerangka itu punya fokus papan tik.

| Tindakan                                    | macOS                          | Windows / Linux                                |
| ------------------------------------------- | ------------------------------ | ---------------------------------------------- |
| Slide sebelumnya / berikutnya               | ↑ ↓ (also Page Up / Page Down) | ArrowUp / ArrowDown (also Page Up / Page Down) |
| Slide pertama / terakhir                    | Home / End                     | Home / End                                     |
| Perluas seleksi ke slide pertama / terakhir | ⇧Home / ⇧End                   | Shift+Home / Shift+End                         |
| Pilih semua slide                           | ⌘A                             | Ctrl+A                                         |
| Hapus slide terpilih                        | ⌫ / Delete                     | Backspace / Delete                             |
| Salin slide                                 | ⌘C                             | Ctrl+C                                         |
| Potong slide                                | ⌘X                             | Ctrl+X                                         |

Berpindah antar slide dengan tombol panah adalah pengecualian — itu juga berlaku dari kanvas, tanpa ada yang terpilih. Menghapus slide tidak berlaku ketika alat pena aktif atau di **Tampilan Baca**.

### Di dalam kotak teks

| Tindakan                                                   | macOS                              | Windows / Linux                        |
| ---------------------------------------------------------- | ---------------------------------- | -------------------------------------- |
| Masuk ke penyuntingan teks pada bentuk teks terpilih       | karakter cetak apa pun, ⏎, atau F2 | karakter cetak apa pun, Enter, atau F2 |
| Kunci teks dan pilih ulang bentuknya                       | Esc atau F2                        | Esc atau F2                            |
| Kunci teks dan pilih ulang bentuknya (pintasan alternatif) | ⌘⏎                                 | Ctrl+⏎                                 |
| Lompat ke placeholder berikutnya                           | ⌃⏎                                 | Ctrl+⏎                                 |
| Turunkan / naikkan tingkat daftar                          | ⇥ / ⇧⇥                             | Tab / Shift+Tab                        |
| Sel tabel berikutnya / sebelumnya                          | ⇥ / ⇧⇥                             | Tab / Shift+Tab                        |
| Pemenggalan baris lunak                                    | ⇧⏎                                 | Shift+⏎                                |

Tab berarti dua hal berbeda: sel berikutnya di dalam tabel, dan menurunkan tingkat daftar di dalam bentuk teks.

### Peragaan Slide

| Tindakan                    | macOS             | Windows / Linux                      |
| --------------------------- | ----------------- | ------------------------------------ |
| Mulai dari awal             | F5                | F5                                   |
| Mulai dari slide sekarang   | ⇧F5 atau ⌘⏎       | ⇧F5                                  |
| Slide berikutnya            | → ↓ ␣ ⏎ Page Down | Right, Down, Space, Enter, Page Down |
| Slide sebelumnya            | ← ↑ Page Up       | Left, Up, Page Up                    |
| Slide pertama               | Home              | Home                                 |
| Slide terakhir              | End               | End                                  |
| Ke slide N                  | ketik N, lalu ⏎   | ketik N, lalu Enter                  |
| Layar hitam                 | B atau .          | B atau .                             |
| Layar putih                 | W atau ,          | W atau ,                             |
| Pulihkan setelah layar mati | tombol apa saja   | tombol apa saja                      |
| Akhiri peragaan             | Esc               | Esc                                  |

Mulai dari slide sekarang memakai ⌘⏎ sekaligus ⇧F5 di macOS. Di Windows tombol yang sama tetap disimpan untuk berpindah antar placeholder.

### Tampilan Baca

| Tindakan                  | macOS             | Windows / Linux                      |
| ------------------------- | ----------------- | ------------------------------------ |
| Slide berikutnya          | → ↓ ␣ ⏎ Page Down | Right, Down, Space, Enter, Page Down |
| Slide sebelumnya          | ← ↑ Page Up       | Left, Up, Page Up                    |
| Keluar dari tampilan baca | Esc               | Esc                                  |

Tekan **Esc** untuk mundur keluar dari apa pun yang sedang Anda masuki — sebuah potongan, bentuk yang sedang Anda gambar, **Edit Titik**, sebuah kelompok, pena, stabilo, atau penghapus, pemutar media, atau **Tampilan Baca** — dan untuk membersihkan seleksi.

Dua yang lain adalah modifier, bukan pintasan: tahan **Shift** saat menggambar untuk mempertahankan proporsi bentuk baru, dan tahan **Shift** saat memutar untuk menempel pada langkah 15°.

## PDF

PDF tidak punya menu sendiri, jadi setiap pintasan di bawah ini mengikuti tab aktif dan hanya berlaku ketika sebuah tab PDF berada di depan.

### Berkas dan Penyuntingan

| Tindakan                 | macOS | Windows / Linux |
| ------------------------ | ----- | --------------- |
| Simpan                   | ⌘S    | Ctrl+S          |
| Urungkan                 | ⌘Z    | Ctrl+Z          |
| Ulangi                   | ⇧⌘Z   | Ctrl+Shift+Z    |
| Temukan (buka pencarian) | ⌘F    | Ctrl+F          |
| Cetak                    | ⌘P    | Ctrl+P          |

Urungkan dan Ulangi melangkah melalui riwayat dokumen; ketika kursor berada di kolom isian, keduanya meninggalkan `⌘Z` untuk kolom itu. `⌘P` membuka dialog cetak dalam aplikasi yang sama dengan tombol Cetak di toolbar.

### Tampilan

| Tindakan            | macOS            | Windows / Linux     |
| ------------------- | ---------------- | ------------------- |
| Perbesar            | ⌘=               | Ctrl+=              |
| Perkecil            | ⌘-               | Ctrl+-              |
| Sesuaikan lebar     | ⌘0               | Ctrl+0              |
| Perbesar (bertahap) | ⌘ + roda tetikus | Ctrl + roda tetikus |

Perbesar juga menerima ⌘+ dan Ctrl++.

### Navigasi Halaman

| Tindakan                   | macOS               | Windows / Linux     |
| -------------------------- | ------------------- | ------------------- |
| Halaman berikutnya         | →                   | →                   |
| Halaman sebelumnya         | ←                   | ←                   |
| Gulir satu layar ke bawah  | PageDown atau Space | PageDown atau Space |
| Gulir satu layar ke atas   | PageUp              | PageUp              |
| Gulir ke bawah             | ↓                   | ↓                   |
| Gulir ke atas              | ↑                   | ↑                   |
| Lompat ke halaman pertama  | Home                | Home                |
| Lompat ke halaman terakhir | End                 | End                 |

Semuanya tanpa modifier, jadi kedua platform membacanya sama. ↓ dan ↑ adalah satu-satunya baris yang berubah mengikuti fokus: di dalam dokumen keduanya menggeser halaman, tetapi dengan fokus di sidebar thumbnail keduanya melangkah satu halaman penuh. Keduanya tidak melakukan apa pun selama Anda mengetik di kolom isian.

### Seleksi dan Penyuntingan

| Tindakan               | macOS                 | Windows / Linux       |
| ---------------------- | --------------------- | --------------------- |
| Hapus anotasi terpilih | Delete atau Backspace | Delete atau Backspace |

Hanya aktif ketika sebuah anotasi, stempel, redaksi, atau gambar terpilih.

### Escape dan Dialog

| Tindakan                       | macOS  | Windows / Linux |
| ------------------------------ | ------ | --------------- |
| Batalkan lapisan teratas       | Escape | Escape          |
| Konfirmasi dialog yang terbuka | Enter  | Enter           |

Escape membongkar satu lapisan setiap kali — lebih dulu dialog yang terbuka, lalu munculan AI, draf teks, panel pencarian, dan terakhir seleksi itu sendiri. Dengan kursor di kotak pencarian, Enter dan ⇧Enter melangkah ke kecocokan berikutnya dan sebelumnya, dan Escape menutup panelnya; dengan kursor di catatan teks baru, `⌘⏎` menguncinya.

## Markdown

GenOffice Markdown menyunting Markdown dengan pratinjau langsung di sampingnya.

### Berkas dan tampilan

| Tindakan                   | macOS      | Windows / Linux    |
| -------------------------- | ---------- | ------------------ |
| Simpan                     | ⌘S         | Ctrl+S             |
| Simpan sebagai             | ⇧⌘S        | Ctrl+Shift+S       |
| Cetak                      | ⌘P         | Ctrl+P             |
| Temukan                    | ⌘F         | Ctrl+F             |
| Alihkan sumber / pratinjau | ⌘E         | Ctrl+E             |
| Perbesar                   | ⌘= atau ⌘+ | Ctrl+= atau Ctrl++ |
| Perkecil                   | ⌘-         | Ctrl+-             |
| Ukuran asli (100%)         | ⌘0         | Ctrl+0             |

Perkecil juga menerima ⌘_ dan Ctrl+_. mencubit trackpad, atau menahan ⌘ atau Ctrl lalu menggulir di atas dokumen, juga memperbesar.

### Blok

| Tindakan                                       | macOS | Windows / Linux      |
| ---------------------------------------------- | ----- | -------------------- |
| Duplikat blok terpilih                         | ⌘D    | Ctrl+D               |
| Hapus blok terpilih                            | ⇧⌘⌫   | Ctrl+Shift+Backspace |
| Pindahkan blok terpilih ke atas                | ⇧⌘↑   | Ctrl+Shift+ArrowUp   |
| Pindahkan blok terpilih ke bawah               | ⇧⌘↓   | Ctrl+Shift+ArrowDown |
| Sisipkan tautan pada seleksi                   | ⌘K    | Ctrl+K               |
| Hapus tautannya (tekan lagi saat tautan aktif) | ⌘K    | Ctrl+K               |

### Di dalam blok matematika

Keduanya hanya berlaku selama editor LaTeX mengambang terbuka.

| Tindakan                      | macOS | Windows / Linux |
| ----------------------------- | ----- | --------------- |
| Terapkan rumusnya             | ⌘⏎    | Ctrl+⏎          |
| Tutup editor tanpa menerapkan | Esc   | Esc             |

Menerapkan rumus kosong akan menghapus bloknya. Editor itu juga menutup pada perubahan apa pun pada dokumen, dan mengeklik di luar menutupnya.

## HTML

GenOffice HTML menampilkan editor kode di samping pratinjau langsung. Tombol yang bekerja pada sebuah elemen halaman hanya berlaku ketika pratinjau itu sendiri yang punya fokus papan tik.

### Berkas dan tampilan sumber

| Tindakan                                       | macOS      | Windows / Linux    |
| ---------------------------------------------- | ---------- | ------------------ |
| Simpan                                         | ⌘S         | Ctrl+S             |
| Simpan sebagai                                 | ⇧⌘S        | Ctrl+Shift+S       |
| Ganti tampilan (sunting / terbagi / pratinjau) | ⌘\         | Ctrl+\             |
| Temukan                                        | ⌘F         | Ctrl+F             |
| Perbesar                                       | ⌘= atau ⌘+ | Ctrl+= atau Ctrl++ |
| Perkecil                                       | ⌘-         | Ctrl+-             |
| Ukuran asli (100%)                             | ⌘0         | Ctrl+0             |

Perkecil juga menerima ⌘_ dan Ctrl+_.

### Format dan riwayat

| Tindakan                     | macOS | Windows / Linux |
| ---------------------------- | ----- | --------------- |
| Tebal                        | ⌘B    | Ctrl+B          |
| Miring                       | ⌘I    | Ctrl+I          |
| Urungkan                     | ⌘Z    | Ctrl+Z          |
| Ulangi                       | ⇧⌘Z   | Ctrl+Shift+Z    |
| Ulangi (pintasan alternatif) | ⌘Y    | Ctrl+Y          |

Di dalam editor kode atau kolom isian formulir, tombol penyuntingan milik kolom itu yang mengambil alih.

### Di dalam bingkai pratinjau, dengan sebuah elemen terpilih

| Tindakan                                | macOS      | Windows / Linux    |
| --------------------------------------- | ---------- | ------------------ |
| Tanyakan elemen itu kepada AI           | ⌘K         | Ctrl+K             |
| Hapus elemennya                         | ⌫ / Delete | Backspace / Delete |
| Sepupu sebelumnya                       | ↑          | ArrowUp            |
| Sepupu berikutnya                       | ↓          | ArrowDown          |
| Pilih induknya                          | ←          | ArrowLeft          |
| Pilih anak pertama                      | →          | ArrowRight         |
| Pindahkan elemen ke atas dalam dokumen  | ⌥↑         | Alt+ArrowUp        |
| Pindahkan elemen ke bawah dalam dokumen | ⌥↓         | Alt+ArrowDown      |
| Batalkan seleksi                        | Esc        | Esc                |

Masing-masing butuh sebuah elemen yang sudah terpilih sebelumnya; tanpa ada yang terpilih, tombolnya tidak melakukan apa pun.

### Di dalam bingkai pratinjau, menyunting teks di tempat

| Tindakan                                 | macOS | Windows / Linux |
| ---------------------------------------- | ----- | --------------- |
| Kunci suntingannya                       | Esc   | Esc             |
| Kunci suntingannya (pintasan alternatif) | ⌘⏎    | Ctrl+⏎          |
| Tebalkan rentang terpilih                | ⌘B    | Ctrl+B          |
| Miringkan rentang terpilih               | ⌘I    | Ctrl+I          |

Menekan kunci menyimpan apa yang Anda ketik dan mengembalikan elemennya ke keadaan terpilih.
