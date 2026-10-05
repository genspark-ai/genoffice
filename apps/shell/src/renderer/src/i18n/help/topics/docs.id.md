# Docs: pemrosesan kata

Docs adalah pengolah kata mirip Word: membaca dan menulis .docx asli dengan penomoran halaman yang benar-benar sesuai dengan yang tampil (WYSIWYG).

## Pita

Tab: **Beranda / Sisipkan / Tata Letak / Desain / Referensi / Tinjau / Tampilan**, ditambah tab kontekstual untuk objek yang dipilih (desain tabel, gambar).

- **Beranda**: papan klip; font (termasuk ukuran CJK dan tanda penekanan); paragraf (perataan/indens/spasi/daftar); gaya (Judul 1-6/Normal/Kutipan, dapat diubah).
- **Sisipkan**: pemisah halaman/seksi, tabel (termasuk tabel cepat), gambar, bentuk, hyperlink, header/footer, nomor halaman, tanggal, kotak teks.
- **Tata Letak**: margin, orientasi dan ukuran kertas, kolom, indentasi dan spasi paragraf.
- **Desain**: tema, set warna, tanda air, batas halaman.
- **Referensi**: daftar isi (dapat diperbarui), catatan kaki/akhir, keterangan, referensi silang.

  ![Tab Referensi](img/docs-references.png)

- **Tinjau**: pemeriksaan ejaan, komentar, lacak perubahan (tampilan Semua/Sederhana), jumlah kata.
- **Tampilan**: penggaris, garis kisi, panel navigasi, zoom, dan **dialog pintasan papan tikik** yang bisa dicari.

## Menu klik kanan

Klik kanan di mana saja pada isi dokumen — menu menyesuaikan dengan yang Anda klik. Kelompok utamanya:

- **Papan klip**: potong / salin / tempel / **tempel sebagai teks biasa**.
- **Font, paragraf**: ubah keluarga dan ukuran font, tebal/miring/garis bawah, perataan/indentasi/spasi tanpa harus membuka pita.
- **Sinonim**: menampilkan sinonim untuk kata yang dipilih; klik salah satu untuk mengganti.
- **Terjemahkan** (AI): terjemahkan bagian yang dipilih ke bahasa target (Inggris, Tionghoa Sederhana, Jepang, Korea, Prancis, Jerman, Spanyol, ...) lewat panel AI.
- **Komentar Baru**: lampirkan komentar pada bagian yang dipilih.
- **Ejaan** (pada kata yang salah eja): pengganti yang disarankan, abaikan semua, tambahkan ke kamus, atur bahasa pemeriksaan.
- **Hyperlink**: buka / edit / salin tautan / hapus hyperlink.
- **Gambar**: lihat gambar, simpan gambar sebagai…, **bungkus teks** (sebaris / persegi kiri & kanan / atas dan bawah / di belakang teks / di depan teks), urutan penataan.
- **Field** (daftar isi, nomor halaman): perbarui field / tampilkan/sembunyikan kode field / edit field.
- **Penomoran daftar** (di dalam daftar): mulai penomoran ulang / lanjutkan penomoran / ubah tingkat daftar / atur nilai penomoran.
- **Tabel** (kursor di dalam tabel): sisipkan baris/kolom, gabungkan / pisahkan sel, pisahkan tabel, sesuaikan otomatis, perataan sel, sebar baris/kolom, properti tabel, menu hapus, pilih.

## Menyunting

- Cari dan ganti (ctrl+F / ctrl+H): sensitif huruf besar-kecil, seluruh kata, regex.
- Penyalin format; urungkan/ulangi yang mendalam; opsi tempel.
- Tabel: gabungkan/pisahkan sel, operasi baris/kolom, batas dan arsir, urutkan, rumus.
- Gambar: pembungkusan teks, pemotongan, kompresi; kanvas gambar.

## Tipografi CJK

- Kompresi tanda baca dan pemenggalan baris kinsoku sesuai dengan Word; konversi lebar penuh/setengah.
- Kandidat font mencakup nama keluarga CJK yang umum di Windows dan macOS.

## AI

- Tombol AI pada pita dan panel samping: tulis ulang, perluas, terjemahkan, ringkas, sisipkan tabel, ditambah instruksi bebas.
- Setiap giliran AI menyimpan snapshot lebih dulu; gulir balik dari daftar versi, dan pengguliran balik itu sendiri bisa dibatalkan.

## Menyimpan dan mengekspor

- Menyimpan .docx menulis ulang hanya paragraf yang berubah — isi yang tidak disentuh tetap identik byte demi byte.
- Mengekspor PDF (sesuai penomoran halaman) dan gambar per halaman.

## Mencetak

ctrl+P melalui dialog sistem, dengan halaman yang tampil persis seperti yang akan tercetak (WYSIWYG).
