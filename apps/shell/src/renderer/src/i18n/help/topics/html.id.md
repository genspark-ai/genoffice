# Editor HTML

Editor HTML membuka .html / .htm dengan dua mode: **Pratinjau** (halaman yang sudah dirender) dan **Sumber**.

- **Pratinjau**: render yang sesungguhnya; stylesheet dan gambar relatif dimuat dari lokasi di samping berkas.
- **Inspektur pratinjau**: klik untuk memilih sebuah elemen, klik dua kali untuk menyunting teksnya di tempat, hapus lewat bilah alat, dan Tanya AI untuk bagian yang dipilih.
- **Mode sumber**: menyunting HTML; ctrl+F untuk mencari, dan Ganti Semua menyimpan penanda yang ditulis ulang.
- **Menyimpan**: setia byte demi byte (BOM/CRLF/baris baru di akhir file dipertahankan); penyimpanan tanpa perubahan tidak menulis ulang.
- **Zoom**: ctrl+roda / cubit untuk memperbesar pratinjau; ctrl+Z di dalam pratinjau membatalkan penyuntingan terakhir.

## Bilah alat

Klik elemen apa pun di pratinjau dan sebuah bilah alat melayang di atasnya:

![Bilah alat melayang di atas elemen yang dipilih](img/html-toolbar.png)

- **File & riwayat**: Simpan, Simpan Sebagai…, Urungkan, Ulangi, Temukan; sakelar **Simpan Otomatis** menulis perubahan secara berkala.
- Sakelar **Pratinjau / Sumber**; **Presentasi** menampilkan halaman di layar penuh.
- **Format**: tebal, miring, tambah/kurangi ukuran font; serta **panel gaya** untuk elemen yang dipilih (warna dan lainnya).
- **Sisipkan**: judul, paragraf, tabel, gambar (dengan tautan), Lainnya.
- **Tindakan gambar** (dengan gambar dipilih): pangkas, **hapus latar belakang**, ganti, kunci rasio aspek.
- **Tindakan elemen** (dengan elemen dipilih di inspektur pratinjau): hapus, duplikat, naikkan/turunkan.
- **Tombol AI**: membuka panel AI; tanyakan apa pun tentang elemen yang dipilih.

## Ekspor

Menu File, semuanya lokal dan semuanya menanyakan tempat untuk menaruh hasilnya:

- **Ekspor sebagai Word…** dan **Ekspor sebagai PDF…** menulis .docx atau .pdf yang sungguhan.
- **Ekspor sebagai HTML satu file…** menulis satu .html dengan gambar tertanam di dalamnya. File ini tidak akan menimpa file yang sedang Anda buka, dan memberi tahu berapa gambar yang tidak bisa ditanam.

## Sisipkan kerangka

Untuk halaman kosong, **Sisipkan ▸ Sisipkan kerangka** menulis dokumen minimal dalam mode standar:

```html
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <title></title>
  </head>
  <body></body>
</html>
```

Setiap bagian ada karena alasan tertentu, dan itulah sebabnya ia berupa perintah, bukan sesuatu yang perlu Anda ketik sendiri:

- **doctype**-nya, atau pratinjau berjalan dalam quirks mode, tempat ukuran kotak dan tata letak tabel mengikuti aturan yang berbeda dari yang Anda harapkan;
- **`lang`**-nya, atau pembaca layar tidak punya bahasa untuk membaca halaman itu, dan peramban memilihkan font serta pemeriksa ejaan untuk bahasa yang salah;
- **charset**-nya, atau halaman berisi teks non-Latin bisa tampil sebagai karakter kacau (mojibake).

Tag meta viewport sengaja tidak disertakan: ini dirender di panel desktop, tanpa ada viewport seluler yang bisa dipengaruhinya.

`lang` mengikuti bahasa antarmuka aplikasi, jadi kerangka yang Anda sisipkan adalah kerangka yang sudah disiapkan untuk perkakas Anda. Suntinglah sesuka hati setelahnya.

Butir ini hanya muncul dalam mode sunting, dan hanya selama dokumennya masih kosong — begitu ada isi, tidak ada lagi yang bisa disisipkan kerangka _ke dalam_.
