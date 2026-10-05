# Editor Markdown

Editor Markdown membuka .md / .markdown dengan pengalaman sumber + pratinjau hasil render.

- **Buka**: dari Beranda atau File ▸ Buka…; perintah baris perintah juga bisa.
- **Sunting**: penyuntingan teks biasa; ekstensi GFM (tabel, daftar tugas, coret, autolink) dirender di pratinjau.
- **Pratinjau**: langsung; aset relatif seperti gambar diselesaikan di samping dokumen.
- **Menyimpan**: setia byte demi byte — BOM, CRLF, dan keberadaan baris baru di akhir file dipertahankan; penyimpanan tanpa perubahan tidak menulis ulang file.
- **Cari dan ganti**: ctrl+F mencari di sumber; Ganti Semua menuliskan hasilnya kembali.
- **AI**: tombol siap pakai membiarkan asisten menulis ulang, memperpanjang, atau menerjemahkan dokumen.

## Bilah alat

Satu baris tombol di atas editor (arahkan kursor untuk melihat keterangan):

![Bilah alat Markdown](img/md-toolbar.png)

- **File & riwayat**: Simpan, Simpan Sebagai…, Urungkan, Ulangi, Temukan; sakelar **Simpan Otomatis** di sebelah kanan menulis perubahan ke disk secara berkala.
- **Tombol AI**: membuka panel AI; di sebelahnya adalah siap-siap untuk menulis ulang / memperpanjang / menerjemahkan.
- **Gaya paragraf** (drop-down): berpindah antara teks isi dan berbagai tingkat judul.
- **Format sebaris**: **tebal**, _miring_, ~~coret~~, `kode sebaris`, tautan.
- **Daftar**: daftar poin, daftar bernomor, daftar tugas.
- **Sisipkan**: tabel, gambar, garis pemisah.
- **Properti**: menyisipkan atau melompat ke blok YAML front matter di bagian atas file.
- **Kerangka**: melompat menurut hierarki judul.
- **Ejaan**: menyalakan atau mematikan pemeriksaan ejaan untuk dokumen ini.

Tiga contoh singkat:

- **Judul**: letakkan kursor di baris itu ▸ drop-down gaya paragraf ▸ "Judul 1".
- **Tabel**: klik **Sisipkan tabel** ▸ seret untuk memilih jumlah baris/kolom ▸ ketik di dalam sel; pratinjau langsung merendernya.
- **Daftar tugas**: pilih beberapa baris ▸ klik **Daftar tugas** ▸ setiap baris menjadi `- [ ]`, dan di pratinjau tampil sebagai kotak centang.
