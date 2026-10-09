# PDF: membaca, memberi anotasi, dan menutupi

Editor PDF punya lima tab di pita: **Beranda / Anotasi / Edit / Halaman / Tampilan**. Ia bisa membaca dan menulis: teks dapat diedit, konten dapat ditutupi, ditandatangani, dan formulir dapat diisi.

## Membaca dan navigasi

- Bilah sisi kiri: **Gambar mini** (klik untuk melompat, rentang yang terlihat disorot) atau **Kerangka** (bookmark, jika ada).
- Zoom: kendali rasio di kanan bawah; ctrl+roda mengubah tingkat zoom.
- Rotasi: per halaman atau semua halaman dari menu Halaman; rotasi ditulis kembali saat menyimpan.
- Cari: ctrl+F untuk mencari seluruh teks dengan semua hasil disorot.
- PDF terenkripsi: akan muncul Kotak Dialog Kata Sandi (jendela kecil tersendiri) untuk membukanya; kata sandi hanya dipakai untuk sesi ini.

## Memilih teks dan memberi markup (Anotasi)

Coba pada paragraf mana pun:

1. **Seret mouse melewati sebuah kalimat** — saat dilepas, sebuah bilah anotasi melayang di atasnya:

![Bilah anotasi setelah memilih teks](img/pdf-highlight.png)

2. Pilih **Sorot** (swatch kuning membuka palet warna), **Garis bawah** atau **Coret**; **Tanya AI** mengirim pilihan beserta pertanyaan Anda ke panel AI.
3. Untuk membatalkan sebuah anotasi, seret-pilih ulang bagian yang sama lalu klik tombol aktif pada bilah itu (sakelar bergaya Word), atau pilih bagian itu dan tekan Delete.

Catatan:

- Seret melewati teks dan sebuah bilah muncul: **Sorot / Garis bawah / Coret / Salin / Tanya AI**.
- Warna diambil dari palet; **menerapkan markup yang sama pada rentang yang sudah ditandai akan menghapusnya** (sakelar bergaya Word).
- Markup yang sudah tersimpan di dalam file dapat dipilih dan dihapus (menu ⋯ atau Delete).
- **Catatan**: selagi alat gambar aktif, lapisan teks tidak dapat dipilih — alatnya menonaktifkan diri sendiri setelah tiap penempatan, sehingga Anda kembali ke mode pilih untuk tindakan berikutnya.

## Alat gambar (Anotasi)

Enam alat: **Gambar, Persegi, Elips, Panah, Catatan**, ditambah **Redaksi area** pada tab Anotasi.

- Tiap alat adalah sakelar: klik untuk mengaktifkan; **alatnya menonaktifkan diri sendiri begitu sebuah bentuk ditempatkan** (klik lagi alatnya untuk melanjutkan); mengklik alat yang aktif juga menonaktifkannya.
- Ik mengikuti lebar goresan; persegi/elips/panah digambar dengan menyeret; warna diambil dari palet gambar.
- Bentuk yang ditempatkan dapat dipilih, dihapus, diseret, dan (persegi/elips) diubah ukurannya.
- **Menutupi konten, alur lengkapnya** (menyembunyikan satu baris teks):

  1. Tab Anotasi ▸ klik **Redaksi area** (alat aktif).
  2. **Seret sebuah kotak di atas konten** — isinya tertutup tanda arsir, dan bilah alat memperoleh tombol **Hapus tanda / Terapkan redaksi**:

  ![Halaman setelah menandai area redaksi](img/pdf-redact.png)

  3. Klik **Terapkan redaksi** lalu konfirmasi — hasilnya adalah salinan kerja tempat teks dan gambar yang tertutup benar-benar dihapus (bukan sekadar ditutupi) dan tidak dapat dibatalkan; dokumen asli tidak tersentuh.

  Terlalu keliru? Tombol hapus tanda menghapus tanda saat ini sehingga Anda bisa menggambar ulang.

## Catatan tempel dan utas komentar

- **Alat Catatan** menjatuhkan pin dan membuka kartu di margin untuk teks (nama penulis dapat diatur); setelah dikonfirmasi tersimpan sebagai anotasi Teks PDF standar.
- Klik sebuah pin untuk membuka utasnya: **Balas** (utas datar gaya WPS/Acrobat), **Edit** komentar Anda, **Hapus** satu komentar atau seluruh utas.
- Edit yang sedang berjalan bertahan sampai penyimpanan menuliskan teks baru ke anotasi yang sama di dalam file, sehingga rantai balasan tetap utuh.

## Mengedit konten PDF (Edit)

- **Edit teks**: klik teks untuk mengeditnya per blok (mesin pdfium; pencocokan font-contain semaksimal mungkin).
- **Sisipkan teks**: tempatkan teks yang bisa dicari dengan pilihan font/ukuran/warna.
- **Sisipkan gambar / stempel**.
- Formulir: bidang AcroForm dapat diisi langsung; nilainya ditulis saat menyimpan.

## Tanda tangan

- **Tanda tangan tinta**: gambar tangan; dapat ditautkan ke bidang tanda tangan formulir.
- **Tanda tangan gambar**: tempatkan sebuah gambar sebagai tanda tangan.
- Tanda tangan yang tersimpan dapat digunakan lagi.

## Operasi halaman (Halaman)

- **Putar / Hapus halaman / Ubah urutan**: seret gambar mini untuk mengubah urutan; penghapusan meminta konfirmasi.
- **Impor halaman**: tarik halaman dari PDF lain ke dalam dokumen. **Sisipkan halaman kosong** menambahkan satu halaman kosong.
- **Ganti halaman** menukar sebuah rentang dengan halaman dari tempat lain; **Pangkas halaman** memotong tepinya, dengan opsi menerapkannya ke semua halaman.
- **Ukuran halaman** menskalakan ulang setiap halaman ke satu ukuran kertas; **Balik urutan** membalik dokumen dari ujung ke ujung.
- **Ekstrak halaman ini**: ekspor halaman yang dipilih ke PDF baru.
- **Pisahkan PDF**: ada dua bentuk — memecah menurut rentang menjadi beberapa file, atau memotong setiap halaman menjadi kisi halaman yang lebih kecil.
- **Gabungkan PDF**: ada dua bentuk — menambahkan PDF lain, atau menggabungkan beberapa halaman ke satu lembar. Ukuran dijumlahkan **sebelum** apa pun dibaca, dan total yang melebihi **1 GiB akan ditolak** dengan pesan yang bisa dibaca (menjaga pemakaian memori tetap terkendali).
- Perubahan tingkat halaman ditulis kembali pada penyimpanan berikutnya; Simpan Sebagai membiarkan dokumen asli tidak tersentuh.

## Ekspor dan pencetakan

- **Ekspor sebagai Word… / PowerPoint… / Excel…** di menu File, atau ketiganya dari **Konversi PDF** pada pita — semuanya lokal, tanpa unggahan. .pptx keluar dengan satu slide per halaman dan .xlsx dengan satu lembar kerja per halaman. Masing-masing menanyakan tempat penyimpanan.
- **Cetak**: urutan dan rotasi saat ini melalui dialog sistem; rentang halaman didukung.

## Menyimpan

- Simpan biasa/otomatis menuliskan anotasi dan editan kembali ke dalam file (secara atomik).
- **Penutupan konten lewat alur Terapkan miliknya sendiri** sehingga menghasilkan salinan dan membiarkan dokumen asli tidak tersentuh, supaya konten sensitif tidak tertinggal di sana.
