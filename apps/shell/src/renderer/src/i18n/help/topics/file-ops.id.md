# Operasi file: ganti nama, hapus, ekspor

Bab ini mencakup operasi file yang dipakai bersama oleh semua editor; pilihan ekspor khas tiap editor ada di babnya masing-masing.

**Menu ⋯** pada baris file (arahkan kursor ke baris file) mengumpulkan tindakan berikut:

![Menu ⋯ pada baris file](img/file-ops.png)

## Mengganti nama

Ada dua pintu masuk, satu set pemeriksaan:

- **⋯ ▸ Ganti nama** dari baris di Beranda.
- **Klik dua kali tab file** untuk mengganti nama di tempat (lihat [Tab dan pengelolaan jendela](help://tabs-and-windows)).

Aturannya: ekstensi dipertahankan otomatis; karakter terlarang, titik di akhir, dan nama cadangan (CON/NUL dan sejenisnya) ditolak dengan pesan; dan nama yang bentrok di folder yang sama juga dicegah. Berkas sungguhan di disk diganti namanya, serta daftar terbaru dan bintang ikut menyesuaikan.

## Menghapus

- **⋯ ▸ Hapus** di Beranda: memindahkan file ke **tempat sampah sistem**, dan dapat dipulihkan dari sistem operasi.
- Setelah penghapusan muncul pemberitahuan dengan aksi undo selama beberapa detik — undo mengembalikan file ke tempat semula.

## Duplikat

**⋯ ▸ Duplikat** membuat salinan <nama> di folder yang sama dan membukanya sebagai tab baru; bila nama bertabrakan, angka ditambahkan otomatis.

## Simpan dan Simpan Sebagai

- **⌘S / ctrl+S** menyimpan file saat ini; file tanpa nama akan lebih dulu meminta lokasi dan nama.
- **Simpan Sebagai…** menulis berkas baru dan membiarkan aslinya tidak tersentuh; penyuntingan setelahnya menyasar berkas baru.
- Setiap penyimpanan bersifat atomik (berkas sementara + ganti nama); berhenti di tengah penulisan tidak akan merusak file.
- Simpan Otomatis baru aktif setelah penyimpanan manual pertama (lihat [Mulai cepat](help://getting-started)).

## Ekspor ke PDF

- **Docs**: Arquivo ▸ Ekspor sebagai PDF… (atau tombol di pita), dengan pemenggalan halaman seperti tata letaknya.
- **Slides**: ekspor meraster halaman demi halaman, dengan progres untuk dek besar.
- **Sheets**: ekspor mengikuti pemenggalan halaman cetak.
- Ekspor dirender di jendela tersembunyi dan hasilnya mendarat di mana pun yang Anda pilih.

## Ekspor ke Word / gambar

- **PDF ▸ PDF ke Word**: mengubah PDF menjadi .docx (dikonversi secara lokal; tata letak kompleks dibuat semaksimal mungkin).
- **Docs** dapat mengekspor halaman sebagai gambar (PNG per halaman).

## Mencetak

Arquivo ▸ Cetak… di tiap editor (⌘P/ctrl+P) membuka kotak dialog cetak sistem; PDF dicetak dengan urutan halaman dan rotasi saat ini.

## Di mana berkas tanpa nama tinggal

Lokasi yang Anda pilih saat penyimpanan pertama menjadi rumah berkas itu; sebelumnya dokumen hanya ada di memori. Simpan Otomatis baru mengambil alih setelah penyimpanan pertama itu.
