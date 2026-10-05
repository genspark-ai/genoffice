# Layar Beranda: tempat file Anda berada

Beranda adalah halaman awal GenOffice: bilah navigasi di sebelah kiri, daftar file dan kartu pembuatan cepat di sebelah kanan.

![Layar Beranda](img/home-screen.png)

## Navigasi bilah sisi

- **Terbaru**: file yang baru saja Anda buka, dikelompokkan menurut waktu (minggu ini / bulan ini / sebelumnya).
- **Berbintang**: file yang Anda beri bintang. Arahkan kursor ke baris file lalu klik bintang untuk menambah atau menghapusnya.
- **Genspark Projects**: setelah masuk ke akun Genspark, menampilkan proyek yang Anda buat dengan Genspark AI di web; klik salah satunya untuk melanjutkan pengeditan di browser. Pencarian, pengurutan menurut waktu, penyegaran, dan muat lebih banyak tersedia.
- **Folder**: sematkan direktori yang sering dipakai ke bilah sisi (Tambah folder…) lalu lompat ke sana seperti penanda. Akar yang tidak tersedia ditampilkan sebagai tidak tersedia dan dapat dihapus dari daftar.
- **Sampah**: menunjuk ke tempat sampah sistem — file yang dihapus pindah ke sana dan dapat dipulihkan dari sistem operasi.

## Daftar file

Setiap baris menampilkan ikon, nama file, waktu modifikasi, dan lainnya. **Menu ⋯** pada baris menyediakan:

- **Ganti nama**: langsung di tempat, ekstensi dipertahankan otomatis.
- **Beri bintang / Hapus bintang**
- **Duplikat**: membuat salinan di folder yang sama.
- **Hapus**: memindahkan file ke tempat sampah sistem — bukan penghapusan permanen.
- **Tampilkan di folder**: menemukan file di aplikasi file Anda.

## Pencarian

Kotak pencarian di atas mencocokkan dua hal sekaligus:

- **Nama file**: penyaringan cepat berdasarkan nama.
- **Isi file**: GenOffice mengindeks file Anda di latar belakang (teks di dalam docx/xlsx/pptx/pdf/md/html, dengan cadangan OCR untuk PDF hasil pindai), sehingga mencari teks isi pun menemukan file. Cakupan dan sakelarnya ada di pengaturan pencarian.

## Kartu Mulai cepat

Kartu di atas daftar membuat dokumen baru dalam satu langkah. Mengklik kartu membuat file bertipe tersebut dan membuka editornya — langsung mulai menulis, atau biarkan AI yang membuat drafannya (setiap editor punya **tombol AI** di pita, dan **Tanya AI** di menu konteks pilihan).

File baru akan masuk ke folder yang sedang dipilih di bilah sisi; jika tidak ada yang dipilih, file masuk ke folder bawaan.

Apa yang dilakukan tiap kartu:

- **AI Docs** (.docx): dokumen teks kosong di editor Docs. File baru ditulis ke disk saat **penyimpanan pertama**; dokumen baru terbuka dengan panel AI terbentang (nonaktifkan lewat Pengaturan → "Buka panel AI di dokumen baru").
- **AI Sheets** (.xlsx): spreadsheet kosong di editor Sheets. Sebelum Anda menyimpan, belum ada file di disk — namanya dicadangkan untuk penyimpanan pertama; setelah generasi AI pertama, file juga bisa diberi nama otomatis dari isinya.
- **AI Slides** (.pptx): presentasi kosong di editor Slides.
- **AI Markdown** (.md): dokumen Markdown kosong di editor Markdown.
- **AI HTML** (.html): halaman web kosong di editor HTML.
- **AI PDF** (.pdf): berbeda dari yang lain — kartu ini **segera** membuat PDF kosong satu halaman yang sungguhan di folder tujuan dan membukanya sebagai file biasa (editor PDF bekerja pada file asli). Bagus untuk memberi anotasi, menutupi bagian, atau menambahkan teks; file bisa diberi nama otomatis dari isinya saat penyimpanan pertama.
- **Buka File Lokal**: pemilih file sistem untuk Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown), dan halaman web (.html/.htm). Bisa memilih banyak file sekaligus; tiap file mendapat tabnya sendiri.

> Tip: menu Arquivo ▸ Novo di bilah menu membuat jenis dokumen yang sama (⌘N/Ctrl+N memakai dokumen teks secara bawaan); menyeret file ke jendela akan membukanya.

## Proyek cloud (Genspark Projects)

- Penggunaan pertama memerlukan masuk ke akun Genspark Anda (alur device-code: GenOffice menampilkan kode, Anda menyelesaikan login di browser).
- Daftar proyek tersinkron dengan versi web; Buka di browser melompat ke sana untuk melanjutkan.
- Tidak masuk tidak memengaruhi satu pun fitur lokal.
