# Sheets: spreadsheet

Sheets adalah editor mirip Excel; perhitungan dijalankan di proses mesin Rust tersendiri (kalau proses itu tumbang, aplikasi tidak ikut tumbang). Membuka dan menyimpan .xlsx asli; .csv dan .tsv terbuka sebagai tabel.

## Antarmuka

- **Pita**: delapan tab, dibahas satu per satu di bawah.
- **Bilah Rumus**: menampilkan dan menyunting rumus sel yang aktif; fungsi umum didukung.
- **Tab lembar** (bawah): menambah / mengganti nama / menghapus / memindahkan lembar.
- **Penyuntingan sel**: klik dua kali atau langsung ketik; Enter mengonfirmasi dan turun, Tab ke kanan, Esc membatalkan (kebiasaan Excel).
- **Pintasan**: sejalan dengan keluarga Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Tab pita

- **Beranda**: font, isian, batas, format angka (mata uang/persen/koma, tambah/kurangi desimal), perataan, penggabungan, penyisipan baris/kolom dan ukurannya, pemformatan bersyarat, format sebagai tabel, gaya sel, papan klip & penyalin format, urutkan & filter.
- **Sisipkan**: bentuk, ikon, simbol, persamaan, cuplikan layar, dan lain-lain.
- **Tata Letak Halaman**: warna & font tema, tombol garis kisi/heading cetak, pratinjau hentian halaman.
- **Rumus**: JumlahOtomatis dan sisipkan fungsi, menentukan nama (juga dari pilihan), lacak presedan/dependen, jendela pemantauan, menghitung ulang lembar/buku kerja.
- **Data**: urutkan & filter (termasuk filter lanjutan, hapus filter), teks ke kolom, gabungkan buku kerja, refresh semua.
- **Tinjau**: telusuri komentar (tampilkan, sebelumnya/berikutnya), terjemahkan.
- **Tampilan**: tombol garis kisi & heading, zoom, pratinjau normal / hentian halaman.
- **Desain Bagan**: muncul saat bagan dipilih — jenis bagan, gaya dan warna, sunting rentang data.

Tab Data, tombol demi tombol (dari kiri ke kanan pada gambar):

![Tab Data](img/sheets-data.png)

- **Tabel Pivot**: bangun pivot dari rentang saat ini; seret bidang untuk mengagregasi.
- **Refresh**: menghitung ulang data pivot saat ini.
- **Dari Teks/CSV**: impor .csv/.txt sebagai lembar baru, dipecah menurut pemisah.
- **Gabungkan Buku Kerja**: menarik lembar dari .xlsx lain ke berkas ini.
- **Refresh Semua**: menghitung ulang setiap pivot dan kumpulan data eksternal.
- **Urutkan** (drop-down): naik / turun / urutan khusus (aturan multikolom).
- **Filter**: menambahkan drop-down ▼ ke baris header; centang nilai yang ingin dipertahankan.
- Tombol kecil bertumpuk di sebelahnya: **Hapus** (kembalikan semua baris), **Terapkan Ulang** (jalankan filter saat ini lagi), **Lanjutan** (filter dengan rentang kriteria).
- **Teks ke Kolom** (drop-down): pecah satu kolom menjadi beberapa menurut pemisah atau lebar tetap.
- **Isi Cepat**: beri satu contoh, sisa kolom terisi mengikuti pola itu (ctrl+E).
- **Hapus Duplikat**: buang baris duplikat menurut kolom yang dipilih.
- **Validasi Data** (drop-down): aturan input untuk bagian yang dipilih (daftar drop-down, rentang angka, ...).
- **Konsolidasikan**: menggabungkan beberapa rentang ke satu tempat menurut kategori.
- **Analisis Bagaimana-Jika** (drop-down): pencarian tujuan — menyelesaikan satu sel input agar sel formula mencapai nilai target.
- **Kelompokkan / Pisahkan Kelompok** (drop-down): kelompok baris/kolom dengan lipat dan buka.
- **Subtotal**: menyisipkan baris subtotal per kategori.

Tab Rumus, tombol demi tombol:

![Tab Rumus](img/sheets-formulas.png)

- **Sisipkan Fungsi** (fx): cari fungsi dengan pemandu argumen.
- **Jumlah Otomatis** (drop-down): SUM sekali klik, ditambah rata-rata/jumlah data/maks/min.
- **Baru Digunakan / Keuangan / Logika / Teks / Tanggal & Waktu / Pencarian & Referensi / Matematika & Trigonometri / Lainnya**: telusuri dan sisipkan fungsi menurut kategori.
- **Manajer Nama**: melihat, membuat, dan menghapus rentasan bernama.
- **Tentukan Nama** (drop-down): beri nama pada bagian yang dipilih; **Gunakan dalam Rumus** menyisipkan nama yang sudah ada; **Buat dari Pilihan** memberi nama rentasan dari baris/kolom heading-nya.
- **Lacak Presedan / Lacak Dependen**: panah biru menunjukkan asal data sebuah rumus dan ke mana hasilnya mengalir; **Hapus Panah** membersihkannya.
- **Tampilkan Rumus**: sel menampilkan rumusnya sendiri, bukan hasilnya.
- **Pemeriksaan Kesalahan**: menemukan dan menjelaskan kesalahan rumus.
- **Jendela Pemantauan**: kunci sel yang Anda pedulikan dan lihat nilainya secara langsung.
- **Opsi Penghitungan** (drop-down): hitung ulang otomatis atau manual; pada mode manual, **Hitung Sekarang / Hitung Lembar** memicu perhitungannya secara manual.

## Angka dan format

- Format angka: umum, angka, mata uang, persen, tanggal/waktu, pecahan, ilmiah, dan lain-lain.
- Perataan, pembungkusan teks, sel gabungan, batas, dan isian.
- Tinggi baris dan lebar kolom dengan menyeret; klik dua kali pada batas untuk menyesuaikan otomatis.

## Data

**Urutkan & filter** (contoh: mengurutkan satu kolom menurun):

1. Klik **sel mana pun di kolom itu** (tidak perlu memilih seluruh kolom).
2. Tab Beranda ▸ **Urutkan & Filter** ▸ **Turun**; seluruh baris akan tersusun ulang bersama-sama (area itu diurutkan sebagai satu kesatuan).
3. Untuk aturan khusus (beberapa kolom, menurut warna): jalur yang sama, pilih **Urut Kustom**.
4. Filter: pilih baris header lalu klik **Urutkan & Filter ▸ Filter** — setiap header mendapat drop-down ▼ untuk mencentang nilai yang dipertahankan; hapus filter untuk mengembalikan semuanya.

- Urutkan dan filter.
- Bekukan panel.
- .csv / .tsv: terbuka langsung sebagai tabel (tsv ber-tab diurai sebagai satu tabel); penyimpanan menulis kembali format aslinya.

## Menu klik kanan

- **Di dalam kisi**: menu milik editor itu sendiri (Univer) — potong/salin/tempel, sisip dan hapus baris/kolom, sembunyikan, gabungkan sel, bekukan panel, dan item harian lainnya.
- **Pada bilah status bawah**: pilih statistik apa yang ditampilkan di bilah status (rata-rata / jumlah data / jumlah, ...); pilihan itu tersimpan.
- **Pada tab lembar di bawah**: tambah / ganti nama / hapus / warnai / sembunyikan lembar (menu tab Univer).
- Menu konteks bilah tab atas dibahas di [Tab dan pengelolaan jendela](help://tabs-and-windows).

## AI

- Panel AI di samping: pilih satu rentang lalu beri instruksi dengan bahasa biasa (format ulang, buat data, tulis rumus).
- Perubahan dari AI dapat dikembalikan dari panel.

## Menyimpan dan mengekspor

- Menyimpan .xlsx (rumus dan format tetap terjaga); Simpan Sebagai; Ekspor ke PDF mengikuti pemenggalan halaman cetak.
- Simpan Otomatis mengikuti aturan global (aktif setelah penyimpanan manual pertama).

## Stabilitas

- Sisi perhitungan Rust terisolasi pada level proses dari antarmuka: jika data ekstrem membuatnya berhenti, Anda akan mendapat pesan dan percobaan pemulihan sesi — bukan aplikasi yang tumbang.
