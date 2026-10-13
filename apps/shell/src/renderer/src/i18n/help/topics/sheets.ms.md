# Sheets: hamparan

Sheets ialah editor yang menyerupai Excel. Pengiraan dijalankan dalam proses berasingan yang menggunakan enjin ditulis dalam Rust, jadipersekitaran yang rosak di sana tidak akan menjatuhkan aplikasi ini. Ia membuka dan menyimpan fail .xlsx yang sebenar, manakala .csv dan .tsv dibuka sebagai jadual.

## Antara muka

- **Reben**: lapan tab, diterangkan satu per satu di bawah.
- **Bar formula**: memaparkan dan menyunting formula sel aktif. Fungsi yang lazim disokong.
- **Tab helaian** (di bawah): tambah / namakan semula / padam / pindahkan helaian.
- **Penyuntingan sel**: klik dua kali atau taip sahaja. Enter mengesahkan dan bergerak ke bawah, Tab bergerak ke kanan, Esc membatalkan, mengikut tabiat Excel.
- **Pintasan**: selaras dengan keluarga Excel (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Tab reben

- **Laman Utama**: fon, isian, sempadan, format nombor (mata wang / peratus / ribuan, tambah atau kurangkan perpuluhan), penjajaran, penggabungan, sisip dan saiz baris serta lajur, format bersyarat, format sebagai jadual, gaya sel, papan keratan dan berus format, isihan dan penapisan.
- **Sisip**: gambar, bentuk, Kotak Teks, pautan, komen, kotak semak, Pengepala & Pengaki, simbol, persamaan; **sembilan jenis carta** (lajur, bar, garis, kawasan, pai, serakan, radar, donat, dan gabungan lajur + garis) serta **Carta Disyorkan**, yang mengesyorkan jenis mengikut bahagian yang dipilih dan memaparkan pratontonnya; **Carta Pangsi** daripada sel tempat anda berada; **Garis Percikan** (garis, lajur, menang/kalah); **Penghiris** dan **Garis Masa** untuk menapis jadual berpivot.
  - Carta, carta pivot dan garis percikan ialah objek sebenar dalam buku kerja yang disimpan.
  - **Penghiris** dan **Garis Masa** ialah kawalan sesi: penapisan yang mereka lakukan disimpan dan Excel memaparkan jadual berpivot yang ditapis dengan cara yang sama, tetapi butang penghiris itu sendiri bukan sebahagian daripada fail.
- **Bentangan Halaman**: warna dan fon tema, suis untuk cetak garisan grid dan tajuk, pratonton pemisah halaman.
- **Formula**: jumlah automatik dan sisip fungsi, takrif nama (juga daripada pilihan), jejak pendahulu dan bergantung, tetingkap pemerhati, kira semula helaian atau buku kerja.
- **Data**: isihan dan penapisan (termasuk penapis lanjutan dan kosongkan penapis), teks ke lajur, gabung buku kerja, muat semula semua.
- **Semakan**: semak komen (tunjukkan, sebelumnya / seterusnya), terjemah, dan kumpulan **Perlindungan** — **Lindungi Helaian**, **Lindungi Buku Kerja** dan **Benarkan Edit Julat**.
  - Perlindungan ditulis ke dalam .xlsx dan tiada kata laluan pada apa yang digunakan oleh aplikasi ini, jadi butang yang sama menjadi **Nyahlindung…** dan membatalkannya. Perlindungan daripada program lain yang _benar-benar_ membawa kata laluan tidak boleh dibuang dari sini.
  - Kedua-dua butang perlindungan berkuat kuasa serta-merta — tiada dialog untuk dibatalkan, hanya nota pada baris status bahawa ia akan ditulis semasa menyimpan.
  - **Benarkan Edit Julat** menanda sel yang kekal boleh disunting sementara bahagian lain helaian dikunci.
- **PANDANGAN**: suis **Bar formula**, Garis Grid, Tajuk dan Serlahkan baris & lajur aktif; zum; paparan **Biasa** dan **Pratonton Pemisah Halaman**. Garis grid dan tajuk disimpan bersama helaian; penandaan ialah pilihan peribadi anda.
- **Reka Bentuk Carta**: muncul apabila carta dipilih, serta jenis carta, gaya dan warna, serta suntingan julat data.

Tab Data, butiran demi butiran (dari kiri ke kanan dalam gambar):

![Tab Data](img/sheets-data.png)

- **Jadual berpivot**: membina jadual berpivot daripada julat semasa. Seret medan untuk mengagregasikan.
- **Muat semula**: mengira semula data jadual berpivot semasa.
- **Daripada teks/CSV**: mengimport fail .csv/.txt sebagai helaian baharu, dipecahkan mengikut pemisah.
- **Gabung buku kerja**: menarik helaian daripada fail .xlsx lain ke dalam fail ini.
- **Muat semula semua**: mengira semula setiap jadual berpivot dan sumber data luar.
- **Isih** (senarai lungsur): menaik / menurun / isih tersuai (peraturan berbilang lajur).
- **Tapis**: menambah lungsur ▼ pada baris tajuk. Tandakan nilai yang mahu dikekalkan.
- Butang kecil bersusun di sebelahnya: **Kosongkan** (memulangkan semua baris), **Guna semula** (menjalankan penapis semasa sekali lagi), **Lanjutan** (menapis menggunakan julat kriteria).
- **Teks ke lajur** (senarai lungsur): membelah satu lajur kepada beberapa mengikut pemisah atau lebar tetap.
- **Isi pantas**: berikan satu contoh dan selebihnya lajur akan diisi mengikut contoh itu (ctrl+E).
- **Buang pendua**: membuang baris berulang berdasarkan lajur yang dipilih.
- **Pengesahan data** (senarai lungsur): peraturan input untuk pilihan tersebut (senarai lungsur, julat nombor, ...).
- **Gabungkan**: menggabungkan beberapa julat ke satu tempat mengikut kategori.
- **Analisis what-if** (senarai lungsur): pencarian sasaran — menyelesaikan satu sel input supaya sel formula sampai pada nilai sasaran.
- **Kumpulan / Nyahkumpulan** (senarai lungsur): kumpulan baris atau lajur dengan lipatan dan pengembangan lipatan.
- **Subjumlah**: menyisip baris subjumlah bagi setiap kategori.

Tab Formula, butiran demi butiran:

![Tab Formula](img/sheets-formulas.png)

- **Sisip fungsi** (fx): mencari fungsi dengan pemandu argumen.
- **Jumlah automatik** (senarai lungsur): SUM dengan satu klik, serta min / bilangan / maksimum / minimum.
- **Digunakan baru-baru ini / Kewangan / Logik / Teks / Tarikh dan masa / Carian dan rujukan / Matematik dan trigonometri / Lain-lain**: semak dan sisip fungsi mengikut kategori.
- **Pengurus nama**: lihat, cipta dan padam julat bernama.
- **Takrif nama** (senarai lungsur): menamakan pilihan; **Guna dalam formula** menyisip nama sedia ada; **Cipta daripada pilihan** menamakan julat berdasarkan baris atau lajur tajuknya.
- **Jejak pendahulu / Jejak bergantung**: anak panah biru yang menunjukkan dari mana data formula datang dan ke mana ia mengalir; **Buang anak panah** menghapusnya.
- **Papar formula**: sel memaparkan formula itu sendiri, bukan hasilnya.
- **Semakan ralat**: mencari dan menerangkan ralat formula.
- **Tetingkap pemerhati**: anda boleh menombak sel yang anda minat dan memerhati nilai langsungnya.
- **Pilihan pengiraan** (senarai lungsur): kira semula automatik atau manual. Dalam mod manual, **Kira sekarang / Kira helaian** mencetusnya secara manual.

## Nombor dan pemformatan

- Format nombor: am, nombor, mata wang, peratus, tarikh/masa, pecahan, saintifik, dan lain-lain.
- Penjajaran, pembalutan teks, sel digabungkan, sempadan dan isian.
- Ketinggian baris dan kelebar lajur dengan menyeret. Klik dua kali pada sempadan untuk melaraskan secara automatik.

## Data

**Isih dan tapis** (contohnya menurun mengikut satu lajur):

1. Klik **sel mana-mana dalam lajur tersebut**. Anda tidak perlu memilih keseluruhan lajur.
2. Tab Laman Utama ▸ **Isih dan Tapis** ▸ **Menurun**. Baris penuh akan tersusun semula bersama-sama, kerana kawasan itu diisihkan sebagai satu keseluruhan.
3. Untuk peraturan tersuai (berbilang lajur, mengikut warna): laluan yang sama, pilih **Isih Tersuai**.
4. Penapisan: pilih baris tajuk dan klik **Isih dan Tapis ▸ Tapis**. Setiap tajuk mendapat lungsur ▼, dan anda menanda nilai yang mahu dikekalkan. Mengosongkan penapis memulangkan semuanya.

- Isih dan tapis.
- Bekalkan beku.
- .csv / .tsv: dibuka terus sebagai jadual, dan fail tsv yang dipisah dengan tab dibaca sebagai satu. Menyimpan menulis semula format asal.

## Menu konteks

- **Dalam grid**: menu editor itu sendiri (Univer), termasuk potong / salin / tampal, sisip dan pemadaman baris serta lajur, sembunyi, gabung sel, bekalkan beku, dan item harian yang lain.
- **Pada bar status di bawah**: pilih statistik yang mahu ditunjukkan pada bar status, iaitu min / bilangan / jumlah, dan pilihan itu akan kekal.
- **Pada tab helaian di bawah**: tambah / namakan semula / padam / warnakan / sembunyikan helaian, iaitu menu tab Univer.
- Menu konteks baris tab di atas diterangkan dalam [Tab dan pengurusan tingkap](help://tabs-and-windows).

## AI

- Panel AI di sisi: pilih julat dan berikan arahan dalam bahasa biasa, seperti memformat semula, menjana data, atau menulis formula.
- Lampirkan fail pada arahan dengan butang 📎, atau seret ke dalam panel tersebut; fail ikut bersama soalan dan gambar kembali sebagai lakaran kecil.
- Satu jawapan boleh memetik sel — klik rujukan itu dan grid akan melompat ke sana.
- Perubahan oleh AI boleh dibatalkan daripada panel tersebut.

## Simpanan dan eksport

- Menyimpan .xlsx, dengan formula dan format dikekalkan; Simpan Sebagai; dan eksport ke PDF mengikut pemenggalan cetakan.
- Autosimpan mengikut peraturan global, iaitu diaktifkan selepas simpanan manual pertama.

## Kestabilan

- Proses pengiraan Rust diasingkan daripada antara muka. Jika data ekstrem, anda akan melihat mesej dan percubaan memulihkan sesi, bukan ranam aplikasi.
