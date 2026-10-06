# PDF: membaca, menganotasi, dan redact

Editor PDF mempunyai lima tab pada reben: **Laman Utama / Anotasi / Edit / Halaman / Pandangan**. Ia boleh membaca dan menulis, yang bermaksud teks boleh disunting, kandungan boleh di-redact dan ditandatangani, serta borang boleh diisi.

## Membaca dan navigasi

- Bar sisi kiri: **lakaran kenit** untuk melompat ke slaid dengan julat yang kelihatan ditanda, atau **kerangka** untuk penanda halaman apabila ia wujud.
- Zum: kawalan nisbah di penjuru kanan bawah. Gunakan ctrl dan roda tetikus untuk menukar zum secara berperingkat.
- Putaran: setiap halaman atau kesemua halaman melalui menu Halaman, dan putaran itu ditulis semula semasa menyimpan.
- Carian: ctrl+F untuk mencari teks penuh, dengan semua hasil ditanda.
- PDF yang disulitkan: prompts pertanyaan kata laluan dalam tetingkap kecil tersendiri, dan kata laluan itu hanya digunakan untuk sesi ini.

## Memilih teks dan menanda (Anotasi)

Cuba pada mana-mana perenggan:

1. **Seret tetikus melepasi ayat** dan selepas dilepaskan, bar anotasi akan muncul di atas teks:

![Bar anotasi selepas memilih teks](img/pdf-highlight.png)

2. 2. Pilih **sorotan** yang contoh kuningnya membuka palet warna, **garisan bawah**, atau **potongan garis**, dan **Tanya AI** menghantar pilihan itu bersama soalan anda ke panel AI.
3. Untuk membatalkan anotasi, pilih semula petikan yang sama dan klik butang yang sedang aktif pada bar tersebut, atau pilih dan tekan Delete.

Butiran yang perlu diketahui:

- Menu lungsur muncul apabila anda menyeret teks, dan ia menyediakan **sorotan / garisan bawah / potongan garis / salin / Tanya AI**.
- Warna dipilih daripada palet tersebut, dan **memasang tanda yang sama pada julat yang sudah ditandakan akan mengeluarkannya**, seperti suis dalam Word.
  **FUNKTION** mengeluarkan tanda itu, seperti suis dalam Word.
- Tanda yang telah disimpan ke dalam fail masih boleh dipilih dan dipadam melalui menu atau kekunci Delete.
- **Nota**: semasa alat lukis diaktifkan, lapisan teks tidak boleh dipilih. Alat itu menekan dirinya sendiri selepas setiap bentuk diletakkan, jadi anda kembali ke mod pilih selepas itu.

## Alat lukis (Anotasi)

Enam alat disediakan: **dakwat, segi empat, elips, anak panah, dan nota**, serta **kotak redact** pada tab Anotasi.

- Setiap alat berfungsi sebagai suis. Klik untuk mengaktifkannya, dan ia akan **mematikan dirinya sendiri sebaik sahaja satu bentuk diletakkan**. Klik semula pada alat itu untuk meneruskan secara berturutan, dan klik pada alat yang sedang aktif akan mematikannya.
- Dakwat mengikut lebar garisan, segi empat, elips dan anak panah dilukis dengan menyeret, dan warna diambil daripada palet lukis.
- Bentuk yang diletakkan boleh dipilih, dipadam, diseret, dan dalam kes segi empat atau elips, diubah saiznya.
- **Redact, aliran kerja penuh** untuk menyembunyikan satu baris teks:

  1. Pergi ke tab Anotasi dan klik **Redaksi kawasan** untuk mengaktifkan alat itu.

  2. **Seret satu kotak di atas kandungan** supaya ia ditutup dengan tanda bercorak, dan bar alat akan memperoleh dua butang **Kosongkan tanda** serta **Gunakan redaksi**.

  ![Halaman selepas menandakan redact](img/pdf-redact.png)

  3. Klik **Gunakan redaksi** dan sahkan. Hasilnya ialah salinan kerja yang teks dan imej yang ditutup itu akan dibuang secara fizikal, dan tindakan ini tidak boleh dibatalkan. Dokumen asal kekal tidak berubah.

  Jika silap, tekan butang kosongkan tanda untuk menghapuskan tanda semasa dan melukis semula.

## Nota melekat dan utas komen

- **Alat nota** menetapkan pin dan membuka kad pada margin untuk teks tersebut, dengan nama pengarang yang boleh dikonfigurasikan. Setelah disahkan, ia disimpan sebagai anotasi teks PDF biasa.
- Klik pin untuk membuka utas, yang membolehkan anda **balas**, **sunting** komen sendiri, atau **padam** satu komen atau keseluruhan utas.
- Suntingan yang sedang berjalan kekal sehingga proses menyimpan menulis teks baharu itu ke dalam anotasi yang sama, dan ini mengekalkan rantaian balasan.

## Menyunting kandungan PDF (Edit)

- **Sunting teks**: klik teks untuk menyuntingnya blok demi blok, menggunakan enjin pdfium dan memadankan fon sebaik mungkin.
- **Sisip teks**: masukkan teks yang boleh dicari dengan pilihan fon, saiz, dan warna.
- **Sisip imej atau set**.
- Borang: medan AcroForm boleh diisi terus, dan nilainya ditulis semasa menyimpan.

## Tandatangan

- **Tandatangan dakwat**: lukis tanda tangan itu dan ia boleh dikaitkan dengan medan tanda tangan borang.
- **Tandatangan imej**: letakkan gambar sebagai tanda tangan.
- Tandatangan yang telah disimpan boleh digunakan semula.

## Operasi halaman (Halaman)

- **Putar, padam, atau susun semula**: seret lakaran kenit untuk menyusun semula, dan pemadaman memerlukan pengesahan.
- **Import halaman**: seret halaman daripada PDF lain ke dalam dokumen. **Sisipkan halaman kosong** menambah satu halaman kosong.
- **Ganti halaman** menukar satu julat dengan halaman dari tempat lain, manakala **Pangkas halaman** memotong tepinya, dengan pilihan untuk menerapkannya kepada semua halaman.
- **Saiz halaman** menskalakan semula setiap halaman kepada satu saiz kertas yang sama, dan **Terbalikkan susunan** membalikkan dokumen dari hujung ke hujung.
- **Kekalkan halaman**: eksport halaman yang dipilih ke PDF baharu.
- **Pisah PDF**: ada dua bentuk — pecahkan mengikut julat menjadi beberapa fail, atau potong setiap halaman menjadi grid halaman yang lebih kecil.
- **Gabung PDF**: ada dua bentuk — tambah PDF lain, atau gabungkan beberapa halaman ke atas satu helaian. Saiz akan dikira **sebelum** apa-apa yang lain dibaca, dan jumlah melebihi 1 GiB akan ditolak dengan ralat yang jelas, bagi memastikan penggunaan memori kekal terhad.
- Perubahan peringkat halaman ditulis semula semasa simpanan seterusnya, dan pilihan Simpan Sebagai membiarkan fail asal tidak berubah.

## Eksport dan cetakan

- **Eksport sebagai Word… / PowerPoint… / Excel…** dalam menu Fail, atau ketiga-tiganya daripada **Tukar PDF** pada reben — semuanya setempat, tiada muat naik. Fail .pptx keluar dengan satu slaid bagi setiap halaman dan fail .xlsx dengan satu lembar kerja bagi setiap halaman. Setiap satu akan bertanya di mana untuk disimpan.
- **Cetak** menggunakan susunan halaman dan putaran semasa melalui dialog sistem, dan julat halaman disokong.

## Simpanan

- Simpanan biasa, sama ada manual atau automatik, menulis anotasi dan suntingan kembali ke dalam fail secara atomik.
- **Penggunaan redaksi** melalui aliran Guna sendiri menghasilkan satu salinan, sambil membiarkan fail asal tidak berubah, jadi kandungan sensitif tidak kekal di dalamnya.
