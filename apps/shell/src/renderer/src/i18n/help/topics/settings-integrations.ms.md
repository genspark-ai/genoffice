# Tetapan, bahasa, tema dan integrasi MCP

## Membuka tetapan

Baris akaun di penjuru kiri bawah pada Laman Utama akan membuka panel tetapan. Ia memaparkan Log masuk apabila anda belum log masuk. Panel ini mempunyai enam bahagian: Akaun, Model AI, Media & Carian AI, Umum, Integrasi dan Perihal.

![Tetapan ▸ Umum, tempat bahasa, tema, autosimpan dan suis statistik penggunaan berada](img/settings-general.png)

Konfigurasi model mempunyai babnya sendiri; dalam **Media & Carian AI** anda menghidupkan penjanaan imej, analisis imej, analisis video, carian web dan carian fail tempatan bagi setiap pembekal.

## Bahasa

- Tetapan menawarkan **21 bahasa antara muka**: Inggeris, Cina ringkas, Jepun, Korea, Perancis, Jerman, Sepanyol, Thai, Indonesia, Rusia, Arab, Portugis, Itali, Poland, Czech, Belanda, Melayu, Ibrani, Hindi, Cina tradisional, dan Vietnam.
- Pertukaran bahasa berkuat kuasa serta-merta dan kekal, manakala bar menu sistem dibina semula dalam bahasa baharu itu.

## Tema

Tiga pilihan ialah Cerah, Gelap, dan Ikut Sistem. Mod Ikut Sistem mengikut rupa sistem pengendalian, dan editor akan mengubah warnanya mengikut tema itu tanpa kelip.

## Umum

- **Hantar statistik penggunaan tanpa nama** — diaktifkan secara lalai. Menggunakan Google Analytics 4 dan menghantar IP awam serta metadata pengangkutan anda; kandungan dokumen dan nama fail tidak pernah dikumpulkan, dan setiap peristiwa hanya membawa satu jenis seperti "membuka .docx". Anda boleh mematikannya di sini pada bila-bila masa.
- **Kedudukan bar sisi AI** (kiri atau kanan), **Saiz teks panel AI** dan **Semakan ejaan dalam sembang AI**.
- **Buka panel AI dalam dokumen baharu** — dimatikan, dokumen baharu bermula dengan panel tertutup, sekali klik sahaja.
- **Simpan automatik semua dokumen** membolehkan AutoSave secara lalai dalam setiap editor; anda masih boleh mematikannya untuk satu tetingkap sahaja.
- **Lokasi simpanan** dengan butang **Tukar**, dan **Apl lalai untuk dokumen Office** untuk menuntut .docx / .xlsx / .pptx bagi GenOffice.

## Media & Carian AI

Ia bukan suis — setiap keupayaan memilih pembekal yang menyediakannya, dan kunci serta URL asas setiap pembekal dimasukkan sekali sahaja lalu dikongsi bersama:

- **Carian web**, **Penjanaan imej**, **Analisis imej** dan **Analisis video**, setiap satu dengan pembekal, model, kunci dan URL asas.
- **Carian fail tempatan** dijalankan pada mesin ini. Di bawahnya terdapat **Susunan semula Jev** yang **dimatikan secara lalai**. Hidupkannya dan petikan 20 hasil tempatan teratas — sehingga 1,200 aksara daripada setiap dokumen, berserta nama fail dan folder — dihantar ke model Jev TypeSafe untuk disusun semula mengikut kaitan. Jika ia dimatikan, tiada apa-apa meninggalkan peranti.

## Perihal

- **Versi**, pautan GitHub projek dan butang **Beri Bintang di GitHub**.
- **Saluran Kemas Kini**: Stabil atau Beta. Menukarnya berkuat kuasa serta-merta dan akan menyemak kemas kini; ia tidak akan menurunkan pemasangan Beta kembali ke Stabil.

## Penetapan sebagai aplikasi lalai

Tetapan boleh mendaftar GenOffice sebagai aplikasi yang mengendalikan fail .docx, .xlsx, .pptx, .pdf, dan yang lain. Pendaftaran ini dilakukan pada peringkat sistem, dan pengesahan diperlukan apabila diminta.

## Notis perisian pihak ketiga dan kemas kini

- Bantuan ▸ Notis Perisian Pihak Kedua memaparkan senarai penuh lesen perisian sumber terbuka yang disertakan dengan aplikasi.
- Bantuan ▸ Semak Kemas Kini memulakan pemeriksaan secara manual. Jika terdapat versi yang lebih baharu, aplikasi akan meminta anda memasangnya.

## Log masuk ke Genspark

- Kemasukan log masuk, sama ada melalui tetapan atau senarai projek awan, menggunakan aliran **kod peranti**. GenOffice memaparkan kod tersebut dan membuka halaman log masuk dalam pelayar, dan aliran ini akan diteruskan secara automatik selepas selesai.
- Log masuk hanya digunakan untuk dua perkara, iaitu senarai projek awan dan model yang dihoskan oleh Genspark. Tanpa log masuk, setiap fungsi setempat dan model tersuai akan terus berfungsi.
- Log keluar boleh dilakukan dengan satu klik dalam tetapan.

## Integrasi MCP

**Integrasi** ialah panel yang menyambungkan GenOffice kepada ejen penulisan kod, dan ia ada babnya sendiri: Menyambungkan ejen penulisan kod. Versi ringkas — pilih satu laluan (baris arahan, atau MCP), ikut bahagian itu, kemudian mulakan sembang baharu dan bertanya.

![Tetapan ▸ Integrasi: tiga langkah, kemudian baris kemahiran dan pilihan MCP](img/settings-integrations.png)

Di bawah **Pelayan HTTP tempatan**, aplikasi ini juga boleh menjalankan pelayan itu sendiri — suis pendayakan dan port — dan **Lanjutan** menambah URL semakan kesihatan serta fail log, dan bukannya meninggalkannya kepada pembantu. Ia hanya mendengarkan pada localhost.

## Ringkasan baris arahan

| Perintah           | Fungsinya                             |
| ------------------ | ------------------------------------- |
| `genoffice <fail>` | membuka fail                          |
| `genoffice mcp`    | memulakan pelayan MCP tempatan        |
| `genoffice --help` | memaparkan semua perintah dan pilihan |
