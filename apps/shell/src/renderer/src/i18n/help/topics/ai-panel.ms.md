# Panel pembantu AI

Setiap editor boleh memanggil panel AI. Pilih sesuatu, berikan arahan, dan lihat hasilnya dipaparkan secara berperingkat.

## Membuka dan menggunakannya

![Panel AI dalam Docs](img/ai-panel.png)

- Akses kepada panel ini ialah melalui **butang AI** pada reben setiap editor, **Tanya AI** dalam menu konteks, atau Tanya AI pada bar anotasi.
- Terangkan tugas itu dalam bahasa biasa, kemudian tekan Enter.
- Balasan dipaparkan secara berperingkat. Apabila AI memerlukan alatan, ia akan menggunakan alatan itu dan meneruskan sehingga selesai.
- **Berhenti** boleh digunakan pada bila-bila masa untuk menghentikan pusingan semasa.
- **Baris gilir**: ketika balasan sedang mengalir, Enter memasukkan mesej anda ke baris gilir dan bukannya membuangnya — jalur di atas kotak input mengiranya; buka untuk sunting, buang, susun semula, jeda atau kosongkan. Mesej seterusnya terhantar sendiri apabila giliran tamat; Esc masih menghentikan giliran.

## Apa yang boleh dibuatnya

- **Docs**: menulis semula, melanjutkan, menterjemah, meringkas, dan memasukkan jadual serta imej. Setiap pusingan mengambil gambar keadaan lebih dahulu.
- **Sheets**: formula, mengisi data, transformasi pukal, dan pemformatan.
- **Slides**: menjana keseluruhan persembahan, menyelaras susun atur, dan menulis semula teks.
- **PDF**: menjawab soalan dan meringkas berdasarkan teks atau halaman yang dipilih.
- **Markdown dan HTML**: menulis semula, melanjutkan, dan menterjemah.
- **Laman Utama ▸ Ask AI** menjawab daripada fail anda yang berbintang dan memetiknya — panel pangkalan pengetahuan ada bahagiannya sendiri dalam bab Laman Utama, dan panel AI Docs membawa kemahiran yang sama.

## Berundur dan keselamatan

- Panel dalam Docs menyimpan **senarai versi** dengan satu gambar keadaan bagi setiap pusingan. Anda boleh berundur ke mana-mana versi dalam senarai itu, dan berundur itu sendiri juga boleh dibatalkan. Gambar keadaan itu kekal walaupun dokumen ditutup dan dibuka semula.
- Suntingan AI menggunakan laluan penyuntingan yang sama seperti suntingan manual, jadi semuanya boleh dibatalkan dan tidak disimpan sehingga anda mengesahkannya.

## Menyembunyikan pemilihan daripada AI

- Dalam Slides, HTML dan Markdown, pemilihan boleh dijauhkan daripada mata model: klik kanan ▸ **Sembunyikan pemilihan daripada AI**, beri nama pemegang tempat, dan model membaca `{{nama}}` sebagai gantinya. (Markdown melengkapkan item ini dengan togol mata-tercorek pada rebennya; milik HTML ada di panel sumber.)
- Fail menyimpan teks sebenar — simpan, eksport dan cetak tidak terganggu, dan tanda itu kekal selepas dibuka semula. Item yang sama membatalkan penyembunyian (Markdown ada **Tunjukkan pemilihan kepada AI semula** yang berasingan).
- HTML menyembunyikan kandungan elemen, nilai atribut dan literal rentetan dalam skrip; halaman yang diterjemahkan tidak pernah berubah.
- Slides turut menyembunyikan gambar, video dan audio secara keseluruhan, dan slaid dengan kandungan tersembunyi tidak pernah dihantar sebagai tangkapan skrin: semakan letak automatik melangkau halaman itu dan menyebutnya.
- Ini hanya untuk AI dalam aplikasi — CLI, alat MCP dan eksport tanpa kepala membaca fail secara terus dan masih nampak teksnya.

## Privasi

- Arahan dan kandungan dokumen yang berkaitan dihantar ke **perkhidmatan model yang anda konfigurasikan**, iaitu Genspark yang dihoskan atau titik akhir tersuai. Tanpa konfigurasi, tiada apa-apa yang dihantar.
- Fail setempat tidak dimuat naik ke tempat lain. Kunci BYOK disimpan dalam fail tetapan aplikasi pada mesin ini dan dihantar hanya dalam pengepala permintaan.
