# Skrin Laman Utama: di mana fail anda berada

Laman Utama ialah halaman permulaan GenOffice: bar navigasi di kiri, senarai fail dan kad penciptaan pantas di kanan.

![Skrin Laman Utama](img/home-screen.png)

## Navigasi bar sisi

- **Terkini**: fail yang baru dibuka oleh anda. Setiap baris dicap dengan waktunya — hari ini, semalam atau tarikh.
- **Berbintang**: fail yang anda tandakan dengan bintang. Halakan kursor ke baris fail dan klik bintang untuk menambah atau membuangnya.
- **Panduan Pengguna**: membuka manual ini.
- **Genspark Projects**: selepas anda log masuk ke akaun Genspark, ia memaparkan projek yang anda cipta di web dengan Genspark AI. Klik satu untuk menyuntingnya terus dalam pelayar. Carian, isihan mengikut masa, muat semula dan muat lagi disokong.
- **Folder**: direktori yang anda tambah ke bar sisi dengan **Tambah folder…**, atau yang anda seret masuk daripada pengurus fail. Setiap satu menjadi lokasi yang boleh dibuka, boleh dicipta subdirektori di dalamnya, boleh dinamakan semula dan boleh dibuang; lokasi yang menjadi luar talian ditandakan sebagai tidak tersedia dan boleh dibuang daripada senarai. **Folder baharu** mencipta satu lagi.

Tiada entri Tong sampah di sini. Fail yang dipadam akan masuk ke tong sampah sistem, dan memulihkannya ialah urusan sistem pengendalian.

## Senarai fail

Setiap baris memaparkan ikon, nama fail, masa diubah suai dan maklumat lain. **⋯ menu** pada baris itu menawarkan:

- **Buka**, dan **Tunjukkan dalam folder** untuk mengesahkan lokasi fail dalam pengurus fail anda.
- **Salin laluan**.
- **Pindah ke folder…**: membuka pemilih folder dan benar-benar memindahkan fail; jika di destino sudah ada fail dengan nama yang sama, anda boleh melangkau, menimpa atau menukar namanya.
- **Namakan semula**: menamakan semula terus di tempat, dengan sambungan dikekalkan secara automatik.
- **Tambah bintang / Buang bintang** — bintang kekal selepas aplikasi ditutup dan mengikut fail apabila anda menamakannya semula.
- **Buat salinan**: mencipta salinan dalam folder yang sama.
- **Padam**: memindahkan fail ke tong sampah sistem. Ini bukan pemadaman kekal.
- **Alih keluar daripada senarai**, pada paparan **Terkini** peringkat atas, untuk membuang satu entri tanpa menyentuh fail itu sendiri.

### Beberapa fail sekali gus

Tandakan kotak pada satu baris, atau ⌘/ctrl-klik, untuk membina satu pilihan; kotak di pengepala memilih segala yang kini disenaraikan, dan satu bar di atas senarai melaporkan berapa yang dipilih (**{n} dipilih**) bersama **Pindah ke folder…** dan **Padam fail** untuk keseluruhan kumpulan. Anda juga boleh menyeret pilihan berbilang ke mana-mana folder pada bar sisi.

## Carian

Kotak carian di atas menapis dua perkara serentak:

- **Nama fail**: penapisan pantas mengikut nama.
- **Kandungan fail**: GenOffice mengindeks fail anda di latar belakang, iaitu teks di dalam docx/xlsx/pptx/pdf/md/html, dengan OCR sebagai sandaran bagi PDF yang diimbas. Oleh itu, mencari dalam teks badan turut menemui fail. Skop dan suis ditetapkan dalam tetapan carian.

## Kad mula pantas

Kad di atas senarai menubuhkan dokumen baharu dalam satu langkah. Klik pada kad akan mencipta fail jenis tersebut dan membuka editornya. Anda boleh terus menulis, atau membiarkan AI menyediakan draf untuk anda. Setiap editor mempunyai **butang AI** pada reben dan **Tanya AI** dalam menu konteks pilihan.

Fail baharu akan mendarat dalam folder yang sedang dipilih dalam bar sisi. Jika tiada apa-apa yang dipilih, ia akan ke folder lalai.

Apa yang setiap kad lakukan:

- **AI Docs** (.docx): dokumen teks kosong dalam editor Docs. Fail hanya ditulis ke cakera semasa **simpanan pertama**. Dokumen baharu dibuka dengan panel AI dikembangkan, dan anda boleh matikan melalui Tetapan → "Buka panel AI dalam dokumen baharu".
- **AI Sheets** (.xlsx): hamparan kosong dalam editor Sheets. Sebelum anda simpan, tiada fail wujud pada cakera; nama itu dikhaskan untuk simpanan pertama. Selepas penjanaan AI yang pertama, fail tersebut juga boleh dinamakan semula secara automatik berdasarkan kandungannya.
- **AI Slides** (.pptx): persembahan kosong dalam editor Slides.
- **AI Markdown** (.md): dokumen Markdown kosong dalam editor Markdown.
- **AI HTML** (.html): laman web kosong dalam editor HTML.
- **AI PDF** (.pdf): berbeza daripada yang lain. Ia **segera** mencipta PDF kosong satu halaman yang sebenar dalam folder sasaran, kemudian membukanya sebagai fail biasa. Editor PDF beroperasi pada fail sebenar. Kad ini sesuai untuk membuat anotasi, redact atau menambah teks, dan semasa simpanan pertama fail itu boleh dinamakan semula secara automatik berdasarkan kandungannya.
- **Buka Fail Setempat**: pemilih fail sistem untuk Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls/.csv/.tsv), PowerPoint (.pptx/.ppt), PDF, Markdown (.md/.markdown) dan laman web (.html/.htm). Pelbagai fail boleh dipilih, dan setiap fail mendapat tabnya sendiri.

> Petua: Fail ▸ Baharu pada bar menu mencipta jenis dokumen yang sama. Pintasan ⌘N/Ctrl+N secara lalai mencipta dokumen teks. Menyeret fail ke dalam tingkap akan membukanya.

## Projek awan (Genspark Projects)

- Kegunaan kali pertama memerlukan anda log masuk ke akaun Genspark melalui aliran kod peranti. GenOffice memaparkan satu kod, dan anda melengkapkan log masuk dalam pelayar.
- Senarai projek disegerakkan dengan web. Arahan Buka dalam pelayar membawa anda ke sana untuk menyambung.
- Tidak log masuk tidak menjejaskan sebarang fungsi setempat.
