# Mula pantas: antara muka dan asas

GenOffice ialah satu pek pejabat yang berjalan sepenuhnya pada mesin anda: satu tingkap, satu baris tab, dengan enam editor — Docs (penyuntingan teks), Sheets (hamparan), Slides (persembahan), PDF, Markdown dan HTML. Fail yang digunakan ialah fail .docx / .xlsx / .pptx / .pdf yang sebenar, serasi sepenuhnya dengan Word, Excel dan PowerPoint. Tiada rangkaian diperlukan.

## Gambaran keseluruhan antara muka

![Skrin Laman Utama](img/home-screen.png)

Tingkap ini mempunyai tiga bahagian:

- **Baris tab (atas)**: setiap fail yang dibuka ialah satu tab. Tab Laman Utama, paling kiri, sentiasa ada dan tidak boleh ditutup; tab lain ialah dokumen anda. Klik dua kali pada tab untuk menamakan semula failnya terus di tempat.
- **Kawasan kandungan**: editor (atau Laman Utama) yang berkaitan dengan tab aktif.
- **Bar menu**: dalam bar menu sistem pada macOS, di atas tingkap pada Windows/Linux. Menu Fail/Edit/Paparan berubah mengikut editor yang aktif.

## Mencipta dokumen

Mana-mana satu daripada berikut:

- Klik pada kad penciptaan pantas dalam bahagian [Mula pantas](help://getting-started) pada **Laman Utama** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **Fail ▸ Baharu**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML atau PDF.
- Seret fail ke dalam tingkap, atau klik dua kali pada fail dalam pengurus fail anda (jika GenOffice ialah aplikasi lalai).

Dokumen baharu dibuka tanpa tajuk; fail pada cakera hanya dicipta semasa simpanan pertama.

## Membuka fail

- Menu **Fail ▸ Buka…** (⌘O/ctrl+O) membuka pemilih sistem: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Klik mana-mana item dalam senarai **Terkini** pada Laman Utama.
- `genoffice <fail>` daripada terminal turut membuka fail.

## Model simpanan

- **Simpanan manual**: ⌘S/ctrl+S, atau Fail ▸ Simpan / Simpan Sebagai… Semasa simpanan pertama, anda akan ditanya tentang lokasi dan nama.
- **Autosimpan** hanya diaktifkan selepas anda menyimpan fail secara manual sekurang-kurangnya sekali — PDF yang anda baca sahaja tidak akan ditulis semula secara senyap. Autosimpan berlaku tidak lama selepas kandungan berubah.
- Menutup tab dengan perubahan yang belum disimpan akan meminta anda memilih Simpan / Buang / Batal dahulu.
- Setiap tulisan bersifat atomik (fail sementara + penamaan semula), jadi gangguan bekalan elektrik tidak akan meninggalkan fail yang separuh jadi.

## Pintasan lazim

| Tindakan        | macOS | Windows / Linux |
| --------------- | ----- | --------------- |
| Dokumen baharu  | ⌘N    | ctrl+N          |
| Buka            | ⌘O    | ctrl+O          |
| Simpan          | ⌘S    | ctrl+S          |
| Tutup tab       | ⌘W    | ctrl+W          |
| Buka manual ini | F1    | F1              |
| Runtuhkan Reben | ⌥⌘R   | Ctrl+F1         |

**Menutupkan reben** berfungsi dalam setiap editor. Baris tab kekal ada dan jalur arahan di bawahnya hilang; tab yang dipilih sekali gus menjadi kawalan lipatan, jadi selagi reben terlipat tiada tab dipilih dan menekan mana-mana tab akan mengembalikan jalur itu. Klik dua kali pada tab melakukan perkara yang sama. Cara anda meninggalkannya diingati bagi setiap editor.

Pintasan dalam setiap editor (berus format, cari & ganti, operasi jadual, ...) diterangkan dalam bab masing-masing; Docs turut menyediakan dialog pintasan papan kekunci yang boleh dicari (**⌘/**) (lihat babnya).

## Pintasan Option+Command

Option+Command ialah lapisan yang dikhaskan Word untuk lompatan berstruktur, dan GenOffice mengisinya dengan cara yang sama. Docs mengambil majoritinya, Sheets mengambil dua miliknya sendiri demi kesetaraan dengan Excel, dan satu pintasan berfungsi di mana-mana sahaja.

**Docs**

| Pintasan (macOS) | Apa yang ia lakukan  | Windows / Linux    |
| ---------------- | -------------------- | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3  | Tajuk 1 / 2 / 3      | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0              | Normal               | Ctrl+Alt+0         |
| ⌥⌘M              | Perenggan            | Ctrl+Alt+M         |
| ⌥⌘A              | Komen Baharu         | Ctrl+Alt+A         |
| ⌥⌘F              | Sisipkan Nota Kaki   | Ctrl+Alt+F         |
| ⌥⌘E              | Sisipkan Nota Hujung | Ctrl+Alt+D         |
| ⌥⌘G              | Pergi ke             | Ctrl+G             |

Dua daripada bertukar pada Windows, atas sebab yang sama seperti yang membuat Word memisahkannya. **macOS memiliki ⌥⌘D** — ia memaparkan dan menyembunyikan Dock — jadi nota hujung ialah ⌥⌘E pada Mac dan Ctrl+Alt+D di tempat lain. Dan **Pergi ke** melepaskan Alt: Ctrl+G, manakala pintasan pada Mac membawanya sekali.

**Sheets**, semasa grid mempunyai fokus

| Pintasan (macOS) | Apa yang ia lakukan | Windows / Linux |
| ---------------- | ------------------- | --------------- |
| ⌥⌘0              | Sempadan luar       | Ctrl+Shift+7    |
| ⌥⌘−              | Tiada sempadan      | Ctrl+Shift+−    |

Windows bukan tulisan semula pasangan Mac. Excel untuk Mac memberi Sheets **keduanya** — ⌘⇧7 dan ⌥⌘0 ialah dua kekunci untuk sempadan luar yang sama — jadi pada Windows arahan itu mengekalkan slot Ctrl+Shift yang memang sudah dimilikinya, manakala lapisan Option itu sahaja tiada.

Perhatikan bahawa **⌥⌘0 bermaksud Normal dalam Docs dan Sempadan luar dalam Sheets**. Keduanya tidak pernah muncul dalam editor yang sama, jadi tiada pertikaian dalam penggunaan, tetapi ⌥⌘0 sudah terpakai dan tidak tersedia sebagai pintasan global.

**Setiap editor**: **⌥⌘R / Ctrl+F1** menutupkan reben, seperti yang diterangkan di atas.

Maka ⌥⌘D kekal bebas untuk GenOffice gunakan pada macOS, jika ada perintah pada masa depan yang memerlukannya.

## Ke mana seterusnya

- Di mana fail berada: [Skrin Laman Utama](help://home-screen).
- Menguruskan ramai fail yang dibuka: [Tab dan pengurusan tingkap](help://tabs-and-windows).
- Membiarkan AI mengendalikan kerja: [Panel pembantu AI](help://ai-panel).
- Bahasa, tema, aplikasi lalai: [Tetapan, bahasa, tema dan integrasi MCP](help://settings-integrations).
