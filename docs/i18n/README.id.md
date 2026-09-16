<p align="center">
  <a href="https://genoffice.ai/">
    <picture>
      <source srcset="../assets/readme/hero-dark.webp" media="(prefers-color-scheme: dark)">
      <img src="../assets/readme/hero.webp" alt="GenOffice — suite Office AI open-source: Docs, Sheets, Slides, PDF, Markdown, dan HTML dengan panel AI bawaan" width="100%">
    </picture>
  </a>
</p>

<h1 align="center">GenOffice</h1>

<p align="center"><b>Suite Office AI open-source pertama di dunia yang paling lengkap.</b><br>
File Word, Excel, PowerPoint, dan PDF, diedit oleh Anda dan AI Anda, disimpan kembali dalam format aslinya.</p>

<p align="center">
  <a href="../../LICENSE"><img src="https://img.shields.io/github/license/genspark-ai/genoffice" alt="License: Apache-2.0"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases/latest"><img src="https://img.shields.io/github/v/release/genspark-ai/genoffice" alt="Latest release"></a>
  <a href="https://github.com/genspark-ai/genoffice/releases"><img src="https://img.shields.io/github/downloads/genspark-ai/genoffice/total" alt="Downloads"></a>
  <a href="https://github.com/genspark-ai/genoffice/stargazers"><img src="https://img.shields.io/github/stars/genspark-ai/genoffice?style=flat" alt="GitHub stars"></a>
</p>

<p align="center"><a href="../../README.md">English</a> · <a href="README.es.md">Español</a> · <a href="README.pt-BR.md">Português (Brasil)</a> · <a href="README.de.md">Deutsch</a> · <a href="README.fr.md">Français</a> · <a href="README.zh-CN.md">简体中文</a> · <a href="README.zh-TW.md">繁體中文</a> · <a href="README.ko.md">한국어</a> · <a href="README.ja.md">日本語</a> · <a href="README.ar.md">العربية</a> · <a href="README.ru.md">Русский</a> · <a href="README.it.md">Italiano</a> · <a href="README.nl.md">Nederlands</a> · <a href="README.pl.md">Polski</a> · <a href="README.cs.md">Čeština</a> · <b>Bahasa Indonesia</b> · <a href="README.ms.md">Bahasa Melayu</a> · <a href="README.th.md">ไทย</a> · <a href="README.hi.md">हिन्दी</a> · <a href="README.he.md">עברית</a></p>

<p align="center">
  <a href="#download"><b>Unduh</b></a> ·
  <a href="#command-line-and-agent-skill"><b>CLI</b></a> ·
  <a href="#mcp-server"><b>MCP</b></a> ·
  <a href="https://genoffice.ai/"><b>Situs Web</b></a> ·
  <a href="https://genoffice.ai/join"><b>Komunitas</b></a> ·
  <a href="../../PRIVACY.md"><b>Privasi</b></a>
</p>

GenOffice adalah alternatif Microsoft Office yang gratis dan open-source
untuk macOS, Windows, dan Linux. Aplikasi ini membuka dan menyimpan file
`.docx`, `.xlsx`, dan `.pptx` asli, mengedit PDF, Markdown, dan HTML, serta
menghadirkan agen AI di samping setiap dokumen — bukan sekadar kotak obrolan
yang ditempelkan di sisi layar, melainkan editor yang membaca file, membuat
perubahan, dan menunjukkan dengan tepat bagian mana yang diubah.

- **Format asli, byte-preserving.** Hanya bagian yang Anda edit yang ditulis
  ulang. Bagian lain dari file tetap sama persis, byte demi byte, sehingga
  dokumen tetap berfungsi normal di Word, Excel, dan PowerPoint.
- **AI yang bisa Anda tinjau.** Perubahan tercatat sebagai tracked changes dan
  diff dengan opsi rollback satu klik. Spreadsheet mendapatkan rumus hidup,
  bukan angka yang ditempel. Deck dan halaman dibuat langsung di kanvas dan
  tetap bisa diedit secara penuh.
- **Lokal secara desain.** File dibuka, diedit, disimpan, dan dikonversi di
  perangkat Anda sendiri. Konversi PDF → Word / Excel / PowerPoint, Markdown →
  Word, dan HTML → Word semuanya berjalan on-device. Hanya panggilan AI yang
  meninggalkan perangkat, menuju penyedia yang Anda pilih.
- **Kunci API Anda sendiri, atau tanpa kunci sama sekali.** Masuk dengan
  Genspark dan lewati urusan kunci API, atau gunakan kunci Anda sendiri untuk
  Claude, OpenAI, Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok,
  Mistral, OpenRouter, Requesty, atau endpoint apa pun yang kompatibel dengan OpenAI,
  termasuk server lokal.
- **Dapat dijalankan lewat skrip dan siap untuk agen.** Aplikasi ini
  menyertakan baris perintah `genoffice` dan sebuah skill untuk Claude Code,
  Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode, dan Windsurf, sehingga
  coding agent bisa membuat, mengonversi, membaca, dan mengedit file Office
  asli di mesin Anda tanpa membuka satu jendela pun.

**Unduh di sini:** [macOS](https://github.com/genspark-ai/genoffice/releases/latest) (Apple Silicon dan Intel) ·
[Windows](https://github.com/genspark-ai/genoffice/releases/latest) (x64 dan Arm) ·
[Linux](https://github.com/genspark-ai/genoffice/releases/latest) (deb, rpm, AppImage) —
detail dan persyaratan ada di bagian [Download](#download).

## Demo

Enam aplikasi, satu panel AI, dan satu baris perintah untuk coding agent Anda.
Setiap tangkapan layar adalah aplikasi asli di macOS, dengan AI yang dijalankan
dari prompt yang bisa Anda baca di panel.

### 1 · Docs — buka dan edit `.docx` dengan AI yang bisa Anda tinjau

<table>
<tr>
<td width="50%"><img src="../assets/readme/docs-report.webp" alt="GenOffice Docs menampilkan halaman laporan tahunan dua kolom dengan gambar sampul selebar halaman, tabel KPI berlatar warna, header dan footer, pada zoom 80% dengan panel AI yang diciutkan"></td>
<td width="50%"><img src="../assets/readme/docs-ai.webp" alt="GenOffice Docs: ringkasan perusahaan dengan gambar banner; AI mempersingkat bagian Overview dan menyisipkan bagian berpoin baru, dan panel menawarkan rollback satu klik"></td>
</tr>
<tr>
<td><b>Membuka file sesuai tata letak Word</b> — bagian dua kolom, gambar full-bleed, tabel berlatar warna, header dan footer, penomoran halaman berdasarkan metrik baris Word. Gaya, komentar, tracked changes, rumus, dan tulisan tangan (ink) tetap utuh saat dibuka lagi.</td>
<td><b>Minta perubahan yang Anda inginkan</b> — AI membaca blok yang diperlukan, menulis ulang bagian Overview, dan menyisipkan bagian berpoin baru. Setiap giliran AI adalah snapshot yang bisa di-rollback; dengan <b>Track changes</b> aktif, perubahan muncul sebagai revisi bergaya Word.</td>
</tr>
</table>

### 2 · Sheets — `.xlsx` dengan rumus dan grafik hidup, bukan angka yang ditempel

<table>
<tr>
<td width="50%"><img src="../assets/readme/sheets-ai.webp" alt="GenOffice Sheets: AI menambahkan sheet Summary berisi pendapatan per wilayah dan kategori menggunakan rumus SUMIF, ditambah grafik kolom, dan melaporkan 43 perubahan yang diterapkan dengan tombol Undo"></td>
<td width="50%"><img src="../assets/readme/sheets-qa.webp" alt="GenOffice Sheets: ditanya wilayah mana yang memimpin pendapatan Q2, AI menjawab Eropa lengkap dengan rincian kategori dan mengutip sel yang digunakan sebagai tautan, di samping sheet Orders"></td>
</tr>
<tr>
<td><b>Bangun sheet-nya</b> — dari satu kalimat, agen menambahkan sheet Summary dengan rumus <code>SUMIF</code> yang sebenarnya per wilayah dan kategori, menyisipkan grafik kolom, dan menerapkan 43 perubahan sebagai satu batch yang bisa dibatalkan sekaligus.</td>
<td><b>Tanyakan langsung</b> — pertanyaan tentang workbook dijawab lengkap dengan penalarannya, dan sel yang dipakai muncul sebagai kutipan yang bisa diklik. Di baliknya: mesin `.xlsx` Rust buatan sendiri, pivot table, slicer, pemformatan bersyarat, dan penelusuran rumus.</td>
</tr>
</table>

### 3 · Slides — dari sebuah prompt menjadi deck `.pptx`

<img src="../assets/readme/slides-generate.webp" alt="Time-lapse GenOffice Slides membuat deck investor Aurora Home: AI menyusun alur cerita di panel, slide muncul satu demi satu di kanvas, dan deck yang selesai diakhiri dengan closing ask" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/slides-cover.webp" alt="GenOffice Slides: slide sampul deck investor Aurora Home yang dibuat AI di kanvas, dengan prompt satu baris aslinya dan ringkasan AI tentang apa yang dibuatnya di panel"></td>
<td width="50%"><img src="../assets/readme/slides-ai.webp" alt="GenOffice Slides: slide penutup yang telah didesain dari deck 11-slide yang sama, dengan strip thumbnail di kiri dan panel AI yang meringkas alur ceritanya"></td>
</tr>
<tr>
<td><b>Satu baris prompt masuk</b> — "Buat deck presentasi investor 10-slide untuk Aurora Home…". GenOffice merancang alur cerita, meriset angka-angkanya, dan menyusun setiap slide langsung ke kanvas sebagai `.pptx` yang sebenarnya.</td>
<td><b>Deck jadi keluar</b> — sebelas slide terdesain dengan tipografi, citra visual yang konsisten, dan call to action penutup; lanjutkan mengedit dengan master, layout, smart guide, dan pemotongan gambar non-destruktif, atau minta panel untuk mengubah gaya, menulis ulang, dan menyusun ulang urutan.</td>
</tr>
</table>

### 4 · PDF — edit teks PDF langsung di tempat, konversi PDF ke Word secara on-device

<table>
<tr>
<td width="50%"><img src="../assets/readme/pdf-edit.webp" alt="GenOffice PDF: mode Edit text menandai setiap blok teks di halaman untuk diedit langsung di tempat, sementara panel AI menjawab pertanyaan tentang laporan dengan kutipan halaman"></td>
<td width="50%"><img src="../assets/readme/pdf-convert.webp" alt="GenOffice Docs menampilkan dokumen Word yang dikonversi secara lokal dari PDF Helios quarterly review, dibuka di tab kedua di samping PDF aslinya"></td>
</tr>
<tr>
<td><b>Edit langsung di dalam halaman</b> — mode Edit text menandai setiap blok teks untuk diketik ulang langsung di tempat; content stream ditulis ulang melalui PDFium dengan font asli, bukan anotasi penutup. Tanyakan sesuatu ke AI tentang laporan panjang dan dapatkan jawaban dengan kutipan halaman.</td>
<td><b>Konversi on-device</b> — <b>PDF Converter → PDF to Word</b> menghasilkan `.docx` yang bisa diedit dan terbuka di Docs berdampingan dengan sumbernya, heading, baris statistik, dan paragraf tetap utuh. Target Excel dan PowerPoint bekerja dengan cara yang sama; halaman hasil pindai (scan) diproses lewat OCR sistem.</td>
</tr>
</table>

### 5 · HTML — pembuat halaman dan UI berbasis AI, dimulai dari design brief

Sampaikan halaman itu untuk apa dan untuk siapa. AI akan mengusulkan **design
brief** terlebih dahulu — hook, palet warna, tipografi, dan arah gaya visual —
lalu membangun satu file `.html` mandiri berdasarkan token-token tersebut.

<img src="../assets/readme/html-restyle-motion.webp" alt="Time-lapse GenOffice HTML mengubah gaya halaman landing Lumen: satu permintaan Restyle di panel mengubah halaman Midnight Studio yang gelap menjadi versi Solar Daybreak yang hangat, sementara setiap bagian dan semua konten tetap di tempatnya" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/html-ai.webp" alt="GenOffice HTML: landing page yang dihasilkan untuk lampu meja tenaga surya dalam arah gaya gelap Midnight Studio, ditampilkan di live preview dengan panel AI meringkas halaman yang baru dibuatnya"></td>
<td width="50%"><img src="../assets/readme/html-restyle.webp" alt="Landing page Lumen yang sama, diubah gayanya oleh AI menjadi arah hangat Solar Daybreak: latar seperti kertas, judul serif, dan aksen oranye, dengan setiap bagian dan seluruh teks tetap dipertahankan"></td>
</tr>
<tr>
<td><b>Dihasilkan dari satu prompt</b> — hero yang berani, kartu fitur, harga, dan formulir waitlist untuk Lumen, dibangun dalam arah Midnight Studio. Klik elemen apa pun untuk mengubah gayanya, klik dua kali untuk mengedit teks, atau beralih ke tampilan sumber CodeMirror.</td>
<td><b>Desain sama, arah baru</b> — satu permintaan <b>Restyle</b> mengganti token-token brief dan halaman ikut menyesuaikan: latar seperti kertas hangat, serif editorial, aksen oranye matahari, tanpa ada yang ditulis ulang. Presentasikan dalam mode layar penuh, atau ekspor sebagai PDF atau dokumen Word asli yang bisa diedit.</td>
</tr>
</table>
<table>
<tr>
<td width="50%"><img src="../assets/readme/html-dashboard.webp" alt="GenOffice HTML: UI dashboard pribadi yang dihasilkan untuk desainer lepas dalam gaya linen hangat, dengan rail kiri, sapaan bergaya serif, dan empat kartu metrik"></td>
<td width="50%"><img src="../assets/readme/html-report.webp" alt="GenOffice HTML: laporan data pasar kendaraan listrik yang dihasilkan dalam gaya broadsheet, dengan masthead serif, angka utama 17,3 juta, dan baris statistik"></td>
</tr>
<tr>
<td><b>Mockup UI</b> — starter "personal dashboard" mengubah sebuah persona menjadi tata letak yang benar-benar berfungsi: rail kiri, sapaan, sparkline jam tertagih, kartu invoice dan utilisasi, semuanya HTML asli yang bisa Anda serahkan ke developer.</td>
<td><b>Data story</b> — starter "data report" membangun broadsheet editorial: masthead serif, satu angka utama, baris statistik yang dipisah garis, grafik SVG inline, dan catatan metodologi.</td>
</tr>
</table>

### 6 · Markdown — block editor di atas `.md` biasa, dengan Ask AI

<table>
<tr>
<td width="50%"><img src="../assets/readme/markdown-ai.webp" alt="GenOffice Markdown: paragraf yang dipilih menampilkan popover Ask AI dengan instruksi yang diketik dan chip saran seperti Polish, Make more concise, Expand, dan Fix grammar, plus tombol Send now dan Add to queue"></td>
<td width="50%"><img src="../assets/readme/markdown-render.webp" alt="GenOffice Markdown merender dokumen catatan peluncuran dengan tabel, flowchart Mermaid, dan daftar tugas, dengan prompt starter panel AI di sebelah kiri"></td>
</tr>
<tr>
<td><b>Ask AI pada bagian yang dipilih</b> — pilih teks apa pun dan chip <b>Ask AI</b> akan muncul: ketik instruksi atau pilih saran, kirim sekarang, atau kumpulkan beberapa edit berjangkar dan jalankan sekaligus dalam satu proses. Entri yang sama tersedia di setiap aplikasi.</td>
<td><b>Dirender, disimpan sebagai Markdown biasa</b> — heading, daftar, tabel, gambar, blok kode, dan diagram Mermaid dalam block editor Tiptap, ditulis kembali sebagai `.md` biasa, dengan ekspor <b>Markdown → Word</b> yang berjalan sepenuhnya lokal.</td>
</tr>
</table>

### 7 · CLI — coding agent Anda menjalankan GenOffice, di mesin Anda

GenOffice menyertakan baris perintah `genoffice` dan sebuah skill agen. Pasang
skill itu, dan Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot, OpenCode,
atau Windsurf bisa membuat, mengonversi, membaca, dan mengedit file Office asli
lewat engine yang sama dengan aplikasinya, tanpa membuka satu jendela pun.

<img src="../assets/readme/cli-deck-in-app.webp" alt="GenOffice Slides menampilkan deck Tata Surya delapan slide yang dibuat coding agent lewat baris perintah genoffice: slide sampul di kanvas, delapan thumbnail di kiri, dan panel AI terbuka" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/cli-slides-grid.webp" alt="Delapan slide hasil render dari deck Tata Surya berdampingan: sampul, linimasa eksplorasi, empat angka kunci, grafik batang diameter planet, planet berbatu versus planet raksasa, angka utama 99,8% milik Matahari, grid empat planet raksasa, dan kesimpulan"></td>
<td width="50%"><img src="../assets/readme/cli-integrations.webp" alt="GenOffice Settings, halaman Integrations: skill genoffice terpasang di Claude Code, dengan tombol Install di samping Codex dan Cursor"></td>
</tr>
<tr>
<td><b>Satu prompt ke agen Anda</b> — "Buat deck delapan slide tentang Tata Surya." Agen membaca skill, menulis style sheet, outline, dan satu spesifikasi halaman per slide, menghasilkan dua foto dengan <code>genoffice image</code>, dan membiarkan <code>genoffice slides check</code> menolak apa pun yang meluap atau tumpang tindih sebelum <code>genoffice create</code> merakit <code>.pptx</code> dan <code>slides render</code> mengembalikan satu PNG per slide untuk dilihat.</td>
<td><b>Pasang sekali, dari Pengaturan → Integrasi</b> — GenOffice menampilkan coding agent yang ditemukannya di komputer ini dan menulis skill ke setiap agen yang Anda pilih. Atau unduh skill sebagai zip, atau jalankan <code>npx skills add genspark-ai/genoffice</code>. Perintah dan alur kerja lengkapnya ada di <a href="#command-line-and-agent-skill">Baris perintah dan skill agen</a>.</td>
</tr>
</table>

### 8 · MCP — alat yang sama lewat Model Context Protocol

Setiap perintah `genoffice` juga merupakan alat MCP. Claude Code, Claude
Desktop, Cursor, dan MCP client lainnya bisa menjalankan sendiri `genoffice
mcp`, tanpa perlu memasang skill dan tanpa membuka jendela, lalu mendapatkan
29 alat plus referensi op sebagai resource. Server HTTP kedua di dalam
aplikasi memungkinkan agen membangun dokumen Word di tab editor yang
terlihat sambil Anda menyaksikannya.

<img src="../assets/readme/mcp-deck-motion.webp" alt="Rekaman time-lapse Claude Code membangun briefing investor delapan slide tentang energi terbarukan lewat MCP server genoffice: mencari figur dan foto, memeriksa setiap foto kandidat dengan media, deck_start menulis style sheet dan outline, deck_page menambahkan satu halaman yang sudah diperiksa setiap kalinya, deck_build merakit .pptx dan slides_render mengembalikan gambar setiap slide; deck yang sudah jadi kemudian terbuka di GenOffice Slides" width="100%">

<table>
<tr>
<td width="50%"><img src="../assets/readme/mcp-deck-in-app.webp" alt="GenOffice Slides menampilkan deck Renewable Energy 2026 delapan slide yang dibangun Claude Code lewat MCP server genoffice: slide sampul dengan foto ladang angin di kanvas dan delapan thumbnail di kiri"></td>
<td width="50%"><img src="../assets/readme/mcp-integrations.webp" alt="GenOffice Settings, halaman Integrations, bagian MCP: perintah satu baris claude mcp add untuk Claude Code, blok JSON untuk Cursor, Claude Desktop, dan MCP client lain, serta opsi server HTTP lokal di bawahnya"></td>
</tr>
<tr>
<td><b>Satu prompt, 38 panggilan alat, tanpa shell</b> — "Buat briefing investor delapan slide tentang energi terbarukan di 2026, dengan foto asli di sampul dan di mana pun foto membantu." Agen mengambil figur dan foto dengan <code>search</code>, menanyakan ke <code>media</code> apakah setiap foto kandidat adalah foto asli, memanggil <code>deck_start</code> dengan style sheet dan outline, lalu <code>deck_page</code> untuk setiap slide; setiap halaman diperiksa terhadap outline dan palet sebelum disimpan, <code>deck_build</code> merakit <code>.pptx</code>, <code>slides_audit</code> mencari overflow, <code>slides_render</code> mengembalikan PNG per slide sebagai konten gambar yang bisa dilihat model, dan <code>deck_replace</code> memperbaiki tiga halaman yang kurang disukainya.</td>
<td><b>Hubungkan sekali, dari Pengaturan → Integrasi</b> — salin baris <code>claude mcp add</code> untuk Claude Code, atau blok JSON ke Cursor, Claude Desktop, atau MCP client lain. Opsi B mengaktifkan server HTTP lokal untuk editor Word yang terlihat. Keduanya dijelaskan di <a href="#mcp-server">Server MCP</a>.</td>
</tr>
</table>

## Mengapa GenOffice

- **Open source**, Apache-2.0, dikembangkan secara terbuka di GitHub.
- **Milik Anda untuk dijalankan.** Aplikasi native untuk macOS, Windows, dan
  Linux; file tetap di disk Anda dan setiap pengeditan, penyimpanan, dan
  konversi terjadi di perangkat Anda sendiri.
- **File Office asli.** `.docx`, `.xlsx`, dan `.pptx` native, byte-preserving:
  bagian file yang tidak Anda sentuh disalin apa adanya.
- **AI yang mengedit dokumen itu sendiri.** Tracked changes di Docs, rumus dan
  grafik langsung di Sheets, slide yang digambar langsung di kanvas, setiap
  giliran AI adalah snapshot yang bisa Anda kembalikan.
- **Model Anda, kunci Anda.** Masuk dengan Genspark, atau bawa kunci untuk
  Claude, OpenAI, Gemini, DeepSeek dan lainnya, termasuk server lokal dan
  endpoint apa pun yang kompatibel dengan OpenAI.
- **PDF yang ditangani dengan benar.** Edit teks langsung di halaman, dan
  konversi PDF ke Word, Excel, atau PowerPoint secara on-device, dengan OCR
  sistem untuk hasil pindaian.
- **Markdown dan HTML juga**, dengan panel AI yang sama dan ekspor lokal ke
  Word.
- **Dapat dijalankan lewat skrip.** Baris perintah `genoffice`, skill agen,
  dan server MCP menempatkan setiap engine untuk melayani Claude Code,
  Claude Desktop, Codex, Cursor, dan agen lainnya, tetap on-device.
- **Gratis**, untuk individu maupun tim.

## Backend AI

**Masuk dengan Genspark** dan tidak ada yang perlu dikonfigurasi: panggilan
model diarahkan lewat proxy Genspark (keluarga Claude, GPT, dan Gemini) dan
agen mendapatkan akses pencarian web dan gambar, generasi gambar, serta
analisis gambar/audio/video.

**Atau bawa kunci Anda sendiri.** Settings → AI menampilkan Claude, OpenAI,
Gemini, DeepSeek, Kimi, GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty,
dan OpenCode Zen/Go, plus slot khusus untuk endpoint apa pun yang kompatibel
dengan OpenAI (base URL + kunci), termasuk server model lokal. Pencarian dan
media memiliki penyedia tersendiri per kemampuan di bawah **AI Media &
Search**: Serper atau Tavily untuk pencarian web, dan OpenAI, Gemini,
Doubao/Seedream, GLM, Grok, Qwen, MiniMax, atau endpoint apa pun yang
kompatibel dengan OpenAI untuk generasi gambar serta analisis gambar/video.

Seluruh suite mendukung tema terang, gelap, dan sistem. Tema hanya mengubah
tampilan di layar: ekspor, cetakan, dan file yang disimpan selalu
mempertahankan warna asli dokumen.

<a id="command-line-and-agent-skill"></a>

## Baris perintah dan skill agen

Semua yang bisa dilakukan aplikasi terhadap sebuah file, bisa dilakukan baris
perintah `genoffice` dari terminal: memeriksa, mengonversi, membuat, membaca,
dan mengedit Word, Excel, PowerPoint, PDF, Markdown, dan HTML dengan engine
yang sama, secara headless. Ia terpasang bersama GenOffice, tidak memerlukan
runtime sendiri, dan tidak pernah mengirim dokumen ke mana pun. Dipadukan
dengan **skill agen** yang disertakan, ia mengubah coding agent menjadi
pekerja dokumen yang menghasilkan file Office asli, bukan tiruan dalam
Markdown.

**Bekerja dengan:** Claude Code, Codex, Cursor, Gemini CLI, GitHub Copilot,
OpenCode, dan Windsurf langsung tanpa konfigurasi, agen lain mana pun yang
membaca skill, dan, lewat [server MCP](#mcp-server), Claude Desktop serta
semua MCP client.

### Memasang skill

| Cara                                       | Yang terjadi                                                                                                                                                                          |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Pengaturan → Integrasi** di aplikasi     | Menampilkan agen yang ditemukan di komputer ini; satu klik menulis skill ke setiap agen yang Anda pilih. Tombol **Update** muncul saat rilis GenOffice membawa skill yang lebih baru. |
| **Unduh sebagai zip** di halaman yang sama | Tata letak yang diterima claude.ai, aplikasi desktop Claude, dan asisten lain sebagai skill yang diunggah.                                                                            |
| `npx skills add genspark-ai/genoffice`     | Memasang dari repositori ini ke agen mana pun yang kompatibel dengan skill.                                                                                                           |

Lalu mulai chat baru dan minta sebuah dokumen. Skill ini mengajari agen kapan
harus menggunakan `genoffice`, bagaimana membaca file sebelum mengeditnya,
dan bagaimana memeriksa hasil kerjanya sendiri.

### Mulai cepat dari terminal

```bash
genoffice --version
genoffice info report.docx --json                  # headings and blocks; or sheets, slides, pages
genoffice convert report.md --to pdf               # md/html/docx/xlsx/pptx → pdf, pdf → docx/xlsx/pptx, …
genoffice create --type docx --from notes.md --out notes.docx
genoffice create --type xlsx --from table.json --out sales.xlsx   # "=SUM(B2:B9)" cells stay live formulas
genoffice docs read report.docx --range 0-9 --json # then `docs apply --ops edits.json` edits in place
genoffice render report.docx --out shots/          # one PNG per page, to look at what you made
genoffice open sales.xlsx                          # hand the result to the editor
```

Setiap perintah mencetak ringkasan satu baris, atau satu objek JSON dengan
`--json`. Pengeditan bersifat atomik: op yang ditolak membiarkan file tak
tersentuh dan mengembalikan pesan kesalahan berpanduan. `genoffice help`
mencantumkan daftar perintah saat ini; referensi lengkapnya ada di
[packages/cli/README.md](../../packages/cli/README.md).

### Apa yang sebenarnya dijalankan agen

Deck Tata Surya di demo di atas hanya butuh satu prompt di Claude Code. Di
baliknya, agen mengikuti alur kerja bertahap dari skill dan CLI memeriksa
setiap tahap sebelum tahap berikutnya dimulai:

```bash
genoffice capabilities --json                        # which cloud tools GenOffice has configured
genoffice guide slides design                        # the deck workflow and layout library
genoffice image "the eight planets in a row …" --aspect 16:9 --out deck/assets/cover.jpg
genoffice slides check deck/outline.json --json      # 8 pages, no findings
genoffice slides check deck/pages/01.json --json     # builds one slide, audits overflow and overlap
…                                                    # one page file per slide, fixed until each check is clean
genoffice create --type pptx --spec deck/pages --outline deck/outline.json --out deck/solar-system.pptx --json
genoffice slides render deck/solar-system.pptx --out deck/shots --json
genoffice slides audit deck/solar-system.pptx --json    # 8 slides, no layout issues
genoffice slides replace deck/solar-system.pptx --slide 4 --spec deck/pages/05.json --json
genoffice open deck/solar-system.pptx
```

Tidak ada panggilan model yang terjadi di dalam `genoffice`: agen yang
berpikir, CLI yang membangun dan memeriksa, dan hasilnya terbuka di GenOffice
atau PowerPoint sebagai `.pptx` biasa.

<a id="mcp-server"></a>

### Server MCP

Perintah yang sama tersedia sebagai alat [Model Context Protocol](https://modelcontextprotocol.io)
untuk asisten yang tidak bisa menjalankan terminal, atau yang tidak ingin
Anda beri akses terminal. Ada dua cara masuk, keduanya ditampilkan dengan
cuplikan siap salin di **Pengaturan → Integrasi → MCP**:

| Cara                                 | Apa itu                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| ------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **A · `genoffice mcp`** (disarankan) | Server stdio yang dijalankan sendiri oleh asisten; GenOffice tidak perlu terbuka. Satu alat per perintah (`info`, `convert`, `create_docx`, `create_xlsx`, `create_pptx`, `create_pdf`, `docs_read` / `docs_apply` / `docs_check`, `sheet_*`, `slides_*`, `render`, `guide`, `search`, `image`, `media`, `open`) plus alur deck bertahap `deck_start` → `deck_page` → `deck_build` → `deck_replace`. Ops, spesifikasi, dan Markdown dikirim inline, jadi client tanpa sistem file pun tetap bisa jalan. |
| **B · Server HTTP lokal**            | Berjalan di dalam aplikasi GenOffice pada `http://127.0.0.1:3093/mcp` (Streamable HTTP, dengan SSE lama). Alatnya menjalankan tab editor Word yang terlihat: `create_session`, `insert_content`, `replace_blocks`, `apply_ops`, `read_document`, `save_session`, dan Anda menyaksikan dokumen terbentuk. Mati secara default; nyalakan di panel pengaturan yang sama.                                                                                                                                   |

```bash
# Claude Code
claude mcp add --transport stdio genoffice -- genoffice mcp
```

```jsonc
// Cursor, Claude Desktop, atau MCP client lain mana pun
{ "mcpServers": { "genoffice": { "command": "genoffice", "args": ["mcp"] } } }
```

`genoffice` di sini adalah CLI yang disertakan di dalam aplikasi (di macOS
`/Applications/GenOffice.app/Contents/Resources/cli/genoffice`; panel
pengaturan menampilkan path persis untuk instalasi Anda). Server ini
membawa instruksi alur kerjanya sendiri dan mengekspos referensi op sebagai
resource `genoffice://guide/*`, jadi skill tidak diperlukan; skill dan
server MCP bisa berdampingan dan asisten memilih salah satunya. Fitur cloud
(`search`, `image`, `media`) tetap lewat provider yang dikonfigurasi di
GenOffice; sisanya berjalan lokal, dan `GENOFFICE_ALLOWED_ROOTS` membatasi
setiap alat hanya ke folder yang Anda daftarkan.

Deck energi terbarukan dalam demo di atas adalah gambaran, dari sisi
protokol, satu prompt di Claude Code yang hanya menyambungkan server MCP
`genoffice`:

```text
capabilities · guide(slides, spec) · guide(slides, design)
search(query) ×4                         → IEA, BNEF and IRENA figures for the slides
search(query, images) ×7 · media(url, ask) ×7
                                         → candidate photos, each one checked to be a real photograph
deck_start(dir, style, outline)          → outline checked: 8 pages to write
deck_page(dir, 0, page) … deck_page(dir, 7, page)
                                         → each page checked against the outline and the palette; one page sent again
deck_build(dir, out)                     → renewables-2026.pptx, no image failures
slides_audit(file) · slides_render(file, out)
                                         → no layout findings; 8 PNGs come back as image content
deck_replace(dir, n, page) ×3 · slides_render(file, out)
                                         → three pages fixed after looking at the renders
```

Tiga puluh delapan panggilan, sekitar tiga belas menit, dan asisten sama
sekali tidak menyentuh shell: figur, foto, semua panduan, pemeriksaan, dan
render berjalan sebagai hasil panggilan alat MCP. Hanya `search` dan `media`
yang keluar dari mesin, ke penyedia yang dikonfigurasi di GenOffice.

<a id="download"></a>

## Download

| Platform                             | Persyaratan                                           | Unduh                                                                                      |
| ------------------------------------ | ----------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| **macOS** — Apple Silicon (arm64)    | macOS 11+                                             | [`.dmg` terbaru (arm64)](https://github.com/genspark-ai/genoffice/releases/latest)         |
| **macOS** — Intel (x64)              | macOS 11+                                             | [`.dmg` terbaru (x64)](https://github.com/genspark-ai/genoffice/releases/latest)           |
| **Windows** (x64, kebanyakan PC)     | Windows 10+, Intel/AMD                                | [Installer `-x64.exe` terbaru](https://github.com/genspark-ai/genoffice/releases/latest)   |
| **Windows** di Arm (ARM64)           | Windows 11 on Arm (Snapdragon X dan sejenisnya)       | [Installer `-arm64.exe` terbaru](https://github.com/genspark-ai/genoffice/releases/latest) |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 atau lebih baru)    | [`.deb` terbaru](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | [`.rpm` terbaru](https://github.com/genspark-ai/genoffice/releases/latest)                 |
| **Linux** — distribusi lainnya       | x86_64, glibc 2.34+, FUSE 2                           | [`.AppImage` terbaru](https://github.com/genspark-ai/genoffice/releases/latest)            |

Semua build berasal dari `main`; installer macOS dan Windows sudah
ditandatangani (signed). Versi lama ada di halaman
[Releases](https://github.com/genspark-ai/genoffice/releases).

<details>
<summary><b>Instalasi di Linux</b></summary>

File deb dipasang dengan apt — perintah ini menarik semua dependensi dan
menambahkan GenOffice ke menu aplikasi:

```bash
sudo apt install ./genoffice_<version>_amd64.deb
```

Di Fedora / keluarga RHEL / openSUSE, pasang rpm sebagai gantinya:

```bash
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

AppImage berjalan langsung di tempat: pasang runtime FUSE 2
(`sudo apt install libfuse2`; di Ubuntu 24.04 paketnya bernama
`libfuse2t64`), jadikan file tersebut executable, lalu jalankan:

```bash
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

</details>

## Cara kerjanya

Tujuh aplikasi Electron — Docs, Sheets, Slides, PDF, Markdown, HTML, dan
shell bertab — berbagi satu lapisan engine berupa paket TypeScript murni,
ditambah sidecar Rust untuk `.xlsx`. File asli selalu menjadi sumber
kebenaran (source of truth): perubahan diterapkan sebagai patch yang sempit,
dan semua bagian yang tidak disentuh editor tetap utuh saat file dibuka
kembali.

```
open docx ─► archive original by hash (never touched)
          ─► parse word/document.xml into a block tree, each block anchored to its original XML
          ─► Tiptap editor (manual + AI editing, dirty tracking)
save      ─► dirty blocks → OOXML fragments (referencing existing styles only)
          ─► splice into the original document.xml; untouched blocks keep their bytes
          ─► repack the zip; every other entry is copied byte-for-byte
```

Tur paket demi paket (engine docx/pptx, `pdf2docx`, `html2docx`, agent core,
dan provider) ada di [CONTRIBUTING.md](../../CONTRIBUTING.md#engine-packages).

## Pengembangan

```bash
npm install
npm run fixtures     # generate test .docx fixtures
npm test             # engine + app unit tests (docs/sheets/slides need no display)
npm run typecheck    # tsc --noEmit across every workspace
npm run dev          # all six editors + shell against Vite dev servers
npm run dev:docs     # a single app (same pattern works per workspace)
npm run dist:mac     # package macOS dmg (regenerates third-party notices)
npm run dist:win     # package Windows nsis installer
npm run dist:linux   # package Linux AppImage + deb + rpm
```

Aplikasi sheets juga membutuhkan toolchain Rust untuk sidecar xlsx-nya
(`cargo` di PATH); `npm run build -w @genoffice/sheets` mengompilasinya
secara otomatis. Lihat [CONTRIBUTING.md](../../CONTRIBUTING.md) untuk
pemeriksaan yang harus dilalui setiap perubahan dan cara pull request
diterima.

## Komunitas

GenOffice terus dikembangkan secara aktif dan masukan Anda membentuk arahnya.

- **Laporkan bug atau ajukan permintaan fitur** di
  [GitHub Issues](https://github.com/genspark-ai/genoffice/issues).
- **Gabung ke grup chat GenOffice** di
  [GenTeam](https://genoffice.ai/join) untuk mengobrol dengan tim dan
  pengguna lain.
- **Beri bintang (star) pada repo ini** jika GenOffice bermanfaat bagi Anda —
  ini cara terbaik untuk mendukung proyek ini.

## FAQ

<details>
<summary><b>Apakah GenOffice gratis?</b></summary>

Ya. GenOffice gratis dan open-source di bawah lisensi Apache-2.0 — tidak ada
masa coba, tidak ada tingkatan berbayar untuk aplikasinya sendiri.

</details>

<details>
<summary><b>Apakah GenOffice bisa membuka file Microsoft Word, Excel, dan PowerPoint?</b></summary>

Ya. GenOffice membuka dan menyimpan file `.docx`, `.xlsx`, dan `.pptx` asli.
Penyimpanan bersifat byte-preserving: bagian file yang tidak Anda sentuh
ditulis kembali byte demi byte, sehingga dokumen tetap berfungsi normal di
Microsoft Office.

</details>

<details>
<summary><b>Apakah GenOffice bisa digunakan secara offline?</b></summary>

Pengeditan dokumen sepenuhnya lokal — file tidak pernah meninggalkan
perangkat Anda untuk dibuka, diedit, disimpan, atau dikonversi. Fitur AI
(agen, pencarian, alat gambar) membutuhkan koneksi jaringan, baik dengan
masuk lewat Genspark atau dengan kunci API model Anda sendiri.

</details>

<details>
<summary><b>Apakah GenOffice bisa mengedit file PDF?</b></summary>

Ya — pengeditan teks dan gambar PDF yang sesungguhnya, yang menulis ulang
content stream halaman dengan font aslinya tetap dipertahankan, bukan
anotasi penutup.

</details>

<details>
<summary><b>Apakah GenOffice bisa mengonversi PDF ke Word, Excel, atau PowerPoint?</b></summary>

Ya — seluruhnya di perangkat Anda sendiri: ekstraksi karakter level PDFium
ditambah analisis layout berbasis geometri, tanpa layanan cloud, tanpa
unggah. Halaman hasil pindai (scan) juga didukung: di macOS dan Windows, OCR
sistem membacanya, sehingga hasilnya berupa teks yang bisa diedit, bukan
gambar halaman.

</details>

<details>
<summary><b>Bisakah saya menggunakan model AI atau kunci API saya sendiri?</b></summary>

Ya. Selain masuk lewat Genspark tanpa kunci API, GenOffice mendukung
penggunaan kunci Anda sendiri untuk Claude, OpenAI, Gemini, DeepSeek, Kimi,
GLM, Qwen, Doubao, MiniMax, Grok, Mistral, OpenRouter, Requesty, dan OpenCode Zen/Go,
plus endpoint apa pun yang kompatibel dengan OpenAI — termasuk server model
lokal. Pencarian, generasi gambar, dan analisis gambar/video menggunakan
kunci tersendiri di Settings → AI Media & Search.

</details>

<details>
<summary><b>Apakah GenOffice bisa mengonversi HTML ke Word?</b></summary>

Ya — Export as Word di aplikasi HTML menghasilkan `.docx` asli yang bisa
diedit, seluruhnya di perangkat Anda sendiri. Halaman dirender di Chromium
bawaan dan diubah menjadi struktur Word yang sesungguhnya: heading,
paragraf, daftar, tabel, kartu, baris KPI, form field, dan latar halaman;
hanya elemen visual yang tidak punya kesetaraan di Word (chart, ikon, kotak
bergaya dekoratif) yang disematkan sebagai gambar.

</details>

<details>
<summary><b>Bisakah saya menjalankan GenOffice dari Claude Code, Codex, Cursor, atau skrip?</b></summary>

Ya. GenOffice memasang baris perintah `genoffice` yang menjalankan engine yang
sama secara headless: memeriksa, mengonversi, membuat, membaca, dan mengedit
dokumen dari terminal atau skrip, dengan output `--json` untuk program. Skill
agen yang disertakan mengajari Claude Code, Codex, Cursor, Gemini CLI, GitHub
Copilot, OpenCode, dan Windsurf untuk menggunakannya; pasang dari
**Pengaturan → Integrasi**. Lihat
[Baris perintah dan skill agen](#command-line-and-agent-skill).

</details>

<details>
<summary><b>Apakah GenOffice mengumpulkan data?</b></summary>

Build resmi yang dipaketkan mengirim data analitik penggunaan terbatas
secara default, dan Anda bisa menonaktifkan pelaporan ini kapan saja lewat
Settings → General. Analitik ini tidak pernah mengirim isi dokumen, nama
file, path file, identitas akun, atau alamat email. Lihat
[GenOffice Privacy](../../PRIVACY.md) untuk daftar lengkap event dan
pengungkapan data.

</details>

## Keamanan

Lihat [SECURITY.md](../../SECURITY.md) untuk postur keamanan proses
(sandboxing renderer, validasi IPC, gating tautan eksternal) dan model
ancaman untuk konten yang dihasilkan AI.

## Ucapan terima kasih

GenOffice tidak akan mungkin ada tanpa proyek-proyek open-source berikut:

- [Electron](https://www.electronjs.org/) — runtime desktop untuk setiap
  aplikasi.
- [Univer](https://github.com/dream-num/univer) (Apache-2.0) — inti UI
  spreadsheet yang diperluas oleh Sheets.
- [PDFium](https://pdfium.googlesource.com/pdfium/) (BSD-3-Clause, disertakan
  lewat [@embedpdf/pdfium](https://github.com/embedpdf/embed-pdf-viewer)) —
  engine content-stream di balik pengeditan teks dan gambar PDF yang
  sesungguhnya.
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0) dan
  [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT) — rendering PDF dan
  penyusunan dokumen.
- [Tiptap](https://tiptap.dev/) / [ProseMirror](https://prosemirror.net/) —
  block editor di Docs dan Markdown.
- [CodeMirror](https://codemirror.net/) (MIT) — editor sumber di HTML.
- [Konva](https://konvajs.org/) — rendering kanvas untuk Slides dan chart
  Sheets.
- [HarfBuzz](https://github.com/harfbuzz/harfbuzz) (wasm) — metrik
  text-shaping untuk skrip kompleks.
- [calamine](https://github.com/tafia/calamine) dan
  [IronCalc](https://github.com/ironcalc/IronCalc) — lapisan baca dan
  kalkulasi dari sidecar xlsx Rust.
- [libeot](https://github.com/umanwizard/libeot) (MPL-2.0) — decoder
  MicroType Express untuk font PowerPoint yang disematkan, diporting ke
  TypeScript.
- [React](https://react.dev/) (MIT) — lapisan UI setiap aplikasi.
- [Mermaid](https://mermaid.js.org/) (MIT) dan [KaTeX](https://katex.org/)
  (MIT) — diagram dan matematika di Markdown dan Docs.
- [opentype.js](https://opentype.js.org/) (MIT) — parsing font untuk metrik
  dan pencarian glyph.
- [JSZip](https://stuk.github.io/jszip/) (MIT) dan
  [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser)
  (MIT) — lapisan container OOXML dan XML.
- [Fluent UI System Icons](https://github.com/microsoft/fluentui-system-icons)
  (MIT) — set ikon di seluruh ribbon.
- [electron-updater](https://www.electron.build/) (MIT) — update dalam
  aplikasi.
- Font Liberation, Carlito, Caladea, dan Noto CJK (OFL/Apache-2.0) — font
  dokumen yang disertakan.

`npm run notices` meregenerasi ringkasan lisensi pihak ketiga yang disertakan
(`tools/gen-third-party-notices.mjs`); semua dependensi runtime berlisensi
MIT/Apache-2.0/BSD-3-Clause/OFL.

## Lisensi

GenOffice dilisensikan di bawah [Apache License 2.0](../../LICENSE), dengan
satu pengecualian: direktori `ee/` dicadangkan untuk modul enterprise di masa
depan dan tercakup oleh [GenOffice Enterprise License](../../ee/LICENSE).

Nama dan logo GenOffice serta Genspark adalah merek dagang milik Mainfunc,
Inc. Lisensi Apache-2.0 tidak memberikan izin untuk menggunakannya (lihat
bagian 6); fork sebaiknya menggunakan branding sendiri.
