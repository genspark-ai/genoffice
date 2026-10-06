# Model AI dan pengaturan

## Penyedia dan model

Model dan kunci dikonfigurasi di Pengaturan (baris akun di kiri bawah Beranda):

![Jendela Pengaturan](img/settings-general.png)

- **Genspark terkelola**: masuk (alur device-code) lalu pakai — tanpa konfigurasi apa pun.
- **Endpoint kustom (BYOK)**: Pengaturan ▸ AI menerima base URL dan kunci API untuk tiap protokol — yang kompatibel dengan OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen), dan lainnya. Kunci hanya ada di header permintaan — tidak pernah di disk, di log, atau di lingkungan proses anak.
- Model berbeda dapat dipilih per kemampuan: obrolan/pembuatan, pembuatan gambar, analisis gambar.
- **Uji koneksi**: memastikan endpoint dapat dijangkau dan model terlihat sebelum disimpan.
- Base URL boleh memuat path dan query string (gaya gateway); path endpoint akan ditambahkan dengan benar.

## Integrasi CLI (semacam Codex)

- Pengaturan menerima path program CLI lokal (direktori beranda non-ASCII dan awalan ~ berfungsi; ~ dikembangkan otomatis); Deteksi otomatis akan memeriksa model yang tersedia pada CLI tersebut.
- Validasi hanya memeriksa keberadaan berkas — tidak ada pembatasan karakter.

## Kapan perubahan berlaku

- Perubahan model dan endpoint berlaku seketika; percakapan yang sedang berjalan memakai konfigurasi lama sampai putaran berikutnya.
