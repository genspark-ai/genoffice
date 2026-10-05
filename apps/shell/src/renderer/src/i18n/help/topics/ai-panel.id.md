# Panel asisten AI

Setiap editor bisa memunculkan panel AI: pilih sesuatu, beri instruksi, dan lihat hasilnya mengalir.

## Membuka dan menggunakannya

![Panel AI di Docs](img/ai-panel.png)

- Titik masuknya: **tombol AI** di pita setiap editor, **Tanya AI** di menu konteks, atau Tanya AI pada bilah anotasi.
- Jelaskan tugas dengan bahasa sehari-hari (tulis ulang ini / jadikan kolom ini persentase / tata ulang halaman ini...) lalu tekan Enter.
- Balasan dirender **secara bertahap**; ketika AI membutuhkan alat (membaca dokumen, mengeditnya, menjalankan skrip), AI menjalankannya dan melanjutkan sampai selesai.
- **Berhenti**: hentikan putaran yang sedang berjalan kapan saja.

## Apa yang dapat dilakukan

- **Docs**: menulis ulang/memperluas/menerjemahkan/merangkum, menyisipkan tabel dan gambar, menyesuaikan format; setiap putaran menyimpan snapshot lebih dulu.
- **Sheets**: rumus, pengisian data, transformasi massal, format.
- **Slides**: pembuatan seluruh dek, penyesuaian tata letak, penulisan ulang naskah.
- **PDF**: tanya jawab dan ringkasan berdasarkan teks atau halaman yang dipilih.
- **Markdown / HTML**: menulis ulang, mengembangkan, menerjemahkan.

## Pengembalian dan keamanan

- Panel Docs menyimpan **daftar versi**: satu snapshot per putaran, Anda bisa mundur ke mana pun, dan rollback itu sendiri bisa dibatalkan dengan Ctrl+Z. Snapshot tetap ada setelah dokumen ditutup lalu dibuka kembali.
- suntingan AI melewati alur penyuntingan yang sama dengan suntingan manual (bisa di-undo, dan tunduk pada simpan) — tidak ada yang melewati konfirmasi simpan Anda.

## Privasi

- Instruksi dan isi dokumen yang relevan dikirim ke **layanan model yang Anda konfigurasikan** (Genspark atau endpoint kustom, lihat bab berikutnya); tanpa konfigurasi, tidak ada yang dikirim.
- File lokal tidak diunggah ke mana pun; kunci BYOK hanya ada di header permintaan — tidak pernah di disk maupun di log.
