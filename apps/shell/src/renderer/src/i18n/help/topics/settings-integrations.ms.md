# Tetapan, bahasa, tema dan integrasi MCP

## Membuka tetapan

Baris akaun di penjuru kiri bawah pada Laman Utama akan membuka panel tetapan. Ia memaparkan Log masuk apabila anda belum log masuk. Pilihan yang berkaitan dengan AI terletak di bahagian Model AI.

![Tetingkap Tetapan](img/settings-integrations.png) — konfigurasi model diterangkan dalam bab Model AI dan tetapan.

## Bahasa

- Tetapan menawarkan **21 bahasa antara muka**: Inggeris, Cina ringkas, Jepun, Korea, Perancis, Jerman, Sepanyol, Thai, Indonesia, Rusia, Arab, Portugis, Itali, Poland, Czech, Belanda, Melayu, Ibrani, Hindi, Cina tradisional, dan Vietnam.
- Pertukaran bahasa berkuat kuasa serta-merta dan kekal, manakala bar menu sistem dibina semula dalam bahasa baharu itu.

## Tema

Tiga pilihan ialah Cerah, Gelap, dan Ikut Sistem. Mod Ikut Sistem mengikut rupa sistem pengendalian, dan editor akan mengubah warnanya mengikut tema itu tanpa kelip.

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

GenOffice menyertakan **pelayan MCP tempatan**. Ini membolehkan klien AI luar, seperti Claude Desktop dan Cursor, membaca dan menulis dokumen anda secara terus.

- Untuk mulakan, jalankan `genoffice mcp` pada baris arahan. Port dan token pengesahan boleh dikonfigurasi, dan pendengarannya hanya kekal pada localhost secara lalai.
- Keupayaan yang disediakan termasuk mencipta, membuka, dan menyunting fail docx, xlsx serta pptx, membaca kandungan, menukar format, dan mengeksport ke PDF. Ini menggunakan set alatan yang sama seperti aplikasi yang berjalan di komputer.
- Pengesahan token adalah pilihan tetapi sangat disyorkannya. Untuk maklumat lanjut, jalankan `genoffice mcp --help`.

## Ringkasan baris arahan

| Perintah            | Fungsinya                              |
| ------------------- | -------------------------------------- |
| `genoffice <fail>`  | membuka fail                           |
| `genoffice mcp`     | memulakan pelayan MCP tempatan          |
| `genoffice --help`  | memaparkan semua perintah dan pilihan   |
