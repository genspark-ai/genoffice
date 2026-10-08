# Model AI dan tetapan

## Pembekal dan model

Model dan kunci dikonfigurasikan dalam Tetapan, melalui baris akaun di penjuru kiri bawah pada Laman Utama:

![Tetingkap Tetapan](img/settings-general.png)

- **Genspark dihoskan**: log masuk dan gunakannya terus. Tiada konfigurasi diperlukan.
- **Titik akhir tersuai (BYOK)**: Tetapan ▸ Model AI menerima URL asas dan kunci API bagi setiap protokol, termasuk yang serasi dengan OpenAI, Anthropic, Gemini, DeepSeek, DashScope (qwen), dan lain-lain. Kunci disimpan dalam fail tetapan aplikasi pada mesin ini dan dihantar hanya dalam pengepala permintaan.
- Anda boleh memilih model yang berbeza bagi setiap keupayaan, iaitu sembang dan penjanaan, penjanaan imej, dan analisis imej.
- **Uji sambungan** menyemak ketercapaian titik akhir dan keterlihatan model sebelum disimpan.
- URL asas boleh mengandungi laluan dan rentetan pertanyaan, dan laluan titik akhir ditambah dengan betul.

## Integrasi baris arahan

- Tetapan menerima laluan program baris arahan tempatan. Direktori rumah yang menggunakan aksara bukan ASCII dan awalan ~ juga berfungsi, dan ~ akan dikembangkan secara automatik. Fungsi semak model akan menyoal baris arahan itu tentang model yang ada.
- Pengesahan hanya menyemak sama ada laluan itu wujud dan tiada sekatan set aksara.

## Bila perubahan berkuat kuasa

- Perubahan pada model dan titik akhir berkuat kuasa serta-merta. Perbualan yang sedang berjalan akan menggunakan konfigurasi lama sehingga pusingan seterusnya.
