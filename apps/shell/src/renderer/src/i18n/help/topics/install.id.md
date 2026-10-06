# Memasang GenOffice

Setiap build dipublikasikan di [halaman Rilis](https://github.com/genspark-ai/genoffice/releases/latest). Semuanya berasal dari `main`; installer untuk macOS dan Windows sudah ditandatangani.

## Pilih berkas untuk komputer Anda

| Platform | Kebutuhan | Berkas |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — kebanyakan PC | Windows 10+, Intel/AMD | installer `-x64.exe` |
| **Windows** di Arm | Windows 11 di Arm (Snapdragon X dan sejenisnya) | installer `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 atau lebih baru) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — lainnya | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Versi lama tetap ada di halaman Rilis.

## Di macOS

Buka `.dmg` lalu seret **GenOffice** ke Applications. Pada build Intel, peluncuran pertama meminta Anda mengonfirmasi pembukaan: klik kanan aplikasinya di Applications ▸ Buka. Itu Gatekeeper bertemu citra disk yang tampak tidak ditandatangani, bukan unduhan yang rusak.

## Di Windows

Jalankan installer `.exe`. Ia menaruh GenOffice di menu Start dan mendaftarkan `.docx`, `.xlsx`, `.pptx`, `.pdf` dan sejenisnya sehingga klik ganda pada berkas langsung membukanya.

## Di Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — berjalan di mana pun diletakkan, tanpa langkah pemasangan:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Ia membutuhkan runtime FUSE 2. Jika Anda lebih suka tidak memasangnya, jalankan dengan `--appimage-extract-and-run` dan sandbox tetap berjalan dari direktori hasil ekstrak.

## Tiga jalur lain

Ketiganya membungkus artefak yang sudah diterbitkan, bukan membangun ulang aplikasinya, sehingga selalu mengikuti biner resmi.

- **Flatpak** — memasang `.deb` lewat mekanisme `extra-data` milik Flatpak dan menjalankannya di bawah Zypak, yang menyesuaikan sandbox bawaan Electron. Pohon `/app` bersifat hanya-baca, jadi pembaruan lewat Flatpak, bukan lewat dialog pembaruan di dalam aplikasi.
- **Nix** — ekspresi paket yang bisa Anda build dari salinan lokal repositori dengan `nix-build packaging/nix`; lihat `packaging/nix`.
- **Docker** — citra konversi massal tanpa antarmuka (`packaging/docker`). Ia mengonversi seluruh pohon dokumen Office, Markdown, dan HTML ke PDF lewat jalur ekspor milik aplikasi itu sendiri, sehingga hasilnya sama dengan yang dihasilkan File ▸ Ekspor. Ini utilitas baris perintah, bukan server.

## Baris perintahnya sudah ikut

Setiap pemasangan membawa perintah `genoffice` yang berbicara dengan mesin yang sama dengan jendela, sehingga keduanya tidak pernah berbeda pendapat tentang sebuah berkas. Di macOS dan Windows perintahnya berada di dalam bundel aplikasi; `genoffice install-cli` menaruhnya di `PATH` Anda. Lihat **Perintah baris perintah dan agen** untuk mengetahui apa saja yang dapat dilaukannya.
