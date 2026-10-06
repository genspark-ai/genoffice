# Pemasangan GenOffice

Setiap keluaran diterbitkan pada [halaman Keluaran](https://github.com/genspark-ai/genoffice/releases/latest). Semuanya datang daripada `main`; pemasang untuk macOS dan Windows adalah bertanda tangan.

## Memilih fail untuk mesin anda

| Platform | Keperluan | Fail |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — kebanyakan PC | Windows 10+, Intel/AMD | pemasang `-x64.exe` |
| **Windows** pada Arm | Windows 11 pada Arm (Snapdragon X dan yang serupa) | pemasang `-arm64.exe` |
| **Linux** — Debian / Ubuntu | x86_64, glibc 2.34+ (Ubuntu 22.04 atau lebih baharu) | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm` |
| **Linux** — apa-apa yang lain | x86_64, glibc 2.34+, FUSE 2 | `.AppImage` |

Versi lama kekal pada halaman Keluaran.

## Pada macOS

Buka `.dmg` itu dan seret **GenOffice** ke dalam Applications. Pada keluaran Intel, pelancaran pertama akan meminta anda mengesahkan membukanya: klik kanan aplikasi itu dalam Applications ▸ Open. Itu Gatekeeper bertemu imej cakera yang kelihatan tidak bertanda tangan, bukan muat turun yang rosak.

## Pada Windows

Jalankan pemasang `.exe`. Ia meletakkan GenOffice dalam menu Start dan mendaftarkan `.docx`, `.xlsx`, `.pptx`, `.pdf` dan yang lain supaya mengklik dua kali pada fail membukanya.

## Pada Linux

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — ia berjalan di tempat ia berada, tanpa langkah pemasangan:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

Ia memerlukan masa larian FUSE 2. Jika anda lebih suka tidak memasangnya, jalankannya dengan `--appimage-extract-and-run` dan sandbox itu berfungsi daripada direktori yang diekstrak.

## Tiga laluan lagi

Ketiganya membalut artifak yang diterbitkan, bukan membina semula aplikasi, jadi semuanya mengikut binari rasmi.

- **Flatpak** — memasang `.deb` itu melalui mekanisme `extra-data` Flatpak dan menjalankannya di bawah Zypak, yang menyesuaikan sandbox Electron sendiri. Pokok `/app` ialah baca sahaja, jadi kemas kini berlaku melalui Flatpak dan bukan melalui dialog kemas kini dalam aplikasi.
- **Nix** — ungkapan pakej yang anda bina daripada salinan sumber kod dengan `nix-build packaging/nix`; lihat `packaging/nix`.
- **Docker** — imej penukaran pukal tanpa antara muka (`packaging/docker`). Ia menukar keseluruhan pokok dokumen Office, Markdown dan HTML kepada PDF melalui saluran eksport aplikasi itu sendiri, jadi hasilnya sama seperti yang akan dihasilkan oleh Fail ▸ Eksport. Ia utiliti CLI, bukan pelayan.

## Baris arahan yang disertakan

Setiap pemasangan meliputi satu perintah `genoffice` yang bercakap dengan enjin yang sama seperti yang digunakan oleh tetingkap itu, jadi kedua-duanya tidak pernah berselisih tentang sesuatu fail. Pada macOS dan Windows ia berada dalam bungkus aplikasi; `genoffice install-cli` meletakkannya pada `PATH` anda. Lihat **Baris arahan dan ejen** untuk apa yang boleh ia lakukan.
