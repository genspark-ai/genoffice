# GenOffice'i yükleme

Her derleme [Sürümler sayfasında](https://github.com/genspark-ai/genoffice/releases/latest) yayımlanır. Hepsi `main` dalından gelir; macOS ve Windows yükleyicileri imzalıdır.

## Makineniz için doğru dosyayı seçin

| Platform                             | Gereksinimler                                         | Dosya                    |
| ------------------------------------ | ----------------------------------------------------- | ------------------------ |
| **macOS** — Apple Silicon            | macOS 11+                                             | `.dmg` (arm64)           |
| **macOS** — Intel                    | macOS 11+                                             | `.dmg` (x64)             |
| **Windows** — çoğu PC                | Windows 10+, Intel/AMD                                | `-x64.exe` yükleyicisi   |
| Arm üzerinde **Windows**             | Arm üzerinde Windows 11 (Snapdragon X ve benzerleri)  | `-arm64.exe` yükleyicisi |
| **Linux** — Debian / Ubuntu          | x86_64, glibc 2.34+ (Ubuntu 22.04 veya daha yeni)     | `.deb`                   |
| **Linux** — Fedora / RHEL / openSUSE | x86_64, glibc 2.34+ (Fedora 35+, RHEL 9+, Leap 15.6+) | `.rpm`                   |
| **Linux** — diğerleri                | x86_64, glibc 2.34+, FUSE 2                           | `.AppImage`              |

Eski sürümler Sürümler sayfasında kalmaya devam eder.

## macOS'te

`.dmg` dosyasını açın ve **GenOffice**'i Uygulamalar klasörüne sürükleyin. Intel derlemesinde ilk açılışta uygulamayı açmayı onaylamanız istenir: Uygulamalar'da uygulamaya sağ tıklayın ▸ Aç. Bu, Gatekeeper'ın imzasız görünen bir disk görüntüsüyle karşılaşmasıdır; bozuk bir indirme değildir.

## Windows'ta

`.exe` yükleyicisini çalıştırın. GenOffice'i Başlat menüsüne ekler ve `.docx`, `.xlsx`, `.pptx`, `.pdf` ve benzeri uzantıları kaydeder; böylece bir dosyaya çift tıklamak onu açar.

## Linux'ta

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — bulunduğu yerden, yükleme adımı olmadan çalışır:

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

FUSE 2 çalışma zamanına ihtiyaç duyar. Onu yüklemek istemiyorsanız `--appimage-extract-and-run` ile çalıştırın; korumalı alan, çıkarılan dizinden çalışır.

## Çalıştırdığınız derleme

**Ayarlar ▸ Hakkında**, uygulamanın sahip olduğunuz derlemeyi, izlediği kanalı ve projenin nerede bulunduğunu bildirdiği yerdir. O kanalda daha yeni bir derleme yayımlandığında aynı bölme güncellemeyi sunar; Flatpak, Nix veya Docker kurulumu sunmaz, çünkü bunlar onları yerleştiren araç tarafından değiştirilir.

![Ayarlar ▸ Hakkında: yüklü sürüm, uygulamanın izlediği güncelleme kanalı ve projenin GitHub bağlantısı](img/install.png)

## Üç yol daha

Üçü de uygulamayı yeniden derlemek yerine yayımlanmış yapıtları sarmalar; dolayısıyla resmi ikili dosyaları izler.

- **Flatpak** — `.deb` dosyasını Flatpak'in `extra-data` mekanizmasıyla yükler ve Electron'un kendi korumalı alanını uyarlayan Zypak altında çalıştırır. `/app` ağacı salt okunurdur; bu yüzden güncellemeler uygulama içi güncelleme iletişim kutusundan değil Flatpak üzerinden yapılır.
- **Nix** — bir kaynak kopyasından `nix-build packaging/nix` ile derlediğiniz bir paket ifadesi; bkz. `packaging/nix`.
- **Docker** — başsız bir toplu dönüştürme görüntüsü (`packaging/docker`). Office, Markdown ve HTML belgelerinden oluşan koca bir ağacı uygulamanın kendi dışa aktarma hattından geçirerek PDF'ye dönüştürür; dolayısıyla çıktı, Dosya ▸ Dışa Aktar'ın üreteceğiyle aynıdır. Bir CLI aracıdır, sunucu değildir.

## Komut satırı da birlikte gelir

Her kurulum, pencerenin konuştuğu motorlarla aynı motorlarla konuşan bir `genoffice` komutuyla gelir; böylece ikisi bir dosya hakkında asla anlaşmazlığa düşmez. macOS ve Windows'ta uygulama paketinin içindedir; `genoffice install-cli` onu `PATH` değişkeninize ekler. Neler yapabildiği için **Komut satırı ve aracılar** bölümüne bakın.
