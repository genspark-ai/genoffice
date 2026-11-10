# 安裝 GenOffice

每個版本都發布在 [Releases 頁面](https://github.com/genspark-ai/genoffice/releases/latest)。它們全部出自 `main`；macOS 與 Windows 的安裝程式已簽署。

## 選擇適合你電腦的檔案

| 平台                                 | 需求                                                   | 檔案                  |
| ------------------------------------ | ------------------------------------------------------ | --------------------- |
| **macOS** — Apple Silicon            | macOS 11+                                              | `.dmg` (arm64)        |
| **macOS** — Intel                    | macOS 11+                                              | `.dmg` (x64)          |
| **Windows** — 大多數 PC              | Windows 10+、Intel/AMD                                 | `-x64.exe` 安裝程式   |
| **Windows** on Arm                   | Arm 版 Windows 11（Snapdragon X 之類）                 | `-arm64.exe` 安裝程式 |
| **Linux** — Debian / Ubuntu          | x86_64、glibc 2.34+（Ubuntu 22.04 或更新）             | `.deb`                |
| **Linux** — Fedora / RHEL / openSUSE | x86_64、glibc 2.34+（Fedora 35+、RHEL 9+、Leap 15.6+） | `.rpm`                |
| **Linux** — 其他                     | x86_64、glibc 2.34+、FUSE 2                            | `.AppImage`           |

舊版本仍保留在 Releases 頁面上。

## 在 macOS 上

開啟 `.dmg`，把 **GenOffice** 拖進「應用程式」資料夾。Intel 版首次啟動時會請你確認開啟：在「應用程式」裡按右鍵點這個應用程式 ▸ 打開。這只是 Gatekeeper 遇到一個看起來未簽署的磁碟映像，並不是下載損毀。

## 在 Windows 上

執行 `.exe` 安裝程式。它會把 GenOffice 放進開始功能表，並註冊 `.docx`、`.xlsx`、`.pptx`、`.pdf` 等格式，按兩下檔案就會用 GenOffice 開啟。

## 在 Linux 上

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — 免安裝，放在哪裡就在哪裡執行：

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

它需要 FUSE 2 執行環境。如果不想安裝它，可以用 `--appimage-extract-and-run` 來執行，沙箱會從解壓出來的目錄運作。

## 你正在執行的建置

**設定 ▸ 關於** 是應用程式回報你正在使用的建置版本、它跟隨的更新頻道，以及專案所在位置的地方。當該頻道發布了更新的建置版本時，同一個面板會提供更新；Flatpak、Nix 或 Docker 安裝則不會，因為這些安裝會被當初把它們裝上去的工具取代。

![設定 ▸ 關於，顯示已安裝的版本、應用程式跟隨的更新頻道，以及專案的 GitHub 連結](img/install.png)

## 另外三條路

這三種方式都只是包裝已發布的建置產物，而不是重新編譯應用程式，所以跟的是官方二進位檔。

- **Flatpak** — 透過 Flatpak 的 `extra-data` 機制安裝 `.deb`，並在 Zypak 沙箱下執行，由 Zypak 調整 Electron 自帶的沙箱。`/app` 目錄是唯讀的，所以更新要走 Flatpak，而不是應用程式內的更新對話方塊。
- **Nix** — 一個套件表示式，用 `nix-build packaging/nix` 從原始碼檢出建置；詳見 `packaging/nix`。
- **Docker** — 一個無頭批次轉換映像（`packaging/docker`）。它會透過應用程式自己的匯出流程，把整棵目錄樹裡的 Office、Markdown 與 HTML 文件轉成 PDF，輸出的結果與 檔案 ▸ 匯出 相同。它是一個命令列工具，不是伺服器。

## 命令列隨著一起安裝

每次安裝都會附帶一個 `genoffice` 命令，它與視窗呼叫的是同一套引擎，因此兩者對同一個檔案不會有分歧。在 macOS 與 Windows 上它位於應用程式套件內；`genoffice install-cli` 會把它放進你的 `PATH`。它能做什麼，見**命令列與代理**。
