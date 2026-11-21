# 安裝 GenOffice

GenOffice 的每一個正式版本與預覽版安裝包都託管在 GitHub 的 [Releases 發布頁面](https://github.com/genspark-ai/genoffice/releases/latest)上。所有安裝包皆由主分支（`main`）自動化流水線編譯打包，其中 macOS 與 Windows 的安裝包皆已完成數位簽章，下載後即可直接安裝使用。

## 選擇適合你電腦架構的安裝包

| 平台                                 | 需求                                                   | 檔案                |
| ------------------------------------ | ------------------------------------------------------ | ------------------- |
| **macOS** — Apple Silicon            | macOS 11+                                              | `.dmg` (arm64)      |
| **macOS** — Intel                    | macOS 11+                                              | `.dmg` (x64)        |
| **Windows** — 大多數 PC              | Windows 10+、Intel/AMD                                 | `-x64.exe` 安裝包   |
| **Windows** on Arm                   | Arm 版 Windows 11（Snapdragon X 等）                   | `-arm64.exe` 安裝包 |
| **Linux** — Debian / Ubuntu          | x86_64、glibc 2.34+（Ubuntu 22.04 或更新）             | `.deb`              |
| **Linux** — Fedora / RHEL / openSUSE | x86_64、glibc 2.34+（Fedora 35+、RHEL 9+、Leap 15.6+） | `.rpm`              |
| **Linux** — 其他                     | x86_64、glibc 2.34+、FUSE 2                            | `.AppImage`         |

如果你需要回退或查閱過往的歷史版本，也可以在同一個 Releases 頁面下方找到完整歸檔。

## 在 macOS 上安裝

雙擊下載好的 `.dmg` 磁碟映像檔，在彈出的視窗中把 **GenOffice** 圖示直接拖進右側的「應用程式（Applications）」資料夾即可完成安裝。需要注意一個小細節：如果你使用的是老款 Intel 晶片的 Mac，首次雙擊啟動時 macOS 的 Gatekeeper 安全機制可能會彈出提示要求確認——此時只需開啟 Finder 的「應用程式」資料夾，在 **GenOffice** 圖示上**點擊滑鼠右鍵並選擇「打開」**，然後在彈窗中點一次確認即可（這只是系統對部分磁碟映像的首次安全詢問，並不代表你下載的安裝包有損壞；確認過一次後，以後正常雙擊或從啟動台開啟就不會再詢問了）。

## 在 Windows 上安裝

直接雙擊下載好的 `-x64.exe`（或高通驍龍筆電對應的 `-arm64.exe`）安裝程式，安裝精靈會自動把 GenOffice 加入系統的「開始」選單與桌面捷徑中，同時在系統裡註冊 `.docx`、`.xlsx`、`.pptx`、`.pdf` 等常見辦公文件的檔案關聯，安裝完成後雙擊電腦裡的任何辦公文件即可直接用 GenOffice 開啟。

## 在 Linux 上安裝

**Debian / Ubuntu 及其衍生發行版**（使用 `apt` 直接安裝本機套件，系統會自動補齊所需相依套件）：

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE 發行版**（分別使用各自的套件管理員安裝 `.rpm` 套件）：

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**通用 AppImage 單一檔案包**（無需管理員權限安裝，下載後賦予可執行權限即可在任意目錄直接執行）：

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

預設情況下，直接雙擊或執行 `.AppImage` 檔案需要系統裝有 `FUSE 2` 執行階段函式庫；如果你所在的 Linux 發行版沒有安裝 FUSE 2、也不想額外改動系統函式庫，只需在啟動命令後面加上 `--appimage-extract-and-run` 參數，程式就會自動在暫存目錄解包並正常啟動內建沙箱。

## 檢視目前版本與更新通道

安裝並開啟軟體後，點擊首頁左下角進入 **設定 ▸ 關於**，就能隨時檢視目前正在執行的精確組建版本號、目前跟隨的更新通道（穩定版或 Beta 版），以及專案的官方 GitHub 倉庫入口。每當所選通道上有新版本發布時，你既可以在這個面板裡直接收到更新提示並一鍵升級，也可以透過頂部選單「說明 ▸ 檢查更新」手動觸發檢測；不過如果你是透過下面提到的 Flatpak、Nix 或 Docker 映像方式安裝的，由於它們的生命週期歸外部套件管理員管轄，應用程式內不會彈出自帶的更新對話框。

![設定 ▸ 關於，其中顯示已安裝的版本、應用程式跟隨的更新通道，以及專案的 GitHub 連結](img/install.png)

## 其他三種封裝與部署方式

除了上面直接下載官方二進位安裝包之外，倉庫還提供了三種面向特定 Linux 桌面環境或伺服器自動化場景的封裝形態——它們底層執行的都是同一套官方組建產物，因此文件排版與匯出結果與桌面版完全一致：

- **Flatpak 沙箱套件**：透過 Flatpak 的 `extra-data` 機制直接拉取並解包官方 `.deb` 組建產物，並在 Zypak 沙箱橋接層下執行以相容 Electron 自帶的安全沙箱；由於 Flatpak 執行階段的 `/app` 目錄是唯讀掛載的，後續版本升級統一透過 `flatpak update` 完成，不走應用程式內的更新彈窗。
- **Nix 宣告式組建**：倉庫的 `packaging/nix` 目錄下內建了完整的 Nix 套件運算式，NixOS 或 Nix 使用者可以在檢出原始碼後直接執行 `nix-build packaging/nix` 組建並納入自己的宣告式環境中。
- **Docker 無頭批次轉換映像**：位於倉庫的 `packaging/docker` 目錄，專門用於伺服器或 CI 流水線上的無人值守文件批次轉換。它會呼叫與桌面端「檔案 ▸ 匯出」完全相同的排版渲染管線，把掛載進去的整棵目錄樹裡的 Word、Excel、PPT、Markdown 和 HTML 文件批次高保真轉換為 PDF 檔案；注意它是一個跑完即退出的批次處理命令列工具，而不是常駐背景的 Web 伺服器。

## 隨包附帶的 `genoffice` 命令列工具

無論你安裝在哪個平台上，每一個 GenOffice 安裝包內部都自帶了一個名為 `genoffice` 的命令列工具，它與圖形介面視窗共用同一套文件剖析器、寫入引擎和排版渲染器，因此用命令列轉換或修改出來的檔案與你在視窗裡手動儲存的結果在位元組與版面上完全一致。在 macOS 和 Windows 上，該二進位檔案預設藏在應用程式安裝目錄內部；你只需在終端機裡執行一次 `genoffice install-cli`，它就會自動把命令連結到你系統的 `PATH` 環境變數中，具體命令用法可參閱**命令列與智慧代理**一章。
