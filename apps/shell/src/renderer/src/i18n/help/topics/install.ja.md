# GenOffice のインストール

すべてのビルドは [Releases ページ](https://github.com/genspark-ai/genoffice/releases/latest) で公開されています。すべて `main` 由来で、macOS と Windows のインストーラーは署名済みです。

## 自分のマシンに合わせてファイルを選ぶ

| プラットフォーム | 必要なもの | ファイル |
| --- | --- | --- |
| **macOS** — Apple Silicon | macOS 11+ | `.dmg` (arm64) |
| **macOS** — Intel | macOS 11+ | `.dmg` (x64) |
| **Windows** — ほとんどの PC | Windows 10+、Intel/AMD | `-x64.exe` インストーラー |
| **Windows** on Arm | Arm 版の Windows 11（Snapdragon X など） | `-arm64.exe` インストーラー |
| **Linux** — Debian / Ubuntu | x86_64、glibc 2.34+（Ubuntu 22.04 以降） | `.deb` |
| **Linux** — Fedora / RHEL / openSUSE | x86_64、glibc 2.34+（Fedora 35+、RHEL 9+、Leap 15.6+） | `.rpm` |
| **Linux** — それ以外 | x86_64、glibc 2.34+、FUSE 2 | `.AppImage` |

古いバージョンは Releases ページにそのまま残っています。

## macOS で

`.dmg` を開いて **GenOffice** を Applications にドラッグします。Intel 版では初回起動時に開いてよいか確認されます。Applications でアプリを右クリック ▸ 開く の順です。これは未署名に見えるディスクイメージに対する Gatekeeper の動作で、ダウンロードの破損ではありません。

## Windows で

`.exe` インストーラーを実行します。GenOffice がスタートメニューに追加され、`.docx`、`.xlsx`、`.pptx`、`.pdf` などが登録されるので、ファイルをダブルクリックすれば開きます。

## Linux で

**Debian / Ubuntu**

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE**

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**AppImage** — インストールはなしで、そのまま動きます：

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

FUSE 2 ランタイムが必要です。インストールしたくない場合は `--appimage-extract-and-run` を付けて実行してください。展開されたディレクトリーからサンドボックスが動作します。

## さらに 3 つの入手方法

いずれもアプリを再ビルドするのではなく、公開されている成果物を包むだけなので、公式のバイナリをそのまま追従します。

- **Flatpak** — Flatpak の `extra-data` 仕組みで `.deb` をインストールし、Electron 本来のサンドボックスをそれに合わせて適応させる Zypak の下で実行します。`/app` ツリーは読み取り専用なので、更新はアプリ内の更新ダイアログではなく Flatpak を経由します。
- **Nix** — チェックアウトを `nix-build packaging/nix` でビルドするパッケージ式です。`packaging/nix` も参照してください。
- **Docker** — ヘッドレスのバッチ変換用イメージ（`packaging/docker`）です。Office、Markdown、HTML の文書ツリー全体を、アプリ自身の書き出しパイプラインで PDF に変換するので、出力は ファイル ▸ 書き出し と同じものになります。サーバーではなく CLI ユーティリティです。

## コマンドラインも同梱されています

どのインストールにも `genoffice` コマンドが含まれ、ウィンドウと同じエンジンを動かします。同じファイルをどちらがどう扱うかが変わることがないので、意见が分かれることはありません。macOS と Windows ではアプリバンドルの中にあり、`genoffice install-cli` で `PATH` に追加できます。何ができるかは **コマンドラインとエージェント** を参照してください。
