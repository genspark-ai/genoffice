# 安装 GenOffice

GenOffice 的每一个正式版本与预览版安装包都托管在 GitHub 的 [Releases 发布页面](https://github.com/genspark-ai/genoffice/releases/latest)上。所有安装包均由主分支（`main`）自动化流水线编译打包，其中 macOS 与 Windows 的安装包均已完成数字签名，下载后即可直接安装使用。

## 选择适合你电脑架构的安装包

| 平台                                 | 需求                                                   | 文件                |
| ------------------------------------ | ------------------------------------------------------ | ------------------- |
| **macOS** — Apple Silicon            | macOS 11+                                              | `.dmg` (arm64)      |
| **macOS** — Intel                    | macOS 11+                                              | `.dmg` (x64)        |
| **Windows** — 大多数 PC              | Windows 10+、Intel/AMD                                 | `-x64.exe` 安装包   |
| **Windows** on Arm                   | Arm 版 Windows 11（Snapdragon X 等）                   | `-arm64.exe` 安装包 |
| **Linux** — Debian / Ubuntu          | x86_64、glibc 2.34+（Ubuntu 22.04 或更新）             | `.deb`              |
| **Linux** — Fedora / RHEL / openSUSE | x86_64、glibc 2.34+（Fedora 35+、RHEL 9+、Leap 15.6+） | `.rpm`              |
| **Linux** — 其他                     | x86_64、glibc 2.34+、FUSE 2                            | `.AppImage`         |

如果你需要回退或查阅过往的历史版本，也可以在同一个 Releases 页面下方找到完整归档。

## 在 macOS 上安装

双击下载好的 `.dmg` 磁盘映像文件，在弹出的窗口中把 **GenOffice** 图标直接拖进右侧的「应用程序（Applications）」文件夹即可完成安装。需要注意一个小细节：如果你使用的是老款 Intel 芯片的 Mac，首次双击启动时 macOS 的 Gatekeeper 安全机制可能会弹出提示要求确认——此时只需打开访达的「应用程序」文件夹，在 **GenOffice** 图标上**点击鼠标右键并选择「打开」**，然后在弹窗中点一次确认即可（这只是系统对部分磁盘映像的首次安全询问，并不代表你下载的安装包有损坏；确认过一次后，以后正常双击或从启动台打开就不会再询问了）。

## 在 Windows 上安装

直接双击下载好的 `-x64.exe`（或高通骁龙笔记本对应的 `-arm64.exe`）安装程序，安装向导会自动把 GenOffice 添加到系统的「开始」菜单与桌面快捷方式中，同时在系统里注册 `.docx`、`.xlsx`、`.pptx`、`.pdf` 等常见办公文档的文件关联，安装完成后双击电脑里的任何办公文档即可直接用 GenOffice 打开。

## 在 Linux 上安装

**Debian / Ubuntu 及其衍生发行版**（使用 `apt` 直接安装本地包，系统会自动补齐所需依赖）：

```sh
sudo apt install ./genoffice_<version>_amd64.deb
```

**Fedora / RHEL / openSUSE 发行版**（分别使用各自的包管理器安装 `.rpm` 包）：

```sh
sudo dnf install ./genoffice-<version>.x86_64.rpm     # Fedora / RHEL family
sudo zypper install ./genoffice-<version>.x86_64.rpm  # openSUSE
```

**通用 AppImage 单文件包**（无需管理员权限安装，下载后赋予可执行权限即可在任意目录直接运行）：

```sh
chmod +x GenOffice-<version>.AppImage
./GenOffice-<version>.AppImage
```

默认情况下，直接双击或运行 `.AppImage` 文件需要系统装有 `FUSE 2` 运行时库；如果你所在的 Linux 发行版没有安装 FUSE 2、也不想额外改动系统库，只需在启动命令后面加上 `--appimage-extract-and-run` 参数，程序就会自动在临时目录解包并正常启动内置沙箱。

## 查看当前版本与更新渠道

安装并打开软件后，点击主屏左下角进入 **设置 ▸ 关于**，就能随时查看当前正在运行的精确构建版本号、当前跟随的更新通道（稳定版或 Beta 版），以及项目的官方 GitHub 仓库入口。每当所选通道上有新版本发布时，你既可以在这个面板里直接收到更新提示并一键升级，也可以通过顶部菜单「帮助 ▸ 检查更新」手动触发检测；不过如果你是通过下面提到的 Flatpak、Nix 或 Docker 镜像方式安装的，由于它们的生命周期归外部包管理器管辖，应用内不会弹出自带的更新对话框。

![设置 ▸ 关于，其中显示已安装的版本、应用跟随的更新渠道，以及项目的 GitHub 链接](img/install.png)

## 其他三种封装与部署方式

除了上面直接下载官方二进制安装包之外，仓库还提供了三种面向特定 Linux 桌面环境或服务器自动化场景的封装形态——它们底层运行的都是同一套官方构建产物，因此文档排版与导出结果与桌面版完全一致：

- **Flatpak 沙箱包**：通过 Flatpak 的 `extra-data` 机制直接拉取并解包官方 `.deb` 构建产物，并在 Zypak 沙箱桥接层下运行以兼容 Electron 自带的安全沙箱；由于 Flatpak 运行时的 `/app` 目录是只读挂载的，后续版本升级统一通过 `flatpak update` 完成，不走应用内的更新弹窗。
- **Nix 声明式构建**：仓库的 `packaging/nix` 目录下内置了完整的 Nix 包表达式，NixOS 或 Nix 用户可以在检出源码后直接运行 `nix-build packaging/nix` 构建并纳入自己的声明式环境中。
- **Docker 无头批量转换镜像**：位于仓库的 `packaging/docker` 目录，专门用于服务器或 CI 流水线上的无人值守文档批量转换。它会调用与桌面端「文件 ▸ 导出」完全相同的排版渲染管线，把挂载进去的整棵目录树里的 Word、Excel、PPT、Markdown 和 HTML 文档批量高保真转换为 PDF 文件；注意它是一个跑完即退出的批处理命令行工具，而不是常驻后台的 Web 服务器。

## 随包附带的 `genoffice` 命令行工具

无论你安装在哪个平台上，每一个 GenOffice 安装包内部都自带了一个名为 `genoffice` 的命令行工具，它与图形界面窗口共用同一套文档解析器、写入引擎和排版渲染器，因此用命令行转换或修改出来的文件与你在窗口里手动保存的结果在字节与版面上完全一致。在 macOS 和 Windows 上，该二进制文件默认藏在应用安装目录内部；你只需在终端里运行一次 `genoffice install-cli`，它就会自动把命令软链接到你系统的 `PATH` 环境变量中，具体命令用法可参阅**命令行与智能体**一章。
