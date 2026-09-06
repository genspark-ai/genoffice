# Capability: web-full-functionality(web 版本整体功能)

## 1. 概述

genoffice 的 web 版本具备与桌面版一致的整体功能:全部 6 个 app(docs、markdown、sheets、slides、pdf、shell)都能以 web 形态真实启动,所有 IPC 通道都经 HTTP 可用。原生专属通道(对话框、打印、剪贴板、字体、全屏、屏幕捕获、窗口/标签页)在 web 端提供浏览器等价实现,不再返回 WEB_UNSUPPORTED。

## 2. 全部 app 的 web bootstrap

- docs、markdown 已有 web-bridge(HTTP 传输构造 window.* API);sheets、slides、pdf、shell 补齐同样的 web-bridge,index.html 在 main.tsx 前加载。
- 每个 app 的 web-bridge 构造与 preload 相同的 window.* API 对象(如 sheets 的 window.desktopApi、slides 的 window.slidesApi、pdf 的 window.pdfApi、shell 的 window.aiOffice/window.tabsApi),通道名、参数形状、返回形状与桌面版一致,仅传输层为 HTTP/SSE。
- 浏览器 bootstrap 保持浏览器安全(不 import electron/node)。

## 3. 全部 IPC 通道经 HTTP 可用

- 桥已包装 ipcMain.handle/on,所有注册通道自动经 HTTP 服务;本 capability 消除 native-only 阻塞:
  - 文件打开/选择(open、pick-image、files-pick、insert-image/media/model3d、workbook:select、home:browse 等)→ 浏览器文件选择器,字节经 `web:write-temp-file` 写入主进程临时文件后走原打开/添加流程;
  - 保存/导出(save-as、export-pdf、export-docx、pick-export-dir、csv-save-confirm 等)→ 主进程返回字节,浏览器 Blob 下载;
  - 打印(print、print-pdf-buffer)→ 浏览器 window.print() 或 PDF 字节下载后打印;
  - 剪贴板(copy-image-to-clipboard、native-clipboard)→ navigator.clipboard;
  - 字体度量(font-metrics)→ canvas measureText;
  - 全屏(show-fullscreen)→ requestFullscreen;
  - 屏幕捕获(capture-screen-sources/source)→ getDisplayMedia;
  - 窗口/标签页(win:new/list/focus、tabs:*)→ window.open 浏览器标签页。
- 主进程提供通用 web 文件通道:`web:write-temp-file`(字节→临时路径)、`web:read-file-bytes`(路径→字节),供各 app web-bridge 复用。

## 4. 核心工作流

- 每个 app:真实浏览器中打开(web 文件选择器)→ 编辑 → 保存(经 HTTP 桥)→ 重开,内容持久化。
- 导出/打印:主进程生成字节,浏览器下载/打印。
- AI 功能:流式分片经 SSE 实时到达,与桌面版一致。

## 5. 文档

- docs/web-electron.md 能力矩阵更新为全功能(原"Unsupported"项改为浏览器等价实现),启动命令覆盖全部 6 个 app 且实测可跑。
