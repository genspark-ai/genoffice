# Outcome

genoffice 的 web 版本具备与桌面版一致的整体功能:全部 6 个 app(docs、markdown、sheets、slides、pdf、shell)都能以 web 形态真实启动,所有 IPC 通道都经 HTTP 可用——原生专属通道提供浏览器等价实现(文件选择、保存/导出下载、打印、剪贴板、字体度量、全屏、屏幕捕获、窗口/标签页),不再是"仅 docs/markdown 可用、其余通道 WEB_UNSUPPORTED"的单独模块形态。

## Source coverage

- 来源:用户请求"实现web版本所有的功能,而不是单独模块,整体功能,复刻有所有的ipc到http"(2026-09-06)。
- 读取状态:complete。
- 保留语义:web 版本整体功能与桌面版一致,全部 IPC 通道经 HTTP 可用。
- 对应 Spec:specs/web-full-functionality/spec.md。
- 对应验收:A1-A13。
- 覆盖状态:covered。

# Scope

- 补齐 sheets/slides/pdf/shell 四个 app 的 web bootstrap(web-bridge.ts + index.html 加载),与 docs/markdown 一致。
- 全部 native-only 通道提供 web 等价实现:
  - 文件打开/选择(open、pick-image、files-pick、insert-image/media/model3d、workbook:select 等)→ 浏览器文件选择器 + 字节上传主进程临时文件;
  - 保存/导出(save-as、export-pdf、export-docx、pick-export-dir、csv-save-confirm 等)→ 主进程返回字节,浏览器下载;
  - 打印(print、print-pdf-buffer)→ 浏览器打印或 PDF 下载后打印;
  - 剪贴板(copy-image-to-clipboard、native-clipboard)→ navigator.clipboard;
  - 字体度量(font-metrics)→ canvas measureText;
  - 全屏(show-fullscreen)→ requestFullscreen;
  - 屏幕捕获(capture-screen-sources/source)→ getDisplayMedia;
  - 窗口/标签页(win:new/list/focus、tabs:*)→ 浏览器标签页(window.open)。
- 移除/放宽 nativeOnlyChannels 中已有 web 等价实现的通道。
- 真实启动全部 6 个 app 的 web 版本并验证核心工作流。

# Non-goals

- 不改变 Electron 桌面版行为(IPC 路径与 preload 语义不变)。
- 不做生产 web 形态(桥静态托管)的部署验证(聚焦 dev 真实启动)。
- 不实现浏览器无法等价的能力(如系统级字体安装、原生通知、真实文件系统目录句柄)。

# Acceptance examples

- A1 全部 6 个 app 的 web 版本可真实启动:执行文档命令后,真实浏览器打开各 app 的 web 地址,页面完整渲染。
- A2 sheets web 版:window.desktopApi 由 HTTP 传输构造,核心工作流(打开/编辑/保存)可用。
- A3 slides web 版:window.slidesApi 由 HTTP 传输构造,核心工作流(打开/编辑/保存)可用。
- A4 pdf web 版:window.pdfApi 由 HTTP 传输构造,核心工作流(打开/编辑/保存)可用。
- A5 shell web 版:window.aiOffice/tabsApi 由 HTTP 传输构造,首页与标签页可用。
- A6 文件打开/选择类通道在 web 端经浏览器文件选择器可用。
- A7 保存/导出类通道在 web 端经浏览器下载可用。
- A8 打印类通道在 web 端经浏览器打印/PDF 下载可用。
- A9 剪贴板/字体/全屏/屏幕捕获类通道在 web 端有浏览器等价实现。
- A10 窗口/标签页管理在 web 端经浏览器标签页可用。
- A11 全部 IPC 通道经 HTTP 可调用:无 IPC_NO_HANDLER,native-only 通道均有 web 等价实现。
- A12 文档更新:docs/web-electron.md 能力矩阵更新为全功能,启动命令实测可跑。
- A13 回归:全部 app typecheck 通过,上一 change 的 E2E 不回归。

# Constraints and invariants

- 真实启动:必须实际运行 Electron 主进程(带桥)与 Vite,并用真实浏览器访问。
- 桥仅绑定 127.0.0.1;web 可用性依赖本机 Electron 进程运行。
- 浏览器 bootstrap 必须保持浏览器安全(不 import electron/node),与 preload 语义一致。
- 桌面版行为不变:preload 仍走 ipcRenderer,通道语义与返回形状一致。

# Decisions

- D1 实现形态 = 共享 web 等价实现 + 每 app web-bridge:在 @genoffice/ipc-bridge 提供浏览器安全的 web 原生等价工具(文件选择/下载/打印/剪贴板/字体/全屏/标签页)与主进程通用 web 文件通道(临时文件写入/字节读取),各 app 的 web-bridge 用它们补齐 native-only 通道。
- D2 覆盖范围 = 全部 6 个 app(docs、markdown、sheets、slides、pdf、shell),与用户"整体功能"要求一致。
- D3 验证形态 = dev 模式真实启动全部 app,真实浏览器验证核心工作流与降级消除。

# Open questions

(无 — 2026-09-06 用户已明确"整体功能、复刻所有 ipc 到 http",进入 Build。)

# Verification expectations

- 真实启动:全部 6 个 app 按文档命令启动,真实浏览器访问,页面渲染与核心工作流实测。
- 通道覆盖:全部 IPC 通道经 HTTP 可调用,无 IPC_NO_HANDLER;native-only 通道均有 web 等价实现。
- 回归:全部 app typecheck 通过;上一 change 的 E2E 不回归。
- 文档中的命令逐条实测可跑后才算通过。
