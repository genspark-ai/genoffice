# Capability: dual-transport(web + Electron 双版本,HTTP + IPC 双协议)

## 1. 概述

genoffice 的每个 app 同时支持两种运行形态与两种传输协议:

- **Electron 版**:renderer 运行在 Electron 中,`window.*` 桥 API 经 `ipcRenderer`(IPC)调用主进程。
- **Web 版**:同一 renderer 构建运行在普通浏览器中,`window.*` 桥 API 经 HTTP(请求)+ SSE(推送)调用**同一台桌面主进程**。

后端是运行中的 Electron 主进程;两种协议调用同一批 `ipcMain.handle` 注册的 handler,同通道同载荷必得同结果。

## 2. 服务端桥(packages/ipc-bridge,Node/Electron 主进程侧)

### 2.1 安装与注册拦截

- `installHttpIpcBridge(options)` 在 app main 入口调用一次;随后**所有**经 `ipcMain.handle(channel, handler)` 注册的通道自动同时获得 IPC 与 HTTP 两种访问形态,无需修改任何既有注册点。
- 安装点晚于/早于业务注册均生效(拦截发生在 `ipcMain.handle` 方法层面)。
- 不改变原 `ipcMain.handle` 的 IPC 行为、校验与错误语义。

### 2.2 HTTP 端点

- `POST /api/ipc/:channel`,请求体为 JSON 载荷(与 `ipcRenderer.invoke` 第二参数一致):
  - 通道存在:执行 handler(注入伪 `event`),返回 `200` + handler 结果 JSON。
  - handler 抛错:返回非 2xx,响应体含错误的 message;错误语义与 IPC 路径一致(rejected invocation 等价)。
  - 通道不存在:返回 404 语义的错误对象。
  - 原生专属通道(dialog/打印/截屏/webContents/窗口管理等):桥不逐一改写 handler;当 handler 因缺少原生环境失败或该通道被标记为原生专属时,web 客户端侧展示结构化"仅桌面版支持"错误(见 2.4/4.3)。
- `GET /api/ipc/events?session=<id>`:SSE 流,承载主进程经 `event.sender.send(channel, payload)` 发出的推送;消息形如 `data: {"channel":"...","payload":...}`,按会话隔离。
- 会话:每个 invoke 请求建立/复用会话;伪 `event.sender.send` 的推送进入该会话的 SSE 流。会话空闲超时回收。

### 2.3 伪 event

- 提供 handler 所需的最小 `IpcMainInvokeEvent` 形状:`sender.send(channel, ...args)` 路由到会话 SSE;不可用的原生属性以显式抛错代替 undefined 误导。
- 不伪装 `BrowserWindow`/`webContents` 全局注册表;依赖它们的 handler 在 web 端触发结构化降级错误。

### 2.4 安全与网络

- 默认仅绑定 `127.0.0.1`,不暴露局域网;可用环境变量覆盖端口,不可配置绑定地址为非回环(本 change 范围内)。
- 端口规则:app renderer dev 端口 +100(docs 5273/sheets 5274/slides 5275/pdf 5276/markdown 5277),环境变量可覆盖。

### 2.5 生产 web 形态

- 桥服务器可静态托管该 app 的 `out/renderer`(Web 版生产形态):访问根路径返回 index.html 与静态资源;API 与静态资源同源,无需 CORS。

## 3. 客户端传输(packages/ipc-bridge,renderer/浏览器侧)

### 3.1 IpcTransport 抽象

- `{ invoke(channel, payload): Promise<unknown>; send?(channel, ...args): void; on(channel, listener): () => void }`。
- **Electron 传输**:封装 `ipcRenderer.invoke/on/send`,行为与现状逐点一致。
- **HTTP 传输**:`invoke` → `POST /api/ipc/:channel`;`on` → SSE 流按 channel 过滤分发;错误还原为可被现有 catch 分支处理的 Error。
- 运行时探测:`isElectronRuntime()`(基于 preload 注入标志/userAgent),web 入口仅在浏览器环境安装 API,Electron 环境下不覆盖 preload 注入的 `window.*`。

## 4. app 接入模式

### 4.1 main(全部 6 app)

- 各 app main 入口一行安装桥(开发与生产均生效)。

### 4.2 preload 工厂化(docs、markdown)

- API 构造逻辑抽出为共享工厂 `createXxxApi(transport)`:通道名、入参校验、返回形状与现 preload 完全一致,仅传输可替换。
- preload = `contextBridge.exposeInMainWorld(name, createXxxApi(electronTransport))`。

### 4.3 web 入口(docs、markdown)

- renderer `index.html` 增加先于主 bundle 的 bootstrap 模块:浏览器环境下以 HTTP 传输构造同名 `window.*` API;原生专属调用经由桥得到结构化"仅桌面版支持"错误(`WEB_UNSUPPORTED` 语义),UI 既有错误处理可见。
- dev:`vite.renderer.config.ts` 将 `/api` 与 `/api/ipc/events` 代理至桥端口;生产:由桥静态托管。

### 4.4 能力矩阵(文档化)

- web 可用:纯数据/计算/AI/存储类通道(后端有文件系统,文件读写可用,交互式选择器除外)。
- web 降级(结构化错误):原生对话框、打印/printToPDF、desktopCapturer、窗口/标签管理、本机字体/OCR、md-asset://、子窗口流程。

## 5. 不变量

- 同通道同载荷:IPC 与 HTTP 结果一致(AI 流式分片顺序一致)。
- Electron 版行为、打包产物、既有测试零回归。
- 桥不改变任何通道的业务语义与 schema;新通道要获得 HTTP 形态无需额外登记。
