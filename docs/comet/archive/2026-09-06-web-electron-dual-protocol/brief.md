# Outcome

genoffice 支持 **web + Electron 双版本**、**HTTP + IPC 双协议**:同一套主进程 handler 与三端共享契约层在两种传输下行为一致;浏览器打开即可用(原生专属能力明确降级),Electron 桌面版行为不回归;实现为最佳最小侵入(不改 300+ 通道定义,桥接自动生效),并以真实进程/真实请求验证,验证通过后更新文档。

## 调查结论(事实基础,已确认)

- monorepo:npm workspaces,apps(docs/markdown/pdf/sheets/slides/shell)+ 13 个 packages(TS 源码直发)。
- 每个均为 electron-vite 三段构建(main/preload/renderer);`dev:renderer`(vite.renderer.config.ts)已是纯 HTTP 页面,但代码硬依赖 preload 注入的 `window.desktop/desktopApi/slidesApi/pdfApi/markdownApi/aiOffice*`,浏览器中无实现。
- 全部约 300+ IPC 通道经 `ipcMain.handle` 注册;preload 为薄封装(invoke/send/on);每 app 的 `src/shared/*.ts` 已是通道名+zod schema+TS 类型三端同源的"传输无关契约层"(sheets/desktop-api.ts 最典型)。
- AI 请求(ai-provider/ai-search)全部在主进程执行,renderer 经 IPC 代理;流式分片经 `event.sender.send('ai:stream-chunk'...)` 推回。
- 纯浏览器不可实现:原生对话框、任意路径文件选择 UI、webContents.print/printToPDF、desktopCapturer、窗口/标签管理、本机字体/OCR、md-asset:// 协议。
- 纯数据/计算型通道(文档引擎运算、recalc、AI、project 存储、home 数据等)可平移 HTTP;文件读写因后端仍是 Electron 主进程,在 web 端依然真实可用(仅交互式选择器不可用)。

# Scope

- 新增共享包(如 `packages/ipc-bridge`):
  - 服务端:包装 `ipcMain.handle` 注册(单点拦截,全部现有通道自动获得 HTTP 形态),提供 `POST /api/ipc/:channel`(invoke)与 `GET /api/ipc/events`(SSE,承载 `event.sender.send` 推送);默认仅绑定 127.0.0.1;伪 `event.sender` 将推送路由到对应会话的 SSE 流。
  - 客户端:`IpcTransport` 抽象(Electron=ipcRenderer;Web=fetch+SSE),与 `isElectron` 运行时探测。
- app 侧最小接入:main 入口一行安装桥;preload 的 API 构造抽出为共享工厂 `createXxxApi(transport)`,preload 用 Electron 传输,web 入口用 HTTP 传输并安装同名 `window.*`;app 的 index.html 增加先于主 bundle 的 web-bridge 模块。
- dev web 模式:复用 `vite.renderer.config.ts` 开发服务器并将 `/api` 代理到桥端口;生产 web 形态:桥服务器静态托管 `out/renderer`。
- 原生专属通道在 web 端的结构化降级与能力矩阵文档。
- 真实验证 + 相关文档更新(README/架构说明,含启动命令、端口、能力矩阵)。

# Non-goals

- 纯云部署(无 Electron 后端)或把 handler 迁出 Electron。
- 移动端适配、多用户/鉴权体系、远程暴露(桥仅本机)。
- 用浏览器 API(File System Access 等)替代原生能力。
- 修改任何现有通道语义/schema;性能优化;AI 供应商改为浏览器直连。

# Acceptance examples

- A1 服务端双协议:Electron 启动后,`curl -X POST http://127.0.0.1:<桥端口>/api/ipc/app:get-language` 返回与 Electron 内一致的 JSON(与 app 实际语言一致)。
- A2 双协议同源一致性:同一通道同一载荷,IPC 与 HTTP 两路调用返回一致结果(≥2 个真实数据通道实测比对)。
- A3 web 版端到端(docs 与 markdown):真实 Chromium 分别打开两 app 的 web 地址,页面完整渲染;核心流程(markdown:编辑→保存→重开;docs:新建→编辑→保存→重开)全程经 HTTP `/api/ipc/*` 完成,结果与 Electron 一致。
- A4 推送事件:至少一个主进程推送场景(如 AI 流式分片)在浏览器端经 SSE 实时收到,顺序与 Electron 一致。
- A5 降级策略:web 端调用原生专属通道(如原生打开对话框)收到结构化"仅桌面版支持"错误;Electron 端同通道行为不变。
- A6 不回归:docs/markdown/新包 typecheck/test/lint 通过,其余 4 app typecheck 通过;现有 electron e2e 不低于改动前基线。
- A7 文档:新增/更新文档含双版本架构、启动命令、端口与能力矩阵,文档命令实测可跑。
- A8 其余 4 app(pdf/sheets/slides/shell)main 已接桥:接线完成、包级桥集成测试(真实 HTTP 服务+真实 fetch)覆盖服务端行为,确保样板可复制。

# Constraints and invariants

- 最小侵入:不逐个改 300+ 通道注册点;桥接经单点包装自动生效;app 侧改动为机械模式(main 一行、preload 工厂化、web 入口)。
- 不破坏现有 Electron 行为、打包产物(shell electron-builder extraResources)与既有测试。
- 桥仅绑定 127.0.0.1;SSE 会话与 invoke 会话绑定,不引入远程访问面。
- 推送事件用 SSE(单向推送足够,避免引入 WebSocket 依赖)。
- 工厂化后的 preload 行为逐字节等价(同通道名、同校验、同返回),仅传输可替换。

# Decisions

- D0 工作区:isolation=current,绑定现有 `web` 分支(用户确认;当前目录未提交改动与本 change 无关,归档时分开提交)。
- D1 web 后端架构 = **Electron 主进程内置 HTTP/SSE 桥**(用户确认):同一批 handler 服务 IPC 与 HTTP 双协议,浏览器连运行中的桌面后端;web 可用性依赖本机 Electron 进程运行;独立纯 Node 服务被否决(需剥离全部 Electron 依赖,违背最小实现)。
- D2 覆盖范围 = **基础设施 + 全部 app main 桥接 + 代表 app(docs、markdown)renderer web 化端到端**(用户确认):共享桥包完整落地;6 个 app 的 main 侧各一行接入桥;docs/markdown 完成 preload 工厂化+web 入口真实可用;其余 app(pdf/sheets/slides/shell)按此样板后续接入,不在本 change 做其 renderer web 化。
- D3 原生能力降级 = **桥层结构化错误**(用户确认):web 端调用原生专属通道返回结构化"仅桌面版支持"错误对象,文档给出能力矩阵;不做逐处 UI 隐藏,不做浏览器 API 替代。
- D-impl 推送通道 = SSE(实现选择)。
- D-impl dev = vite 代理 /api→桥;生产 web = 桥静态托管 out/renderer(实现选择)。
- D-impl 桥端口 = 各 app renderer dev 端口 +100(docs 5273/sheets 5274/slides 5275/pdf 5276/markdown 5277),支持环境变量覆盖(实现选择)。

# Open questions

(无 — 2026-09-05 用户已确认 Shape:目标/范围/D1-D3 决定/8 条验收项/非目标,进入 Build。)

# Verification expectations

- 真实验证:真实启动 Electron(带桥)+ 真实 Chromium(playwright)访问 web 版;curl 实测 HTTP 通道;SSE 推送实测;与 IPC 路径结果比对;拒绝仅单测/mock 的验证。
- 回归:受影响 app 与新包 typecheck/test/lint;现有 e2e 基线比对。
- 文档中的命令逐条实测可跑后才算通过。
