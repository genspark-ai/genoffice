# Outcome

真实启动 genoffice 的 web 版本并验证可用:在真实进程中启动 web 版(docs 与 markdown),用真实浏览器打开页面,验证页面完整渲染、核心流程(编辑→保存→重开)经 HTTP 桥完成、原生专属能力返回结构化降级错误;验证通过后更新相关文档(启动命令实测可跑)。

## Source coverage

- 来源:用户请求"真实启动web版本验证"(2026-09-06)。
- 读取状态:complete。
- 保留语义:真实启动 web 版本(非仅自动化测试),真实浏览器访问,核心流程可用。
- 对应 Spec:specs/web-launch/spec.md。
- 对应验收:A1-A5。
- 覆盖状态:covered。

# Scope

- 真实启动 web 版本:docs 与 markdown 两个 app,dev 模式(真实 Electron 主进程 + Vite + 浏览器)。
- 真实浏览器(Chromium/Chrome)打开 web 地址,验证页面完整渲染。
- 核心流程:markdown 编辑→保存→重开;docs 新建→编辑→保存→重开,全程经 HTTP `/api/ipc/*`。
- 原生专属通道(如原生打开对话框)在 web 端返回结构化"仅桌面版支持"错误。
- 验证通过后更新文档:启动命令实测可跑,补充真实启动验证说明。

# Non-goals

- 不修改桥实现与通道语义(上一 change 已归档)。
- 不做生产 web 形态(桥静态托管)的部署验证(本 change 聚焦 dev 真实启动)。
- 不覆盖其余 4 app(pdf/sheets/slides/shell)的 renderer web 化。

# Acceptance examples

- A1 真实启动:执行文档中的启动命令,docs 与 markdown 的 web 地址在真实浏览器中可打开,页面完整渲染(编辑器可见)。
- A2 核心流程(markdown):真实浏览器中编辑→保存→重开,内容持久化,全程经 HTTP 桥。
- A3 核心流程(docs):真实浏览器中新建→编辑→保存→重开,内容持久化,全程经 HTTP 桥。
- A4 降级:web 端调用原生专属通道收到结构化"仅桌面版支持"错误。
- A5 文档:启动命令与真实启动验证说明已更新,命令实测可跑。

# Constraints and invariants

- 真实启动:必须实际运行 Electron 主进程(带桥)与 Vite,并用真实浏览器访问,拒绝仅单测/mock。
- 不修改上一 change 已归档的实现与文档语义。
- 桥仅绑定 127.0.0.1;web 可用性依赖本机 Electron 进程运行。

# Decisions

- D1 验证形态 = dev 模式真实启动(用户确认):`npm run dev -w @genoffice/<app>` 启动 Electron 主进程 + Vite,真实浏览器打开 Vite 地址;与上一 change 的自动化 E2E 互补,聚焦"真实启动"这一用户可见要求。
- D2 覆盖范围 = docs 与 markdown 两个代表 app(用户确认):与上一 change 的 web 化范围一致。

# Open questions

(无 — 2026-09-06 用户已确认 Shape:真实启动 web 版本验证,进入 Build。)

# Verification expectations

- 真实启动:实际运行启动命令,真实浏览器访问,页面渲染与核心流程实测。
- 回归:docs/markdown typecheck 通过;上一 change 的 E2E 不回归。
- 文档中的命令逐条实测可跑后才算通过。
